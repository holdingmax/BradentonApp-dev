"""
Persistencia + "vista de reporte" de Lottery dentro de Carga de Datos (ver
CLAUDE.md, "Carga de Datos vs. Excels" y el segundo módulo, Lottery).
Mismo espíritu que reportes_db.py (capa de datos pura, sin conocimiento de
PDFs -- reporte_diario.py hace la extracción), pero en un archivo aparte
porque en el mundo real Lottery ya es un Excel completamente distinto al de
Reporte Diario, no la misma base de datos.

Todo vive en reportes_data/ (mismo directorio ya gitignored que usa
reportes_db.py): reportes_data/lottery.db es el archivo SQLite;
reportes_data/lottery_pdfs/{año}/{mes}/ guarda una copia de cada PDF de
"Daily Sales Report" ya subido.

## Los bloques de 7 días (igual que el Excel real)

Confirmado contra un Excel de Lottery real (`LOTTERY. Analisis 08.2026`):
cada bloque son 7 días de **semana calendario ISO, lunes a domingo** (nunca
se reinician al empezar un mes -- el último bloque de un mes sigue
corriendo hacia el mes que viene, tal cual hace el Excel real), seguidos de
una fila de Subtotal (suma de esos 7 días) y una fila de Debito (el pago
semanal real vía Chase Bank). build_month_blocks() arma esto mismo a partir
de lo guardado en lottery_days -- agrupando por (isoyear, isoweek) en vez
de por mes calendario -- y calcula en Python los mismos resultados que en
el Excel calculan las fórmulas de esas dos filas (ver _DAY_FORMULAS /
_SUBTOTAL_SUM_COLUMNS / _compute_debito_row más abajo, cada uno con la
fórmula real de la que sale, confirmada contra un Excel de Lottery real).
"""

import calendar
import json
import os
import shutil
import sqlite3
from datetime import date, datetime, timedelta

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "lottery.db")
_PDF_DIR = os.path.join(_BASE_DIR, "lottery_pdfs")

_DEPARTMENT_FIELDS = ("online_count", "online_net_sales", "skoff_count", "skoff_net_sales")
_SALES_REPORT_FIELDS = (
    "sales", "pagos", "cash_balance", "comis", "prize_free_plays",
    "pays_units", "pays_amount", "skoff_sales_amount", "sales_comm",
)
ALL_DAY_FIELDS = _DEPARTMENT_FIELDS + _SALES_REPORT_FIELDS


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    # timeout=30 + WAL -- ver reportes_db.py: necesario desde que las cargas
    # en segundo plano (jobs.py) pueden escribir de verdad en paralelo.
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    _ensure_schema(conn)
    return conn


def _ensure_schema(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS lottery_days (
            date TEXT PRIMARY KEY,
            online_count INTEGER, online_net_sales REAL,
            department_source TEXT, department_pdf_filename TEXT,
            sales REAL, pagos REAL, cash_balance REAL,
            comis REAL, prize_free_plays REAL,
            skoff_count INTEGER, skoff_net_sales REAL,
            pays_units REAL, pays_amount REAL, skoff_sales_amount REAL, sales_comm REAL,
            sales_report_source TEXT, sales_report_pdf_filename TEXT,
            updated_at TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS lottery_blocks (
            iso_year INTEGER NOT NULL,
            iso_week INTEGER NOT NULL,
            chase_bank_date TEXT,
            updated_at TEXT,
            PRIMARY KEY (iso_year, iso_week)
        )
        """
    )
    conn.commit()


def _date_key(value):
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    return str(value)


def _parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return datetime.strptime(str(value), "%Y-%m-%d").date()


def _to_float(value):
    if value is None or value == "":
        return None
    return float(value)


def _now():
    return datetime.now().isoformat(timespec="seconds")


def upsert_department_fields(report_date, online_count, online_net_sales, skoff_count, skoff_net_sales, source, pdf_filename=None):
    """
    D/E/N/O -- lo que sale del Department Sales Report del PDF de cierre
    diario, el MISMO PDF que ya alimenta Reporte Diario (ver
    webapp.py: _persist_reporte_diario_departments, que ahora también llama
    a esto -- primer paso del "interconectado" que pidió el usuario:
    subir un solo PDF alimenta los dos módulos, igual que ya hacía el
    Excel). Nunca toca las columnas de Sales Report (F/G/H/I/K/P/Q/R/S).
    """
    key = _date_key(report_date)
    now = _now()
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO lottery_days
                (date, online_count, online_net_sales, skoff_count, skoff_net_sales,
                 department_source, department_pdf_filename, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(date) DO UPDATE SET
                online_count = excluded.online_count,
                online_net_sales = excluded.online_net_sales,
                skoff_count = excluded.skoff_count,
                skoff_net_sales = excluded.skoff_net_sales,
                department_source = excluded.department_source,
                department_pdf_filename = COALESCE(excluded.department_pdf_filename, lottery_days.department_pdf_filename),
                updated_at = excluded.updated_at
            """,
            (key, online_count, online_net_sales, skoff_count, skoff_net_sales, source, pdf_filename, now),
        )
        conn.commit()
    finally:
        conn.close()


def upsert_sales_report_fields(report_date, fields, source, pdf_filename=None):
    """
    F/G/H/I/K/P/Q/R/S -- lo que sale del "Daily Sales Report" del portal de
    Florida Lottery, un PDF completamente aparte (no comparte nada con
    Reporte Diario) subido directo acá. `fields` trae las mismas claves que
    devuelve reporte_diario.extract_lottery_receipt_fields_from_sales_report
    (sales/pagos/cash_balance/comis/prize_free_plays/pays_units/
    pays_amount/skoff_sales_amount/sales_comm). Nunca toca las columnas de
    Department (online_*/skoff_count/skoff_net_sales).
    """
    key = _date_key(report_date)
    now = _now()
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO lottery_days
                (date, sales, pagos, cash_balance, comis, prize_free_plays,
                 pays_units, pays_amount, skoff_sales_amount, sales_comm,
                 sales_report_source, sales_report_pdf_filename, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(date) DO UPDATE SET
                sales = excluded.sales, pagos = excluded.pagos, cash_balance = excluded.cash_balance,
                comis = excluded.comis, prize_free_plays = excluded.prize_free_plays,
                pays_units = excluded.pays_units, pays_amount = excluded.pays_amount,
                skoff_sales_amount = excluded.skoff_sales_amount, sales_comm = excluded.sales_comm,
                sales_report_source = excluded.sales_report_source,
                sales_report_pdf_filename = COALESCE(excluded.sales_report_pdf_filename, lottery_days.sales_report_pdf_filename),
                updated_at = excluded.updated_at
            """,
            (
                key, fields.get("sales"), fields.get("pagos"), fields.get("cash_balance"),
                fields.get("comis"), fields.get("prize_free_plays"), fields.get("pays_units"),
                fields.get("pays_amount"), fields.get("skoff_sales_amount"), fields.get("sales_comm"),
                source, pdf_filename, now,
            ),
        )
        conn.commit()
    finally:
        conn.close()


def upsert_manual_day(report_date, fields):
    """
    Carga/corrección manual de un día completo (pantalla de edición) --
    pisa TODOS los campos de golpe (a diferencia de los dos upsert de
    arriba, que solo tocan su propio subconjunto), marcando las dos fuentes
    como "manual". `fields` acepta cualquiera de las claves de
    ALL_DAY_FIELDS -- una ausente se guarda como None (en blanco), igual
    que ya hace la fila de un día en el Excel cuando no hay dato.
    """
    key = _date_key(report_date)
    now = _now()
    row = {name: _to_float(fields.get(name)) for name in ALL_DAY_FIELDS}
    row["date"] = key
    row["department_source"] = "manual"
    row["sales_report_source"] = "manual"
    row["updated_at"] = now
    columns = list(row.keys())
    placeholders = ", ".join("?" for _ in columns)
    update_clause = ", ".join(f"{col} = excluded.{col}" for col in columns if col != "date")
    conn = _connect()
    try:
        conn.execute(
            f"INSERT INTO lottery_days ({', '.join(columns)}) VALUES ({placeholders}) "
            f"ON CONFLICT(date) DO UPDATE SET {update_clause}",
            [row[col] for col in columns],
        )
        conn.commit()
    finally:
        conn.close()


def blank_day(report_date):
    """
    Fila "vacía" completa (todas las claves de ALL_DAY_FIELDS presentes en
    None, no ausentes) para un día sin ninguna fila en la base -- evita que
    Jinja reciba un atributo inexistente (Undefined, no None) al armar un
    bloque con días todavía sin cargar, o la pantalla de un día nuevo.
    """
    row = {name: None for name in ALL_DAY_FIELDS}
    row.update(
        date=_date_key(report_date), department_source=None, department_pdf_filename=None,
        sales_report_source=None, sales_report_pdf_filename=None,
    )
    return row


def get_day(report_date):
    key = _date_key(report_date)
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM lottery_days WHERE date = ?", (key,)).fetchone()
    finally:
        conn.close()
    return dict(row) if row is not None else None


def store_pdf_copy(report_date, source_path, original_filename):
    d = _parse_date(report_date)
    dest_dir = os.path.join(_PDF_DIR, f"{d.year:04d}", f"{d.month:02d}")
    os.makedirs(dest_dir, exist_ok=True)
    dest_path = os.path.join(dest_dir, original_filename)
    shutil.copyfile(source_path, dest_path)
    return os.path.relpath(dest_path, _BASE_DIR)


def absolute_pdf_path(relative_path):
    return os.path.join(_BASE_DIR, relative_path)


def get_month_pdf_list(year, month):
    """
    Un renglón por cada PDF guardado este mes (Department y/o Sales Report,
    un día puede tener los dos) -- para la lista "PDFs cargados" de la
    barra lateral (ver CLAUDE.md, "quiero que se puedan guardar también los
    PDF diarios de Lottery"). Cada renglón ya trae el `kind` que espera
    /carga-datos/lottery/dia/<fecha>/pdf/<kind>.
    """
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT date, department_pdf_filename, sales_report_pdf_filename FROM lottery_days "
            "WHERE date LIKE ? AND (department_pdf_filename IS NOT NULL OR sales_report_pdf_filename IS NOT NULL) "
            "ORDER BY date ASC",
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()
    items = []
    for row in rows:
        if row["department_pdf_filename"]:
            items.append({"date": row["date"], "kind": "department", "label": "PDF Diario (Reporte Diario)"})
        if row["sales_report_pdf_filename"]:
            items.append({"date": row["date"], "kind": "sales_report", "label": "Daily Sales Report"})
    return items


def search_pdfs(query, limit=20):
    """
    PDFs de Lottery (Department o Sales Report) cuyo nombre de archivo
    contiene `query` -- pedido explícito del usuario (2026-09-16), buscador
    del header por nombre de PDF.
    """
    like = f"%{query}%"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT date, department_pdf_filename, sales_report_pdf_filename FROM lottery_days "
            "WHERE department_pdf_filename LIKE ? OR sales_report_pdf_filename LIKE ? "
            "ORDER BY date DESC LIMIT ?",
            (like, like, limit),
        ).fetchall()
    finally:
        conn.close()
    items = []
    for row in rows:
        if row["department_pdf_filename"] and query.lower() in row["department_pdf_filename"].lower():
            items.append({"date": row["date"], "kind": "department", "pdf_filename": row["department_pdf_filename"]})
        if row["sales_report_pdf_filename"] and query.lower() in row["sales_report_pdf_filename"].lower():
            items.append({"date": row["date"], "kind": "sales_report", "pdf_filename": row["sales_report_pdf_filename"]})
    return items[:limit]


def set_block_chase_date(iso_year, iso_week, chase_bank_date):
    """
    chase_bank_date en None borra lo cargado para ese bloque (vuelve a la
    sugerencia automática, encadenada -- ver _nearest_confirmed_chase_date).
    Si se carga una fecha real, se usa como ancla para auto-corregir
    cualquier OTRO bloque POSTERIOR ya confirmado que no siga la cadencia
    de 7 días entre bloques consecutivos -- pedido explícito del usuario
    (2026-09-12): "de forma automatica si se detecto que esta mal que se
    corrijan el resto de fechas". Los bloques ANTERIORES al que se acaba de
    confirmar nunca se tocan (ver _realign_chase_dates, corregido
    2026-09-14: cambiar una fecha de septiembre no puede alterar años
    anteriores). Devuelve la lista de bloques que se corrigieron solos
    (puede estar vacía) para que la ruta pueda avisarlo.
    """
    now = _now()
    key = _date_key(chase_bank_date) if chase_bank_date else None
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO lottery_blocks (iso_year, iso_week, chase_bank_date, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(iso_year, iso_week) DO UPDATE SET
                chase_bank_date = excluded.chase_bank_date, updated_at = excluded.updated_at
            """,
            (iso_year, iso_week, key, now),
        )
        corrected = []
        if key:
            corrected = _realign_chase_dates(conn, iso_year, iso_week, _parse_date(key), now)
        conn.commit()
        return corrected
    finally:
        conn.close()


def _realign_chase_dates(conn, anchor_iso_year, anchor_iso_week, anchor_date, now):
    """
    Recorre los demás bloques ya confirmados y los realinea a +/-7 días por
    semana de diferencia contra el ancla recién guardada -- SOLO los
    bloques posteriores al ancla. Pedido explícito del usuario (2026-09-14):
    "si alguien cambia una fecha en un cuadro de septiembre, terminaria
    cambiando todas las fechas de todos los anos anteriores, eso no podria
    ser" -- antes de este fix, un bloque confirmado hace mucho (ej. de
    2024) que por el motivo que fuera no seguía la cadencia de 7 días
    contra el bloque recién tocado terminaba reescribiéndose igual, sin
    importar si estaba antes o después en el tiempo. Ahora un bloque
    anterior al ancla nunca se toca -- solo bloques con lunes posterior al
    del ancla.
    """
    anchor_monday, _ = _iso_week_bounds(anchor_iso_year, anchor_iso_week)
    rows = conn.execute(
        "SELECT iso_year, iso_week, chase_bank_date FROM lottery_blocks WHERE chase_bank_date IS NOT NULL"
    ).fetchall()
    corrected = []
    for row in rows:
        if row["iso_year"] == anchor_iso_year and row["iso_week"] == anchor_iso_week:
            continue
        row_monday, _ = _iso_week_bounds(row["iso_year"], row["iso_week"])
        if row_monday <= anchor_monday:
            continue
        expected_key = _date_key(anchor_date + (row_monday - anchor_monday))
        if row["chase_bank_date"] != expected_key:
            conn.execute(
                "UPDATE lottery_blocks SET chase_bank_date = ?, updated_at = ? WHERE iso_year = ? AND iso_week = ?",
                (expected_key, now, row["iso_year"], row["iso_week"]),
            )
            corrected.append(
                {
                    "iso_year": row["iso_year"], "iso_week": row["iso_week"],
                    "old": row["chase_bank_date"], "new": expected_key,
                }
            )
    return corrected


def _nearest_confirmed_chase_date(conn, iso_year, iso_week):
    """
    Entre todos los bloques con una fecha de Chase Bank ya confirmada,
    busca el más cercano (en semanas) a (iso_year, iso_week) -- para
    encadenar la sugerencia de un bloque sin confirmar todavía a +/-7 días
    por semana de diferencia, en vez de adivinar con una fórmula fija.
    Devuelve (monday_del_bloque_ancla, fecha_confirmada) o None si todavía
    no hay ningún bloque confirmado en toda la base.
    """
    rows = conn.execute(
        "SELECT iso_year, iso_week, chase_bank_date FROM lottery_blocks WHERE chase_bank_date IS NOT NULL"
    ).fetchall()
    if not rows:
        return None
    target_monday = date.fromisocalendar(iso_year, iso_week, 1)
    best = None
    best_distance = None
    for row in rows:
        row_monday = date.fromisocalendar(row["iso_year"], row["iso_week"], 1)
        distance = abs((row_monday - target_monday).days)
        if best_distance is None or distance < best_distance:
            best = (row_monday, _parse_date(row["chase_bank_date"]))
            best_distance = distance
    return best


# ---------------------------------------------------------------------------
# Vista de reporte: bloques de 7 días + Subtotal + Debito, calculados igual
# que las fórmulas reales del Excel -- confirmado línea por línea contra un
# Excel real antes de escribir esto.
# ---------------------------------------------------------------------------

# columna Excel -> (clave, fórmula real) para lo que se computa por día.
_DAY_FORMULAS = {
    "comis_ratio": lambda d: _safe_div(d.get("comis"), d.get("sales")),          # J =+I/F
    "prize_ratio": lambda d: _safe_div(d.get("prize_free_plays"), d.get("sales")),  # L =+K/F
    "total_comm": lambda d: _safe_add(d.get("comis"), d.get("prize_free_plays")),   # M =+I+K
    "sales_comm_ratio": lambda d: _safe_div(d.get("sales_comm"), d.get("skoff_sales_amount")),  # T =+S/R
    "net_total": lambda d: _safe_add3(d.get("pays_amount"), d.get("skoff_sales_amount"), d.get("sales_comm")),  # U =+Q+R+S
    "cuenta_final": lambda d: _safe_sub(0, _safe_add(d.get("pagos"), d.get("pays_amount"))),  # X =-G-Q
}

# D, E, F, G, I, K, M, P, Q, R, S -- las que suma la fila de Subtotal (H y
# N/O quedan afuera a propósito, así viene la plantilla real).
_SUBTOTAL_SUM_FIELDS = (
    "online_count", "online_net_sales", "sales", "pagos", "comis",
    "prize_free_plays", "total_comm", "pays_units", "pays_amount",
    "skoff_sales_amount", "sales_comm",
)


def _safe_div(a, b):
    if a is None or b in (None, 0):
        return None
    return a / b


def _safe_add(a, b):
    if a is None and b is None:
        return None
    return (a or 0) + (b or 0)


def _safe_sub(a, b):
    if a is None and b is None:
        return None
    return (a or 0) - (b or 0)


def _safe_add3(a, b, c):
    if a is None and b is None and c is None:
        return None
    return (a or 0) + (b or 0) + (c or 0)


def decorate_day(row):
    """Agrega las columnas calculadas (J/L/M/T/U/X) a una fila cruda de lottery_days."""
    day = dict(row)
    for key, fn in _DAY_FORMULAS.items():
        day[key] = fn(day)
    return day


def _has_any_data(day):
    return any(day.get(field) is not None for field in ALL_DAY_FIELDS)


def _iso_week_bounds(iso_year, iso_week):
    monday = date.fromisocalendar(iso_year, iso_week, 1)
    return monday, monday + timedelta(days=6)


def _build_block(conn, iso_year, iso_week, days_by_date):
    monday, sunday = _iso_week_bounds(iso_year, iso_week)
    day_rows = []
    for offset in range(7):
        d = monday + timedelta(days=offset)
        raw = days_by_date.get(d.isoformat())
        day_rows.append(decorate_day(raw) if raw is not None else decorate_day(blank_day(d)))

    subtotal = {}
    for field in _SUBTOTAL_SUM_FIELDS:
        values = [d.get(field) for d in day_rows if d.get(field) is not None]
        subtotal[field] = round(sum(values), 2) if values else None
    subtotal["prize_ratio"] = _safe_div(subtotal.get("prize_free_plays"), subtotal.get("sales"))

    debito = {}
    debito["online_net_sales"] = _safe_sub(subtotal.get("online_net_sales"), subtotal.get("sales"))  # E =+E-F
    debito["sales"] = _safe_add(  # F =+F+G+I+K+10
        _safe_add(_safe_add(subtotal.get("sales"), subtotal.get("pagos")), subtotal.get("comis")),
        _safe_add(subtotal.get("prize_free_plays"), 10),
    )
    debito["pays_amount"] = _safe_add3(subtotal.get("pays_amount"), subtotal.get("skoff_sales_amount"), subtotal.get("sales_comm"))  # Q
    debito["net_debit"] = _safe_add(debito["sales"], debito["pays_amount"])  # V =+F+Q (de la fila Debito)

    block_row = conn.execute(
        "SELECT chase_bank_date FROM lottery_blocks WHERE iso_year = ? AND iso_week = ?",
        (iso_year, iso_week),
    ).fetchone()
    chase_bank_date = block_row["chase_bank_date"] if block_row else None

    # Sugerencia autocompletada -- pedido explícito del usuario (2026-09-12):
    # "que se autocomplete el dia que en teoria pagarian... el patron de
    # cada tantos dias luego del bloque es cuando pagan". Se encadena desde
    # el bloque confirmado más cercano (+/-7 días por semana de diferencia)
    # en vez de adivinar siempre "miércoles siguiente" -- ese heurístico
    # queda solo como último recurso, para cuando TODAVÍA no hay ningún
    # bloque confirmado en toda la base (primera vez que se usa el módulo).
    anchor = _nearest_confirmed_chase_date(conn, iso_year, iso_week)
    if anchor is not None:
        anchor_monday, anchor_date = anchor
        suggested_chase_date = anchor_date + (monday - anchor_monday)
    else:
        suggested_chase_date = sunday + timedelta(days=3)  # miércoles siguiente -- mismo patrón visto en el Excel real

    return {
        "iso_year": iso_year,
        "iso_week": iso_week,
        "start": monday.isoformat(),
        "end": sunday.isoformat(),
        "days": day_rows,
        "has_data": any(_has_any_data(d) for d in day_rows),
        "subtotal": subtotal,
        "debito": debito,
        "chase_bank_date": chase_bank_date,
        "suggested_chase_date": suggested_chase_date.isoformat(),
    }


def build_month_blocks(year, month):
    """
    Todos los bloques de 7 días (semana ISO) que tocan el mes pedido -- el
    primero y el último normalmente se pasan unos días al mes vecino,
    exactamente como en el Excel real. Devuelve una lista de bloques
    ordenada, cada uno con sus 7 días + Subtotal + Debito ya calculados.
    """
    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])
    first_iso_year, first_iso_week, _ = first_day.isocalendar()
    last_iso_year, last_iso_week, _ = last_day.isocalendar()

    range_start, _ = _iso_week_bounds(first_iso_year, first_iso_week)
    _, range_end = _iso_week_bounds(last_iso_year, last_iso_week)

    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM lottery_days WHERE date >= ? AND date <= ? ORDER BY date",
            (range_start.isoformat(), range_end.isoformat()),
        ).fetchall()
        days_by_date = {row["date"]: dict(row) for row in rows}

        blocks = []
        iso_year, iso_week = first_iso_year, first_iso_week
        while (iso_year, iso_week) <= (last_iso_year, last_iso_week):
            blocks.append(_build_block(conn, iso_year, iso_week, days_by_date))
            monday, _ = _iso_week_bounds(iso_year, iso_week)
            next_monday = monday + timedelta(days=7)
            iso_year, iso_week, _ = next_monday.isocalendar()
    finally:
        conn.close()
    return blocks


def monthly_debit_total(year, month):
    """
    Suma de V (Debito) de cada bloque cuya fecha de Chase Bank -- la real,
    tipeada a mano -- cae dentro de (year, month). Un bloque sin fecha de
    Chase Bank cargada todavía queda afuera de cualquier mes (no se adivina
    con la sugerencia) -- mismo criterio de no inventar un dato que
    representa un hecho bancario real todavía no ocurrido/registrado.
    """
    total = 0.0
    any_value = False
    for block in build_month_blocks(year, month):
        if not block["chase_bank_date"]:
            continue
        d = _parse_date(block["chase_bank_date"])
        if d.year == year and d.month == month and block["debito"]["net_debit"] is not None:
            total += block["debito"]["net_debit"]
            any_value = True
    return round(total, 2) if any_value else None
