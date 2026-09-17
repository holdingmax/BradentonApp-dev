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
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS lottery_month_closing (
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            gastos_admin REAL,
            caja_skoff_ajuste REAL,
            updated_at TEXT,
            PRIMARY KEY (year, month)
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


# ---------------------------------------------------------------------------
# Export a Excel/PDF -- pedido explícito del usuario (2026-09-19, misma
# tanda que Store Info/Caja): "empezemos a trabajar en lo mismo de exportar
# excel y pdf de lottery". Investigado contra el Excel real (LOTTERY.
# Analisis 08.2026, columna por columna con openpyxl antes de escribir
# nada, nunca asumido) -- misma estructura de siempre (bloques de 7 días
# sin espacio entre ellos, Subtotal/Debito con SUS MISMAS fórmulas reales
# confirmadas, ver _DAY_FORMULAS/_build_block más arriba) y los mismos
# colores reales (gris D9D9D9 para el grupo ONLINE/Terminal, amarillo
# FFFF00 para SKOFF, naranja FFC000 para la columna "TOTAL", verde 92D050
# para "CUENTA FINAL"/Debito).
# ---------------------------------------------------------------------------

# La columna "C" (TOTAL) del Excel real venía como una columna oculta,
# sin datos -- confirmado que en la práctica se puede borrar del todo,
# nunca hace falta (2026-09-16, pedido explícito del usuario). Se sacó
# por completo del export -- ya no existe ningún hueco/columna escondida,
# todo el layout de acá en más corre un lugar a la izquierda respecto a
# las notas viejas de esta sección (D pasa a ser C, E pasa a ser D, etc.)
# -- confirmado contra el inventario técnico real que mandó el usuario
# ("LOTTERY_estructura_tecnica.json", generado con esa columna ya
# eliminada) para no perder ningún ancho/color/fórmula en el camino.
_LOTTERY_EXPORT_HEADERS = (
    # (texto fila 3, ancho de columna en el Excel real)
    ("Fecha", 11.9), ("Fecha", 11.4),
    ("COUNT", 6.7), ("NET SALES REPORT", 9.6), ("SALES", 10.1), ("Pagos", 9.9),
    ("Cash Balance", 11.1), ("Comis", 11.6), ("En %", 8.6), ("Prize Free Plays", 10.7),
    ("En %", 8.3), ("TOTAL SALES COM", 9.7),
    ("COUNT", 6.7), ("NET SALES", 10.6), ("Pays", 6.3), ("Pagos", 11.1),
    ("Books Settled", 10.4), ("Sales Com", 11.1), ("En %", 9.1), ("NET TOTAL", 10.7),
    (None, 12.1), (None, 11.1), (None, 13.0),
)


# Columnas en negrita en las filas de día -- corrección final del usuario
# (2026-09-16): fechas + D/F/H/I/J/K/Q/S/U/V/W (ya no C/E/G/P/R/T, que
# habían quedado de una lectura de columnas equivocada en la ronda anterior).
_LOTTERY_DAY_BOLD_COLS = {1, 2, 4, 6, 8, 9, 10, 11, 17, 19, 21, 22, 23}

# Comis/Prize Free Plays (online) y Pagos/Sales Com (skoff) van con la
# fuente en rojo -- confirmado contra "LOTTERY. Analisis 09.2026
# WEB.xlsx" (font.color real FFFF0000 en esas 4 columnas, no un simple
# "negativo en rojo" del formato numérico). F se agregó a pedido explícito
# del usuario (2026-09-16, cuarta ronda de correcciones de formato).
_LOTTERY_RED_FONT_COLS = {6, 8, 10, 16, 18}

# Los ratios "En %" (antes J/L/T) van con formato de porcentaje real, no
# el formato contable con [Red] que se les aplicaba antes por error.
_LOTTERY_PERCENT_COLS = {9, 11, 19}

# C/M/O (los tres "COUNT"/"Pays" del layout) van sin decimales -- pedido
# explícito del usuario (2026-09-16): "8, 90" en vez de "8,00 90,00". El
# resto de las columnas numéricas crudas (D/E/F/G/H/J/N/P/Q/R) sí llevan
# 2 decimales, como siempre.
_LOTTERY_INT_COLS = {3, 13, 15}
_LOTTERY_INT_FMT = "#,##0_ ;[Red]\\-#,##0\\ "

# Alineación, cuarta ronda de correcciones (2026-09-16): C..M más O
# centradas; N/P/Q/R/S/T alineadas a la derecha (reemplaza la ronda
# anterior, que tenía Q centrada y P/R/S/T a la izquierda). Se aplica a
# cada fila (día, Subtotal, Debito) donde estas columnas tengan contenido.
_LOTTERY_CENTER_COLS = set(range(3, 14)) | {15}
_LOTTERY_RIGHT_COLS = {14, 16, 17, 18, 19, 20}


def _lottery_apply_column_alignment(sheet, row, styles):
    for col in _LOTTERY_CENTER_COLS:
        sheet.cell(row=row, column=col).alignment = styles["center"]
    for col in _LOTTERY_RIGHT_COLS:
        sheet.cell(row=row, column=col).alignment = styles["right"]


def _lottery_write_day_row(sheet, row, day, styles):
    """Una fila de día (C..H, J, M..R crudos + I/K/L/S/T/W calculados con
    la misma fórmula real -- ver _DAY_FORMULAS)."""
    XlFont, THIN_BORDER = styles["font"], styles["border"]
    BOLD_FONT = XlFont(bold=True)
    RED_FONT = XlFont(bold=True, color="FFFF0000")
    business_date = _parse_date(day["date"])
    values = {
        1: business_date - timedelta(days=1), 2: business_date,
        3: day.get("online_count"), 4: day.get("online_net_sales"),
        5: day.get("sales"), 6: day.get("pagos"), 7: day.get("cash_balance"),
        8: day.get("comis"), 10: day.get("prize_free_plays"),
        13: day.get("skoff_count"), 14: day.get("skoff_net_sales"),
        15: day.get("pays_units"), 16: day.get("pays_amount"),
        17: day.get("skoff_sales_amount"), 18: day.get("sales_comm"),
    }
    formulas = {
        9: f"=+H{row}/E{row}", 11: f"=+J{row}/E{row}", 12: f"=+H{row}+J{row}",
        19: f"=+R{row}/Q{row}", 20: f"=+P{row}+Q{row}+R{row}", 23: f"=-F{row}-P{row}",
    }
    for col, value in values.items():
        cell = sheet.cell(row=row, column=col, value=value)
        cell.border = THIN_BORDER
        if col in (1, 2):
            cell.number_format = "mm-dd-yy"
            cell.alignment = styles["center"]
        else:
            cell.number_format = styles["int_fmt"] if col in _LOTTERY_INT_COLS else styles["plain_fmt"]
        if col in _LOTTERY_RED_FONT_COLS:
            cell.font = RED_FONT
        elif col in _LOTTERY_DAY_BOLD_COLS:
            cell.font = BOLD_FONT
    for col, formula in formulas.items():
        cell = sheet.cell(row=row, column=col, value=formula)
        # La columna W ("Cuenta Final" del día) va sin borde en el Excel
        # real -- confirmado contra "LOTTERY. Analisis 09.2026 WEB.xlsx".
        if col != 23:
            cell.border = THIN_BORDER
        cell.number_format = "0.00%" if col in _LOTTERY_PERCENT_COLS else styles["plain_fmt"]
        if col in _LOTTERY_DAY_BOLD_COLS:
            cell.font = BOLD_FONT
        if col == 23:
            # "Centrados al medio" pedido explícito del usuario (2026-09-16).
            cell.alignment = styles["center"]
    for col, fill in styles["day_fill_by_col"].items():
        sheet.cell(row=row, column=col).fill = fill
    # U ("Cuenta Final" del bloque) va sin valor en las filas de día -- solo
    # el relleno verde -- pero pedido explícito del usuario (2026-09-16,
    # quinta ronda): igual necesitan su propio borde, como el resto de la
    # tabla.
    sheet.cell(row=row, column=21).border = THIN_BORDER
    sheet.cell(row=row, column=20).font = XlFont(bold=True)
    _lottery_apply_column_alignment(sheet, row, styles)
    _lottery_apply_group_borders(sheet, row, styles)
    # Filas de día ligeramente más altas que el default -- pedido explícito
    # del usuario (2026-09-16), "solo un poco".
    sheet.row_dimensions[row].height = 16.5


def _lottery_apply_group_borders(sheet, row, styles):
    """Los tres grupos de columnas (ONLINE C-L / SKOFF M-T / CUENTA FINAL
    U) van separados por un borde más grueso ("medium") en el Excel real
    -- confirmado columna por columna contra "LOTTERY. Analisis 09.2026
    WEB.xlsx". Se aplica encima del borde fino ya escrito, sin pisar los
    otros tres lados de cada celda."""
    for col, sides in styles["group_border_sides"].items():
        cell = sheet.cell(row=row, column=col)
        base = cell.border
        cell.border = styles["border_factory"](
            left=sides.get("left", base.left), right=sides.get("right", base.right),
            top=base.top, bottom=base.bottom,
        )


def _lottery_write_block(sheet, start_row, block, styles):
    """Escribe los 7 días + Subtotal + Debito de un bloque, arrancando en
    `start_row` -- devuelve la fila siguiente (sin ningún espacio entre
    bloques, igual que el archivo real)."""
    XlFont, THIN_BORDER = styles["font"], styles["border"]
    RED_BOLD = XlFont(bold=True, color="FFFF0000")
    for offset, day in enumerate(block["days"]):
        _lottery_write_day_row(sheet, start_row + offset, day, styles)

    sub_row = start_row + 7
    d1, d2 = start_row, start_row + 6
    for col, letter in ((3, "C"), (4, "D"), (5, "E"), (6, "F"), (8, "H"), (10, "J"), (12, "L"), (15, "O"), (16, "P"), (17, "Q"), (18, "R")):
        cell = sheet.cell(row=sub_row, column=col, value=f"=SUM({letter}{d1}:{letter}{d2})")
        cell.border = THIN_BORDER
        cell.number_format = styles["int_fmt"] if col in _LOTTERY_INT_COLS else styles["plain_fmt"]
        cell.font = RED_BOLD if col in _LOTTERY_RED_FONT_COLS else XlFont(bold=True)
    lratio = sheet.cell(row=sub_row, column=11, value=f"=+J{sub_row}/E{sub_row}")
    lratio.border = THIN_BORDER
    lratio.font = XlFont(bold=True)
    for col, fill in styles["subtotal_fill_by_col"].items():
        sheet.cell(row=sub_row, column=col).fill = fill
    _lottery_apply_column_alignment(sheet, sub_row, styles)
    _lottery_apply_group_borders(sheet, sub_row, styles)

    deb_row = sub_row + 1
    deb_values = {
        4: f"=+D{sub_row}-E{sub_row}",
        5: f"=+E{sub_row}+F{sub_row}+H{sub_row}+J{sub_row}+10",
        16: f"=+P{sub_row}+Q{sub_row}+R{sub_row}",
        20: "Debito",
        21: f"=+E{deb_row}+P{deb_row}",
    }
    for col, value in deb_values.items():
        cell = sheet.cell(row=deb_row, column=col, value=value)
        cell.border = THIN_BORDER
        if col in (4, 5, 16, 21):
            cell.number_format = styles["plain_fmt"]
        # Columna U ("Cuenta Final" del bloque) -- pedido explícito del
        # usuario (2026-09-16): son pocos valores (uno por bloque), así que
        # se les sube el tamaño de texto para que resalten (ajustado a 12
        # en la corrección siguiente, misma fecha -- 14 quedaba muy grande).
        cell.font = XlFont(bold=True, size=12) if col == 21 else XlFont(bold=True)
    if block.get("chase_bank_date"):
        d = _parse_date(block["chase_bank_date"])
        chase_cell = sheet.cell(row=deb_row, column=22, value=f"Chase Bank {d.strftime('%d/%m/%Y')}")
        chase_cell.font = XlFont(bold=True)
        # "Ajustar el texto" pedido explícito del usuario (2026-09-16) -- el
        # texto de dos líneas (Chase Bank + fecha) no debe sobresalir de la
        # celda.
        chase_cell.alignment = styles["wrap_chase"]
    for col, fill in styles["debito_fill_by_col"].items():
        sheet.cell(row=deb_row, column=col).fill = fill
    # Las celdas del Debito entre E:J y P:R se ven en el Excel real como
    # una sola celda unificada (más ancha, centrada) -- y esa fila queda
    # más alta (30 en vez de 15.75) para que se note el cambio de bloque.
    # El merge tiene que hacerse ANTES de terminar de poner bordes: openpyxl
    # borra el borde de cualquier celda interior de un rango recién
    # mergeado, así que asignarlo después es lo único que lo deja visible.
    sheet.merge_cells(start_row=deb_row, start_column=5, end_row=deb_row, end_column=10)
    sheet.merge_cells(start_row=deb_row, start_column=16, end_row=deb_row, end_column=18)
    # Todas las celdas vacías que rodean el resultado del bloque (entre C y
    # L, y entre M y T) llevan borde completo -- pedido explícito del
    # usuario (2026-09-16), no solo las que ya tenían un valor propio.
    for col in range(3, 21):
        sheet.cell(row=deb_row, column=col).border = THIN_BORDER
    _lottery_apply_column_alignment(sheet, deb_row, styles)
    sheet.cell(row=deb_row, column=5).alignment = styles["center"]
    sheet.cell(row=deb_row, column=16).alignment = styles["center"]
    sheet.row_dimensions[deb_row].height = 30.0
    _lottery_apply_group_borders(sheet, deb_row, styles)

    return deb_row + 1


def build_lottery_export_workbook(year, month, dest_path):
    """
    Excel NUEVO (nunca toca el archivo real) con los bloques de Lottery del
    mes, mismo formato/colores/fórmulas que el Excel real -- confirmado
    columna por columna contra `LOTTERY. Analisis 08.2026.xlsx` antes de
    escribir esto (ver el comentario de arriba de esta sección).
    """
    import openpyxl
    from openpyxl.styles import Alignment as XlAlignment, Border as XlBorder, Font as XlFont, PatternFill, Side as XlSide

    GRAY = PatternFill("solid", fgColor="FFD9D9D9")
    YELLOW = PatternFill("solid", fgColor="FFFFFF00")
    ORANGE = PatternFill("solid", fgColor="FFFFC000")
    GREEN = PatternFill("solid", fgColor="FF92D050")
    PLAIN_FMT = '#,##0.00_ ;[Red]\\-#,##0.00\\ '
    MONEY_FMT = '"$"\\ #,##0.00'
    THIN_SIDE = XlSide(style="thin", color="FF000000")
    MEDIUM_SIDE = XlSide(style="medium", color="FF000000")
    THIN_BORDER = XlBorder(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
    CENTER = XlAlignment(horizontal="center", vertical="center")
    LEFT = XlAlignment(horizontal="left", vertical="center")
    RIGHT = XlAlignment(horizontal="right", vertical="center")
    WRAP_CHASE = XlAlignment(wrap_text=True, vertical="center")

    styles = {
        "font": XlFont, "border": THIN_BORDER, "center": CENTER, "left": LEFT, "right": RIGHT,
        "wrap_chase": WRAP_CHASE, "plain_fmt": PLAIN_FMT, "int_fmt": _LOTTERY_INT_FMT,
        "money_fmt": MONEY_FMT, "yellow": YELLOW, "green": GREEN, "orange": ORANGE,
        # L (TOTAL SALES COM) va con relleno gris en las filas de día --
        # pedido explícito del usuario (2026-09-16), quinta ronda.
        "day_fill_by_col": {12: GRAY, 20: YELLOW, 21: GREEN, 23: ORANGE},
        "subtotal_fill_by_col": {5: GREEN, 6: GREEN, 12: GREEN, 15: YELLOW, 16: GREEN, 17: YELLOW, 18: YELLOW, 21: GREEN},
        "debito_fill_by_col": {4: YELLOW, 5: GRAY, 16: YELLOW, 21: GREEN},
        # Separadores más gruesos entre los 3 grupos de columnas (ONLINE
        # C-L / SKOFF M-T / CUENTA FINAL U) -- confirmado contra el Excel
        # real, 2026-09-16. D (NET SALES REPORT) además queda enmarcada
        # con borde medium a los dos lados -- pedido explícito del usuario,
        # quinta ronda de correcciones de formato (misma fecha).
        "border_factory": XlBorder,
        "group_border_sides": {
            3: {"left": MEDIUM_SIDE}, 4: {"left": MEDIUM_SIDE, "right": MEDIUM_SIDE},
            5: {"left": MEDIUM_SIDE},
            12: {"right": MEDIUM_SIDE},
            21: {"left": MEDIUM_SIDE, "right": MEDIUM_SIDE},
        },
    }

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = f"{month:02d}.{year}"
    sheet.row_dimensions[1].height = 16.5
    # Pedido explícito del usuario (2026-09-16): que el Excel se abra
    # siempre con un zoom del 85%, sin que el usuario tenga que ajustarlo.
    sheet.sheet_view.zoomScale = 85

    sheet.cell(row=2, column=3, value="ONLINE").fill = GRAY
    sheet.cell(row=2, column=3).font = XlFont(bold=True)
    sheet.cell(row=2, column=3).alignment = CENTER
    sheet.merge_cells(start_row=2, start_column=3, end_row=2, end_column=11)
    sheet.cell(row=2, column=13, value="SKOFF = SCRATCH-OFF ").fill = YELLOW
    sheet.cell(row=2, column=13).font = XlFont(bold=True)
    sheet.cell(row=2, column=13).alignment = CENTER
    sheet.merge_cells(start_row=2, start_column=13, end_row=2, end_column=20)
    sheet.cell(row=2, column=21, value="CUENTA FINAL").fill = GREEN
    sheet.cell(row=2, column=21).font = XlFont(bold=True)
    # "Ajustar el texto" pedido explícito del usuario (2026-09-16) -- el
    # banner "CUENTA FINAL" en una columna angosta necesita wrap_text para
    # no cortarse, igual que el Excel real (confirmado ahí también).
    sheet.cell(row=2, column=21).alignment = XlAlignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.merge_cells(start_row=2, start_column=21, end_row=3, end_column=21)
    sheet.row_dimensions[2].height = 19.95

    no_fill_cols = {8, 10}  # Comis/Prize Free Plays -- sin relleno en el real
    for col, (text, width) in enumerate(_LOTTERY_EXPORT_HEADERS, start=1):
        if text is not None:
            cell = sheet.cell(row=3, column=col, value=text)
            cell.font = XlFont(bold=True, size=10)
            cell.alignment = XlAlignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER
            if 13 <= col <= 20:
                cell.fill = YELLOW
            elif col not in no_fill_cols:
                cell.fill = GRAY
            if col == 4:
                # D (NET SALES REPORT) enmarcada -- ver group_border_sides.
                cell.border = XlBorder(left=MEDIUM_SIDE, right=MEDIUM_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
        sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    sheet.row_dimensions[3].height = 39.0

    row = 4
    for block in build_month_blocks(year, month):
        row = _lottery_write_block(sheet, row, block, styles)

    _lottery_write_closing_section(sheet, row - 1, styles)

    workbook.save(dest_path)
    return dest_path


def _lottery_write_closing_section(sheet, last_debito_row, styles):
    """
    Escribe, justo debajo del último bloque semanal, las dos tablas de
    cierre de mes -- "LIQUIDACION CIERRE LOTTERY" / "LIQUIDACION CIERRE
    RECAUDACION COMISIONES" -- con las MISMAS fórmulas y el mismo diseño
    que el Excel real (confirmado contra "LOTTERY. Analisis 09.2026
    WEB.xlsx", fila por fila, antes de escribir esto).

    Las dos filas de "TOTAL DEL MES" (acá llamadas r_total1/r_total2) se
    arman sumando los mismos Subtotal de cada bloque -- a diferencia del
    archivo real de referencia (que tenía algunas de estas fórmulas
    "atrasadas", sin actualizar tras agregarse un bloque nuevo a mano),
    acá siempre suman TODOS los bloques del mes, para que el total sea
    correcto sin importar cuántos bloques tenga un mes en particular.
    """
    import openpyxl
    from openpyxl.styles import Border as XlBorder, Side as XlSide

    XlFont, THIN_BORDER = styles["font"], styles["border"]
    YELLOW, GREEN, CENTER = styles["yellow"], styles["green"], styles["center"]
    MONEY_FMT, PLAIN_FMT = styles["money_fmt"], styles["plain_fmt"]
    THIN_SIDE = XlSide(style="thin", color="FF000000")

    # Un bloque son 9 filas (7 días + Subtotal + Debito) sin espacio entre
    # ellos, arrancando en la fila 4 -- de ahí se derivan los Subtotal.
    subtotal_rows = list(range(11, last_debito_row, 9))
    first_day_row = 4

    r_total1 = last_debito_row + 1
    r_total2 = last_debito_row + 2
    r_helper = last_debito_row + 5
    r_title1 = last_debito_row + 6
    r_online, r_skoff, r_gastos = last_debito_row + 7, last_debito_row + 8, last_debito_row + 9
    r_pagar1, r_pagar2 = last_debito_row + 10, last_debito_row + 11
    r_sum1, r_diff1 = last_debito_row + 12, last_debito_row + 13
    r_title2 = last_debito_row + 14
    r_pagarcom, r_pagaronline, r_pagarskoff = last_debito_row + 15, last_debito_row + 16, last_debito_row + 17
    r_chase, r_comonline, r_comskoff, r_caja = (
        last_debito_row + 18, last_debito_row + 19, last_debito_row + 20, last_debito_row + 21,
    )
    r_blank2, r_sum2, r_diff2 = last_debito_row + 22, last_debito_row + 23, last_debito_row + 24

    def joined(letter):
        return "=+" + "+".join(f"{letter}{r}" for r in subtotal_rows)

    # --- TOTAL DEL MES (dos filas justo después del último bloque) ---
    total1 = {
        "C": f"=SUM(C{first_day_row}:C{last_debito_row})",
        "D": joined("D"), "E": joined("E"), "F": joined("F"), "L": joined("L"),
        "N": f"=SUM(N{first_day_row}:N{last_debito_row})",
        "P": joined("P"), "Q": joined("Q"), "R": joined("R"),
        "U": f"=SUM(U{first_day_row}:U{last_debito_row})",
        "W": f"=SUM(W{first_day_row}:W{last_debito_row})+0",
    }
    fill_by_col_total1 = {16: YELLOW, 17: YELLOW, 18: YELLOW, 23: styles["orange"]}
    INT_FMT = "#,##0_ ;[Red]\\-#,##0\\ "
    for letter, formula in total1.items():
        col = openpyxl.utils.column_index_from_string(letter)
        cell = sheet.cell(row=r_total1, column=col, value=formula)
        cell.font = XlFont(bold=True)
        cell.number_format = INT_FMT if letter == "C" else PLAIN_FMT
        if col in fill_by_col_total1:
            cell.fill = fill_by_col_total1[col]

    total2 = {
        "D": f"=+E{r_total1}-D{r_total1}",
        "E": joined("E"),
        "L": f"=+L{r_total1}/E{r_total1}",
        "N": f"=SUM(N{first_day_row}:N{last_debito_row})+0",
    }
    for letter, formula in total2.items():
        col = openpyxl.utils.column_index_from_string(letter)
        cell = sheet.cell(row=r_total2, column=col, value=formula)
        cell.font = XlFont(bold=True)
        cell.number_format = "0.00%" if letter == "L" else PLAIN_FMT
        if letter in ("E", "N"):
            cell.fill = styles["orange"]

    # --- Fila auxiliar (helper, sin etiqueta visible en el Excel real) ---
    g = sheet.cell(row=r_helper, column=7, value=f"=+G{r_online}+G{r_skoff}")
    g.font = XlFont(bold=True)
    g.number_format = MONEY_FMT
    h = sheet.cell(row=r_helper, column=8, value=f"=+G{r_helper}-N{r_title1}")
    h.number_format = "_-* #,##0.00_-;\\-* #,##0.00_-;_-* \"-\"??_-;_-@_-"

    # --- Título 1 ---
    title1 = sheet.cell(row=r_title1, column=4, value="LIQUIDACION CIERRE LOTTERY ")
    title1.font = XlFont(bold=True, underline="single")
    title1.alignment = CENTER
    sheet.merge_cells(start_row=r_title1, start_column=4, end_row=r_title1, end_column=6)
    for col in (4, 5, 6):
        sheet.cell(row=r_title1, column=col).border = XlBorder(bottom=THIN_BORDER.bottom)

    def box_row(row_num, label_col4=None, label_col5=None, fill_by_col=None, values=None):
        fill_by_col = fill_by_col or {}
        values = values or {}
        for col in range(4, 9):
            cell = sheet.cell(row=row_num, column=col)
            cell.font = XlFont(bold=True)
            cell.border = THIN_BORDER
            cell.fill = fill_by_col.get(col, YELLOW)
        if label_col4 is not None:
            sheet.cell(row=row_num, column=4, value=label_col4)
        if label_col5 is not None:
            sheet.cell(row=row_num, column=5, value=label_col5)
        for col, value in values.items():
            cell = sheet.cell(row=row_num, column=col, value=value)
            cell.number_format = MONEY_FMT

    box_row(r_online, "Caja-On Line", values={7: f"=+E{r_total2}+0"})
    box_row(r_skoff, "Caja-Skoff", values={7: f"=+N{r_total1}+4774.2"})
    box_row(r_gastos, "Gastos Adminits-Loteria", values={7: 150})
    box_row(r_pagar1, "a", "Lottery a Pagar", fill_by_col={5: GREEN, 6: GREEN, 8: GREEN}, values={8: f"=+E{r_total2}+0"})
    box_row(r_pagar2, "a", "Lottery a Pagar", fill_by_col={5: GREEN, 6: GREEN, 8: GREEN}, values={8: f"=+N{r_total1}+G{r_gastos}+4774.2"})

    sum1_g = sheet.cell(row=r_sum1, column=7, value=f"=SUM(G{r_online}:G{r_gastos+2})")
    sum1_g.font, sum1_g.border, sum1_g.fill, sum1_g.number_format = XlFont(bold=True), THIN_BORDER, YELLOW, MONEY_FMT
    sum1_h = sheet.cell(row=r_sum1, column=8, value=f"=SUM(H{r_online}:H{r_gastos+2})")
    sum1_h.font, sum1_h.border, sum1_h.fill, sum1_h.number_format = XlFont(bold=True), THIN_BORDER, YELLOW, MONEY_FMT
    for col in (4, 5, 6):
        c = sheet.cell(row=r_sum1, column=col)
        c.border, c.fill = THIN_BORDER, YELLOW

    diff1 = sheet.cell(row=r_diff1, column=8, value=f"=+G{r_sum1}-H{r_sum1}")
    diff1.number_format = MONEY_FMT

    # --- Título 2 ---
    title2 = sheet.cell(row=r_title2, column=4, value="LIQUIDACION CIERRE RECAUDACION COMISIONES")
    title2.font = XlFont(bold=True, underline="single")

    box_row(r_pagarcom, "Lottery a Pagar", fill_by_col={4: GREEN}, values={7: f"=+H{r_comonline}+H{r_comskoff}+H{r_chase}"})
    box_row(r_pagaronline, "Lottery a Pagar-Online", values={7: f"=-F{r_total1}"})
    box_row(r_pagarskoff, "Lottery a Pagar-skoff", values={7: f"=-P{r_total1}+0+W3"})
    box_row(r_chase, "a", "Chase Bank", values={8: f"=+U{r_total1}"})
    box_row(r_comonline, "a", "Comision On-Line", values={8: f"=-L{r_total1}"})
    box_row(r_comskoff, "a", "Comision Skoff", values={8: f"=-R{r_total1}"})
    box_row(r_caja, "a", "Caja", fill_by_col={5: GREEN, 6: GREEN, 8: GREEN}, values={8: f"=-F{r_total1}-P{r_total1}"})
    box_row(r_blank2)

    sum2_g = sheet.cell(row=r_sum2, column=7, value=f"=SUM(G{r_pagarcom}:G{r_blank2})")
    sum2_g.font, sum2_g.fill, sum2_g.border, sum2_g.number_format = XlFont(bold=True), YELLOW, THIN_BORDER, MONEY_FMT
    sum2_h = sheet.cell(row=r_sum2, column=8, value=f"=SUM(H{r_pagarcom}:H{r_blank2})")
    sum2_h.font, sum2_h.fill, sum2_h.border, sum2_h.number_format = XlFont(bold=True), YELLOW, THIN_BORDER, MONEY_FMT
    for col in (4, 5, 6):
        c = sheet.cell(row=r_sum2, column=col)
        c.fill, c.border = YELLOW, THIN_BORDER

    diff2 = sheet.cell(row=r_diff2, column=8, value=f"=+G{r_sum2}-H{r_sum2}")
    diff2.number_format = MONEY_FMT

    # W3 -- celda auxiliar que usa la fórmula de "Lottery a Pagar-skoff"
    # de arriba (siempre 0 en el Excel real, pero la fórmula la referencia
    # de todos modos -- se replica tal cual para no romper esa fórmula).
    sheet.cell(row=3, column=23, value=0)


def _fmt_money_pdf(value):
    return "—" if value is None else "{:,.2f}".format(value)


def _fmt_int_pdf(value):
    return "—" if value is None else "{:,.0f}".format(value)


def build_lottery_export_pdf(year, month, dest_path):
    """
    Versión PDF (con los mismos colores del export a Excel, sin fórmulas)
    -- mismas columnas que ya muestra /carga-datos/lottery/historial, con
    los bloques de 7 días + Subtotal + Debito uno abajo del otro.
    """
    from pdf_export import build_simple_table_pdf

    GRAY, YELLOW, ORANGE, GREEN = "#D9D9D9", "#FFFF00", "#FFC000", "#92D050"
    headers = [
        "Día", "Count\n(Online)", "Sales $\n(Online)", "Sales\n(Terminal)", "Pagos\n(Terminal)",
        "Cash Bal.", "Comis", "Prize FP", "Total Comm",
        "Count\n(Skoff)", "Sales\n(Skoff)", "Pays U", "Pays $", "Sales\nAmt", "Sales\nComm", "Net\nTotal",
        "Chase Bank", "Total Pagos",
    ]
    header_fill_by_col = {
        1: GRAY, 2: GRAY, 3: GRAY, 4: GRAY, 5: GRAY, 8: GRAY,
        9: YELLOW, 10: YELLOW, 11: YELLOW, 12: YELLOW, 13: YELLOW, 14: YELLOW, 15: YELLOW,
        16: GREEN,
    }
    data_fill_by_col = {15: YELLOW, 16: GREEN}
    # Anchos escalados ~1.2x sobre los originales -- pedido explícito del
    # usuario (2026-09-16): "se ve medio apretado" -- usan casi todo el
    # ancho disponible de la página (legal apaisado, 335.6mm entre
    # márgenes) en vez de dejar ~57mm libres sin usar.
    col_widths_mm = [16.8, 16.8, 21.6, 19.2, 16.8, 19.2, 16.8, 16.8, 19.2, 16.8, 21.6, 14.4, 18.0, 18.0, 16.8, 18.0, 28.8, 19.2]

    rows = []
    for block in build_month_blocks(year, month):
        for day in block["days"]:
            d = datetime.strptime(day["date"], "%Y-%m-%d")
            rows.append([
                d.strftime("%d-%m"),
                _fmt_int_pdf(day.get("online_count")), _fmt_money_pdf(day.get("online_net_sales")),
                _fmt_money_pdf(day.get("sales")), _fmt_money_pdf(day.get("pagos")), _fmt_money_pdf(day.get("cash_balance")),
                _fmt_money_pdf(day.get("comis")), _fmt_money_pdf(day.get("prize_free_plays")), _fmt_money_pdf(day.get("total_comm")),
                _fmt_int_pdf(day.get("skoff_count")), _fmt_money_pdf(day.get("skoff_net_sales")),
                _fmt_int_pdf(day.get("pays_units")), _fmt_money_pdf(day.get("pays_amount")), _fmt_money_pdf(day.get("skoff_sales_amount")),
                _fmt_money_pdf(day.get("sales_comm")), _fmt_money_pdf(day.get("net_total")),
                "", _fmt_money_pdf(day.get("cuenta_final")),
            ])
        sub, deb = block["subtotal"], block["debito"]
        rows.append([
            "Subtotal",
            _fmt_int_pdf(sub.get("online_count")), _fmt_money_pdf(sub.get("online_net_sales")),
            _fmt_money_pdf(sub.get("sales")), _fmt_money_pdf(sub.get("pagos")), "",
            _fmt_money_pdf(sub.get("comis")), _fmt_money_pdf(sub.get("prize_free_plays")), _fmt_money_pdf(sub.get("total_comm")),
            "", "",
            _fmt_int_pdf(sub.get("pays_units")), _fmt_money_pdf(sub.get("pays_amount")), _fmt_money_pdf(sub.get("skoff_sales_amount")),
            _fmt_money_pdf(sub.get("sales_comm")), "",
            "", "",
        ])
        chase_text = ""
        if block.get("chase_bank_date"):
            d = _parse_date(block["chase_bank_date"])
            chase_text = f"${_fmt_money_pdf(deb.get('net_debit'))}\n{d.strftime('%d/%m/%Y')}"
        else:
            chase_text = f"${_fmt_money_pdf(deb.get('net_debit'))}\n(sin confirmar)"
        rows.append([
            "Debito",
            "", _fmt_money_pdf(deb.get("online_net_sales")),
            _fmt_money_pdf(deb.get("sales")), "", "",
            "", "", "",
            "", "", "", _fmt_money_pdf(deb.get("pays_amount")), "",
            "", "",
            chase_text, "",
        ])

    title = f"Lottery — {month:02d}/{year}"
    build_simple_table_pdf(
        dest_path, title, headers, rows,
        col_widths_mm=col_widths_mm,
        header_fill_by_col=header_fill_by_col,
        data_fill_by_col=data_fill_by_col,
        font_size=9.5,
        cell_padding=6,
    )
    return dest_path


# --- Cierre mensual: "LIQUIDACION CIERRE LOTTERY" / "LIQUIDACION CIERRE
# RECAUDACION COMISIONES" (pedido explícito del usuario, 2026-09-16) -------
#
# Estas dos tablas viven al pie de la hoja del Excel de Lottery real -- un
# asiento contable que, mirando sus fórmulas reales (confirmado contra los
# Excel de Julio/Agosto/Septiembre 2026 con openpyxl, columna por columna,
# antes de escribir código), resultó ser ENTERAMENTE derivable de datos que
# esta página ya tiene guardados -- nunca hace falta subir ningún Excel
# nuevo (primer intento de esta sesión, descartado a pedido explícito del
# usuario: "lo ideal es no tener que subir ningún excel").
#
# Mapeo real de cada celda (fila E55:I71 del Excel de septiembre-2026, el
# mismo layout en los 3 meses reales revisados):
#   Tabla 1 "LIQUIDACION CIERRE LOTTERY" (Debe=H, Haber=I):
#     Caja-On Line     (Debe)  = +F50+0            -> SUM mensual de "sales" (F, ONLINE/SALES)
#     Caja-Skoff       (Debe)  = +O49+4774.2        -> SUM mensual de "skoff_net_sales" (O) + una constante fija
#     Gastos Adminits-Loteria (Debe) = 150 (tipeado a mano en el Excel -- ACÁ es el único campo editable)
#     a Lottery a Pagar (Haber) = +F50+0            -> = Caja-On Line (mismo monto, asiento espejo)
#     a Lottery a Pagar (Haber) = +O49+H57+4774.2   -> = Caja-Skoff + Gastos Adminits-Loteria
#   Tabla 2 "LIQUIDACION CIERRE RECAUDACION COMISIONES" (Debe=H, Haber=I):
#     Lottery a Pagar         (Debe)  = +I67+I68+I66 -> suma de las 3 filas "Haber" de abajo (Chase Bank+Comision On-Line+Comision Skoff)
#     Lottery a Pagar-Online  (Debe)  = -G49          -> -SUM mensual de "pagos" (G, ONLINE/Pagos)
#     Lottery a Pagar-skoff   (Debe)  = -Q49+0+X3(=0) -> -SUM mensual de "pays_amount" (Q, SKOFF/Pagos)
#     a Chase Bank    (Haber) = +V49  -> SUM del "Debito"/net_debit (V) de cada bloque semanal del mes
#     a Comision On-Line (Haber) = -M49 -> -SUM mensual de "total_comm" (M=I+K, ONLINE/TOTAL SALES COM)
#     a Comision Skoff   (Haber) = -S49 -> -SUM mensual de "sales_comm" (S, SKOFF/Sales Com)
#     a Caja             (Haber) = -G49-Q49 -> -(Pagos ONLINE) - (Pagos SKOFF)
#
# La cifra que en el Excel real sumaba Caja-Skoff (4774.2 en Agosto Y
# Septiembre 2026, sin fórmula en los dos) pasó por dos vueltas: primero se
# replicó como constante fija, después -- a pedido del usuario ("no es
# fija") -- se volvió un campo editable por mes ("Ajuste Caja-Skoff"). El
# usuario terminó pidiendo sacar ese cuadrito de la página del todo -- así
# que Caja-Skoff quedó SIN ningún ajuste, solo la suma real de
# `skoff_net_sales` -- ver el pendiente en CLAUDE.md si en algún momento
# hiciera falta reincorporar esa cifra de otra forma.
#
# "Gastos Adminits-Loteria" es fijo -- confirmado explícitamente por el
# usuario ("los 150 sí son fijos y no hace falta editar") -- una constante
# de código, sin ningún campo editable ni fila en la base.
#
# Todo lo demás se recalcula EN VIVO cada vez que se entra a la página,
# sumando los mismos días que ya muestra "Cuadro del mes" para ese mes
# (build_month_blocks) -- incluida la cola de días que se pasa al mes
# vecino, igual que hace el Excel real. Si el mes todavía no tiene ningún
# PDF cargado, todo sale en 0 y la página lo muestra vacío (has_data=False).

_GASTOS_ADMIN_FIJO = 150.0


def compute_month_closing(year, month):
    """
    Arma en el momento las dos tablas de cierre de mes -- ver el comentario
    de arriba para el mapeo completo. Nunca lee ni escribe ningún Excel.
    """
    blocks = build_month_blocks(year, month)
    days = [d for block in blocks for d in block["days"]]
    has_data = any(_has_any_data(d) for d in days)

    def total(field):
        values = [d.get(field) for d in days if d.get(field) is not None]
        return round(sum(values), 2) if values else 0.0

    total_sales = total("sales")
    total_pagos_online = total("pagos")
    total_total_comm = total("total_comm")
    total_skoff_net_sales = total("skoff_net_sales")
    total_pagos_skoff = total("pays_amount")
    total_sales_comm = total("sales_comm")
    total_net_debit = round(sum((b["debito"].get("net_debit") or 0.0) for b in blocks), 2)

    gastos_admin = _GASTOS_ADMIN_FIJO
    days_with_data = sum(1 for d in days if _has_any_data(d))

    caja_online = total_sales
    caja_skoff = total_skoff_net_sales
    lottery_a_pagar_1 = caja_online
    lottery_a_pagar_2 = round(caja_skoff + gastos_admin, 2)

    com_chase_bank = total_net_debit
    com_comision_online = round(-total_total_comm, 2)
    com_comision_skoff = round(-total_sales_comm, 2)
    com_caja = round(-total_pagos_online - total_pagos_skoff, 2)
    com_lottery_a_pagar_online = round(-total_pagos_online, 2)
    com_lottery_a_pagar_skoff = round(-total_pagos_skoff, 2)
    com_lottery_a_pagar = round(com_chase_bank + com_comision_online + com_comision_skoff, 2)

    result = {
        "has_data": has_data,
        "days_with_data": days_with_data,
        "gastos_admin": gastos_admin,
        "caja_online": caja_online,
        "caja_skoff": caja_skoff,
        "lottery_a_pagar_1": lottery_a_pagar_1,
        "lottery_a_pagar_2": lottery_a_pagar_2,
        "com_lottery_a_pagar": com_lottery_a_pagar,
        "com_lottery_a_pagar_online": com_lottery_a_pagar_online,
        "com_lottery_a_pagar_skoff": com_lottery_a_pagar_skoff,
        "com_chase_bank": com_chase_bank,
        "com_comision_online": com_comision_online,
        "com_comision_skoff": com_comision_skoff,
        "com_caja": com_caja,
        # Totales crudos, expuestos solo para que la página pueda mostrar
        # "de dónde sale" cada celda (mismo criterio que los popovers de
        # verificación de Store Info/Lottery historial) -- no forman parte
        # del asiento en sí.
        "total_sales": total_sales,
        "total_pagos_online": total_pagos_online,
        "total_total_comm": total_total_comm,
        "total_skoff_net_sales": total_skoff_net_sales,
        "total_pagos_skoff": total_pagos_skoff,
        "total_sales_comm": total_sales_comm,
        "total_net_debit": total_net_debit,
    }
    result["debe_total"] = round(caja_online + caja_skoff + gastos_admin, 2)
    result["haber_total"] = round(lottery_a_pagar_1 + lottery_a_pagar_2, 2)
    result["diferencia"] = round(result["debe_total"] - result["haber_total"], 2)
    result["com_debe_total"] = round(
        com_lottery_a_pagar + com_lottery_a_pagar_online + com_lottery_a_pagar_skoff, 2
    )
    result["com_haber_total"] = round(
        com_chase_bank + com_comision_online + com_comision_skoff + com_caja, 2
    )
    result["com_diferencia"] = round(result["com_debe_total"] - result["com_haber_total"], 2)
    return result
