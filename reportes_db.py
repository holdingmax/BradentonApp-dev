"""
Persistencia de los Reportes Diarios guardados dentro de la propia página
(ver CLAUDE.md -> "Guardar y ver los Reportes Diarios dentro de la
página"). Capa de datos pura, sin ningún conocimiento de PDFs ni de Excel
-- reporte_diario.py hace la extracción, este módulo solo guarda/consulta.

Todo vive en reportes_data/ (gitignored, mismo criterio que users.json):
reportes_data/reportes_diarios.db es el archivo SQLite;
reportes_data/pdfs/{año}/{mes}/ guarda una copia de cada PDF ya subido.
"""

import calendar
import json
import os
import shutil
import sqlite3
from datetime import date, datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "reportes_diarios.db")
_PDF_DIR = os.path.join(_BASE_DIR, "pdfs")

_STORE_INFO_NUMERIC_FIELDS = (
    "volume", "sales_fuel", "desc_comb", "non_fuel_total", "desc_otros",
    "tax_collect", "total_sales", "cash", "local_accounts", "other_amount",
    "network_revenue", "total_revenue",
)


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    # timeout=30 + WAL -- antes de las cargas en segundo plano (ver jobs.py,
    # CLAUDE.md "sexta tanda") el servidor de desarrollo atendía una sola
    # request a la vez, así que dos escrituras a la misma base nunca podían
    # solaparse de verdad. Con threaded=True y trabajos corriendo en hilos
    # propios, sí puede pasar (ej. Ventas y Store Info del mismo PDF,
    # guardándose casi al mismo tiempo) -- sin esto, la segunda escritura
    # podía fallar con "database is locked" tras los 5s de timeout default
    # de sqlite3. WAL deja lectores y un escritor conviviendo sin bloquearse
    # entre sí; timeout=30 le da margen de sobra a la que sí tiene que
    # esperar (una carga real nunca escribe más de una vez por segundo).
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    _ensure_schema(conn)
    return conn


def _ensure_schema(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS daily_reports (
            date TEXT PRIMARY KEY,
            from_time TEXT, to_time TEXT,
            volume REAL, sales_fuel REAL, desc_comb REAL,
            non_fuel_total REAL, desc_otros REAL, tax_collect REAL,
            cash REAL, credit_terms_json TEXT,
            local_accounts REAL, network_revenue REAL, total_revenue REAL,
            pdf_filename TEXT,
            store_info_source TEXT,
            printed_total_sales REAL,
            printed_total_units REAL,
            updated_at TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS daily_report_departments (
            date TEXT NOT NULL,
            department TEXT NOT NULL,
            count INTEGER,
            amount REAL,
            source TEXT,
            updated_at TEXT,
            PRIMARY KEY (date, department)
        )
        """
    )
    # Columnas agregadas después de la primera versión de esta tabla -- una
    # base ya existente en disco no las gana solas con CREATE TABLE IF NOT
    # EXISTS, hace falta este mini-migrador (ver "printed_total_sales/units"
    # en CLAUDE.md). Idempotente: en una base nueva ya vienen del CREATE de
    # arriba, así que no hace nada acá.
    _ensure_columns(conn, "daily_reports", {"printed_total_sales": "REAL", "printed_total_units": "REAL"})
    # "Other" del Method of Payment Totals (columna U de Store Info) y
    # "Total Sales" (impreso tal cual por el reporte, columna R) --
    # agregadas 2026-09-12, ver CLAUDE.md.
    _ensure_columns(conn, "daily_reports", {"other_amount": "REAL", "total_sales": "REAL"})
    # Monto del departamento "LOCAL ACCT" del Department Sales Report (2026-
    # 09-15, aclaración del usuario) -- se guarda acá, aparte de
    # daily_report_departments, porque ese departamento NUNCA debe
    # aparecer como una fila más (ver reporte_diario.extract_department_
    # sales_for_day) -- este campo alimenta SOLO la categoría "Gettel" de
    # Ventas por Departamento (ver group_department_sales/get_day_local_
    # acct_amount/get_month_local_acct_amount).
    _ensure_columns(conn, "daily_reports", {"local_acct_amount": "REAL"})
    conn.commit()


def _ensure_columns(conn, table, columns):
    existing = {row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    for name, col_type in columns.items():
        if name not in existing:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {col_type}")


def _date_key(value):
    """Acepta date/datetime o un string 'YYYY-MM-DD' ya armado."""
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    return str(value)


def _to_float(value):
    if value is None or value == "":
        return None
    return float(value)


def _time_str(value):
    if value is None or value == "":
        return None
    if hasattr(value, "strftime"):
        return value.strftime("%H:%M")
    return str(value)


def _now():
    return datetime.now().isoformat(timespec="seconds")


def replace_department_sales(report_date, records, pdf_filename=None, local_acct_amount=None):
    """
    Reemplaza TODO el día de golpe con lo recién leído del PDF (source=
    "ocr") -- mismo criterio que ya usa hoy la escritura del Excel
    (reprocesar un día reescribe sus celdas sin importar qué hubiera antes,
    sea una carga vieja o una corrección manual). `records` es la lista tal
    cual la devuelve reporte_diario.extract_department_sales_for_day:
    [{"department", "count", "amount"}, ...].

    `local_acct_amount` (mismo `extract_department_sales_for_day`, aparte
    de `records`) reemplaza también el campo del mismo nombre en
    daily_reports -- `None` si el PDF no trajo "LOCAL ACCT" ese día, para
    que reprocesar un día que ya no lo tiene lo borre en vez de dejar un
    valor viejo pegado.
    """
    key = _date_key(report_date)
    now = _now()
    conn = _connect()
    try:
        conn.execute("DELETE FROM daily_report_departments WHERE date = ?", (key,))
        for record in records:
            conn.execute(
                """
                INSERT INTO daily_report_departments (date, department, count, amount, source, updated_at)
                VALUES (?, ?, ?, ?, 'ocr', ?)
                """,
                (key, record["department"], int(record["count"]), float(record["amount"]), now),
            )
        conn.execute(
            """
            INSERT INTO daily_reports (date, pdf_filename, local_acct_amount, updated_at) VALUES (?, ?, ?, ?)
            ON CONFLICT(date) DO UPDATE SET
                pdf_filename = COALESCE(excluded.pdf_filename, daily_reports.pdf_filename),
                local_acct_amount = excluded.local_acct_amount,
                updated_at = excluded.updated_at
            """,
            (key, pdf_filename, _to_float(local_acct_amount), now),
        )
        conn.commit()
    finally:
        conn.close()


def upsert_department_row(report_date, department, count, amount, source="manual"):
    """Corrige o agrega UN departamento puntual, sin tocar el resto del día."""
    key = _date_key(report_date)
    now = _now()
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO daily_report_departments (date, department, count, amount, source, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(date, department) DO UPDATE SET
                count = excluded.count, amount = excluded.amount,
                source = excluded.source, updated_at = excluded.updated_at
            """,
            (key, department.strip(), int(count), float(amount), source, now),
        )
        conn.execute(
            """
            INSERT INTO daily_reports (date, updated_at) VALUES (?, ?)
            ON CONFLICT(date) DO UPDATE SET updated_at = excluded.updated_at
            """,
            (key, now),
        )
        conn.commit()
    finally:
        conn.close()


def delete_department_row(report_date, department):
    key = _date_key(report_date)
    conn = _connect()
    try:
        conn.execute(
            "DELETE FROM daily_report_departments WHERE date = ? AND department = ?",
            (key, department),
        )
        conn.commit()
    finally:
        conn.close()


def delete_departments_for_date(report_date):
    """
    Borra TODOS los departamentos de un día (para poder recargarlo desde
    cero -- pedido explícito del usuario 2026-09-15, checkbox de eliminar en
    Ventas por Departamento). No toca Store Info de ese día -- son dos
    extracciones independientes, mismo criterio que el resto del módulo.
    """
    key = _date_key(report_date)
    conn = _connect()
    try:
        cur = conn.execute("DELETE FROM daily_report_departments WHERE date = ?", (key,))
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def upsert_store_info(report_date, fields, source, pdf_filename=None):
    """
    Reemplaza la fila completa de Store Info de ese día -- un solo
    formulario con todos los campos juntos, tanto para lo leído del PDF
    como para una corrección manual. `fields` acepta las mismas claves que
    devuelve reporte_diario.extract_store_info_from_pdf (from_time/to_time
    como datetime.time o como texto "HH:MM"; el resto números;
    "credit_terms" una lista de montos).
    """
    key = _date_key(report_date)
    now = _now()

    row = {"date": key}
    row["from_time"] = _time_str(fields.get("from_time"))
    row["to_time"] = _time_str(fields.get("to_time"))
    for field in _STORE_INFO_NUMERIC_FIELDS:
        row[field] = _to_float(fields.get(field))
    row["credit_terms_json"] = json.dumps([_to_float(v) for v in (fields.get("credit_terms") or [])])
    row["store_info_source"] = source
    row["updated_at"] = now
    if pdf_filename:
        row["pdf_filename"] = pdf_filename

    columns = list(row.keys())
    placeholders = ", ".join("?" for _ in columns)
    update_clause = ", ".join(f"{col} = excluded.{col}" for col in columns if col != "date")
    conn = _connect()
    try:
        conn.execute(
            f"INSERT INTO daily_reports ({', '.join(columns)}) VALUES ({placeholders}) "
            f"ON CONFLICT(date) DO UPDATE SET {update_clause}",
            [row[col] for col in columns],
        )
        conn.commit()
    finally:
        conn.close()


def get_month_overview(year, month):
    """
    Un resumen por cada día del mes (haya datos o no, para que salte a la
    vista qué falta) -- usado por /reporte/historial. Cada día trae además
    su propio desglose de departamentos (nombre + monto, de mayor a menor)
    para un panel al pasar el mouse -- pedido explícito del usuario
    (2026-09-12): la fila resumida ("N depto.") mostraba muy poco, y solo
    apretando "Ver/editar" se veía el detalle real.
    """
    days_in_month = calendar.monthrange(year, month)[1]
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        dept_rows = conn.execute(
            """
            SELECT date, COUNT(*) AS n, SUM(amount) AS total, GROUP_CONCAT(DISTINCT source) AS sources
            FROM daily_report_departments
            WHERE date LIKE ?
            GROUP BY date
            """,
            (f"{prefix}%",),
        ).fetchall()
        dept_by_date = {row["date"]: row for row in dept_rows}

        detail_rows = conn.execute(
            "SELECT date, department, count, amount FROM daily_report_departments WHERE date LIKE ? ORDER BY amount DESC",
            (f"{prefix}%",),
        ).fetchall()
        detail_by_date = {}
        for row in detail_rows:
            detail_by_date.setdefault(row["date"], []).append(
                {"department": row["department"], "count": row["count"], "amount": row["amount"]}
            )

        info_rows = conn.execute(
            "SELECT date, store_info_source FROM daily_reports WHERE date LIKE ? AND store_info_source IS NOT NULL",
            (f"{prefix}%",),
        ).fetchall()
        info_by_date = {row["date"]: row["store_info_source"] for row in info_rows}
    finally:
        conn.close()

    overview = []
    for day in range(1, days_in_month + 1):
        key = f"{prefix}{day:02d}"
        dept_row = dept_by_date.get(key)
        overview.append(
            {
                "date": key,
                "day": day,
                "department_count": dept_row["n"] if dept_row else 0,
                "department_total": dept_row["total"] if dept_row else None,
                "department_sources": (dept_row["sources"].split(",") if dept_row and dept_row["sources"] else []),
                "department_detail": detail_by_date.get(key, []),
                "store_info_source": info_by_date.get(key),
            }
        )
    return overview


def get_month_department_totals(year, month):
    """
    Un renglón por departamento con la suma de Cantidad/Monto de TODO el
    mes -- pedido explícito del usuario (2026-09-13): a diferencia del
    Excel (que tiene la autosuma de cada departamento al pie de "CARGA
    AQUI"), acá solo se podía ver un día a la vez. De mayor a menor monto
    (el departamento que más vendió primero).
    """
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            """
            SELECT department, SUM(count) AS count, SUM(amount) AS amount
            FROM daily_report_departments
            WHERE date LIKE ?
            GROUP BY department
            ORDER BY amount DESC
            """,
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()
    return [{"department": row["department"], "count": row["count"], "amount": row["amount"]} for row in rows]


def get_day(report_date):
    """{"date", "departments": [...], "store_info": {...} | None} para /reporte/dia/<date>."""
    key = _date_key(report_date)
    conn = _connect()
    try:
        dept_rows = conn.execute(
            "SELECT department, count, amount, source FROM daily_report_departments "
            "WHERE date = ? ORDER BY department",
            (key,),
        ).fetchall()
        info_row = conn.execute("SELECT * FROM daily_reports WHERE date = ?", (key,)).fetchone()
    finally:
        conn.close()

    departments = [dict(row) for row in dept_rows]
    store_info = None
    if info_row is not None and info_row["store_info_source"] is not None:
        store_info = dict(info_row)
        store_info["credit_terms"] = json.loads(store_info.get("credit_terms_json") or "[]")
    pdf_filename = info_row["pdf_filename"] if info_row is not None else None
    # Los totales impresos (ver upsert_printed_totals) no dependen de si ya
    # hay Store Info cargado -- son un chequeo aparte contra la suma de
    # Departamentos, así que viajan sueltos y no solo dentro de store_info.
    printed_total_sales = info_row["printed_total_sales"] if info_row is not None else None
    printed_total_units = info_row["printed_total_units"] if info_row is not None else None
    return {
        "date": key,
        "departments": departments,
        "store_info": store_info,
        "pdf_filename": pdf_filename,
        "printed_total_sales": printed_total_sales,
        "printed_total_units": printed_total_units,
    }


def upsert_printed_totals(report_date, total_sales, total_units):
    """
    Guarda el monto total y las unidades totales tal cual figuran IMPRESOS
    en el reporte de ventas -- lo tipea el usuario a mano para poder
    comparar contra la suma de lo ya cargado en Departamentos y detectar si
    algo quedó mal leído/cargado (mismo espíritu que la columna BS del
    Excel real, ver CLAUDE.md). `None` en cualquiera de los dos borra ese
    chequeo puntual sin tocar el otro.
    """
    key = _date_key(report_date)
    now = _now()
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO daily_reports (date, printed_total_sales, printed_total_units, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(date) DO UPDATE SET
                printed_total_sales = excluded.printed_total_sales,
                printed_total_units = excluded.printed_total_units,
                updated_at = excluded.updated_at
            """,
            (key, _to_float(total_sales), _to_float(total_units), now),
        )
        conn.commit()
    finally:
        conn.close()


def get_month_store_info(year, month):
    """
    Un renglón por CADA día del mes, haya datos o no -- para el reporte
    mensual tipo Excel de /reporte/historial. Pedido explícito del usuario
    (2026-09-14): antes solo se devolvían los días ya cargados, quedando
    todos pegados unos con otros sin ningún espacio -- mareaba no poder ver
    de un vistazo qué días faltan completar. Un día sin nada cargado
    todavía se devuelve con todos los campos en None (la plantilla ya
    muestra "—" para eso) pero con su fecha real, para poder marearlo
    visualmente como un salto Y para poder linkear igual a "Editar".
    """
    days_in_month = calendar.monthrange(year, month)[1]
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM daily_reports WHERE date LIKE ? AND store_info_source IS NOT NULL ORDER BY date",
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()

    by_date = {}
    for row in rows:
        item = dict(row)
        item["credit_terms"] = json.loads(item.get("credit_terms_json") or "[]")
        by_date[item["date"]] = item

    blank_fields = (
        "from_time", "to_time", "volume", "sales_fuel", "desc_comb", "non_fuel_total",
        "desc_otros", "tax_collect", "total_sales", "cash", "local_accounts",
        "other_amount", "network_revenue", "total_revenue",
    )
    result = []
    for day in range(1, days_in_month + 1):
        day_key = f"{year:04d}-{month:02d}-{day:02d}"
        if day_key in by_date:
            result.append(by_date[day_key])
        else:
            blank = {field: None for field in blank_fields}
            blank["date"] = day_key
            blank["credit_terms"] = []
            result.append(blank)
    return result


def get_month_department_amounts(year, month, department):
    """
    {"YYYY-MM-DD": amount} de UN departamento puntual, para todo el mes --
    usado por Gettel/Toyota para cruzar "LOCAL ACCT" (ver CLAUDE.md,
    "Gettel/Toyota -- fila de Local Account y DIF"). No todos los días
    tienen este departamento -- solo aparecen los que sí.
    """
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT date, amount FROM daily_report_departments WHERE date LIKE ? AND department = ?",
            (f"{prefix}%", department),
        ).fetchall()
    finally:
        conn.close()
    return {row["date"]: row["amount"] for row in rows}


def get_day_local_acct_amount(report_date):
    """
    Monto de "LOCAL ACCT" (departamento del Department Sales Report,
    guardado aparte -- ver replace_department_sales) para un solo día --
    alimenta ÚNICAMENTE la categoría "Gettel" del resumen por categoría de
    ese día (reporte_diario.group_department_sales). 0.0 si el día no tiene
    nada guardado (nunca None, mismo criterio que get_day_gettel_amount de
    gettel_db, que este campo reemplaza para este propósito puntual).
    """
    key = _date_key(report_date)
    conn = _connect()
    try:
        row = conn.execute("SELECT local_acct_amount FROM daily_reports WHERE date = ?", (key,)).fetchone()
    finally:
        conn.close()
    return (row["local_acct_amount"] if row else None) or 0.0


def get_month_local_acct_amount(year, month):
    """Suma de "LOCAL ACCT" del mes -- alimenta la categoría "Gettel" del resumen mensual de Ventas por Departamento."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT SUM(local_acct_amount) AS total FROM daily_reports WHERE date LIKE ?",
            (f"{prefix}%",),
        ).fetchone()
    finally:
        conn.close()
    return row["total"] or 0.0


def set_local_acct_amount(report_date, amount):
    """
    Corrige/backfillea SOLO este campo de un día, sin tocar nada más --
    usado para reprocesar días ya guardados antes de que este campo
    existiera (ver replace_department_sales para el camino normal, que lo
    reemplaza junto con el resto del día).
    """
    key = _date_key(report_date)
    now = _now()
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO daily_reports (date, local_acct_amount, updated_at) VALUES (?, ?, ?)
            ON CONFLICT(date) DO UPDATE SET
                local_acct_amount = excluded.local_acct_amount,
                updated_at = excluded.updated_at
            """,
            (key, _to_float(amount), now),
        )
        conn.commit()
    finally:
        conn.close()


def get_month_local_accounts(year, month):
    """
    {"YYYY-MM-DD": local_accounts} de Store Info para todo el mes -- usado
    por Gettel/Toyota para cruzar contra la columna real "Local Account"
    del Excel Gettel-Toyota (ver CLAUDE.md, corrección 2026-09-15). Este es
    el campo chico de Store Info (sección Method of Payment Totals), NO el
    departamento "LOCAL ACCT" del Department Sales Report -- confirmado
    contra el Excel real (Gettel-Toyota 08.2026, columna "Local Account")
    que sus valores coinciden EXACTOS con este campo, día por día, mientras
    que el departamento "LOCAL ACCT" del PDF solo aparece esporádicamente
    (cargos puntuales grandes, no una cifra diaria) y nunca coincide.
    """
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT date, local_accounts FROM daily_reports WHERE date LIKE ? AND local_accounts IS NOT NULL",
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()
    return {row["date"]: row["local_accounts"] for row in rows}


def list_known_departments():
    """Nombres de departamento ya vistos alguna vez -- para sugerir en la carga manual (datalist)."""
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT DISTINCT department FROM daily_report_departments ORDER BY department"
        ).fetchall()
    finally:
        conn.close()
    return [row["department"] for row in rows]


def store_pdf_copy(report_date, source_path, original_filename):
    """
    Copia el PDF ya subido (nunca el original del usuario -- source_path ya
    es una copia en el workspace temporal de la request) a
    reportes_data/pdfs/{año}/{mes}/. Devuelve la ruta relativa a
    reportes_data/, para guardar en daily_reports.pdf_filename.
    """
    if isinstance(report_date, (date, datetime)):
        d = report_date
    else:
        d = datetime.strptime(str(report_date), "%Y-%m-%d").date()
    dest_dir = os.path.join(_PDF_DIR, f"{d.year:04d}", f"{d.month:02d}")
    os.makedirs(dest_dir, exist_ok=True)
    dest_path = os.path.join(dest_dir, original_filename)
    shutil.copyfile(source_path, dest_path)
    return os.path.relpath(dest_path, _BASE_DIR)


def absolute_pdf_path(relative_path):
    return os.path.join(_BASE_DIR, relative_path)


def get_month_pdf_list(year, month):
    """Un renglón por cada PDF de cierre diario guardado este mes -- para la lista "Documentos" de la barra lateral."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT date, pdf_filename FROM daily_reports WHERE date LIKE ? AND pdf_filename IS NOT NULL ORDER BY date ASC",
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()
    return [{"date": row["date"], "pdf_filename": row["pdf_filename"]} for row in rows]


def search_pdfs(query, limit=20):
    """
    PDFs de cierre diario cuyo nombre de archivo contiene `query` -- pedido
    explícito del usuario (2026-09-16), buscador del header por nombre de
    PDF. `pdf_filename` es la ruta relativa completa (incluye año/mes), así
    que el LIKE matchea igual contra el nombre real del archivo.
    """
    like = f"%{query}%"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT date, pdf_filename FROM daily_reports WHERE pdf_filename LIKE ? "
            "ORDER BY date DESC LIMIT ?",
            (like, limit),
        ).fetchall()
    finally:
        conn.close()
    return [{"date": row["date"], "pdf_filename": row["pdf_filename"]} for row in rows]
