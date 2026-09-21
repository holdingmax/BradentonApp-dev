"""
Gettel / Toyota -- Carga de Datos: guarda los totales diarios de cupones de
combustible de las dos empresas (Monto + Galones), sin escribir ningún
Excel -- mismo patrón que chase_db.py/lottery_db.py.

Los datos ya se leen con funciones puras existentes (`gettel_toyota_parser.
_summarize_origin_workbook`/`summarize_pdf_report`), pensadas originalmente
para fusionarse en la hoja "Gettel-Toyota MM.YYYY" del Excel Cierre -- acá
se guardan directo en base, un renglón por día con las 4 columnas (Gettel
Monto/Galones, Toyota Monto/Galones) que esa hoja ya tenía.

Pedido explícito del usuario (2026-09-12, cuarta tanda): "deberias
considerarlas en ponerlos en el cuadro de ventas" -- el monto de Gettel de
acá es lo que alimenta la categoría "Gettel" de Ventas por Departamento
(ver reporte_diario.group_department_sales, parámetro gettel_amount), que
antes siempre quedaba en $0 porque esa venta nunca sale del PDF de cierre
diario -- viene de estos reportes de cupones.
"""

import os
import sqlite3
from datetime import date, datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "gettel_toyota.db")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    # timeout=30 + WAL -- ver reportes_db.py: necesario desde que las cargas
    # en segundo plano (jobs.py) pueden escribir de verdad en paralelo.
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS gettel_toyota_days (
            date TEXT PRIMARY KEY,
            gettel_amount REAL,
            gettel_gallons REAL,
            gettel_source TEXT,
            toyota_amount REAL,
            toyota_gallons REAL,
            toyota_source TEXT,
            updated_at TEXT
        )
        """
    )
    # -- Gettel: Pagos de Cupones (2026-09-19, pedido explícito del usuario)
    # -- reemplaza la lectura automática de PDF de pagos (gettel_toyota_
    # parser.extract_pago_batch_from_pdf/process_gettel_pagos, la
    # herramienta vieja de /gettel/pagos que escribe directo sobre el
    # Excel Cierre real) porque "no esta leyendo bien los pagos que subi
    # en pdf" -- ahora se carga a mano, un renglón por cupón/transacción
    # (mismo criterio "Guardar en base" ya usado en Combustible/Físico).
    # Un pago (envío/depósito de Gettel) puede traer varias transacciones
    # -- comparten Fecha + Pago N°, cada una con su propio Transc N° y
    # Total Cupón (así lo agrupa visualmente la hoja real, con celdas
    # unificadas de Fecha/Pago N° -- ver gettel_pagos.grouped_pagos).
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS gettel_pagos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            pago_n INTEGER,
            transc_n TEXT,
            total_cupon REAL NOT NULL,
            created_at TEXT,
            updated_at TEXT
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_gettel_pagos_fecha ON gettel_pagos(fecha)")
    # `source` ("pdf"/"manual") y `empresa` (Toyota/Kia, leída del nombre del
    # archivo -- ver gettel_pagos_parser.py) -- agregadas 2026-09-21 cuando la
    # carga pasó de manual a lectura automática de PDF, mismo patrón
    # _ensure_columns que fisico_db.py (ALTER TABLE seguro sobre bases ya
    # existentes, sin romper filas viejas cargadas a mano que quedan con
    # source/empresa en NULL -- se muestran como "Manual" en la UI).
    _ensure_columns(conn, "gettel_pagos", {"source": "TEXT", "empresa": "TEXT"})
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS gettel_pagos_months (
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            pendiente_anterior_override REAL,
            updated_at TEXT,
            PRIMARY KEY (year, month)
        )
        """
    )
    return conn


def _day_key(value):
    if isinstance(value, str):
        return value
    if isinstance(value, (date, datetime)):
        return value.isoformat() if isinstance(value, date) and not isinstance(value, datetime) else value.date().isoformat()
    return str(value)


def _ensure_columns(conn, table, columns):
    """Mismo patrón que fisico_db._ensure_columns/reportes_db._ensure_columns -- agrega columnas nuevas a una tabla ya existente sin romper bases viejas."""
    existing = {row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    for name, col_type in columns.items():
        if name not in existing:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {col_type}")


def upsert_vendor_totals(vendor, totals_by_date, source="excel"):
    """
    `vendor` es "gettel" o "toyota". `totals_by_date` es el dict {date:
    {"amount","gallons"}} que ya devuelven las funciones de extracción de
    gettel_toyota_parser -- se pisa (reemplaza) el día completo de ese
    vendor, nunca se acumula (mismo criterio que reemplazar un día entero
    de Reporte Diario al volver a subir el mismo PDF).
    """
    if vendor not in ("gettel", "toyota"):
        raise ValueError(f"Vendor desconocido: {vendor}")
    if not totals_by_date:
        return 0

    amount_col = f"{vendor}_amount"
    gallons_col = f"{vendor}_gallons"
    source_col = f"{vendor}_source"
    now = datetime.now().isoformat(timespec="seconds")

    conn = _connect()
    try:
        for day, totals in totals_by_date.items():
            key = _day_key(day)
            conn.execute(
                f"""
                INSERT INTO gettel_toyota_days (date, {amount_col}, {gallons_col}, {source_col}, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(date) DO UPDATE SET
                    {amount_col} = excluded.{amount_col},
                    {gallons_col} = excluded.{gallons_col},
                    {source_col} = excluded.{source_col},
                    updated_at = excluded.updated_at
                """,
                (key, float(totals.get("amount") or 0.0), float(totals.get("gallons") or 0.0), source, now),
            )
        conn.commit()
    finally:
        conn.close()
    return len(totals_by_date)


def get_day(report_date):
    key = _day_key(report_date)
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM gettel_toyota_days WHERE date = ?", (key,)).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None


def get_month_days(year, month):
    """Un renglón por día del mes que tenga algo cargado (Gettel y/o Toyota) -- no rellena los días vacíos, a diferencia de reporte_historial (acá no hay "carga a mano" pendiente que mostrar, es puramente lo ya subido)."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM gettel_toyota_days WHERE date LIKE ? ORDER BY date",
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def get_month_gettel_amount(year, month):
    """Suma del Monto de Gettel del mes -- lo que alimenta la categoría "Gettel" del resumen mensual de Ventas por Departamento."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT SUM(gettel_amount) AS total FROM gettel_toyota_days WHERE date LIKE ?",
            (f"{prefix}%",),
        ).fetchone()
    finally:
        conn.close()
    return row["total"] or 0.0


def get_toyota_days_years():
    """Años con algo cargado en gettel_toyota_days -- pedido por gettel_reportes.get_available_years para poblar el selector de año del reporte Gettel de Reportes (2026-09-21)."""
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT DISTINCT substr(date, 1, 4) AS y FROM gettel_toyota_days ORDER BY y"
        ).fetchall()
    finally:
        conn.close()
    return [int(row["y"]) for row in rows if row["y"]]


def get_day_gettel_amount(report_date):
    """Monto de Gettel de un solo día -- lo que alimenta la categoría "Gettel" del resumen por categoría de un día puntual."""
    day = get_day(report_date)
    return (day.get("gettel_amount") if day else None) or 0.0
def add_pago(fecha, pago_n, transc_n, total_cupon, source="manual", empresa=None):
    """
    `source` ("pdf"/"manual") y `empresa` (Toyota/Kia) -- pedido explícito
    del usuario (2026-09-21): la carga pasó de manual a lectura automática
    de PDF (ver gettel_pagos_parser.extract_pagos_from_pdf), mismo criterio
    que fisico_db.add_invoice(source=...). `empresa` es solo informativo --
    no participa de ningún cálculo de gettel_pagos.build_month_report.
    """
    now = datetime.now().isoformat(timespec="seconds")
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO gettel_pagos (fecha, pago_n, transc_n, total_cupon, source, empresa, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                _day_key(fecha), pago_n, (transc_n or "").strip() or None, float(total_cupon),
                source, (empresa or "").strip() or None, now, now,
            ),
        )
        conn.commit()
        return conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]
    finally:
        conn.close()


def update_pago(pago_id, fecha, pago_n, transc_n, total_cupon, empresa=None):
    """
    Corrige a mano una fila ya guardada (leída de PDF, posiblemente con un
    campo mal leído) -- mismo criterio "editar la fila completa" que
    eft_db.update_coupon_row. `source` no se toca -- sigue siendo "pdf",
    solo se corrige el valor, no el origen.
    """
    now = datetime.now().isoformat(timespec="seconds")
    conn = _connect()
    try:
        cur = conn.execute(
            """
            UPDATE gettel_pagos
            SET fecha = ?, pago_n = ?, transc_n = ?, total_cupon = ?, empresa = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                _day_key(fecha), pago_n, (transc_n or "").strip() or None, float(total_cupon),
                (empresa or "").strip() or None, now, pago_id,
            ),
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def find_pago_by_transaction(transc_n):
    """
    Busca un pago ya guardado por N° de Transacción, en CUALQUIER mes --
    para detectar duplicados al subir PDFs (mismo criterio que
    fisico_db.find_invoice_by_number: el N° de Transacción del registro es
    un contador corrido único de esa caja, nunca se repite). None si no hay
    N° para buscar o no se encontró.
    """
    number = (transc_n or "").strip()
    if not number:
        return None
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT * FROM gettel_pagos WHERE transc_n = ? COLLATE NOCASE", (number,)
        ).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None


def delete_pago(pago_id):
    conn = _connect()
    try:
        conn.execute("DELETE FROM gettel_pagos WHERE id = ?", (pago_id,))
        conn.commit()
    finally:
        conn.close()


def get_month_pagos(year, month):
    """Pagos/cupones del mes, en el mismo orden en que los agrupa la hoja real (Fecha, Pago N°, y por último el orden de carga)."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM gettel_pagos WHERE fecha LIKE ? ORDER BY fecha, pago_n, id",
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def get_pago_years():
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT DISTINCT substr(fecha, 1, 4) AS y FROM gettel_pagos ORDER BY y"
        ).fetchall()
        return [int(row["y"]) for row in rows if row["y"]]
    finally:
        conn.close()


def get_pago_month_settings(year, month):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT * FROM gettel_pagos_months WHERE year = ? AND month = ?", (year, month)
        ).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None


def get_pago_settings_years():
    conn = _connect()
    try:
        rows = conn.execute("SELECT DISTINCT year AS y FROM gettel_pagos_months ORDER BY y").fetchall()
        return [int(row["y"]) for row in rows if row["y"] is not None]
    finally:
        conn.close()


def set_month_pendiente_anterior_override(year, month, value):
    """value en None borra el override -- vuelve a encadenarse del Pendiente que pasó del mes anterior."""
    now = datetime.now().isoformat(timespec="seconds")
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO gettel_pagos_months (year, month, pendiente_anterior_override, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(year, month) DO UPDATE SET
                pendiente_anterior_override = excluded.pendiente_anterior_override,
                updated_at = excluded.updated_at
            """,
            (year, month, value, now),
        )
        conn.commit()
    finally:
        conn.close()
