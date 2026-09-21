"""
Módulo "Físico" -- pedido explícito del usuario (2026-09-18, misma sesión
que Proyecciones): reconciliación mensual de combustible (Inventario
Teórico vs. lectura física real de los tanques), decodificada leyendo un
ejemplo real (`hoja_fisico.xlsx`) con openpyxl -- no adivinada. Capa de
datos pura, sin ningún conocimiento de PDF/Excel -- ver fisico.py para el
cálculo del reporte del mes.

Todo vive en reportes_data/fisico.db (gitignored, mismo criterio que el
resto de los módulos de Carga de Datos):
  - `fuel_invoices`: un renglón por factura de compra de combustible,
    cargada a mano o leída de un PDF (ver carga_datos_combustible en
    webapp.py). El usuario confirmó (2026-09-18) que prefería carga
    manual mientras no hubiera un ejemplo real de factura para construir
    un lector automático confiable -- el 2026-09-21 subió 5 facturas
    reales del proveedor de combustible (J.H. Williams Oil Company,
    formato "Qty/Item/Description/.../Unit Price w/o Tax/Unit Price with
    Tax" + "FREIGHT SUMMARY" + "TOTAL INVOICE AMOUNT DUE"), decodificadas
    con pdfplumber (`fisico_invoice_parser.py`) -- ahora la carga
    preferida es subir el PDF y que el físico se actualice solo; la carga
    manual queda como respaldo para cuando no hay un PDF limpio. `gallons`
    guarda la suma de los galones de FREIGHT-FUELS (lo que el camión trae
    de verdad, no el galonaje de facturación -- son distintos, ver
    fisico_invoice_parser.py) y `amount` guarda el TOTAL INVOICE AMOUNT
    DUE de abajo de todo (con flete e impuestos, el costo real pagado) --
    mismo criterio que ya usaba la carga manual (`=7600+1199` en el
    ejemplo real). `fuel_invoice_lines` guarda el detalle por grado
    (87/93 octanos) de cada factura leída por PDF -- solo a título
    informativo/histórico, no participa del cálculo de Físico (que sigue
    usando el agregado `gallons`/`amount` de `fuel_invoices`, sin split
    por grado).
  - `fisico_months`: un renglón por mes con dos cosas editables a mano
    que NO se pueden calcular solas:
      * el override del "Inventario Inicial Teórico" (gal + $) -- hace
        falta para el primer mes que se usa este módulo (no hay mes
        anterior del que encadenar), y queda disponible siempre por si
        hace falta corregirlo a mano en cualquier otro mes (mismo
        criterio que caja_db.set_month_opening_balance).
      * la lectura física real de los tanques ("Inventario Final Real",
        columna H del Excel real) -- SIEMPRE es un dato externo (una
        medición manual/de gauge), nunca se puede derivar de nada ya
        cargado en la app.
"""

import os
import sqlite3
from datetime import date, datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "fisico.db")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    _ensure_schema(conn)
    return conn


def _ensure_schema(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS fuel_invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_date TEXT NOT NULL,
            due_date TEXT,
            invoice_number TEXT,
            gallons REAL NOT NULL,
            amount REAL NOT NULL,
            source TEXT,
            created_at TEXT,
            updated_at TEXT
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_fuel_invoices_date ON fuel_invoices(invoice_date)")
    _ensure_columns(conn, "fuel_invoices", {"bol_number": "TEXT"})
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS fuel_invoice_lines (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_id INTEGER NOT NULL,
            product_code TEXT,
            description TEXT,
            qty_billing REAL,
            qty_freight REAL,
            unit_price_wo_tax REAL,
            unit_price_with_tax REAL,
            line_total REAL
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_fuel_invoice_lines_invoice ON fuel_invoice_lines(invoice_id)"
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS fisico_months (
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            initial_gallons_override REAL,
            initial_amount_override REAL,
            real_ending_gallons REAL,
            real_reading_date TEXT,
            updated_at TEXT,
            PRIMARY KEY (year, month)
        )
        """
    )
    conn.commit()


def _ensure_columns(conn, table, columns):
    """Mismo patrón que reportes_db._ensure_columns -- agrega columnas nuevas a una tabla ya existente sin romper bases viejas."""
    existing = {row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    for name, col_type in columns.items():
        if name not in existing:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {col_type}")


def _now():
    return datetime.now().isoformat(timespec="seconds")


def _date_key(value):
    if value is None or value == "":
        return None
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    return str(value)


def add_invoice(invoice_date, due_date, invoice_number, gallons, amount, source="manual", bol_number=None, lines=None):
    """
    `lines` (opcional): detalle por grado leído del PDF -- lista de dicts
    con "product_code"/"description"/"qty_billing"/"qty_freight"/
    "unit_price_wo_tax"/"unit_price_with_tax"/"line_total" (ver
    fisico_invoice_parser.extract_fuel_invoice). La carga manual no manda
    `lines` (None) -- el detalle por grado solo existe cuando se lee de un
    PDF real. `gallons`/`amount` siguen siendo el agregado que usa
    fisico.build_month_report, sin importar si vinieron de `lines` ya
    sumadas o de la carga manual de siempre.
    """
    now = _now()
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO fuel_invoices
                (invoice_date, due_date, invoice_number, gallons, amount, source, bol_number, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                _date_key(invoice_date),
                _date_key(due_date),
                (invoice_number or "").strip() or None,
                float(gallons),
                float(amount),
                source,
                (bol_number or "").strip() or None,
                now,
                now,
            ),
        )
        invoice_id = conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]
        for line in (lines or []):
            conn.execute(
                """
                INSERT INTO fuel_invoice_lines
                    (invoice_id, product_code, description, qty_billing, qty_freight,
                     unit_price_wo_tax, unit_price_with_tax, line_total)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    invoice_id,
                    line.get("product_code"),
                    line.get("description"),
                    line.get("qty_billing"),
                    line.get("qty_freight"),
                    line.get("unit_price_wo_tax"),
                    line.get("unit_price_with_tax"),
                    line.get("line_total"),
                ),
            )
        conn.commit()
        return invoice_id
    finally:
        conn.close()


def update_invoice(invoice_id, invoice_date, due_date, invoice_number, gallons, amount):
    """
    Corrige a mano una factura ya guardada (leída de PDF, posiblemente con
    un campo mal leído) -- pedido explícito del usuario (2026-09-21): al
    sacar la carga manual, la única forma de arreglar un dato mal leído es
    editar la fila ya guardada. Mismo criterio "editar la fila completa"
    que eft_db.update_coupon_row/gettel_db.update_pago. No toca `source`
    (sigue siendo "pdf") ni `bol_number`/`lines` (detalle por grado,
    informativo, no se edita acá).
    """
    now = _now()
    conn = _connect()
    try:
        cur = conn.execute(
            """
            UPDATE fuel_invoices
            SET invoice_date = ?, due_date = ?, invoice_number = ?, gallons = ?, amount = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                _date_key(invoice_date), _date_key(due_date), (invoice_number or "").strip() or None,
                float(gallons), float(amount), now, invoice_id,
            ),
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def delete_invoice(invoice_id):
    conn = _connect()
    try:
        conn.execute("DELETE FROM fuel_invoice_lines WHERE invoice_id = ?", (invoice_id,))
        conn.execute("DELETE FROM fuel_invoices WHERE id = ?", (invoice_id,))
        conn.commit()
    finally:
        conn.close()


def get_invoice_lines(invoice_id):
    """Detalle por grado de una factura (vacío para las cargadas a mano, que nunca tienen lines)."""
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM fuel_invoice_lines WHERE invoice_id = ? ORDER BY id", (invoice_id,)
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def find_invoice_by_number(invoice_number):
    """
    Busca una factura ya guardada por N° de factura, en CUALQUIER mes --
    para detectar duplicados al subir PDFs (mismo criterio que
    proveedores_db.save_invoice: un N° de factura repetido nunca se
    vuelve a insertar). None si no hay N° para buscar o no se encontró.
    """
    number = (invoice_number or "").strip()
    if not number:
        return None
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT * FROM fuel_invoices WHERE invoice_number = ? COLLATE NOCASE", (number,)
        ).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None


def get_month_invoices(year, month):
    """Facturas del mes (por fecha de factura), más viejas primero -- mismo criterio que el Excel real."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM fuel_invoices WHERE invoice_date LIKE ? ORDER BY invoice_date, id",
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()
    invoices = [dict(row) for row in rows]
    for invoice in invoices:
        invoice["lines"] = get_invoice_lines(invoice["id"])
    return invoices


def get_invoice_years():
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT DISTINCT substr(invoice_date, 1, 4) AS y FROM fuel_invoices ORDER BY y"
        ).fetchall()
        return [int(row["y"]) for row in rows if row["y"]]
    finally:
        conn.close()


def get_month_settings(year, month):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT * FROM fisico_months WHERE year = ? AND month = ?", (year, month)
        ).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None


def get_settings_years():
    conn = _connect()
    try:
        rows = conn.execute("SELECT DISTINCT year AS y FROM fisico_months ORDER BY y").fetchall()
        return [int(row["y"]) for row in rows if row["y"] is not None]
    finally:
        conn.close()


def set_month_initial_override(year, month, gallons, amount):
    """gallons/amount en None borra el override -- vuelve a encadenarse del Inventario Final Teórico del mes anterior."""
    now = _now()
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO fisico_months (year, month, initial_gallons_override, initial_amount_override, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(year, month) DO UPDATE SET
                initial_gallons_override = excluded.initial_gallons_override,
                initial_amount_override = excluded.initial_amount_override,
                updated_at = excluded.updated_at
            """,
            (year, month, gallons, amount, now),
        )
        conn.commit()
    finally:
        conn.close()


def set_month_real_ending(year, month, gallons, reading_date=None):
    """gallons en None borra la lectura -- el mes vuelve a quedar "pendiente de lectura física"."""
    now = _now()
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO fisico_months (year, month, real_ending_gallons, real_reading_date, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(year, month) DO UPDATE SET
                real_ending_gallons = excluded.real_ending_gallons,
                real_reading_date = excluded.real_reading_date,
                updated_at = excluded.updated_at
            """,
            (year, month, gallons, _date_key(reading_date), now),
        )
        conn.commit()
    finally:
        conn.close()
