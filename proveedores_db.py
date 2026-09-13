"""
Proveedores -- Carga de Datos: guarda las facturas ya extraídas de cada PDF
(fecha/N°/monto/proveedor) sin escribir ningún Excel Ledger -- mismo patrón
que chase_db.py/eft_db.py. Primer paso de este módulo (2026-09-14, pedido
explícito del usuario: "quiero que empieces con el modulo de proveedores y
donde pueda guardar sus facturas") -- reusa el motor de detección/extracción
de 32 proveedores + el dinámico tal cual está en proveedores.py
(`extract_invoices_from_pdf`), sin tocar ni un carácter de esa lógica; acá
solo se guarda el resultado.

reportes_data/proveedores.db (mismo directorio gitignored de siempre):
- supplier_invoices: una fila por factura, clave natural (supplier_key,
  invoice_no) para el mismo criterio de duplicado que ya usa el Ledger real
  ("un mismo N° de factura ya cargado se omite").

El PDF original se guarda aparte con documents_db.py (módulo "proveedores"),
igual que EFT/Gettel/CMV -- no se duplica esa pieza acá.
"""

import os
import sqlite3
from datetime import date, datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "proveedores.db")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    # timeout=30 + WAL -- ver reportes_db.py: necesario desde que las cargas
    # en segundo plano (jobs.py) pueden escribir de verdad en paralelo.
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS supplier_invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_key TEXT NOT NULL,
            supplier_label TEXT NOT NULL,
            invoice_date TEXT NOT NULL,
            invoice_no TEXT NOT NULL,
            amount REAL NOT NULL,
            source_filename TEXT,
            uploaded_at TEXT,
            UNIQUE (supplier_key, invoice_no)
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_supplier_invoices_period ON supplier_invoices (invoice_date)"
    )
    return conn


def _date_key(value):
    if isinstance(value, str):
        return value
    if isinstance(value, (date, datetime)):
        return value.isoformat() if isinstance(value, date) and not isinstance(value, datetime) else value.date().isoformat()
    return str(value)


def invoice_exists(supplier_key, invoice_no):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT 1 FROM supplier_invoices WHERE supplier_key = ? AND invoice_no = ?",
            (supplier_key, invoice_no),
        ).fetchone()
    finally:
        conn.close()
    return row is not None


def save_invoice(supplier_key, supplier_label, invoice_date, invoice_no, amount, source_filename=None):
    """
    Guarda una factura -- devuelve False sin tocar nada si ese N° de
    factura ya estaba cargado para este proveedor (mismo criterio de
    duplicado que el Ledger real: por N° de factura, no por fecha/monto).
    """
    conn = _connect()
    try:
        try:
            conn.execute(
                """
                INSERT INTO supplier_invoices
                    (supplier_key, supplier_label, invoice_date, invoice_no, amount, source_filename, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    supplier_key, supplier_label, _date_key(invoice_date), str(invoice_no),
                    float(amount), source_filename, datetime.now().isoformat(timespec="seconds"),
                ),
            )
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
    finally:
        conn.close()


def get_month_invoices(year, month):
    """Todas las facturas de un mes (por fecha de factura), agrupadas por proveedor -- más reciente primero dentro de cada uno."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM supplier_invoices WHERE invoice_date LIKE ? ORDER BY supplier_label, invoice_date DESC",
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()

    by_supplier = {}
    for row in rows:
        item = dict(row)
        by_supplier.setdefault(item["supplier_key"], {"label": item["supplier_label"], "invoices": [], "total": 0.0})
        by_supplier[item["supplier_key"]]["invoices"].append(item)
        by_supplier[item["supplier_key"]]["total"] += item["amount"]
    groups = sorted(by_supplier.values(), key=lambda g: g["label"])
    for group in groups:
        group["total"] = round(group["total"], 2)
    return groups


def list_suppliers_with_counts():
    """Un renglón por proveedor con al menos una factura guardada -- para el selector de "ver historial completo de este proveedor"."""
    conn = _connect()
    try:
        rows = conn.execute(
            """
            SELECT supplier_key, supplier_label, COUNT(*) AS n, SUM(amount) AS total
            FROM supplier_invoices GROUP BY supplier_key ORDER BY supplier_label
            """
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def get_supplier_invoices(supplier_key):
    """Historial completo (todos los meses) de un proveedor puntual, más reciente primero."""
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM supplier_invoices WHERE supplier_key = ? ORDER BY invoice_date DESC",
            (supplier_key,),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def delete_invoice(invoice_id):
    conn = _connect()
    try:
        cur = conn.execute("DELETE FROM supplier_invoices WHERE id = ?", (invoice_id,))
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()
