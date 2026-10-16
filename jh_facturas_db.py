"""
Facturas varias de J.H. Williams (pedido del usuario, 2026-10-13): Mako,
VPN, Chevron Network Fee, gift cards y las que vengan sueltas, leídas por
jh_facturas.py. Una fila por factura; el mes de cada una es el de su fecha.
Próximos EFT (eft_pronostico.py) suma las que ningún EFT cargado pagó.

<carpeta de datos>/jh_facturas.db (app_paths.DATA_DIR), o el esquema
jh_facturas de PostgreSQL con DATABASE_URL (db.py; la tabla la crea la
migración migraciones/versions/0001_jh_facturas.py).
"""

import os
from datetime import datetime

import app_paths
import db

_SCHEMA = "jh_facturas"
_DB_PATH = os.path.join(app_paths.DATA_DIR, "jh_facturas.db")
_FIELDS = ("invoice_number", "invoice_date", "due_date", "amount", "kind", "description", "detail", "source_filename", "uploaded_at")


def _init_sqlite(conn):
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS jh_misc_invoices (
            invoice_number TEXT PRIMARY KEY,
            invoice_date TEXT NOT NULL,
            due_date TEXT NOT NULL,
            amount REAL NOT NULL,
            kind TEXT NOT NULL,
            description TEXT,
            detail TEXT,
            source_filename TEXT,
            uploaded_at TEXT
        )
        """
    )


def _connect():
    return db.connect(_SCHEMA, _DB_PATH, _init_sqlite)


def _row(row):
    return {field: row[field] for field in _FIELDS}


def save_invoice(invoice, source_filename=None):
    """
    Guarda una factura de jh_facturas.extract_invoice: "new" si es nueva,
    "same" si ya estaba igual. ValueError si ya estaba cargada con otro
    importe o fecha (se deja la cargada).
    """
    number = invoice["invoice_number"]
    conn = _connect()
    try:
        stored = conn.execute("SELECT * FROM jh_misc_invoices WHERE invoice_number = ?", (number,)).fetchone()
        if stored:
            if round(stored["amount"], 2) == round(invoice["amount"], 2) and stored["invoice_date"] == invoice["invoice_date"].isoformat():
                return "same"
            raise ValueError(
                f"La factura {number} ya estaba cargada con otro importe o fecha "
                f"(${stored['amount']:,.2f}): se dejó la cargada. Si está mal, eliminala y subila de nuevo."
            )
        conn.execute(
            "INSERT INTO jh_misc_invoices (invoice_number, invoice_date, due_date, amount, kind, description, detail, "
            "source_filename, uploaded_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (number, invoice["invoice_date"].isoformat(), invoice["due_date"].isoformat(), round(invoice["amount"], 2),
             invoice["kind"], invoice.get("description"), invoice.get("detail"), source_filename,
             datetime.now().isoformat(timespec="seconds")),
        )
        conn.commit()
        return "new"
    finally:
        conn.close()


def get_month_invoices(year, month):
    """Las facturas con fecha en ese mes, de la más vieja a la más nueva."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM jh_misc_invoices WHERE invoice_date LIKE ? ORDER BY invoice_date, invoice_number", (prefix + "%",)
        ).fetchall()
    finally:
        conn.close()
    return [_row(r) for r in rows]


def list_all_invoices():
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM jh_misc_invoices ORDER BY invoice_date, invoice_number").fetchall()
    finally:
        conn.close()
    return [_row(r) for r in rows]


def delete_invoice(invoice_number):
    conn = _connect()
    try:
        conn.execute("DELETE FROM jh_misc_invoices WHERE invoice_number = ?", (invoice_number,))
        conn.commit()
    finally:
        conn.close()
