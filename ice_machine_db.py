"""
Persistencia de Ice Machine y Food Truck (pedido del usuario, 2026-10-07):
los Payment Summary de Cantaloupe, uno por semana. Los recibos de depósito
(Ice Machine, Food Truck, Vaccumms) son los de depositos_db, compartidos con
el módulo Depósitos. Capa de datos pura: lectura y control en ice_machine.py.

reportes_data/ice_machine.db (gitignored, como las demás bases).
"""

import app_paths
import os
import sqlite3
from datetime import datetime

_BASE_DIR = app_paths.DATA_DIR
_DB_PATH = os.path.join(_BASE_DIR, "ice_machine.db")

FIELDS = ("summary_no", "reference", "from_date", "to_date", "transactions", "gross",
          "process_fees", "service_fees", "net", "notes")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS payment_summaries (
            summary_no TEXT PRIMARY KEY,
            reference TEXT,
            from_date TEXT,
            to_date TEXT NOT NULL,
            transactions INTEGER,
            gross REAL,
            process_fees REAL,
            service_fees REAL,
            net REAL,
            notes TEXT,
            filename TEXT,
            uploaded_at TEXT
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_payment_summaries_to ON payment_summaries (to_date)")
    return conn


def save_summary(summary, filename=None):
    """Guarda (o reemplaza, mismo número) un Payment Summary. Devuelve True si ya estaba."""
    conn = _connect()
    try:
        with conn:
            existed = conn.execute("SELECT 1 FROM payment_summaries WHERE summary_no = ?",
                                   (summary["summary_no"],)).fetchone() is not None
            columns = (*FIELDS, "filename", "uploaded_at")
            conn.execute(
                f"INSERT OR REPLACE INTO payment_summaries ({', '.join(columns)}) VALUES ({', '.join('?' for _ in columns)})",
                [*(summary.get(f) for f in FIELDS), filename, datetime.now().isoformat(timespec="seconds")],
            )
        return existed
    finally:
        conn.close()


def list_month(year, month):
    """Los Payment Summary pagados en el mes (por la fecha "To", que es la del pago en Chase)."""
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM payment_summaries WHERE to_date LIKE ? ORDER BY to_date",
                            (f"{year:04d}-{month:02d}-%",)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def delete_summary(summary_no):
    conn = _connect()
    try:
        with conn:
            row = conn.execute("SELECT to_date FROM payment_summaries WHERE summary_no = ?", (summary_no,)).fetchone()
            conn.execute("DELETE FROM payment_summaries WHERE summary_no = ?", (summary_no,))
        return row["to_date"] if row else None
    finally:
        conn.close()
