"""
Persistencia de Ice Machine y Food Truck (pedido del usuario, 2026-10-07):
los Payment Summary de Cantaloupe, uno por semana. Los recibos de depósito
(Ice Machine, Food Truck, Vaccumms) son los de depositos_db, compartidos con
el módulo Depósitos. Capa de datos pura: lectura y control en ice_machine.py.

reportes_data/ice_machine.db (gitignored, como las demás bases), o el esquema
ice_machine de PostgreSQL con DATABASE_URL (db.py; la tabla la crea la
migración migraciones/versions/0002_ice_machine.py).
"""

import app_paths
import os
from datetime import datetime

import db

_SCHEMA = "ice_machine"
_DB_PATH = os.path.join(app_paths.DATA_DIR, "ice_machine.db")

FIELDS = ("summary_no", "reference", "from_date", "to_date", "transactions", "gross",
          "process_fees", "service_fees", "net", "notes")


def _init_sqlite(conn):
    conn.execute("PRAGMA journal_mode=WAL")
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


def _connect():
    return db.connect(_SCHEMA, _DB_PATH, _init_sqlite)


def save_summary(summary, filename=None):
    """Guarda (o reemplaza, mismo número) un Payment Summary. Devuelve True si ya estaba."""
    conn = _connect()
    try:
        with conn:
            existed = conn.execute("SELECT 1 FROM payment_summaries WHERE summary_no = ?",
                                   (summary["summary_no"],)).fetchone() is not None
            # Todas las columnas van en el INSERT: actualizar la fila que ya
            # estaba deja lo mismo que el INSERT OR REPLACE de antes, y
            # ON CONFLICT lo entienden SQLite y Postgres.
            columns = (*FIELDS, "filename", "uploaded_at")
            updates = ", ".join(f"{c} = excluded.{c}" for c in columns if c != "summary_no")
            conn.execute(
                f"INSERT INTO payment_summaries ({', '.join(columns)}) VALUES ({', '.join('?' for _ in columns)}) "
                f"ON CONFLICT (summary_no) DO UPDATE SET {updates}",
                [*(summary.get(f) for f in FIELDS), filename, datetime.now().isoformat(timespec="seconds")],
            )
        return existed
    finally:
        conn.close()


def list_month(year, month):
    """Los Payment Summary pagados en el mes (por la fecha "To", que es la del pago en Chase)."""
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM payment_summaries WHERE to_date LIKE ? ORDER BY to_date, rowid",
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
