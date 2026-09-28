"""
Depósitos guardados (Documentos -> Depósitos) -- capa de datos pura. Un
renglón por recibo de depósito (no por PDF subido): un PDF con 3 recibos
deja 3 filas, cada una con su propio PDF de una página (ver depositos.py).

Fecha, monto y descripción salen del OCR y se pueden editar a mano. El mes
en que se lista cada depósito es el de su fecha (se mueve solo si se corrige
la fecha).

reportes_data/depositos.db + reportes_data/depositos/{año}/{mes}/*.pdf
"""

import os
import sqlite3
import uuid
from datetime import datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "depositos.db")
_FILES_DIR = os.path.join(_BASE_DIR, "depositos")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS deposits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            deposit_date TEXT,
            amount REAL,
            description TEXT,
            tx_number INTEGER,
            kind TEXT,
            pdf_path TEXT NOT NULL,
            source_filename TEXT,
            page_index INTEGER,
            edited INTEGER DEFAULT 0,
            uploaded_at TEXT
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_deposits_period ON deposits (year, month)")
    conn.commit()
    return conn


def _now():
    return datetime.now().isoformat(timespec="seconds")


def absolute_path(relpath):
    return os.path.join(_FILES_DIR, relpath)


def new_pdf_relpath(year, month):
    return os.path.join(f"{year:04d}", f"{month:02d}", f"{uuid.uuid4().hex[:12]}.pdf")


def find_duplicate(tx_number, deposit_date, amount):
    """Mismo recibo ya cargado (el N° de transacción se repite entre días, por eso se pide todo)."""
    if tx_number is None or deposit_date is None or amount is None:
        return None
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT id FROM deposits WHERE tx_number = ? AND deposit_date = ? AND ABS(amount - ?) < 0.005",
            (tx_number, deposit_date, amount),
        ).fetchone()
        return row["id"] if row else None
    finally:
        conn.close()


def add_deposit(year, month, deposit_date, amount, description, tx_number, kind, pdf_relpath, source_filename, page_index):
    conn = _connect()
    try:
        cur = conn.execute(
            """
            INSERT INTO deposits (year, month, deposit_date, amount, description, tx_number, kind,
                                  pdf_path, source_filename, page_index, uploaded_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (year, month, deposit_date, amount, description, tx_number, kind,
             pdf_relpath, source_filename, page_index, _now()),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def find_by_source_page(source_filename, page_index):
    """Mismo PDF y misma página ya cargados (para recibos con algún dato sin leer, donde find_duplicate no alcanza)."""
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT id FROM deposits WHERE source_filename = ? AND page_index = ?",
            (source_filename, page_index),
        ).fetchone()
        return row["id"] if row else None
    finally:
        conn.close()


def update_deposit(deposit_id, deposit_date, amount, description, kind):
    """
    Corrección a mano. Si cambia la fecha, el depósito pasa al mes de la
    fecha nueva. `kind` (Food Truck/Ice Machine/otra aclaración, o None =
    depósito normal) es lo que decide si cuenta en el control contra Caja.
    """
    fields = {"deposit_date": deposit_date, "amount": amount, "description": description,
              "kind": kind, "edited": 1}
    if deposit_date:
        fields["year"], fields["month"] = int(deposit_date[:4]), int(deposit_date[5:7])
    conn = _connect()
    try:
        conn.execute(
            f"UPDATE deposits SET {', '.join(f'{k} = ?' for k in fields)} WHERE id = ?",
            tuple(fields.values()) + (deposit_id,),
        )
        conn.commit()
    finally:
        conn.close()


def list_month(year, month):
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM deposits WHERE year = ? AND month = ? "
            "ORDER BY deposit_date IS NULL, deposit_date, tx_number, id",
            (year, month),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_deposit(deposit_id):
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM deposits WHERE id = ?", (deposit_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def delete_deposit(deposit_id):
    deposit = get_deposit(deposit_id)
    if not deposit:
        return None
    conn = _connect()
    try:
        conn.execute("DELETE FROM deposits WHERE id = ?", (deposit_id,))
        conn.commit()
    finally:
        conn.close()
    try:
        os.remove(absolute_path(deposit["pdf_path"]))
    except OSError:
        pass
    return deposit
