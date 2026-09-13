"""
Persistencia de lo único que Caja (Carga de Datos) necesita guardar por su
cuenta -- ver CLAUDE.md, "Módulo Caja -- Carga de Datos". Todo lo demás
(Depósitos/Food Truck-Ice/Lottery Cuenta Final/Total Sales/Cash/Other/Total
Revenue) sale en el momento de chase_db/lottery_db/reportes_db, sin guardado
propio -- pero los GASTOS de caja (columna M del Excel real) no vienen de
ningún lado más, el usuario los tipea a mano; y el Saldo Inicial/Final de
cada mes también necesita poder editarse (encadenado del mes anterior por
default, ver caja.py: _resolve_opening_balance).

reportes_data/caja.db (mismo directorio gitignored que las demás bases):
- caja_expense_items: uno o más gastos en efectivo por día, cada uno con su
  propio detalle (a quién se le pagó) -- pedido explícito del usuario
  (2026-09-12, cuarta tanda): "poder escribirlos a mano en el sistema con
  un detalle de a quien pertenecen", mismo espíritu que el comentario de
  Excel que ya lleva cada celda de la columna M real. Reemplaza el modelo
  viejo (un solo monto por día, sin detalle) -- no había ningún dato real
  cargado todavía, así que no hizo falta migrar nada.
- caja_attachments: archivos de referencia (Excel de gastos, etc.) subidos
  para un mes -- "sistema de almacenamiento", pedido explícito del usuario:
  "pasarle todos los excels de gastos con caja y que esten ahi a mano para
  ver cuando estes parado en tal mes en particular". Nunca se leen/parsean,
  quedan solo para consulta mientras se tipean los gastos a mano de arriba.
- caja_months: Saldo Inicial (editable) y un override opcional de Saldo
  Final por mes -- si no hay override, el Saldo Final "real" es el último
  Saldo corrido que calcula caja.py, nunca algo que se guarde acá aparte.
"""

import os
import sqlite3
from datetime import date, datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "caja.db")
_ATTACHMENTS_DIR = os.path.join(_BASE_DIR, "caja_attachments")


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
        CREATE TABLE IF NOT EXISTS caja_expense_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            amount REAL NOT NULL,
            detail TEXT,
            created_at TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS caja_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            filename TEXT NOT NULL,
            stored_path TEXT NOT NULL,
            uploaded_at TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS caja_months (
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            opening_balance REAL,
            closing_balance_override REAL,
            updated_at TEXT,
            PRIMARY KEY (year, month)
        )
        """
    )
    # Índices por fecha -- pensado para escalar a años de gastos/documentos
    # sin que "los del mes tal" se vuelva un recorrido completo de la tabla
    # (ver CLAUDE.md, "pensar en una base de datos grande y confiable").
    conn.execute("CREATE INDEX IF NOT EXISTS idx_caja_expense_items_date ON caja_expense_items (date)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_caja_attachments_period ON caja_attachments (year, month)")
    conn.commit()


def _now():
    return datetime.now().isoformat(timespec="seconds")


def _date_key(value):
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    return str(value)


def add_expense_item(report_date, amount, detail):
    key = _date_key(report_date)
    now = _now()
    conn = _connect()
    try:
        conn.execute(
            "INSERT INTO caja_expense_items (date, amount, detail, created_at) VALUES (?, ?, ?, ?)",
            (key, amount, (detail or "").strip() or None, now),
        )
        conn.commit()
    finally:
        conn.close()


def delete_expense_item(item_id):
    conn = _connect()
    try:
        conn.execute("DELETE FROM caja_expense_items WHERE id = ?", (item_id,))
        conn.commit()
    finally:
        conn.close()


def get_month_expense_items(year, month):
    """{"YYYY-MM-DD": [{"id","amount","detail"}, ...]} -- puede haber varios por día."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM caja_expense_items WHERE date LIKE ? ORDER BY date, id",
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()
    by_date = {}
    for row in rows:
        by_date.setdefault(row["date"], []).append(
            {"id": row["id"], "amount": row["amount"], "detail": row["detail"]}
        )
    return by_date


def get_month_expenses(year, month):
    """{"YYYY-MM-DD": total} -- suma de los ítems de cada día, para DIF EFECT (caja.py)."""
    items_by_date = get_month_expense_items(year, month)
    return {
        day: round(sum(item["amount"] or 0.0 for item in items), 2)
        for day, items in items_by_date.items()
    }


def store_attachment(year, month, upload):
    """
    Guarda un archivo de referencia (Excel de gastos, etc.) subido para un
    mes -- nunca se lee/parsea, solo queda a mano para consultar. Mismo
    criterio de nombre-tal-cual que reportes_db.store_pdf_copy (una
    subcarpeta por archivo evita colisiones sin tocar el nombre original).
    """
    filename = os.path.basename(upload.filename)
    month_dir = os.path.join(_ATTACHMENTS_DIR, f"{year:04d}", f"{month:02d}")
    os.makedirs(month_dir, exist_ok=True)
    file_dir = os.path.join(month_dir, datetime.now().strftime("%Y%m%d%H%M%S%f"))
    os.makedirs(file_dir, exist_ok=True)
    stored_path = os.path.join(file_dir, filename)
    upload.save(stored_path)

    now = _now()
    conn = _connect()
    try:
        conn.execute(
            "INSERT INTO caja_attachments (year, month, filename, stored_path, uploaded_at) VALUES (?, ?, ?, ?, ?)",
            (year, month, filename, stored_path, now),
        )
        conn.commit()
    finally:
        conn.close()


def get_month_attachments(year, month):
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM caja_attachments WHERE year = ? AND month = ? ORDER BY uploaded_at",
            (year, month),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def get_attachment(attachment_id):
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM caja_attachments WHERE id = ?", (attachment_id,)).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None


def delete_attachment(attachment_id):
    attachment = get_attachment(attachment_id)
    conn = _connect()
    try:
        conn.execute("DELETE FROM caja_attachments WHERE id = ?", (attachment_id,))
        conn.commit()
    finally:
        conn.close()
    if attachment:
        try:
            os.remove(attachment["stored_path"])
        except OSError:
            pass


def get_month_settings(year, month):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT * FROM caja_months WHERE year = ? AND month = ?", (year, month)
        ).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None


def set_month_opening_balance(year, month, opening_balance):
    """opening_balance en None borra el override -- vuelve a encadenarse del saldo final del mes anterior."""
    now = _now()
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO caja_months (year, month, opening_balance, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(year, month) DO UPDATE SET
                opening_balance = excluded.opening_balance, updated_at = excluded.updated_at
            """,
            (year, month, opening_balance, now),
        )
        conn.commit()
    finally:
        conn.close()


def set_month_closing_override(year, month, closing_balance):
    """closing_balance en None borra el ajuste manual -- vuelve al Saldo corrido calculado."""
    now = _now()
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO caja_months (year, month, closing_balance_override, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(year, month) DO UPDATE SET
                closing_balance_override = excluded.closing_balance_override, updated_at = excluded.updated_at
            """,
            (year, month, closing_balance, now),
        )
        conn.commit()
    finally:
        conn.close()
