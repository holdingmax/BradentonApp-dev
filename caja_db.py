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
- caja_months: Saldo Inicial (editable) y un override opcional de Saldo
  Final por mes -- si no hay override, el Saldo Final "real" es el último
  Saldo corrido que calcula caja.py, nunca algo que se guarde acá aparte.

Los documentos de referencia del mes (Excel de gastos, comprobantes de
depósito, etc.) vivían acá como `caja_attachments` -- pedido del usuario
(2026-09-14) fue que Caja tuviera su propio apartado de Documentos "como
los demás módulos" (página propia, separada del cuadro principal, ver
webapp.py: _DOCUMENTS_MODULES) -- se migraron a documents_db.py (módulo
"caja", mismo mecanismo genérico que ya usan EFT/Gettel/CMV) y esta tabla
se dejó de usar.
"""

import os
import sqlite3
from datetime import date, datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "caja.db")


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
        cur = conn.execute(
            "INSERT INTO caja_expense_items (date, amount, detail, created_at) VALUES (?, ?, ?, ?)",
            (key, amount, (detail or "").strip() or None, now),
        )
        conn.commit()
        return cur.lastrowid
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




def get_expense_years():
    """
    Años distintos con al menos un gasto de Caja cargado a mano -- pedido
    explícito del usuario (2026-09-17, módulo "Reportes"): Caja no guarda
    casi ningún dato propio (todo lo demás sale de Chase/Lottery/Store
    Info al vuelo, ver caja.build_month_report_from_db), así que esto
    cubre el caso de un mes con gastos cargados pero sin ninguna otra
    fuente todavía -- ver caja.get_available_years, que une esto con los
    años de las otras 3 fuentes.
    """
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT DISTINCT substr(date, 1, 4) AS y FROM caja_expense_items ORDER BY y"
        ).fetchall()
        return [int(row["y"]) for row in rows if row["y"]]
    finally:
        conn.close()


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
