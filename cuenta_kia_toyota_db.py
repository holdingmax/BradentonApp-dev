"""
Persistencia de la cuenta corriente de Kia y Toyota (pedido del usuario,
2026-10-07): lo leído de las hojas de Gettel-Toyota de los Excel de Cierre
(cargas por día y empresa, pagos, Local Account y VS de Store Info). Subir
de nuevo un Excel solo completa: un día ya cargado se reemplaza si el nuevo
trae las empresas separadas, un pago repetido (fecha, N° de transacción e
importe) no se duplica, y la empresa asignada a mano nunca se pisa.

reportes_data/cuenta_kia_toyota.db (gitignored, como las demás bases).
"""

import os
import sqlite3
from datetime import datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "cuenta_kia_toyota.db")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS charges (
            date TEXT PRIMARY KEY,
            la REAL, kia REAL, kia_gal REAL, toyota REAL, toyota_gal REAL,
            main INTEGER, source TEXT, updated_at TEXT
        );
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL, transc TEXT NOT NULL, amount REAL NOT NULL,
            company TEXT, company_manual INTEGER DEFAULT 0,
            source TEXT, updated_at TEXT,
            UNIQUE (date, transc, amount)
        );
        CREATE TABLE IF NOT EXISTS pos_only_assignments (
            date TEXT PRIMARY KEY, company TEXT, updated_at TEXT
        );
        CREATE TABLE IF NOT EXISTS pos_days (
            date TEXT PRIMARY KEY, la REAL, vs REAL, source TEXT
        );
        """
    )
    return conn


def _now():
    return datetime.now().isoformat(timespec="seconds")


def save_import(data, source):
    """Guarda lo leído de un Excel (cuenta_kia_toyota.read_control_workbook). Devuelve (días, pagos nuevos)."""
    from cuenta_kia_toyota import best_days

    conn = _connect()
    new_days = new_payments = 0
    try:
        with conn:
            for day, d in best_days(data["days"]).items():
                current = conn.execute("SELECT kia, toyota, main FROM charges WHERE date = ?", (day,)).fetchone()
                has = bool((d["kia"] or 0) or (d["toyota"] or 0))
                if current is not None:
                    current_has = bool((current["kia"] or 0) or (current["toyota"] or 0))
                    if (current_has, bool(current["main"])) > (has, bool(d.get("main", True))):
                        continue
                else:
                    new_days += 1
                conn.execute(
                    "INSERT OR REPLACE INTO charges (date, la, kia, kia_gal, toyota, toyota_gal, main, source, updated_at) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (day, d["la"], d["kia"], d["kia_gal"], d["toyota"], d["toyota_gal"], int(d.get("main", True)),
                     source, _now()),
                )
            for p in data["payments"]:
                current = conn.execute("SELECT id, company FROM payments WHERE date = ? AND transc = ? AND amount = ?",
                                       (p["date"], p["transc"], p["amount"])).fetchone()
                if current is None:
                    conn.execute("INSERT INTO payments (date, transc, amount, company, source, updated_at) "
                                 "VALUES (?, ?, ?, ?, ?, ?)",
                                 (p["date"], p["transc"], p["amount"], p["company"], source, _now()))
                    new_payments += 1
                elif current["company"] is None and p["company"]:
                    conn.execute("UPDATE payments SET company = ?, updated_at = ? WHERE id = ? AND company_manual = 0",
                                 (p["company"], _now(), current["id"]))
            for d in data["pos"]:
                if d["la"] is None and d["vs"] is None:
                    continue
                conn.execute("INSERT OR REPLACE INTO pos_days (date, la, vs, source) VALUES (?, ?, ?, ?)",
                             (d["date"], d["la"], d["vs"], source))
    finally:
        conn.close()
    return new_days, new_payments


def get_charges():
    conn = _connect()
    try:
        return {r["date"]: dict(r) for r in conn.execute("SELECT * FROM charges")}
    finally:
        conn.close()


def get_payments():
    conn = _connect()
    try:
        return [dict(r) for r in conn.execute("SELECT * FROM payments ORDER BY date, transc")]
    finally:
        conn.close()


def get_pos_days():
    conn = _connect()
    try:
        return {r["date"]: dict(r) for r in conn.execute("SELECT * FROM pos_days")}
    finally:
        conn.close()


def get_pos_only_assignments():
    """{fecha: empresa} de lo cobrado en el POS sin recibo, elegido a mano."""
    conn = _connect()
    try:
        return {r["date"]: r["company"] for r in conn.execute("SELECT * FROM pos_only_assignments")}
    finally:
        conn.close()


def assign_pos_only(day, company):
    conn = _connect()
    try:
        with conn:
            conn.execute("INSERT OR REPLACE INTO pos_only_assignments (date, company, updated_at) VALUES (?, ?, ?)",
                         (day, company, _now()))
    finally:
        conn.close()


def assign_company(payment_ids, company):
    """Empresa elegida a mano para los cobros de un pago (no la pisa ningún Excel)."""
    conn = _connect()
    try:
        with conn:
            conn.executemany("UPDATE payments SET company = ?, company_manual = 1, updated_at = ? WHERE id = ?",
                             [(company, _now(), pid) for pid in payment_ids])
    finally:
        conn.close()
