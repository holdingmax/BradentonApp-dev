"""
Horas de Trabajo -- guarda los reportes semanales de "Clock In/Out Detail
Report" ya extraídos por horas_trabajo.py: una semana = un reporte, con un
renglón por empleado (horas trabajadas leídas del PDF, tarifa por hora y
descuento editables). El monto a pagar por empleado y el total de la semana
se calculan siempre en el momento (hours * rate - deduct) a partir de estos
tres campos -- nunca se guarda un total ya calculado, mismo criterio que
DIF EFECT/Saldo en Caja.

reportes_data/horas_trabajo.db (mismo directorio gitignored de siempre).
"""

import os
import sqlite3
from datetime import datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "horas_trabajo.db")

DEFAULT_HOURLY_RATE = 15.0


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
        CREATE TABLE IF NOT EXISTS horas_trabajo_weeks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_date TEXT NOT NULL UNIQUE,
            period_from TEXT,
            period_to TEXT,
            source TEXT,
            document_id INTEGER,
            created_at TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS horas_trabajo_employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            week_id INTEGER NOT NULL,
            employee_name TEXT NOT NULL,
            hours_label TEXT,
            hours REAL NOT NULL DEFAULT 0,
            rate REAL NOT NULL DEFAULT 15.0,
            deduct REAL NOT NULL DEFAULT 0
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_horas_trabajo_weeks_date ON horas_trabajo_weeks (report_date)"
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_horas_trabajo_employees_week ON horas_trabajo_employees (week_id)"
    )
    conn.commit()


def _now():
    return datetime.now().isoformat(timespec="seconds")


def _date_key(value):
    if value is None:
        return None
    if isinstance(value, str):
        return value
    return value.isoformat()


def upsert_week(report_date, period_from, period_to, employees, source="ocr", document_id=None):
    """
    Reemplaza por completo la semana de `report_date` -- mismo criterio que
    reprocesar un día ya cargado de Reporte Diario: si ya existía, se
    borran sus empleados y se reinsertan los nuevos (así resubir el mismo
    PDF no duplica nada); si no, se crea. `employees` es una lista de
    {"employee_name","hours_label","hours"} -- rate/deduct arrancan en su
    default ($15/hora, sin descuento). Devuelve el week_id.
    """
    key = _date_key(report_date)
    conn = _connect()
    try:
        row = conn.execute("SELECT id FROM horas_trabajo_weeks WHERE report_date = ?", (key,)).fetchone()
        if row:
            week_id = row["id"]
            conn.execute(
                "UPDATE horas_trabajo_weeks SET period_from=?, period_to=?, source=?, "
                "document_id=COALESCE(?, document_id) WHERE id=?",
                (_date_key(period_from), _date_key(period_to), source, document_id, week_id),
            )
            conn.execute("DELETE FROM horas_trabajo_employees WHERE week_id = ?", (week_id,))
        else:
            cur = conn.execute(
                "INSERT INTO horas_trabajo_weeks (report_date, period_from, period_to, source, document_id, created_at) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (key, _date_key(period_from), _date_key(period_to), source, document_id, _now()),
            )
            week_id = cur.lastrowid

        for emp in employees:
            conn.execute(
                "INSERT INTO horas_trabajo_employees (week_id, employee_name, hours_label, hours, rate, deduct) "
                "VALUES (?, ?, ?, ?, ?, 0)",
                (week_id, emp["employee_name"], emp.get("hours_label"), float(emp.get("hours") or 0.0), DEFAULT_HOURLY_RATE),
            )
        conn.commit()
    finally:
        conn.close()
    return week_id


def set_week_document(week_id, document_id):
    conn = _connect()
    try:
        conn.execute("UPDATE horas_trabajo_weeks SET document_id = ? WHERE id = ?", (document_id, week_id))
        conn.commit()
    finally:
        conn.close()


def add_employee(week_id, employee_name, hours=0.0, rate=DEFAULT_HOURLY_RATE, deduct=0.0):
    conn = _connect()
    try:
        cur = conn.execute(
            "INSERT INTO horas_trabajo_employees (week_id, employee_name, hours_label, hours, rate, deduct) "
            "VALUES (?, ?, NULL, ?, ?, ?)",
            (week_id, employee_name.strip(), float(hours or 0.0), float(rate or 0.0), float(deduct or 0.0)),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def update_employee(employee_id, employee_name=None, hours=None, rate=None, deduct=None):
    fields, values = [], []
    if employee_name is not None:
        fields.append("employee_name = ?")
        values.append(employee_name.strip())
    if hours is not None:
        fields.append("hours = ?")
        values.append(float(hours))
    if rate is not None:
        fields.append("rate = ?")
        values.append(float(rate))
    if deduct is not None:
        fields.append("deduct = ?")
        values.append(float(deduct))
    if not fields:
        return
    values.append(employee_id)
    conn = _connect()
    try:
        conn.execute(f"UPDATE horas_trabajo_employees SET {', '.join(fields)} WHERE id = ?", values)
        conn.commit()
    finally:
        conn.close()


def delete_employee(employee_id):
    conn = _connect()
    try:
        conn.execute("DELETE FROM horas_trabajo_employees WHERE id = ?", (employee_id,))
        conn.commit()
    finally:
        conn.close()


def get_employee_week(employee_id):
    """week_id del empleado, o None -- para poder volver a la página del mes correcto tras editar/borrar."""
    conn = _connect()
    try:
        row = conn.execute("SELECT week_id FROM horas_trabajo_employees WHERE id = ?", (employee_id,)).fetchone()
        return row["week_id"] if row else None
    finally:
        conn.close()


def delete_week(week_id):
    conn = _connect()
    try:
        conn.execute("DELETE FROM horas_trabajo_employees WHERE week_id = ?", (week_id,))
        conn.execute("DELETE FROM horas_trabajo_weeks WHERE id = ?", (week_id,))
        conn.commit()
    finally:
        conn.close()


def get_week(week_id):
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM horas_trabajo_weeks WHERE id = ?", (week_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_month_weeks(year, month):
    """Semanas cuyo report_date cae en este mes, más antigua primero (mismo
    orden ascendente ya elegido para EFT) -- cada una con sus empleados."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        weeks = conn.execute(
            "SELECT * FROM horas_trabajo_weeks WHERE report_date LIKE ? ORDER BY report_date ASC",
            (f"{prefix}%",),
        ).fetchall()
        result = []
        for week in weeks:
            employees = conn.execute(
                "SELECT * FROM horas_trabajo_employees WHERE week_id = ? ORDER BY id ASC",
                (week["id"],),
            ).fetchall()
            result.append({"week": dict(week), "employees": [dict(e) for e in employees]})
        return result
    finally:
        conn.close()
