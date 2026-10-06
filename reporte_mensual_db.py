"""
Persistencia del Reporte Mensual del POS (pedido del usuario, 2026-10-06):
un reporte por mes, con los totales de Store Info y los departamentos del
mes tal cual los imprime el "Resumen Ventas". Capa de datos pura: la lectura
del PDF y el cruce viven en reporte_mensual.py.

reportes_data/reporte_mensual.db (gitignored, como las demás bases).
"""

import json
import os
import sqlite3
from datetime import datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "reporte_mensual.db")

STORE_INFO_FIELDS = (
    "volume", "sales_fuel", "desc_comb", "non_fuel_total", "desc_otros",
    "tax_collect", "total_sales", "cash", "local_accounts", "other_amount",
    "network_revenue", "total_revenue",
)


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS monthly_reports (
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            from_date TEXT, to_date TEXT,
            {", ".join(f"{field} REAL" for field in STORE_INFO_FIELDS)},
            credit_terms_json TEXT,
            printed_department_amount REAL,
            printed_department_count INTEGER,
            warnings_json TEXT,
            filename TEXT,
            updated_at TEXT,
            PRIMARY KEY (year, month)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS monthly_report_departments (
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            department TEXT NOT NULL,
            count INTEGER,
            amount REAL,
            PRIMARY KEY (year, month, department)
        )
        """
    )
    return conn


def save_report(report, filename=None):
    """
    Reemplaza el reporte del mes entero con lo recién leído (`report` como lo
    devuelve reporte_mensual.extract_monthly_report). Los departamentos se
    guardan solo si se pudieron leer; si no, el mes queda sin departamentos.
    """
    info = report["store_info"]
    printed = report.get("printed_department_total") or {}
    conn = _connect()
    try:
        with conn:
            conn.execute("DELETE FROM monthly_reports WHERE year = ? AND month = ?", (report["year"], report["month"]))
            conn.execute("DELETE FROM monthly_report_departments WHERE year = ? AND month = ?", (report["year"], report["month"]))
            columns = ["year", "month", "from_date", "to_date", *STORE_INFO_FIELDS, "credit_terms_json",
                       "printed_department_amount", "printed_department_count", "warnings_json", "filename", "updated_at"]
            values = [report["year"], report["month"], report["from_date"], report["to_date"],
                      *(info.get(field) for field in STORE_INFO_FIELDS),
                      json.dumps(info.get("credit_terms")) if info.get("credit_terms") is not None else None,
                      printed.get("amount"), printed.get("count"), json.dumps(report.get("warnings") or []),
                      filename, datetime.now().isoformat(timespec="seconds")]
            conn.execute(
                f"INSERT INTO monthly_reports ({', '.join(columns)}) VALUES ({', '.join('?' for _ in columns)})", values,
            )
            for record in report.get("departments") or []:
                conn.execute(
                    "INSERT OR REPLACE INTO monthly_report_departments (year, month, department, count, amount) VALUES (?, ?, ?, ?, ?)",
                    (report["year"], report["month"], record["department"], record.get("count"), record.get("amount")),
                )
    finally:
        conn.close()


def get_report(year, month):
    """El reporte guardado del mes, o None. `departments` es [] si no se pudieron leer."""
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM monthly_reports WHERE year = ? AND month = ?", (year, month)).fetchone()
        if row is None:
            return None
        departments = conn.execute(
            "SELECT department, count, amount FROM monthly_report_departments WHERE year = ? AND month = ? ORDER BY amount DESC",
            (year, month),
        ).fetchall()
    finally:
        conn.close()
    store_info = {field: row[field] for field in STORE_INFO_FIELDS}
    store_info["credit_terms"] = json.loads(row["credit_terms_json"]) if row["credit_terms_json"] else None
    return {
        "year": row["year"],
        "month": row["month"],
        "from_date": row["from_date"],
        "to_date": row["to_date"],
        "store_info": store_info,
        "departments": [dict(d) for d in departments],
        "printed_department_total": (
            {"amount": row["printed_department_amount"], "count": row["printed_department_count"]}
            if row["printed_department_amount"] is not None else None
        ),
        "warnings": json.loads(row["warnings_json"]) if row["warnings_json"] else [],
        "filename": row["filename"],
        "updated_at": row["updated_at"],
    }


def delete_report(year, month):
    conn = _connect()
    try:
        with conn:
            conn.execute("DELETE FROM monthly_reports WHERE year = ? AND month = ?", (year, month))
            conn.execute("DELETE FROM monthly_report_departments WHERE year = ? AND month = ?", (year, month))
    finally:
        conn.close()
