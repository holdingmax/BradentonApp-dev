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
    # Edición a mano (pedido del usuario, 2026-10-07): cuándo se corrigió.
    columns = {row[1] for row in conn.execute("PRAGMA table_info(monthly_reports)")}
    if "edited_at" not in columns:
        conn.execute("ALTER TABLE monthly_reports ADD COLUMN edited_at TEXT")
    # Completado con la hoja pedida de nuevo (pedido del usuario, 2026-10-07).
    if "sheet_added_at" not in columns:
        conn.execute("ALTER TABLE monthly_reports ADD COLUMN sheet_added_at TEXT")
    return conn


def update_report(year, month, store_info, departments, printed_amount, printed_count, warnings, from_sheet=False):
    """
    Corrección a mano del reporte guardado (pedido del usuario, 2026-10-07):
    Store Info (`store_info` con las claves de STORE_INFO_FIELDS y
    "credit_terms"), departamentos [{department, count, amount}], total
    impreso de departamentos y los avisos ya recalculados. Devuelve False si
    el mes no tiene reporte. from_sheet=True: lo completó la hoja pedida de
    nuevo (sheet_added_at), no una edición a mano (edited_at).
    """
    conn = _connect()
    try:
        with conn:
            exists = conn.execute(
                "SELECT 1 FROM monthly_reports WHERE year = ? AND month = ?", (year, month)
            ).fetchone()
            if exists is None:
                return False
            assignments = ", ".join(f"{field} = ?" for field in STORE_INFO_FIELDS)
            credit_terms = store_info.get("credit_terms")
            conn.execute(
                f"UPDATE monthly_reports SET {assignments}, credit_terms_json = ?, printed_department_amount = ?, "
                "printed_department_count = ?, warnings_json = ?, "
                f"{'sheet_added_at' if from_sheet else 'edited_at'} = ? WHERE year = ? AND month = ?",
                [*(store_info.get(field) for field in STORE_INFO_FIELDS),
                 json.dumps(credit_terms) if credit_terms is not None else None,
                 printed_amount, printed_count, json.dumps(warnings or []),
                 datetime.now().isoformat(timespec="seconds"), year, month],
            )
            conn.execute("DELETE FROM monthly_report_departments WHERE year = ? AND month = ?", (year, month))
            for record in departments:
                conn.execute(
                    "INSERT OR REPLACE INTO monthly_report_departments (year, month, department, count, amount) VALUES (?, ?, ?, ?, ?)",
                    (year, month, record["department"], record.get("count"), record.get("amount")),
                )
        return True
    finally:
        conn.close()


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
        "edited_at": row["edited_at"],
        "sheet_added_at": row["sheet_added_at"],
    }


def list_months():
    """[(year, month)] de los reportes guardados, del más nuevo al más viejo."""
    conn = _connect()
    try:
        return [(r["year"], r["month"]) for r in
                conn.execute("SELECT year, month FROM monthly_reports ORDER BY year DESC, month DESC")]
    finally:
        conn.close()


def delete_report(year, month):
    conn = _connect()
    try:
        with conn:
            conn.execute("DELETE FROM monthly_reports WHERE year = ? AND month = ?", (year, month))
            conn.execute("DELETE FROM monthly_report_departments WHERE year = ? AND month = ?", (year, month))
    finally:
        conn.close()
