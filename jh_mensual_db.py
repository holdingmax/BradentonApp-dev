"""
Persistencia de los reportes mensuales de J.H. Williams (pedido del usuario,
2026-10-06, chat 21): un reporte de cada tipo por mes (EFT History, Invoice
History, Credit Card Daily Summary), con sus renglones tal cual. Capa de
datos pura: la lectura del PDF y los cruces viven en jh_mensual.py.

reportes_data/jh_mensual.db (gitignored, como las demás bases).
"""

import json
import os
import sqlite3
from datetime import datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "jh_mensual.db")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS jh_monthly_reports (
            kind TEXT NOT NULL,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            from_date TEXT, to_date TEXT,
            rows_json TEXT, totals_json TEXT,
            filename TEXT, updated_at TEXT,
            PRIMARY KEY (kind, year, month)
        )
        """
    )
    return conn


# Qué identifica a cada renglón, qué es su importe (no se pisa nunca) y qué
# es un estado que un reporte más nuevo puede actualizar (una factura que se
# pagó, un EFT que pasó a Posted).
# TOTAL_FIELDS: lo que suma el total impreso de cada reporte, en su orden.
ROW_KEY = {"eft": "reference", "invoices": "invoice", "coupons": "date"}
_ROW_AMOUNTS = {"eft": ("date", "amount"), "invoices": ("date", "amount"), "coupons": ("coupons", "gross", "fees", "net")}
_ROW_STATUS = {"eft": ("status",), "invoices": ("balance",), "coupons": ()}
TOTAL_FIELDS = {"eft": ("amount",), "invoices": ("amount", "balance"), "coupons": ("gross", "fees", "net")}


def totals_for(kind, rows):
    return {field: round(sum(r.get(field) or 0.0 for r in rows), 2) for field in TOTAL_FIELDS[kind]}


def merge_report(kind, month_report, filename=None):
    """
    Suma al reporte guardado de ese tipo y mes los renglones nuevos de
    `month_report` (un mes de jh_mensual.extract_report). Un renglón ya
    cargado se deja como está (si el reporte nuevo trae otro importe, se
    avisa en `conflicts`); solo se actualiza su estado (saldo de la factura,
    estado del EFT). Devuelve {"added", "updated", "conflicts": [clave]}.
    """
    key = ROW_KEY[kind]
    year, month = month_report["year"], month_report["month"]
    stored = get_reports(year, month).get(kind)
    rows = {r[key]: dict(r) for r in (stored["rows"] if stored else [])}
    added = updated = 0
    conflicts = []
    for r in month_report["rows"]:
        old = rows.get(r[key])
        if old is None:
            rows[r[key]] = dict(r)
            added += 1
        elif any(old.get(f) != r.get(f) for f in _ROW_AMOUNTS[kind]):
            conflicts.append(r[key])
        elif any(old.get(f) != r.get(f) for f in _ROW_STATUS[kind]):
            old.update({f: r.get(f) for f in _ROW_STATUS[kind]})
            updated += 1
    merged = sorted(rows.values(), key=lambda r: (r["date"], r[key]))
    from_date = min(filter(None, [month_report["from_date"], stored["from_date"] if stored else None]))
    to_date = max(filter(None, [month_report["to_date"], stored["to_date"] if stored else None]))
    if added or updated or not stored:
        conn = _connect()
        try:
            with conn:
                conn.execute(
                    "INSERT OR REPLACE INTO jh_monthly_reports "
                    "(kind, year, month, from_date, to_date, rows_json, totals_json, filename, updated_at) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (kind, year, month, from_date, to_date, json.dumps(merged), json.dumps(totals_for(kind, merged)),
                     filename, datetime.now().isoformat(timespec="seconds")),
                )
        finally:
            conn.close()
    return {"added": added, "updated": updated, "conflicts": conflicts}


def _row_to_report(row):
    return {
        "kind": row["kind"],
        "year": row["year"],
        "month": row["month"],
        "from_date": row["from_date"],
        "to_date": row["to_date"],
        "rows": json.loads(row["rows_json"] or "[]"),
        "totals": json.loads(row["totals_json"] or "{}"),
        "filename": row["filename"],
        "updated_at": row["updated_at"],
    }


def get_reports(year, month):
    """{tipo: reporte} de los reportes guardados del mes (los que falten no están)."""
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM jh_monthly_reports WHERE year = ? AND month = ?", (year, month)).fetchall()
    finally:
        conn.close()
    return {row["kind"]: _row_to_report(row) for row in rows}


def get_all_invoice_numbers():
    """Todas las facturas de los Invoice History guardados, de cualquier mes."""
    conn = _connect()
    try:
        rows = conn.execute("SELECT rows_json FROM jh_monthly_reports WHERE kind = 'invoices'").fetchall()
    finally:
        conn.close()
    return {r["invoice"] for row in rows for r in json.loads(row["rows_json"] or "[]")}


def get_rows(kind):
    """Todos los renglones guardados de un tipo de reporte, de cualquier mes."""
    conn = _connect()
    try:
        rows = conn.execute("SELECT rows_json FROM jh_monthly_reports WHERE kind = ?", (kind,)).fetchall()
    finally:
        conn.close()
    return [r for row in rows for r in json.loads(row["rows_json"] or "[]")]


def get_coupon_rows():
    """Todas las filas guardadas de los Credit Card Daily Summary, de cualquier mes: para identificar el detalle de cupones."""
    return get_rows("coupons")


def delete_report(kind, year, month):
    conn = _connect()
    try:
        with conn:
            conn.execute("DELETE FROM jh_monthly_reports WHERE kind = ? AND year = ? AND month = ?", (kind, year, month))
    finally:
        conn.close()
