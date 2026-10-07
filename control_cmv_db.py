"""
Persistencia del Control CMV (pedido del usuario, 2026-10-07): los reportes
de ventas de Elistar ("Depts Report" o "P & L Report"), uno de cada tipo por
mes, guardados ya leídos (control_cmv.read_elistar_report). Subir otro del
mismo mes y tipo lo reemplaza.

reportes_data/control_cmv.db (gitignored, como las demás bases).
"""

import json
import os
import sqlite3
from datetime import datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "control_cmv.db")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS elistar_reports (
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            kind TEXT NOT NULL,
            data_json TEXT NOT NULL,
            filename TEXT,
            uploaded_at TEXT,
            PRIMARY KEY (year, month, kind)
        )
        """
    )
    return conn


def save_report(report, filename=None):
    conn = _connect()
    try:
        with conn:
            conn.execute(
                "INSERT OR REPLACE INTO elistar_reports (year, month, kind, data_json, filename, uploaded_at) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (report["year"], report["month"], report["kind"], json.dumps(report), filename,
                 datetime.now().isoformat(timespec="seconds")),
            )
    finally:
        conn.close()


def get_month(year, month):
    """{kind: reporte} del mes ("depts" y/o "pl"), con filename y uploaded_at."""
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM elistar_reports WHERE year = ? AND month = ?", (year, month)).fetchall()
    finally:
        conn.close()
    return {row["kind"]: dict(json.loads(row["data_json"]), filename=row["filename"], uploaded_at=row["uploaded_at"])
            for row in rows}


def delete_report(year, month, kind):
    conn = _connect()
    try:
        with conn:
            conn.execute("DELETE FROM elistar_reports WHERE year = ? AND month = ? AND kind = ?", (year, month, kind))
    finally:
        conn.close()
