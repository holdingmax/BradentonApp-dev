"""
Persistencia del Control CMV (pedido del usuario, 2026-10-07): los reportes
de ventas de Elistar ("Depts Report" o "P & L Report"), uno de cada tipo por
mes, guardados ya leídos (control_cmv.read_elistar_report). Subir otro del
mismo mes y tipo lo reemplaza.

reportes_data/control_cmv.db (gitignored, como las demás bases).
"""

import json
import app_paths
import os
from datetime import datetime

import db

_BASE_DIR = app_paths.DATA_DIR
_DB_PATH = os.path.join(_BASE_DIR, "control_cmv.db")


def _init_sqlite(conn):
    conn.execute("PRAGMA journal_mode=WAL")
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


def _connect():
    # Sin DATABASE_URL, la base SQLite de siempre; con ella, el esquema control_cmv de
    # PostgreSQL (db.py; tablas de migraciones/versions/0005_control_cmv.py).
    return db.connect("control_cmv", _DB_PATH, _init_sqlite)


def save_report(report, filename=None):
    conn = _connect()
    try:
        with conn:
            conn.execute(
                "INSERT INTO elistar_reports (year, month, kind, data_json, filename, uploaded_at) "
                "VALUES (?, ?, ?, ?, ?, ?) ON CONFLICT (year, month, kind) DO UPDATE SET "
                "data_json = excluded.data_json, filename = excluded.filename, uploaded_at = excluded.uploaded_at",
                (report["year"], report["month"], report["kind"], json.dumps(report), filename,
                 datetime.now().isoformat(timespec="seconds")),
            )
    finally:
        conn.close()


def get_month(year, month):
    """{kind: reporte} del mes ("depts" y/o "pl"), con filename y uploaded_at."""
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM elistar_reports WHERE year = ? AND month = ? ORDER BY kind", (year, month)).fetchall()
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
