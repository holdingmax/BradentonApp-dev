"""
Gettel / Toyota -- Carga de Datos: guarda los totales diarios de cupones de
combustible de las dos empresas (Monto + Galones), sin escribir ningún
Excel -- mismo patrón que chase_db.py/lottery_db.py.

Los datos ya se leen con funciones puras existentes (`gettel_toyota_parser.
_summarize_origin_workbook`/`summarize_pdf_report`), pensadas originalmente
para fusionarse en la hoja "Gettel-Toyota MM.YYYY" del Excel Cierre -- acá
se guardan directo en base, un renglón por día con las 4 columnas (Gettel
Monto/Galones, Toyota Monto/Galones) que esa hoja ya tenía.

Pedido explícito del usuario (2026-09-12, cuarta tanda): "deberias
considerarlas en ponerlos en el cuadro de ventas" -- el monto de Gettel de
acá es lo que alimenta la categoría "Gettel" de Ventas por Departamento
(ver reporte_diario.group_department_sales, parámetro gettel_amount), que
antes siempre quedaba en $0 porque esa venta nunca sale del PDF de cierre
diario -- viene de estos reportes de cupones.
"""

import os
import sqlite3
from datetime import date, datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "gettel_toyota.db")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    # timeout=30 + WAL -- ver reportes_db.py: necesario desde que las cargas
    # en segundo plano (jobs.py) pueden escribir de verdad en paralelo.
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS gettel_toyota_days (
            date TEXT PRIMARY KEY,
            gettel_amount REAL,
            gettel_gallons REAL,
            gettel_source TEXT,
            toyota_amount REAL,
            toyota_gallons REAL,
            toyota_source TEXT,
            updated_at TEXT
        )
        """
    )
    return conn


def _day_key(value):
    if isinstance(value, str):
        return value
    if isinstance(value, (date, datetime)):
        return value.isoformat() if isinstance(value, date) and not isinstance(value, datetime) else value.date().isoformat()
    return str(value)


def upsert_vendor_totals(vendor, totals_by_date, source="excel"):
    """
    `vendor` es "gettel" o "toyota". `totals_by_date` es el dict {date:
    {"amount","gallons"}} que ya devuelven las funciones de extracción de
    gettel_toyota_parser -- se pisa (reemplaza) el día completo de ese
    vendor, nunca se acumula (mismo criterio que reemplazar un día entero
    de Reporte Diario al volver a subir el mismo PDF).
    """
    if vendor not in ("gettel", "toyota"):
        raise ValueError(f"Vendor desconocido: {vendor}")
    if not totals_by_date:
        return 0

    amount_col = f"{vendor}_amount"
    gallons_col = f"{vendor}_gallons"
    source_col = f"{vendor}_source"
    now = datetime.now().isoformat(timespec="seconds")

    conn = _connect()
    try:
        for day, totals in totals_by_date.items():
            key = _day_key(day)
            conn.execute(
                f"""
                INSERT INTO gettel_toyota_days (date, {amount_col}, {gallons_col}, {source_col}, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(date) DO UPDATE SET
                    {amount_col} = excluded.{amount_col},
                    {gallons_col} = excluded.{gallons_col},
                    {source_col} = excluded.{source_col},
                    updated_at = excluded.updated_at
                """,
                (key, float(totals.get("amount") or 0.0), float(totals.get("gallons") or 0.0), source, now),
            )
        conn.commit()
    finally:
        conn.close()
    return len(totals_by_date)


def get_day(report_date):
    key = _day_key(report_date)
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM gettel_toyota_days WHERE date = ?", (key,)).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None


def get_month_days(year, month):
    """Un renglón por día del mes que tenga algo cargado (Gettel y/o Toyota) -- no rellena los días vacíos, a diferencia de reporte_historial (acá no hay "carga a mano" pendiente que mostrar, es puramente lo ya subido)."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM gettel_toyota_days WHERE date LIKE ? ORDER BY date",
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def get_month_gettel_amount(year, month):
    """Suma del Monto de Gettel del mes -- lo que alimenta la categoría "Gettel" del resumen mensual de Ventas por Departamento."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT SUM(gettel_amount) AS total FROM gettel_toyota_days WHERE date LIKE ?",
            (f"{prefix}%",),
        ).fetchone()
    finally:
        conn.close()
    return row["total"] or 0.0


def get_day_gettel_amount(report_date):
    """Monto de Gettel de un solo día -- lo que alimenta la categoría "Gettel" del resumen por categoría de un día puntual."""
    day = get_day(report_date)
    return (day.get("gettel_amount") if day else None) or 0.0
