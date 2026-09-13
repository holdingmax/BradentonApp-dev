"""
CMV -- Carga de Datos: guarda Costo (por UPC/departamento) y Ventas
mensuales (por UPC/departamento/mes), sin escribir ningún Excel -- mismo
patrón que chase_db.py/reportes_db.py.

Costo (cmv_costs): reemplaza los departamentos presentes en cada carga
(preserva los demás) -- mismo criterio que `_merge_with_existing_
departments` (cmv_costo.py) para el Excel real: subir un solo
departamento no borra los otros 23. Un cambio de precio contra lo ya
guardado queda anotado en cmv_price_changes (mismo espíritu que la
columna "Price Change" del Excel real).

Ventas (cmv_monthly_sales): reemplaza TODO un (año, mes, departamento) en
cada carga -- mismo criterio que la hoja real ("CMV Ventas reemplaza toda
la hoja del departamento en cada carga, no acumula, no hace merge por
UPC" -- ver CLAUDE.md), para no arrastrar filas viejas que ya no venden.
El margen se calcula en el momento contra cmv_costs (por UPC, con
fallback por nombre -- mismo criterio que la fórmula real de la columna
G, `_cost_lookup_formula_with_name_fallback` en monthly_sales.py), nunca
guardado como valor fijo.
"""

import os
import sqlite3
from datetime import datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "cmv.db")


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
        CREATE TABLE IF NOT EXISTS cmv_costs (
            upc TEXT PRIMARY KEY,
            upc_mod TEXT,
            name TEXT,
            dept_name TEXT,
            dept_id TEXT,
            cost REAL,
            price REAL,
            updated_at TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cmv_price_changes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            upc TEXT NOT NULL,
            name TEXT,
            dept_name TEXT,
            old_price REAL,
            new_price REAL,
            changed_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cmv_monthly_sales (
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            dept_name TEXT NOT NULL,
            upc TEXT NOT NULL,
            name TEXT,
            count INTEGER,
            amount REAL,
            PRIMARY KEY (year, month, dept_name, upc)
        )
        """
    )
    return conn


def _now():
    return datetime.now().isoformat(timespec="seconds")


def replace_costs_for_departments(records):
    """
    `records`: lista de dicts {upc, upc_mod, name, cost, price, dept_id,
    dept_name} (ya limpios -- salida de cmv_costo._consolidate_department_
    files, una fila por df.to_dict('records')). Agrupa por dept_name,
    borra lo que ya había de CADA departamento presente, e inserta lo
    nuevo -- los departamentos NO incluidos en esta carga quedan intactos.
    """
    by_dept = {}
    for row in records:
        dept = (row.get("dept_name") or row.get("DeptName") or "").strip()
        if not dept:
            continue
        by_dept.setdefault(dept, []).append(row)

    if not by_dept:
        return {"departments": 0, "rows": 0, "price_changes": 0}

    now = _now()
    price_changes = 0
    conn = _connect()
    try:
        existing_prices = {
            row["upc"]: row["price"]
            for row in conn.execute("SELECT upc, price FROM cmv_costs").fetchall()
        }
        for dept, rows in by_dept.items():
            conn.execute(
                "DELETE FROM cmv_costs WHERE dept_name = ? COLLATE NOCASE", (dept,)
            )
            for row in rows:
                upc = str(row.get("upc") or row.get("UPC") or "").strip()
                if not upc:
                    continue
                new_price = row.get("price") if row.get("price") is not None else row.get("Price")
                old_price = existing_prices.get(upc)
                if old_price is not None and new_price is not None and round(float(old_price), 2) != round(float(new_price), 2):
                    conn.execute(
                        """
                        INSERT INTO cmv_price_changes (upc, name, dept_name, old_price, new_price, changed_at)
                        VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (upc, row.get("name") or row.get("Name"), dept, old_price, new_price, now),
                    )
                    price_changes += 1
                conn.execute(
                    """
                    INSERT INTO cmv_costs (upc, upc_mod, name, dept_name, dept_id, cost, price, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(upc) DO UPDATE SET
                        upc_mod = excluded.upc_mod,
                        name = excluded.name,
                        dept_name = excluded.dept_name,
                        dept_id = excluded.dept_id,
                        cost = excluded.cost,
                        price = excluded.price,
                        updated_at = excluded.updated_at
                    """,
                    (
                        upc,
                        row.get("upc_mod") or row.get("UPCMod"),
                        row.get("name") or row.get("Name"),
                        dept,
                        row.get("dept_id") or row.get("DeptID"),
                        row.get("cost") if row.get("cost") is not None else row.get("Cost"),
                        new_price,
                        now,
                    ),
                )
        conn.commit()
    finally:
        conn.close()

    return {
        "departments": len(by_dept),
        "rows": sum(len(rows) for rows in by_dept.values()),
        "price_changes": price_changes,
    }


def list_departments():
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT dept_name, COUNT(*) AS n FROM cmv_costs GROUP BY dept_name ORDER BY dept_name"
        ).fetchall()
    finally:
        conn.close()
    return [{"dept_name": row["dept_name"], "count": row["n"]} for row in rows]


def get_costs_by_department(dept_name):
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM cmv_costs WHERE dept_name = ? COLLATE NOCASE ORDER BY name",
            (dept_name,),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def get_recent_price_changes(limit=30):
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM cmv_price_changes ORDER BY changed_at DESC, id DESC LIMIT ?",
            (limit,),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def _cost_lookup_index(conn):
    by_upc = {}
    by_name = {}
    for row in conn.execute("SELECT upc, name, cost FROM cmv_costs").fetchall():
        if row["cost"] is None:
            continue
        by_upc[row["upc"]] = row["cost"]
        name_key = (row["name"] or "").strip().upper()
        if name_key:
            by_name.setdefault(name_key, row["cost"])
    return by_upc, by_name


def replace_month_department_sales(year, month, dept_name, rows):
    """
    `rows`: lista de dicts {upc, name, count, amount}. Reemplaza TODO lo
    que había para (year, month, dept_name) -- mismo criterio que la hoja
    real, nunca acumula ni hace merge por UPC.
    """
    conn = _connect()
    try:
        conn.execute(
            "DELETE FROM cmv_monthly_sales WHERE year = ? AND month = ? AND dept_name = ? COLLATE NOCASE",
            (year, month, dept_name),
        )
        for row in rows:
            upc = str(row.get("upc") or "").strip()
            if not upc:
                continue
            conn.execute(
                """
                INSERT INTO cmv_monthly_sales (year, month, dept_name, upc, name, count, amount)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (year, month, dept_name, upc, row.get("name"), row.get("count") or 0, row.get("amount") or 0.0),
            )
        conn.commit()
    finally:
        conn.close()


def get_month_department_totals(year, month):
    """
    Un renglón por departamento con Count/Amount/Costo/Margen del mes --
    el costo se resuelve en el momento por UPC (con fallback por nombre,
    mismo criterio que la fórmula real) contra cmv_costs, nunca guardado.
    """
    conn = _connect()
    try:
        by_upc, by_name = _cost_lookup_index(conn)
        rows = conn.execute(
            "SELECT dept_name, upc, name, count, amount FROM cmv_monthly_sales WHERE year = ? AND month = ?",
            (year, month),
        ).fetchall()
    finally:
        conn.close()

    totals = {}
    for row in rows:
        dept = row["dept_name"]
        bucket = totals.setdefault(
            dept, {"dept_name": dept, "count": 0, "amount": 0.0, "cost_total": 0.0, "cost_known": True}
        )
        bucket["count"] += row["count"] or 0
        bucket["amount"] += row["amount"] or 0.0
        cost = by_upc.get(row["upc"])
        if cost is None:
            cost = by_name.get((row["name"] or "").strip().upper())
        if cost is None:
            bucket["cost_known"] = False
        else:
            bucket["cost_total"] += cost * (row["count"] or 0)

    result = []
    for dept, bucket in sorted(totals.items()):
        margin = bucket["amount"] - bucket["cost_total"] if bucket["cost_known"] else None
        margin_pct = (margin / bucket["amount"] * 100) if margin is not None and bucket["amount"] else None
        result.append(
            {
                "dept_name": dept,
                "count": bucket["count"],
                "amount": round(bucket["amount"], 2),
                "cost_total": round(bucket["cost_total"], 2) if bucket["cost_known"] else None,
                "margin": round(margin, 2) if margin is not None else None,
                "margin_pct": round(margin_pct, 1) if margin_pct is not None else None,
            }
        )
    return result


def get_month_years_with_data():
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT DISTINCT year, month FROM cmv_monthly_sales ORDER BY year DESC, month DESC"
        ).fetchall()
    finally:
        conn.close()
    return [(row["year"], row["month"]) for row in rows]
