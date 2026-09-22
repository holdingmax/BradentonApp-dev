"""
Proveedores -- Carga de Datos: guarda las facturas ya extraídas de cada PDF
(fecha/N°/monto/proveedor) sin escribir ningún Excel Ledger -- mismo patrón
que chase_db.py/eft_db.py. Primer paso de este módulo (2026-09-14, pedido
explícito del usuario: "quiero que empieces con el modulo de proveedores y
donde pueda guardar sus facturas") -- reusa el motor de detección/extracción
de 32 proveedores + el dinámico tal cual está en proveedores.py
(`extract_invoices_from_pdf`), sin tocar ni un carácter de esa lógica; acá
solo se guarda el resultado.

reportes_data/proveedores.db (mismo directorio gitignored de siempre):
- supplier_invoices: una fila por factura, clave natural (supplier_key,
  invoice_no) para el mismo criterio de duplicado que ya usa el Ledger real
  ("un mismo N° de factura ya cargado se omite").
- supplier_hidden (2026-09-22, pedido explícito del usuario: "quiero que se
  pueda ocultar a los proveedores que no se quiera ver en la pagina"): un
  supplier_key por fila mientras esté oculto de la grilla de "Guardado" --
  ocultar/desocultar no borra ni toca ninguna factura/pago, solo cambia si
  el módulo aparece en la grilla.

El PDF original se guarda aparte con documents_db.py (módulo "proveedores"),
igual que EFT/Gettel/CMV -- no se duplica esa pieza acá.
"""

import os
import sqlite3
from datetime import date, datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "proveedores.db")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    # timeout=30 + WAL -- ver reportes_db.py: necesario desde que las cargas
    # en segundo plano (jobs.py) pueden escribir de verdad en paralelo.
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS supplier_invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_key TEXT NOT NULL,
            supplier_label TEXT NOT NULL,
            invoice_date TEXT NOT NULL,
            invoice_no TEXT NOT NULL,
            amount REAL NOT NULL,
            source_filename TEXT,
            uploaded_at TEXT,
            UNIQUE (supplier_key, invoice_no)
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_supplier_invoices_period ON supplier_invoices (invoice_date)"
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS supplier_hidden (
            supplier_key TEXT PRIMARY KEY,
            hidden_at TEXT
        )
        """
    )
    # Credit memos (2026-09-22, pedido explícito del usuario -- "en H.T. se
    # pueden cargar credits memo que disminuyen lo que hay que pagar de las
    # facturas"): una nota de crédito reduce el saldo igual que un pago
    # (Haber en el cuadro de cuenta corriente, ver webapp.py:
    # _build_supplier_ledger) pero NO es un pago de Chase -- se guarda
    # aparte para no mezclarla con chase_db. Genérica (cualquier
    # supplier_key), aunque el pedido puntual era para H.T. Hackney -- la
    # UI de carga queda scopeada a ese proveedor en el template.
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS supplier_credit_memos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_key TEXT NOT NULL,
            supplier_label TEXT NOT NULL,
            credit_date TEXT NOT NULL,
            credit_no TEXT,
            amount REAL NOT NULL,
            note TEXT,
            uploaded_at TEXT
        )
        """
    )
    # Pagos a mano con Caja (2026-09-22, pedido explícito del usuario --
    # "que se puedan hacer cargas manuales de pago para proveedores como
    # Bimbo, Flori gas, Sam's... deben ir conectados a caja y sumarse en la
    # columna de gastos el dia que fueran cargados"). `caja_expense_item_id`
    # guarda el id del gasto ya sumado en caja_db.caja_expense_items -- así
    # borrar este pago puede borrar también ese gasto reflejado en Caja.
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS supplier_manual_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_key TEXT NOT NULL,
            supplier_label TEXT NOT NULL,
            payment_date TEXT NOT NULL,
            amount REAL NOT NULL,
            note TEXT,
            caja_expense_item_id INTEGER,
            uploaded_at TEXT
        )
        """
    )
    # Configuración individual por proveedor (2026-09-22, pedido explícito
    # del usuario: "cada proveedor tenga un tipo de configuracion
    # individual en la que tocando un boton se pueda agregar de que se le
    # hacen pagos en efectivo o recibe credits memo") -- reemplaza el
    # gateo hardcodeado de la ronda anterior (pago a mano para cualquiera,
    # credit memo solo para ht_hackney): ahora el usuario prende/apaga
    # cada capacidad por proveedor desde la propia página, sin tocar
    # código. Sin fila = las dos capacidades apagadas (default).
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS supplier_settings (
            supplier_key TEXT PRIMARY KEY,
            allow_manual_payments INTEGER NOT NULL DEFAULT 0,
            allow_credit_memos INTEGER NOT NULL DEFAULT 0
        )
        """
    )
    # Semilla de una sola vez -- los proveedores que el usuario ya había
    # pedido explícitamente (Bimbo/Flori-Gas/Sam's para pago a mano,
    # H.T. Hackney para credit memos) arrancan habilitados; `INSERT OR
    # IGNORE` hace que esto no pise nunca una configuración que el usuario
    # ya haya cambiado a mano después.
    for _key in ("bimbo", "flori_gas", "sams_club"):
        conn.execute(
            "INSERT OR IGNORE INTO supplier_settings (supplier_key, allow_manual_payments, allow_credit_memos) VALUES (?, 1, 0)",
            (_key,),
        )
    for _key in ("ht_hackney",):
        conn.execute(
            "INSERT OR IGNORE INTO supplier_settings (supplier_key, allow_manual_payments, allow_credit_memos) VALUES (?, 0, 1)",
            (_key,),
        )
    conn.commit()
    return conn


def _date_key(value):
    if isinstance(value, str):
        return value
    if isinstance(value, (date, datetime)):
        return value.isoformat() if isinstance(value, date) and not isinstance(value, datetime) else value.date().isoformat()
    return str(value)


def invoice_exists(supplier_key, invoice_no):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT 1 FROM supplier_invoices WHERE supplier_key = ? AND invoice_no = ?",
            (supplier_key, invoice_no),
        ).fetchone()
    finally:
        conn.close()
    return row is not None


def save_invoice(supplier_key, supplier_label, invoice_date, invoice_no, amount, source_filename=None):
    """
    Guarda una factura -- devuelve False sin tocar nada si ese N° de
    factura ya estaba cargado para este proveedor (mismo criterio de
    duplicado que el Ledger real: por N° de factura, no por fecha/monto).
    """
    conn = _connect()
    try:
        try:
            conn.execute(
                """
                INSERT INTO supplier_invoices
                    (supplier_key, supplier_label, invoice_date, invoice_no, amount, source_filename, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    supplier_key, supplier_label, _date_key(invoice_date), str(invoice_no),
                    float(amount), source_filename, datetime.now().isoformat(timespec="seconds"),
                ),
            )
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
    finally:
        conn.close()


def get_available_years():
    """Años distintos con al menos una factura guardada -- selector de mes/año del reporte de Reportes."""
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT DISTINCT substr(invoice_date, 1, 4) AS y FROM supplier_invoices ORDER BY y"
        ).fetchall()
        return [int(row["y"]) for row in rows if row["y"]]
    finally:
        conn.close()


def get_month_invoices(year, month):
    """Todas las facturas de un mes (por fecha de factura), agrupadas por proveedor -- más reciente primero dentro de cada uno."""
    prefix = f"{year:04d}-{month:02d}-"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM supplier_invoices WHERE invoice_date LIKE ? ORDER BY supplier_label, invoice_date DESC",
            (f"{prefix}%",),
        ).fetchall()
    finally:
        conn.close()

    by_supplier = {}
    for row in rows:
        item = dict(row)
        by_supplier.setdefault(item["supplier_key"], {"label": item["supplier_label"], "invoices": [], "total": 0.0})
        by_supplier[item["supplier_key"]]["invoices"].append(item)
        by_supplier[item["supplier_key"]]["total"] += item["amount"]
    groups = sorted(by_supplier.values(), key=lambda g: g["label"])
    for group in groups:
        group["total"] = round(group["total"], 2)
    return groups


def list_suppliers_with_counts():
    """Un renglón por proveedor con al menos una factura guardada -- para el selector de "ver historial completo de este proveedor"."""
    conn = _connect()
    try:
        rows = conn.execute(
            """
            SELECT supplier_key, supplier_label, COUNT(*) AS n, SUM(amount) AS total
            FROM supplier_invoices GROUP BY supplier_key ORDER BY supplier_label
            """
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def get_supplier_invoices(supplier_key):
    """Historial completo (todos los meses) de un proveedor puntual, más reciente primero."""
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM supplier_invoices WHERE supplier_key = ? ORDER BY invoice_date DESC",
            (supplier_key,),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def delete_invoice(invoice_id):
    conn = _connect()
    try:
        cur = conn.execute("DELETE FROM supplier_invoices WHERE id = ?", (invoice_id,))
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def set_supplier_hidden(supplier_key, hidden):
    """Oculta/desoculta un proveedor de la grilla de "Guardado" -- nunca toca ninguna factura/pago."""
    conn = _connect()
    try:
        if hidden:
            conn.execute(
                "INSERT OR IGNORE INTO supplier_hidden (supplier_key, hidden_at) VALUES (?, ?)",
                (supplier_key, datetime.now().isoformat(timespec="seconds")),
            )
        else:
            conn.execute("DELETE FROM supplier_hidden WHERE supplier_key = ?", (supplier_key,))
        conn.commit()
    finally:
        conn.close()


def list_hidden_suppliers():
    """{"supplier_key", ...} -- todos los proveedores ocultos ahora mismo."""
    conn = _connect()
    try:
        rows = conn.execute("SELECT supplier_key FROM supplier_hidden").fetchall()
    finally:
        conn.close()
    return {row["supplier_key"] for row in rows}


def get_supplier_settings(supplier_key):
    """{"allow_manual_payments", "allow_credit_memos"} -- False/False si el proveedor no tiene fila todavía."""
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT allow_manual_payments, allow_credit_memos FROM supplier_settings WHERE supplier_key = ?",
            (supplier_key,),
        ).fetchone()
    finally:
        conn.close()
    if row is None:
        return {"allow_manual_payments": False, "allow_credit_memos": False}
    return {"allow_manual_payments": bool(row["allow_manual_payments"]), "allow_credit_memos": bool(row["allow_credit_memos"])}


def set_supplier_settings(supplier_key, allow_manual_payments, allow_credit_memos):
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO supplier_settings (supplier_key, allow_manual_payments, allow_credit_memos)
            VALUES (?, ?, ?)
            ON CONFLICT (supplier_key) DO UPDATE SET
                allow_manual_payments = excluded.allow_manual_payments,
                allow_credit_memos = excluded.allow_credit_memos
            """,
            (supplier_key, 1 if allow_manual_payments else 0, 1 if allow_credit_memos else 0),
        )
        conn.commit()
    finally:
        conn.close()


def save_credit_memo(supplier_key, supplier_label, credit_date, credit_no, amount, note=None):
    conn = _connect()
    try:
        cur = conn.execute(
            """
            INSERT INTO supplier_credit_memos
                (supplier_key, supplier_label, credit_date, credit_no, amount, note, uploaded_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                supplier_key, supplier_label, _date_key(credit_date), credit_no, float(amount),
                note, datetime.now().isoformat(timespec="seconds"),
            ),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def get_supplier_credit_memos(supplier_key):
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM supplier_credit_memos WHERE supplier_key = ? ORDER BY credit_date DESC",
            (supplier_key,),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def delete_credit_memo(credit_id):
    conn = _connect()
    try:
        cur = conn.execute("DELETE FROM supplier_credit_memos WHERE id = ?", (credit_id,))
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def save_manual_payment(supplier_key, supplier_label, payment_date, amount, note=None, caja_expense_item_id=None):
    conn = _connect()
    try:
        cur = conn.execute(
            """
            INSERT INTO supplier_manual_payments
                (supplier_key, supplier_label, payment_date, amount, note, caja_expense_item_id, uploaded_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                supplier_key, supplier_label, _date_key(payment_date), float(amount),
                note, caja_expense_item_id, datetime.now().isoformat(timespec="seconds"),
            ),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def get_supplier_manual_payments(supplier_key):
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM supplier_manual_payments WHERE supplier_key = ? ORDER BY payment_date DESC",
            (supplier_key,),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def delete_manual_payment(payment_id):
    """Borra el pago y devuelve la fila borrada (con caja_expense_item_id) -- el caller usa ese id para borrar también el gasto reflejado en Caja."""
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM supplier_manual_payments WHERE id = ?", (payment_id,)).fetchone()
        if row is None:
            return None
        conn.execute("DELETE FROM supplier_manual_payments WHERE id = ?", (payment_id,))
        conn.commit()
        return dict(row)
    finally:
        conn.close()


def _fmt_money_pdf(value):
    """Mismo criterio de signo que chase_rules._fmt_money_pdf -- "$" adelante, negativo con "-" antes del "$"."""
    if value is None:
        return "—"
    if value < 0:
        return "-${:,.2f}".format(abs(value))
    return "${:,.2f}".format(value)


def build_proveedores_pdf_report(year, month, dest_path):
    """
    PDF del módulo "Reportes" (2026-09-22, pedido explícito del usuario:
    "tambien quiero que se agregue lo que teniamos en el excel que era un
    resumen mensual en reportes, donde se va a mostrar la cantidad de
    facturas que llegaron de un proveedor y cual fue el total del mes") --
    mismo espíritu que la vieja hoja "RESUMEN COMPRAS" del Excel Ledger,
    ahora sacado de lo ya guardado en supplier_invoices en vez de sumar
    fórmulas de Excel. Una fila por proveedor (N° de facturas + total del
    mes) + fila TOTAL en negrita al pie -- mismo patrón que
    chase_rules.build_chase_pdf_report/eft_db.build_eft_pdf_report.
    """
    from pdf_export import build_simple_table_pdf

    groups = get_month_invoices(year, month)
    table_rows = []
    grand_total = 0.0
    grand_count = 0
    for group in groups:
        count = len(group["invoices"])
        total = group["total"]
        grand_total += total
        grand_count += count
        table_rows.append([group["label"], str(count), _fmt_money_pdf(total)])
    table_rows.append(["TOTAL", str(grand_count), _fmt_money_pdf(grand_total)])

    if groups:
        period_label = f"Período: {month:02d}/{year}"
    else:
        period_label = f"Período: sin facturas cargadas todavía en {month:02d}/{year}"

    title = f"Proveedores — Resumen de compras — {month:02d}/{year}"
    build_simple_table_pdf(
        dest_path,
        title,
        ["Proveedor", "Facturas", "Total"],
        table_rows,
        col_widths_mm=[180, 40, 60],
        bold_last_row=True,
        company_header=True,
        period_label=period_label,
    )
    return dest_path
