"""
Persistencia de EFT (Cta Cte J.H.Williams) y Cupones guardados dentro de la
propia página (ver CLAUDE.md -> "Conversión de EFT y Cupones a Carga de
Datos"). Capa de datos pura, sin ningún conocimiento de PDF/Excel --
eft_cta_cte.py/cupones_append.py hacen la lectura, este módulo solo guarda/
consulta y resuelve el cruce EFT<->Cupón (lo que en el Excel vivía como la
fórmula de la columna F de Cupones, acá se calcula en SQL al leer, nunca se
guarda un valor fijo que pueda quedar desactualizado).

reportes_data/eft.db (gitignored, mismo directorio que las demás bases):
- eft_deposits: un EFT (RCV-#####) por fila.
- eft_coupons: los cupones/facturas que ESE EFT trae adentro (columna
  "coupon" puede quedar NULL si el PDF no traía el DDC -- ver
  set_manual_coupon_id, la corrección manual que pidió el usuario).
- eft_paid_invoices: facturas pagadas por el EFT sin desglose de cupón
  propio (la lista de encabezado del PDF).
- cupones: los DDC del reporte mensual de J.H. Williams -- "pendientes"
  hasta que algún eft_coupons.coupon los matchee (join en tiempo de
  lectura, nunca un valor guardado que se desactualice).
"""

import os
import sqlite3
from datetime import date, datetime

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "eft.db")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    # timeout=30 + WAL -- ver reportes_db.py: necesario desde que las cargas
    # en segundo plano (jobs.py) pueden escribir de verdad en paralelo.
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    _ensure_schema(conn)
    return conn


def _ensure_schema(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS eft_deposits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rcv_number TEXT NOT NULL,
            eft_date TEXT,
            gross_total REAL,
            fees_total REAL,
            net_total REAL,
            source_filename TEXT,
            updated_at TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS eft_coupons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            deposit_id INTEGER NOT NULL REFERENCES eft_deposits(id),
            date TEXT,
            invoice TEXT,
            coupon TEXT,
            gross_amount REAL,
            fees_amount REAL,
            paid_amount REAL,
            coupon_manual INTEGER DEFAULT 0
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS eft_paid_invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            deposit_id INTEGER NOT NULL REFERENCES eft_deposits(id),
            invoice TEXT,
            paid_amount REAL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cupones (
            coupon_id TEXT PRIMARY KEY,
            date TEXT,
            gross REAL,
            fees REAL,
            net REAL,
            reported_group_text TEXT,
            reported_group_total REAL,
            source_filename TEXT,
            updated_at TEXT
        )
        """
    )
    conn.commit()


def _amounts_close(left, right, tolerance=0.01):
    try:
        return abs(float(left) - float(right)) <= tolerance
    except (TypeError, ValueError):
        return False


def find_existing_deposit(rcv_number, eft_date, net_total):
    """
    Mismo criterio que eft_cta_cte.eft_already_loaded_in_workbook (RCV +
    fecha + neto total coinciden) pero contra la base en vez de un Excel --
    evita cargar el mismo EFT dos veces.
    """
    if not rcv_number:
        return None
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM eft_deposits WHERE rcv_number = ? AND eft_date = ?",
            (rcv_number, eft_date),
        ).fetchall()
        for row in rows:
            if row["net_total"] is None or _amounts_close(row["net_total"], net_total):
                return dict(row)
        return None
    finally:
        conn.close()


def save_eft(header_data, paid_invoices, credit_coupons, source_filename=None):
    """
    Guarda un EFT ya extraído (ver eft_cta_cte.extract_eft_data) -- llamador
    ya debe haber chequeado find_existing_deposit si quiere evitar duplicados
    (acá no se vuelve a chequear, guarda siempre).

    Devuelve el id del depósito creado.
    """
    rcv_number = header_data.get("draft_no") or "UNKNOWN"
    eft_date = header_data.get("eft_date")
    gross_total = sum(float(c.get("gross_amount") or 0.0) for c in credit_coupons)
    fees_total = sum(float(c.get("fees_amount") or 0.0) for c in credit_coupons)
    net_total = sum(float(c.get("paid_amount") or 0.0) for c in credit_coupons)
    now = datetime.utcnow().isoformat()

    conn = _connect()
    try:
        cur = conn.execute(
            """
            INSERT INTO eft_deposits (rcv_number, eft_date, gross_total, fees_total, net_total, source_filename, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (rcv_number, eft_date, gross_total, fees_total, net_total, source_filename, now),
        )
        deposit_id = cur.lastrowid
        for coupon in credit_coupons:
            conn.execute(
                """
                INSERT INTO eft_coupons (deposit_id, date, invoice, coupon, gross_amount, fees_amount, paid_amount, coupon_manual)
                VALUES (?, ?, ?, ?, ?, ?, ?, 0)
                """,
                (
                    deposit_id, coupon.get("date"), coupon.get("invoice"), coupon.get("coupon"),
                    coupon.get("gross_amount"), coupon.get("fees_amount"), coupon.get("paid_amount"),
                ),
            )
        for entry in paid_invoices or []:
            conn.execute(
                "INSERT INTO eft_paid_invoices (deposit_id, invoice, paid_amount) VALUES (?, ?, ?)",
                (deposit_id, entry.get("invoice"), entry.get("paid_amount")),
            )
        conn.commit()
        return deposit_id
    finally:
        conn.close()


def set_manual_coupon_id(eft_coupon_id, coupon_id):
    """
    Agrega (o corrige) a mano el DDC de una línea de EFT que el PDF no traía
    -- pedido explícito del usuario (2026-09-12): "en caso de que haya un
    cupon en un EFT que no tiene un DDC yo le pueda agregar ese numero que
    falta". Una vez puesto, esa línea ya se puede cruzar contra Cupones
    (get_cupones_with_status hace el join en el momento, no hace falta
    "resincronizar" nada aparte).
    """
    coupon_id = (coupon_id or "").strip().upper() or None
    conn = _connect()
    try:
        cur = conn.execute(
            "UPDATE eft_coupons SET coupon = ?, coupon_manual = 1 WHERE id = ?",
            (coupon_id, eft_coupon_id),
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def delete_deposit(deposit_id):
    """
    Borra un EFT ya guardado, junto con sus cupones/facturas asociadas --
    pedido explícito del usuario (2026-09-14): "los EFT subidos no tienen
    forma de ser eliminados, deberia poder cambiarse eso". eft_coupons/
    eft_paid_invoices no tienen ON DELETE CASCADE, así que se borran a mano
    antes que el depósito, todo en una sola conexión.
    """
    conn = _connect()
    try:
        conn.execute("DELETE FROM eft_coupons WHERE deposit_id = ?", (deposit_id,))
        conn.execute("DELETE FROM eft_paid_invoices WHERE deposit_id = ?", (deposit_id,))
        cur = conn.execute("DELETE FROM eft_deposits WHERE id = ?", (deposit_id,))
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def get_month_deposits(year, month):
    """
    EFT cargados con fecha dentro de ese mes, cada uno con sus líneas de
    cupón/factura -- ordenados por la FECHA del EFT (no por el orden en que
    se cargaron, pedido explícito del usuario 2026-09-12: "los eft deberian
    mostrarse en datos no por el orden en el que se los cargo, sino por la
    fecha que tiene el EFT"). El más antiguo primero, el más reciente al
    final (corregido el mismo día, segunda vuelta: "deberia ir de por la
    fecha mas antigua desde arriba hasta la fecha actual mas hacia abajo")
    -- si algún EFT quedó con una fecha que no se pudo parsear, va al final.
    """
    conn = _connect()
    try:
        deposits = conn.execute("SELECT * FROM eft_deposits ORDER BY id DESC").fetchall()
        month_deposits = [dep for dep in deposits if _date_in_month(dep["eft_date"], year, month)]
        # Separado en dos grupos (en vez de una sola clave compuesta) para
        # que "sin fecha" quede siempre al final, nunca mezclado por orden
        # de fecha con los que sí tienen una.
        with_date = sorted(
            (dep for dep in month_deposits if _parse_eft_date(dep["eft_date"]) is not None),
            key=lambda dep: _parse_eft_date(dep["eft_date"]),
        )
        without_date = [dep for dep in month_deposits if _parse_eft_date(dep["eft_date"]) is None]
        ordered = with_date + without_date

        result = []
        for dep in ordered:
            coupons = conn.execute(
                "SELECT * FROM eft_coupons WHERE deposit_id = ? ORDER BY id", (dep["id"],)
            ).fetchall()
            paid = conn.execute(
                "SELECT * FROM eft_paid_invoices WHERE deposit_id = ? ORDER BY id", (dep["id"],)
            ).fetchall()
            result.append({
                "deposit": dict(dep),
                "coupons": [dict(c) for c in coupons],
                "paid_invoices": [dict(p) for p in paid],
            })
        return result
    finally:
        conn.close()


def _parse_eft_date(eft_date):
    if not eft_date:
        return None
    try:
        return datetime.strptime(str(eft_date), "%m/%d/%Y")
    except ValueError:
        return None


def _date_in_month(eft_date, year, month):
    """eft_date queda guardada como MM/DD/YYYY (el formato de origen del PDF) -- se parsea, nunca se compara como texto."""
    parsed = _parse_eft_date(eft_date)
    if parsed is None:
        return False
    return parsed.year == year and parsed.month == month


def list_all_deposits():
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM eft_deposits ORDER BY id DESC").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_cupones_with_status(limit=None):
    """
    Todos los cupones guardados, cada uno con su estado de cruce contra
    eft_coupons (mismo criterio que la columna F/H/I del Excel, calculado
    en el momento vía join en vez de guardado -- nunca puede quedar
    desactualizado). "pending" = True si ningún eft_coupons.coupon todavía
    coincide con este DDC.
    """
    conn = _connect()
    try:
        query = "SELECT * FROM cupones ORDER BY date IS NULL, date DESC, coupon_id"
        if limit:
            query += f" LIMIT {int(limit)}"
        cupones = conn.execute(query).fetchall()
        result = []
        for row in cupones:
            match = conn.execute(
                """
                SELECT ec.*, ed.rcv_number, ed.eft_date
                FROM eft_coupons ec JOIN eft_deposits ed ON ed.id = ec.deposit_id
                WHERE ec.coupon = ? LIMIT 1
                """,
                (row["coupon_id"],),
            ).fetchone()
            entry = dict(row)
            entry["match"] = dict(match) if match else None
            result.append(entry)
        return result
    finally:
        conn.close()


def _parse_cupon_date(value):
    """
    cupones.date puede haber quedado guardada en varios formatos según el
    tipo de reporte mensual que la trajo (un objeto datetime de Excel
    serializado por sqlite3, o texto crudo de un CSV con cualquier
    formato) -- se prueban varios formatos conocidos antes de rendirse.
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def get_cupones_grouped_by_month():
    """
    Los mismos cupones de get_cupones_with_status (históricos, nunca
    filtrados por mes -- pedido explícito del usuario 2026-09-12: "esto
    quiero que sea un excel el cual incluya todos los cupones que se cargan
    de forma historica, no por mes") agrupados por mes calendario para poder
    mostrar un "diferenciador de meses" -- más reciente primero. La
    agrupación se hace con la fecha ya parseada en Python (nunca confiando
    en el ORDER BY date de SQL, que es un simple orden de texto y podría
    quedar mal si dos cargas guardaron el formato de fecha distinto).
    Cupones sin fecha parseable quedan en un grupo aparte, al final.
    """
    cupones = get_cupones_with_status()
    for cp in cupones:
        cp["_parsed_date"] = _parse_cupon_date(cp.get("date"))
    dated = sorted(
        (cp for cp in cupones if cp["_parsed_date"] is not None),
        key=lambda cp: cp["_parsed_date"],
        reverse=True,
    )
    undated = [cp for cp in cupones if cp["_parsed_date"] is None]

    groups = []
    index_by_key = {}
    for cp in dated:
        d = cp.pop("_parsed_date")
        key = (d.year, d.month)
        if key not in index_by_key:
            index_by_key[key] = {"year": d.year, "month": d.month, "cupones": []}
            groups.append(index_by_key[key])
        index_by_key[key]["cupones"].append(cp)
    for cp in undated:
        cp.pop("_parsed_date", None)
    if undated:
        groups.append({"year": None, "month": None, "cupones": undated})
    return groups


def get_unmatched_eft_coupons():
    """Líneas de EFT sin DDC (el PDF no lo traía) -- candidatas a agregarlo a mano."""
    conn = _connect()
    try:
        rows = conn.execute(
            """
            SELECT ec.*, ed.rcv_number, ed.eft_date
            FROM eft_coupons ec JOIN eft_deposits ed ON ed.id = ec.deposit_id
            WHERE ec.coupon IS NULL OR ec.coupon = ''
            ORDER BY ed.id DESC, ec.id
            """
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def upsert_cupones(records, source_filename=None):
    """
    Guarda (o actualiza) cupones ya expandidos por coupon_id -- ver
    cupones_append._expand_records_by_coupon_split, reusado tal cual desde
    webapp.py antes de llamar acá. Cada dict: {"coupon", "date", "gross",
    "fees", "net", "reported_group_text", "reported_group_total"}.

    Un DDC ya guardado se actualiza (mismo criterio que Chase: la última
    carga del reporte mensual manda) -- devuelve (inserted, updated).
    """
    conn = _connect()
    try:
        inserted = 0
        updated = 0
        now = datetime.utcnow().isoformat()
        for rec in records:
            coupon_id = (rec.get("coupon") or "").strip().upper()
            if not coupon_id:
                continue
            exists = conn.execute(
                "SELECT 1 FROM cupones WHERE coupon_id = ?", (coupon_id,)
            ).fetchone() is not None
            conn.execute(
                """
                INSERT INTO cupones (coupon_id, date, gross, fees, net, reported_group_text, reported_group_total, source_filename, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(coupon_id) DO UPDATE SET
                    date = excluded.date,
                    gross = excluded.gross,
                    fees = excluded.fees,
                    net = excluded.net,
                    reported_group_text = excluded.reported_group_text,
                    reported_group_total = excluded.reported_group_total,
                    source_filename = excluded.source_filename,
                    updated_at = excluded.updated_at
                """,
                (
                    coupon_id, rec.get("date"), rec.get("gross"), rec.get("fees"), rec.get("net"),
                    rec.get("reported_group_text"), rec.get("reported_group_total"), source_filename, now,
                ),
            )
            if exists:
                updated += 1
            else:
                inserted += 1
        conn.commit()
        return inserted, updated
    finally:
        conn.close()
