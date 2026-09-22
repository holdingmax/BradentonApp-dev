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
  update_coupon_row, la corrección manual completa que pidió el usuario).
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


def update_coupon_row(eft_coupon_id, *, date=None, invoice=None, coupon=None, gross_amount=None, fees_amount=None, paid_amount=None):
    """
    Corrige a mano cualquier campo de una línea de EFT ya guardada -- pedido
    explícito del usuario (2026-09-16): "corregir" solo dejaba editar el
    DDC, ahora deja editar la fila completa (fecha/factura/DDC/gross/fees/
    net) -- el PDF puede haber traído algo mal leído, no solo el DDC
    faltante. Reemplaza a la vieja set_manual_coupon_id (solo DDC). El
    caller (webapp.py) siempre manda las 6 columnas juntas -- no se admiten
    actualizaciones parciales, para no arriesgar pisar con NULL un campo
    que el formulario no mandó. `coupon_manual` se sigue marcando en 1
    (mismo criterio de antes) -- ya no decide ningún color en la UI (ver
    carga_datos_eft_historial.html), pero queda como registro de que esta
    línea fue corregida a mano al menos una vez.
    """
    coupon_value = (coupon or "").strip().upper() or None
    conn = _connect()
    try:
        cur = conn.execute(
            """
            UPDATE eft_coupons
            SET date = ?, invoice = ?, coupon = ?, gross_amount = ?, fees_amount = ?, paid_amount = ?, coupon_manual = 1
            WHERE id = ?
            """,
            (date, invoice, coupon_value, gross_amount, fees_amount, paid_amount, eft_coupon_id),
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


def get_deposit_years():
    """
    Años distintos con al menos un EFT cargado -- pedido explícito del
    usuario (2026-09-17, módulo nuevo "Reportes"): el selector de mes/año
    del reporte de EFT solo debe ofrecer años que de verdad tengan EFT
    cargados, no un rango arbitrario. `eft_date` se guarda como texto
    MM/DD/YYYY (formato de origen del PDF), así que se parsea con
    `_parse_eft_date` en vez de asumir que el texto ordena bien.
    """
    conn = _connect()
    try:
        rows = conn.execute("SELECT DISTINCT eft_date FROM eft_deposits").fetchall()
        years = {
            parsed.year
            for parsed in (_parse_eft_date(row["eft_date"]) for row in rows)
            if parsed is not None
        }
        return sorted(years)
    finally:
        conn.close()


def list_all_deposits():
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM eft_deposits ORDER BY id DESC").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def backfill_grouped_cupones_from_eft():
    """
    Completa cupones "agrupados" con los montos reales de gross/fees/net
    que ese DDC individual trae en un EFT ya cargado -- pedido explícito
    del usuario (2026-09-17: "hay un montón de cupones que muestran
    0.00... quiero que estos se completen en caso de que falten con los
    datos de los DDC que vienen en los EFT").

    Cuando el reporte mensual de Cupones trae una fila que combina varios
    DDC con un solo total (un "batch"), `cupones_append.
    expand_monthly_records_for_storage` guarda cada DDC como su propia
    fila pero con gross/fees/net en 0 -- no hay forma de saber cómo se
    reparte el total entre ellos SOLO con el reporte mensual. Si más
    adelante se carga un EFT que sí trae ESE DDC puntual con su propio
    monto (columna Reference del EFT, ver eft_cta_cte.py), ahí SÍ se sabe
    el valor real de ese DDC individual -- se usa para completar la fila
    en vez de dejarla en 0 para siempre.

    Solo toca cupones que siguen en 0 Y que vinieron de un grupo
    (`reported_group_text` no nulo) -- un cupón cargado individual (no
    agrupado) nunca se pisa, y uno ya completado (por una corrida
    anterior de esto mismo, o a mano) tampoco se vuelve a tocar, así que
    es seguro llamarla en cada carga de la página (ver la ruta en
    webapp.py) sin necesidad de ningún botón aparte -- se autocompleta
    solo a medida que se van cargando más EFT.

    Devuelve la cantidad de cupones completados en esta corrida.
    """
    conn = _connect()
    try:
        candidates = conn.execute(
            """
            SELECT coupon_id FROM cupones
            WHERE reported_group_text IS NOT NULL AND gross = 0 AND fees = 0 AND net = 0
            """
        ).fetchall()
        filled = 0
        for row in candidates:
            match = conn.execute(
                "SELECT gross_amount, fees_amount, paid_amount FROM eft_coupons WHERE coupon = ? LIMIT 1",
                (row["coupon_id"],),
            ).fetchone()
            if match is None:
                continue
            conn.execute(
                "UPDATE cupones SET gross = ?, fees = ?, net = ? WHERE coupon_id = ?",
                (match["gross_amount"] or 0.0, match["fees_amount"] or 0.0, match["paid_amount"] or 0.0, row["coupon_id"]),
            )
            filled += 1
        if filled:
            conn.commit()
        return filled
    finally:
        conn.close()


def get_cupones_with_status(limit=None):
    """
    Todos los cupones guardados, cada uno con su estado de cruce contra
    eft_coupons (mismo criterio que la columna F/H/I del Excel, calculado
    en el momento vía join en vez de guardado -- nunca puede quedar
    desactualizado). "pending" = True si ningún eft_coupons.coupon todavía
    coincide con este DDC. `diff` (pedido explícito del usuario, 2026-09-16
    -- "poner despues de si quedaron diferencias entre los EFT como se
    tenia la columna G en el excel") es el mismo cálculo que esa columna G
    real: Net Amount reportado menos lo que el EFT efectivamente pagó por
    esa línea -- None si todavía no hay ningún EFT que lo cruce.

    `group_remaining`/`group_pending_count` (2026-09-17, "que le
    descuenten al cupón padre con el que vinieron en el batch, así queden
    las diferencias en 0"): para un cupón agrupado que TODAVÍA sigue en 0
    (ningún EFT lo resolvió todavía -- ver backfill_grouped_cupones_from_eft
    arriba), se calcula cuánto le queda al grupo entero descontando lo que
    ya resolvieron sus hermanos (`reported_group_total` menos la suma de
    los hermanos que ya se completaron) -- así el cupón pendiente muestra
    el saldo REAL que le queda al batch, no el total original del grupo
    entero (que ya no aplica una vez que otros DDC del mismo batch se
    fueron resolviendo). Ambos quedan en None para un cupón no agrupado, o
    para uno agrupado que ya se resolvió (tiene su propio monto real).
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
            if entry["match"] is not None and entry["match"].get("paid_amount") is not None:
                entry["diff"] = round((entry.get("net") or 0) - entry["match"]["paid_amount"], 2)
            else:
                entry["diff"] = None
            entry["group_remaining"] = None
            entry["group_pending_count"] = None
            result.append(entry)

        groups = {}
        for entry in result:
            group_text = entry.get("reported_group_text")
            if group_text:
                groups.setdefault(group_text, []).append(entry)
        for members in groups.values():
            group_total = next((m["reported_group_total"] for m in members if m.get("reported_group_total") is not None), None)
            if group_total is None:
                continue
            resolved_sum = sum((m.get("net") or 0.0) for m in members if (m.get("net") or 0.0) != 0.0)
            pending = [m for m in members if (m.get("net") or 0.0) == 0.0]
            if not pending:
                continue
            remaining = round(group_total - resolved_sum, 2)
            for m in pending:
                m["group_remaining"] = remaining
                m["group_pending_count"] = len(pending)

        return result
    finally:
        conn.close()


def get_cupones_flat():
    """
    Todos los cupones guardados en un único listado plano, del más antiguo
    al más nuevo (pedido explícito del usuario, 2026-09-16 -- "quiero que
    los cupones se muestren desde el mas antiguo primero al mas nuevo... y
    que cada vez que se abra la parte de cupones, que te lo abra a lo que
    seria al final de la pagina"). Reemplaza a get_cupones_grouped_by_month
    -- ya no hay divisor de mes visual, la columna "Mes EFT" cumple ese rol
    ahora. Cupones sin fecha parseable quedan primero (antes que cualquier
    fecha real conocida), para que la vista "al final de la página" siga
    mostrando siempre los cupones fechados más recientes.

    `date_display` (pedido explícito del usuario, 2026-09-17 -- "quiero
    que la fecha de los cupones sea dd/mm/yyyy") reusa el mismo parser
    tolerante que ya usa el ordenamiento (`_parse_cupon_date`, entiende
    los varios formatos crudos con los que quedó guardada `cupones.date`
    según el reporte que la trajo) -- `date` en sí NO se toca, sigue
    crudo tal cual se guardó, por si algo más lo llega a necesitar así.
    """
    cupones = get_cupones_with_status()
    for cp in cupones:
        cp["_parsed_date"] = _parse_cupon_date(cp.get("date"))
    cupones.sort(key=lambda cp: cp["_parsed_date"] or datetime.min)
    for cp in cupones:
        parsed = cp.pop("_parsed_date", None)
        cp["date_display"] = parsed.strftime("%d/%m/%Y") if parsed else None
    return cupones


def eft_month_and_year(eft_date):
    """(año, mes) del EFT que pagó este cupón, o None si no hay cruce/fecha parseable."""
    parsed = _parse_eft_date(eft_date)
    if parsed is None:
        return None
    return (parsed.year, parsed.month)


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


def delete_cupones_month(year, month):
    """
    Borra todos los cupones guardados cuya fecha caiga en (year, month) --
    pedido explícito del usuario (2026-09-17: "no hay forma de borrar los
    cupones cargados... debería haber una forma de borrar todos los
    cupones del mes"). Usa el mismo parser tolerante que el resto de este
    módulo (_parse_cupon_date) para decidir qué filas matchean, nunca un
    LIKE contra el texto crudo -- la fecha puede haber quedado guardada en
    más de un formato según el reporte que la trajo (ver get_cupones_flat).
    Devuelve la cantidad de cupones borrados.
    """
    conn = _connect()
    try:
        rows = conn.execute("SELECT coupon_id, date FROM cupones").fetchall()
        to_delete = []
        for row in rows:
            parsed = _parse_cupon_date(row["date"])
            if parsed is not None and parsed.year == year and parsed.month == month:
                to_delete.append(row["coupon_id"])
        if to_delete:
            conn.executemany(
                "DELETE FROM cupones WHERE coupon_id = ?",
                [(cid,) for cid in to_delete],
            )
            conn.commit()
        return len(to_delete)
    finally:
        conn.close()


def get_cupones_years():
    """Años distintos con al menos un cupón guardado, ascendente -- para poblar el selector de \"borrar mes\"."""
    conn = _connect()
    try:
        rows = conn.execute("SELECT date FROM cupones").fetchall()
    finally:
        conn.close()
    years = set()
    for row in rows:
        parsed = _parse_cupon_date(row["date"])
        if parsed is not None:
            years.add(parsed.year)
    return sorted(years)


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


def _fmt_money_pdf(value):
    """Mismo criterio de signo que chase_rules._fmt_money_pdf -- "-$" antes del monto, nunca "$-"."""
    if value is None:
        return "—"
    if value < 0:
        return "-${:,.2f}".format(abs(value))
    return "${:,.2f}".format(value)


def build_eft_pdf_report(year, month, dest_path):
    """
    PDF del módulo nuevo "Reportes" (pedido explícito del usuario,
    2026-09-17): "Lo mismo quiero que hagas con el Reporte de los EFT
    incluyendo todos los datos que se tengan del mes que se selecciono, y
    lo mismo estar incluido en ese PDF los cupones cargados hasta ese
    momento y cuanto acumulan". A diferencia de Chase (que sí se resume
    por Detalle), acá "todos los datos" significa una fila por CADA EFT
    (RCV) cargado ese mes -- no se agrupa nada, se listan todos.

    Dos secciones (ver `pdf_export.build_multi_section_pdf`):
    1. "EFT del mes" -- un renglón por depósito (RCV/Fecha/Gross/Fees/Net)
       más una fila TOTAL. El período que se imprime en el membrete usa la
       fecha MÍNIMA/MÁXIMA real de los EFT de ese mes (mismo criterio que
       el PDF de Chase) -- nunca asume que el mes esté completo.
    2. "Cupones" -- NO es un listado (serían miles de filas, va contra "que
       se vea bien y agradable a la vista") sino el acumulado HISTÓRICO
       completo -- "cupones cargados hasta ese momento" es today, no solo
       los de este mes -- mismo total que ya muestra `/carga-datos/eft/
       cupones/historial` (get_cupones_with_status, sin filtrar por mes).
    """
    from pdf_export import build_multi_section_pdf

    deposits = get_month_deposits(year, month)
    eft_rows = []
    total_gross = total_fees = total_net = 0.0
    parsed_dates = []
    for entry in deposits:
        dep = entry["deposit"]
        parsed = _parse_eft_date(dep.get("eft_date"))
        if parsed:
            parsed_dates.append(parsed)
        gross = dep.get("gross_total") or 0.0
        fees = dep.get("fees_total") or 0.0
        net = dep.get("net_total") or 0.0
        total_gross += gross
        total_fees += fees
        total_net += net
        eft_rows.append(
            [
                dep.get("rcv_number") or "—",
                parsed.strftime("%d/%m/%Y") if parsed else (dep.get("eft_date") or "—"),
                _fmt_money_pdf(gross),
                _fmt_money_pdf(fees),
                _fmt_money_pdf(net),
            ]
        )
    eft_rows.append(["TOTAL", "", _fmt_money_pdf(total_gross), _fmt_money_pdf(total_fees), _fmt_money_pdf(total_net)])

    if parsed_dates:
        period_label = (
            f"Período: {min(parsed_dates).strftime('%d/%m/%Y')} al "
            f"{max(parsed_dates).strftime('%d/%m/%Y')}"
        )
    else:
        period_label = f"Período: sin EFT cargados todavía en {month:02d}/{year}"

    # Cupones -- pedido explícito del usuario (2026-09-22): reemplaza el
    # resumen histórico de una sola línea por DOS cuadros reales, los dos
    # acotados al mes que se está reportando: (1) los cupones que un EFT
    # de ESTE mes efectivamente pagó (match.eft_date, no cupon.date -- un
    # EFT de septiembre puede pagar un cupón cargado con fecha de agosto,
    # lo que importa acá es cuándo se cobró) y (2) los cupones con fecha
    # de este mes que TODAVÍA no se aplicaron a ningún EFT (pendientes),
    # para poder ver de un vistazo qué falta cobrar de lo que se cargó.
    cupones = get_cupones_with_status()

    applied_rows = []
    total_applied_net = total_applied_paid = 0.0
    for c in cupones:
        match = c.get("match")
        if not match:
            continue
        eft_parsed = _parse_eft_date(match.get("eft_date"))
        if not eft_parsed or eft_parsed.year != year or eft_parsed.month != month:
            continue
        cupon_parsed = _parse_cupon_date(c.get("date"))
        net = c.get("net") or 0.0
        paid = match.get("paid_amount") or 0.0
        total_applied_net += net
        total_applied_paid += paid
        applied_rows.append(
            [
                c.get("coupon_id") or "—",
                cupon_parsed.strftime("%d/%m/%Y") if cupon_parsed else (c.get("date") or "—"),
                match.get("rcv_number") or "—",
                eft_parsed.strftime("%d/%m/%Y"),
                _fmt_money_pdf(net),
                _fmt_money_pdf(paid),
            ]
        )
    applied_rows.sort(key=lambda row: row[3])
    if applied_rows:
        applied_rows.append(
            ["TOTAL", "", "", "", _fmt_money_pdf(total_applied_net), _fmt_money_pdf(total_applied_paid)]
        )

    pending_rows = []
    total_pending = 0.0
    for c in cupones:
        if c.get("match"):
            continue
        cupon_parsed = _parse_cupon_date(c.get("date"))
        if not cupon_parsed or cupon_parsed.year != year or cupon_parsed.month != month:
            continue
        net = c.get("net") or 0.0
        total_pending += net
        pending_rows.append(
            [
                c.get("coupon_id") or "—",
                cupon_parsed.strftime("%d/%m/%Y"),
                _fmt_money_pdf(c.get("gross")),
                _fmt_money_pdf(c.get("fees")),
                _fmt_money_pdf(net),
            ]
        )
    pending_rows.sort(key=lambda row: row[1])
    if pending_rows:
        pending_rows.append(["TOTAL", "", "", "", _fmt_money_pdf(total_pending)])

    title = f"EFT — {month:02d}/{year}"
    sections = [
        {
            "heading": "EFT del mes",
            "headers": ["RCV", "Fecha", "Gross", "Fees", "Net"],
            "rows": eft_rows,
            "col_widths_mm": [55, 45, 55, 55, 55],
            "bold_last_row": True,
        },
        {
            "heading": "Cupones aplicados a EFT de este mes",
            "headers": ["Cupón", "Fecha Cupón", "RCV", "Fecha EFT", "Net", "Pagado"],
            "rows": applied_rows,
            "col_widths_mm": [40, 40, 40, 40, 40, 40],
            "bold_last_row": True,
        }
        if applied_rows
        else {
            "heading": "Cupones aplicados a EFT de este mes",
            "note": "Ningún EFT de este mes pagó un cupón todavía.",
        },
        {
            "heading": "Cupones pendientes (fecha de este mes, sin aplicar a ningún EFT)",
            "headers": ["Cupón", "Fecha", "Gross", "Fees", "Net"],
            "rows": pending_rows,
            "col_widths_mm": [48, 48, 48, 48, 48],
            "bold_last_row": True,
        }
        if pending_rows
        else {
            "heading": "Cupones pendientes (fecha de este mes, sin aplicar a ningún EFT)",
            "note": "No hay cupones pendientes con fecha de este mes.",
        },
    ]
    build_multi_section_pdf(dest_path, title, sections, period_label=period_label, company_header=True)
    return dest_path
