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

import calendar
import hashlib
import os
import sqlite3
from datetime import date, datetime, timedelta

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
    # Detalle de cupones (pedido del usuario, 2026-10-06, ver
    # cupones_detalle.py): un grupo por depósito (fila del reporte mensual de
    # cupones), con los DDC que lo forman si se conocen, y sus batches con la
    # fecha real de venta. `signature` (los batches ordenados) evita cargar
    # dos veces el mismo grupo.
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cupon_detail_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            signature TEXT NOT NULL UNIQUE,
            gross REAL, fees REAL, net REAL,
            first_date TEXT, last_date TEXT,
            coupons TEXT,
            coupons_source TEXT,
            source_filename TEXT,
            updated_at TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cupon_detail_batches (
            group_id INTEGER NOT NULL REFERENCES cupon_detail_groups(id),
            idx INTEGER NOT NULL,
            batch_date TEXT NOT NULL,
            batch TEXT NOT NULL,
            gross REAL, fees REAL, net REAL,
            PRIMARY KEY (group_id, idx)
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
            deposit = dict(dep)
            # Los totales que se muestran salen de las líneas, no de lo
            # guardado al extraer: así una corrección a mano de una línea se
            # ve en el encabezado y en el PDF (auditoría 2026-09). Lo guardado
            # queda intacto porque find_existing_deposit compara contra el
            # neto original del PDF para no cargar dos veces el mismo EFT.
            if coupons:
                deposit["gross_total"] = round(sum(c["gross_amount"] or 0.0 for c in coupons), 2)
                deposit["fees_total"] = round(sum(c["fees_amount"] or 0.0 for c in coupons), 2)
                deposit["net_total"] = round(sum(c["paid_amount"] or 0.0 for c in coupons), 2)
            result.append({
                "deposit": deposit,
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
            # Un mismo DDC suele pagarse en 2-4 líneas del EFT (una por
            # factura SI-): hay que sumarlas todas, no tomar la primera
            # (auditoría 2026-09, faltaban $68k de Net).
            match = conn.execute(
                """
                SELECT COUNT(*) AS n, SUM(gross_amount) AS gross_amount,
                       SUM(fees_amount) AS fees_amount, SUM(paid_amount) AS paid_amount
                FROM eft_coupons WHERE coupon = ?
                """,
                (row["coupon_id"],),
            ).fetchone()
            if not match["n"]:
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
                SELECT MAX(ec.id) AS id, ed.rcv_number, ed.eft_date,
                       SUM(ec.gross_amount) AS gross_amount, SUM(ec.fees_amount) AS fees_amount,
                       SUM(ec.paid_amount) AS paid_amount, COUNT(*) AS lines
                FROM eft_coupons ec JOIN eft_deposits ed ON ed.id = ec.deposit_id
                WHERE ec.coupon = ?
                """,
                (row["coupon_id"],),
            ).fetchone()
            entry = dict(row)
            # Suma de todas las líneas del DDC; RCV/fecha de la última línea
            # (SQLite toma las columnas sueltas de la fila con MAX(id)).
            entry["match"] = dict(match) if match and match["lines"] else None
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


def order_cupones_by_eft(cupones):
    """
    Mismo listado de get_cupones_flat, reordenado para el historial (pedido
    del usuario, 2026-10-02): "a todos los cupones que fueron aplicados en
    una misma fecha de EFT que los pongas continuados... así se puede ver de
    forma más clara qué cupones fueron aplicados y cuáles no". Por año (el
    del EFT si ya se aplicó, si no el del cupón): primero los aplicados,
    juntos por EFT (fecha + RCV, del más viejo al más nuevo), y al final los
    pendientes por su fecha. Agrega a cada cupón `eft_date_display`
    (DD/MM/YYYY), `group_year` y `eft_band` (0/1, alterna en cada EFT para
    que la plantilla los distinga a simple vista).
    """
    for cp in cupones:
        match = cp.get("match")
        eft_parsed = _parse_eft_date(match.get("eft_date")) if match else None
        cupon_parsed = _parse_cupon_date(cp.get("date"))
        cp["eft_date_display"] = eft_parsed.strftime("%d/%m/%Y") if eft_parsed else None
        anchor = eft_parsed or cupon_parsed
        cp["group_year"] = anchor.year if anchor else None
        cp["_sort"] = (
            cp["group_year"] if cp["group_year"] is not None else -1,
            0 if eft_parsed else 1,  # aplicados primero, pendientes al final del año
            eft_parsed or datetime.min,
            (match or {}).get("rcv_number") or "",
            cupon_parsed or datetime.min,
            cp.get("coupon_id") or "",
        )
    ordered = sorted(cupones, key=lambda cp: cp.pop("_sort"))
    band = 0
    previous = None
    for cp in ordered:
        match = cp.get("match")
        block = (cp.get("eft_date_display"), (match or {}).get("rcv_number")) if match else None
        cp["eft_block_start"] = block != previous
        if block != previous:
            band = 1 - band
            previous = block
        cp["eft_band"] = band
        cp["eft_block_total"] = None

    # Renglón de total debajo de cada EFT (pedido del usuario, 2026-10-02):
    # lo que suman sus cupones contra lo que el EFT trae en sus líneas de
    # cupón (todas, también las que no tienen DDC o cuyo DDC no está cargado).
    eft_totals = _eft_coupon_totals()
    block_members = []
    for index, cp in enumerate(ordered):
        if not cp.get("match"):
            continue
        block_members.append(cp)
        following = ordered[index + 1] if index + 1 < len(ordered) else None
        if following is not None and following.get("match") and not following["eft_block_start"]:
            continue
        sums = {key: round(sum((m.get(key) or 0.0) for m in block_members), 2) for key in ("gross", "fees", "net")}
        eft = eft_totals.get(((cp["match"].get("rcv_number") or ""), cp["match"].get("eft_date") or ""))
        total = {"count": len(block_members), **sums, "eft": eft}
        if eft is not None:
            # Los EFT de ene-may no detallan el Fee (Gross = Net, Fee 0):
            # ahí solo se puede comparar el Net.
            fees_detailed = not (abs(eft["fees"]) < 0.01 and abs(eft["gross"] - eft["net"]) < 0.01)
            total["fees_detailed"] = fees_detailed
            total["gross_ok"] = abs(sums["gross"] - eft["gross"]) < 0.01 if fees_detailed else None
            total["fees_ok"] = abs(sums["fees"] - eft["fees"]) < 0.01 if fees_detailed else None
            total["net_ok"] = abs(sums["net"] - eft["net"]) < 0.01
            total["all_ok"] = total["net_ok"] and total["gross_ok"] is not False and total["fees_ok"] is not False
            total["diff"] = round(sums["net"] - eft["net"], 2)
        cp["eft_block_total"] = total
        block_members = []
    return ordered


def _eft_coupon_totals():
    """
    {(RCV, fecha del EFT): Gross/Fee/Net sumados de todas sus líneas de
    cupón}, más lo que explica una diferencia: líneas sin DDC y líneas con
    un DDC que no está en Cupones.
    """
    conn = _connect()
    try:
        rows = conn.execute(
            """
            SELECT ed.rcv_number, ed.eft_date, COUNT(ec.id) AS lines,
                   SUM(ec.gross_amount) AS gross, SUM(ec.fees_amount) AS fees, SUM(ec.paid_amount) AS net,
                   SUM(CASE WHEN ec.coupon IS NULL OR TRIM(ec.coupon) = '' THEN 1 ELSE 0 END) AS no_ddc_lines,
                   SUM(CASE WHEN ec.coupon IS NULL OR TRIM(ec.coupon) = '' THEN ec.paid_amount ELSE 0 END) AS no_ddc_net,
                   SUM(CASE WHEN TRIM(ec.coupon) <> '' AND c.coupon_id IS NULL THEN 1 ELSE 0 END) AS unknown_lines,
                   SUM(CASE WHEN TRIM(ec.coupon) <> '' AND c.coupon_id IS NULL THEN ec.paid_amount ELSE 0 END) AS unknown_net
            FROM eft_deposits ed
            JOIN eft_coupons ec ON ec.deposit_id = ed.id
            LEFT JOIN cupones c ON c.coupon_id = ec.coupon
            GROUP BY ed.rcv_number, ed.eft_date
            """
        ).fetchall()
    finally:
        conn.close()
    return {
        ((row["rcv_number"] or ""), row["eft_date"] or ""): {
            "lines": row["lines"],
            "gross": round(row["gross"] or 0.0, 2),
            "fees": round(row["fees"] or 0.0, 2),
            "net": round(row["net"] or 0.0, 2),
            "no_ddc_lines": row["no_ddc_lines"] or 0,
            "no_ddc_net": round(row["no_ddc_net"] or 0.0, 2),
            "unknown_lines": row["unknown_lines"] or 0,
            "unknown_net": round(row["unknown_net"] or 0.0, 2),
        }
        for row in rows
    }


def autolink_missing_ddc_by_amount():
    """
    Línea de EFT sin DDC (el PDF no lo traía o no se leyó) + un cupón
    pendiente con EXACTAMENTE el mismo monto: se completa el DDC de esa
    línea (pedido del usuario, 2026-10-02). Con eso el cupón queda cruzado
    con su EFT (RCV, fecha, pagado) igual que si el DDC hubiera venido en
    el PDF -- el cruce de get_cupones_with_status es por DDC.

    Para no adivinar: Gross y Net iguales al centavo (y Fee también, si los
    dos lo tienen), ningún monto en 0, la fecha del cupón no posterior a la
    del EFT, y la pareja tiene que ser única de los dos lados -- si dos
    cupones pendientes o dos líneas sin DDC comparten el monto, no se toca y
    queda para completar a mano. coupon_manual = 2 marca "completado solo
    por monto". Devuelve la cantidad de líneas completadas.

    Cupón de un batch (2026-10-02): el reporte trae un total para varios
    DDC y el que falta queda en 0 (ver backfill_grouped_cupones_from_eft).
    Si es el ÚNICO que falta de su batch, su monto es lo que le queda al
    batch (total menos los hermanos ya resueltos): ese saldo tiene que ser
    igual al centavo al Net de la línea y la fecha de la línea igual a la
    del cupón. Al completarse el DDC, el backfill le pasa Gross/Fee/Net del
    EFT al cupón.
    """
    conn = _connect()
    completed = 0
    try:
        lines = conn.execute(
            """
            SELECT ec.id, ec.date, ec.gross_amount, ec.fees_amount, ec.paid_amount, ed.eft_date
            FROM eft_coupons ec JOIN eft_deposits ed ON ed.id = ec.deposit_id
            WHERE ec.coupon IS NULL OR TRIM(ec.coupon) = ''
            """
        ).fetchall()
        if not lines:
            return 0
        linked = {
            row[0] for row in conn.execute(
                "SELECT DISTINCT coupon FROM eft_coupons WHERE coupon IS NOT NULL AND TRIM(coupon) <> ''"
            )
        }
        cupones = conn.execute(
            "SELECT coupon_id, date, gross, fees, net, reported_group_text, reported_group_total FROM cupones"
        ).fetchall()

        def cents(value):
            return None if value is None else round(value, 2)

        # Saldo del batch para el único cupón en 0 de cada grupo.
        group_remaining = {}
        groups = {}
        for cupon in cupones:
            if cupon["reported_group_text"]:
                groups.setdefault(cupon["reported_group_text"], []).append(cupon)
        for members in groups.values():
            total = next((m["reported_group_total"] for m in members if m["reported_group_total"] is not None), None)
            zero = [m for m in members if not cents(m["net"])]
            if total is None or len(zero) != 1 or cents(zero[0]["gross"]) or cents(zero[0]["fees"]):
                continue
            remaining = cents(total - sum(m["net"] for m in members if cents(m["net"])))
            if remaining and remaining > 0:
                group_remaining[zero[0]["coupon_id"]] = remaining

        def matches(line, cupon):
            eft_day = _parse_eft_date(line["eft_date"])
            cupon_day = _parse_cupon_date(cupon["date"])
            if eft_day is not None and cupon_day is not None and cupon_day > eft_day:
                return False
            line_gross, line_fees, line_net = cents(line["gross_amount"]), cents(line["fees_amount"]), cents(line["paid_amount"])
            if not line_net:
                return False
            if cents(cupon["gross"]) and cents(cupon["net"]):
                if line_gross != cents(cupon["gross"]) or line_net != cents(cupon["net"]):
                    return False
                return line_fees is None or cupon["fees"] is None or line_fees == cents(cupon["fees"])
            remaining = group_remaining.get(cupon["coupon_id"])
            if remaining is None or line_net != remaining:
                return False
            line_day = _parse_eft_date(line["date"]) or _parse_cupon_date(line["date"])
            return line_day is not None and cupon_day is not None and line_day == cupon_day

        pending = [c for c in cupones if c["coupon_id"] not in linked]
        pairs = [(line, cupon) for line in lines for cupon in pending if matches(line, cupon)]
        per_line, per_cupon = {}, {}
        for line, cupon in pairs:
            per_line[line["id"]] = per_line.get(line["id"], 0) + 1
            per_cupon[cupon["coupon_id"]] = per_cupon.get(cupon["coupon_id"], 0) + 1
        for line, cupon in pairs:
            # Única de los dos lados; si no, no se adivina.
            if per_line[line["id"]] != 1 or per_cupon[cupon["coupon_id"]] != 1:
                continue
            conn.execute(
                "UPDATE eft_coupons SET coupon = ?, coupon_manual = 2 WHERE id = ? AND (coupon IS NULL OR TRIM(coupon) = '')",
                (cupon["coupon_id"], line["id"]),
            )
            completed += 1
        if completed:
            conn.commit()
    finally:
        conn.close()
    if completed:
        backfill_grouped_cupones_from_eft()
    return completed


def get_coupon_gross_by_date():
    """
    {fecha ISO: {"gross", "count", "unknown"}} de los cupones por su fecha
    (control Tarjetas y Cupones). `unknown` cuenta los cupones de un batch
    que todavía no tienen monto propio (Gross en 0): ese día suma de menos.
    """
    conn = _connect()
    try:
        rows = conn.execute("SELECT date, gross, reported_group_text FROM cupones").fetchall()
    finally:
        conn.close()
    result = {}
    for row in rows:
        parsed = _parse_cupon_date(row["date"])
        if parsed is None:
            continue
        day = result.setdefault(parsed.date().isoformat(), {"gross": 0.0, "count": 0, "unknown": 0})
        day["gross"] = round(day["gross"] + (row["gross"] or 0.0), 2)
        day["count"] += 1
        if row["reported_group_text"] and not row["gross"]:
            day["unknown"] += 1
    return result


def get_cupones_detail_by_date():
    """
    {fecha ISO: {"gross", "fees", "net", "coupons": [DDC], "unknown"}} de los
    cupones cargados (cruce con el reporte mensual de J.H.). `unknown` cuenta
    los DDC de un grupo que todavía no tienen monto propio (en 0 hasta que
    un EFT los paga): ese día suma de menos.
    """
    conn = _connect()
    try:
        rows = conn.execute("SELECT coupon_id, date, gross, fees, net, reported_group_text FROM cupones").fetchall()
    finally:
        conn.close()
    result = {}
    for row in rows:
        parsed = _parse_cupon_date(row["date"])
        if parsed is None:
            continue
        day = result.setdefault(parsed.date().isoformat(), {"gross": 0.0, "fees": 0.0, "net": 0.0, "coupons": [], "unknown": 0})
        for key in ("gross", "fees", "net"):
            day[key] = round(day[key] + (row[key] or 0.0), 2)
        day["coupons"].append(row["coupon_id"])
        if row["reported_group_text"] and not any(row[key] for key in ("gross", "fees", "net")):
            day["unknown"] += 1
    return result


def insert_new_cupones(records, source_filename=None):
    """
    Agrega a Cupones solo los DDC que no estaban (records como los de
    cupones_append.expand_monthly_records_for_storage): lo ya cargado no se
    toca (pedido del usuario, 2026-10-06, para el Credit Card Daily Summary
    en PDF). Después completa como upsert_cupones (EFT, DDC por monto,
    detalle de cupones). Devuelve (agregados, ya estaban).
    """
    conn = _connect()
    inserted = skipped = regrouped = 0
    now = datetime.utcnow().isoformat()
    try:
        with conn:
            for rec in records:
                coupon_id = (rec.get("coupon") or "").strip().upper()
                if not coupon_id:
                    continue
                cur = conn.execute(
                    "INSERT OR IGNORE INTO cupones (coupon_id, date, gross, fees, net, reported_group_text, reported_group_total, source_filename, updated_at) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (coupon_id, rec.get("date"), rec.get("gross"), rec.get("fees"), rec.get("net"),
                     rec.get("reported_group_text"), rec.get("reported_group_total"), source_filename, now),
                )
                if cur.rowcount:
                    inserted += 1
                    continue
                skipped += 1
                # Un DDC que quedó cargado solo y ahora el reporte lo trae en
                # un grupo: se le pone el grupo (los montos no se tocan), así
                # el que falta del grupo sale por diferencia y se enlaza con
                # su línea de EFT (autolink_missing_ddc_by_amount).
                if rec.get("reported_group_text"):
                    regrouped += conn.execute(
                        "UPDATE cupones SET reported_group_text = ?, reported_group_total = ? "
                        "WHERE coupon_id = ? AND reported_group_text IS NULL",
                        (rec["reported_group_text"], rec.get("reported_group_total"), coupon_id),
                    ).rowcount
    finally:
        conn.close()
    if inserted or regrouped:
        backfill_grouped_cupones_from_eft()
        autolink_missing_ddc_by_amount()
        identify_detail_groups()
    return inserted, skipped


# ---------------------------------------------------------------------------
# Detalle de cupones (pedido del usuario, 2026-10-06; lectura en
# cupones_detalle.py, control día por día en control_tarjetas.py)
# ---------------------------------------------------------------------------

# Un grupo se identifica con la fila del reporte mensual de cupones cuya
# fecha cae entre su último batch y estos días después (en agosto-septiembre
# 2026 fue siempre el mismo día del último batch).
_DETAIL_MATCH_DAYS = 3


def _detail_signature(batches):
    lines = sorted(f'{b["batch_date"]}|{b["batch"]}|{b["gross"]:.2f}|{b["fees"]:.2f}|{b["net"]:.2f}' for b in batches)
    return hashlib.sha1("\n".join(lines).encode("utf-8")).hexdigest()


def save_coupon_detail(groups, source_filename=None):
    """
    Guarda los grupos leídos (cupones_detalle.extract_detail_groups). Un
    grupo ya cargado (mismos batches) no se duplica; si ahora trae los DDC
    (el PDF de un solo cupón los imprime), se completan. Los DDC de los que
    no los traen se buscan después con identify_detail_groups.
    Devuelve [{"id", "status": new|repeated, "first_date", "last_date", "gross", "coupons"}].
    """
    conn = _connect()
    results = []
    now = datetime.utcnow().isoformat()
    try:
        with conn:
            for group in groups:
                batches = group["batches"]
                signature = _detail_signature(batches)
                dates = sorted(b["batch_date"] for b in batches)
                coupons = ",".join(group["coupons"]) if group.get("coupons") else None
                existing = conn.execute("SELECT id, coupons FROM cupon_detail_groups WHERE signature = ?", (signature,)).fetchone()
                if existing is None:
                    # Un PDF de un solo cupón trae solo sus batches: si todos ya
                    # están en un grupo cargado (el Credit Card Detail del depósito
                    # entero), es ese mismo grupo y no otro; antes se guardaba
                    # aparte y el control día por día los contaba dos veces
                    # (revisión 2026-10-08).
                    holders = set()
                    for b in batches:
                        row = conn.execute(
                            "SELECT group_id FROM cupon_detail_batches WHERE batch_date = ? AND batch = ? "
                            "AND ROUND(gross, 2) = ROUND(?, 2) AND ROUND(fees, 2) = ROUND(?, 2) AND ROUND(net, 2) = ROUND(?, 2)",
                            (b["batch_date"], b["batch"], b["gross"], b["fees"], b["net"]),
                        ).fetchone()
                        holders.add(row["group_id"] if row else None)
                    if batches and None not in holders and len(holders) == 1:
                        existing = conn.execute("SELECT id, coupons FROM cupon_detail_groups WHERE id = ?",
                                                (holders.pop(),)).fetchone()
                        coupons = None  # los DDC de un cupón no son los del grupo entero
                if existing is not None:
                    if coupons and existing["coupons"] != coupons:
                        conn.execute(
                            "UPDATE cupon_detail_groups SET coupons = ?, coupons_source = 'pdf', updated_at = ? WHERE id = ?",
                            (coupons, now, existing["id"]),
                        )
                    status, group_id = "repeated", existing["id"]
                else:
                    totals = group["totals"]
                    cur = conn.execute(
                        "INSERT INTO cupon_detail_groups (signature, gross, fees, net, first_date, last_date, coupons, coupons_source, source_filename, updated_at) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (signature, totals["gross"], totals["fees"], totals["net"], dates[0], dates[-1],
                         coupons, "pdf" if coupons else None, source_filename, now),
                    )
                    for idx, b in enumerate(batches):
                        conn.execute(
                            "INSERT INTO cupon_detail_batches (group_id, idx, batch_date, batch, gross, fees, net) VALUES (?, ?, ?, ?, ?, ?, ?)",
                            (cur.lastrowid, idx, b["batch_date"], b["batch"], b["gross"], b["fees"], b["net"]),
                        )
                    status, group_id = "new", cur.lastrowid
                results.append({"id": group_id, "status": status, "first_date": dates[0], "last_date": dates[-1],
                                "gross": group["totals"]["gross"], "coupons": coupons})
    finally:
        conn.close()
    return results


def identify_detail_groups(report_rows=()):
    """
    Pone los DDC a los grupos de detalle que no los traen: el grupo es una
    fila del reporte mensual de cupones, así que se busca esa fila con la
    misma fecha (o hasta _DETAIL_MATCH_DAYS días después del último batch)
    y los mismos Gross, Fees y Net. Se busca en Cupones (los DDC de un mismo
    `reported_group_text`, o un DDC solo; si sus montos individuales todavía
    están en 0, con el total del grupo) y en `report_rows`, las filas del
    Credit Card Daily Summary en PDF (jh_mensual_db.get_coupon_rows). Solo
    si hay una sola coincidencia. Se corre al cargar detalle, el reporte
    mensual de Cupones o el Credit Card Daily Summary.
    """
    conn = _connect()
    try:
        pending = conn.execute(
            "SELECT id, gross, fees, net, last_date FROM cupon_detail_groups WHERE coupons IS NULL"
        ).fetchall()
        if not pending:
            return 0
        candidates = {}
        for r in conn.execute("SELECT coupon_id, date, gross, fees, net, reported_group_text, reported_group_total FROM cupones"):
            parsed = _parse_cupon_date(r["date"])
            if parsed is None:
                continue
            group_text = r["reported_group_text"]
            c = candidates.setdefault(group_text or r["coupon_id"], {
                # Los DDC del grupo son los que nombra el reporte, aunque
                # alguno haya quedado cargado aparte de una carga anterior.
                "members": sorted({d.strip().upper() for d in group_text.split(",") if d.strip()}) if group_text else [r["coupon_id"]],
                "coupons": [], "gross": 0.0, "fees": 0.0, "net": 0.0, "complete": True,
                "group_total": r["reported_group_total"], "day": parsed.date(),
            })
            c["coupons"].append(r["coupon_id"])
            for key in ("gross", "fees", "net"):
                c[key] += r[key] or 0.0
            if not any(r[key] for key in ("gross", "fees", "net")):
                c["complete"] = False
        # El mismo depósito puede estar en Cupones y en el reporte en PDF:
        # cuenta una sola vez (por sus DDC).
        candidates = {",".join(c["members"]): c for c in candidates.values()}
        for r in report_rows:
            coupons = ",".join(sorted(r["coupons"]))
            if coupons not in candidates or not candidates[coupons]["complete"]:
                candidates[coupons] = {
                    "members": sorted(r["coupons"]), "gross": r["gross"], "fees": r["fees"], "net": r["net"],
                    "complete": True, "group_total": None, "day": date.fromisoformat(r["date"]),
                }
        identified = 0
        for g in pending:
            last = date.fromisoformat(g["last_date"])
            matches = []
            for c in candidates.values():
                if not (last <= c["day"] <= last + timedelta(days=_DETAIL_MATCH_DAYS)):
                    continue
                if c["complete"]:
                    same = all(abs(c[k] - (g[k] or 0.0)) < 0.005 for k in ("gross", "fees", "net"))
                else:
                    # reported_group_total es el neto del grupo (ver cupones_append).
                    same = c["group_total"] is not None and abs(c["group_total"] - (g["net"] or 0.0)) < 0.005
                if same:
                    matches.append(c)
            if len(matches) == 1:
                conn.execute(
                    "UPDATE cupon_detail_groups SET coupons = ?, coupons_source = 'monto' WHERE id = ?",
                    (",".join(matches[0]["members"]), g["id"]),
                )
                identified += 1
        if identified:
            conn.commit()
        return identified
    finally:
        conn.close()


def get_detail_groups():
    """Los grupos de detalle cargados, con sus batches: [{"id", "coupons" (lista o None), "gross", "fees", "net", "first_date", "last_date", "batches"}]."""
    conn = _connect()
    try:
        groups = [dict(g) for g in conn.execute("SELECT * FROM cupon_detail_groups ORDER BY last_date, id")]
        batches = conn.execute(
            "SELECT group_id, batch_date, batch, gross, fees, net FROM cupon_detail_batches ORDER BY batch_date, idx"
        ).fetchall()
    finally:
        conn.close()
    by_group = {}
    for b in batches:
        by_group.setdefault(b["group_id"], []).append({k: b[k] for k in ("batch_date", "batch", "gross", "fees", "net")})
    for g in groups:
        g["coupons"] = g["coupons"].split(",") if g["coupons"] else None
        g["batches"] = by_group.get(g["id"], [])
    return groups


def get_detail_batches_between(start, end):
    """Batches del detalle con fecha de venta entre start y end (ISO), cada uno con los DDC de su grupo."""
    conn = _connect()
    try:
        rows = conn.execute(
            """
            SELECT b.batch_date, b.batch, b.gross, b.fees, b.net, g.coupons
            FROM cupon_detail_batches b JOIN cupon_detail_groups g ON g.id = b.group_id
            WHERE b.batch_date BETWEEN ? AND ? ORDER BY b.batch_date, b.idx
            """,
            (start, end),
        ).fetchall()
    finally:
        conn.close()
    return [dict(r) for r in rows]


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
    carga del reporte mensual manda) -- devuelve (inserted, updated,
    repeated), donde `repeated` son los DDC que el propio reporte trae más
    de una vez (auditoría 2026-09, webapp.py:4391): antes la última fila
    pisaba a la primera sin aviso. Si una de las repetidas es la fila de un
    batch (montos en 0) y otra trae montos reales, gana la de montos reales.

    Al final corre backfill_grouped_cupones_from_eft: el reporte es
    acumulativo y cada resubida vuelve a poner en 0 a los miembros de un
    batch, que así recuperan enseguida el monto que trae su EFT.
    """
    unique = {}
    repeated = []
    for rec in records:
        coupon_id = (rec.get("coupon") or "").strip().upper()
        if not coupon_id:
            continue
        previous = unique.get(coupon_id)
        if previous is None:
            unique[coupon_id] = rec
            continue
        if coupon_id not in repeated:
            repeated.append(coupon_id)
        previous_is_zero = not any((previous.get(k) or 0.0) for k in ("gross", "fees", "net"))
        if previous_is_zero:
            unique[coupon_id] = rec

    conn = _connect()
    try:
        inserted = 0
        updated = 0
        now = datetime.utcnow().isoformat()
        for coupon_id, rec in unique.items():
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
    finally:
        conn.close()
    backfill_grouped_cupones_from_eft()
    autolink_missing_ddc_by_amount()
    identify_detail_groups()
    return inserted, updated, repeated


def _fmt_money_pdf(value):
    """Mismo criterio de signo que chase_rules._fmt_money_pdf -- "-$" antes del monto, nunca "$-"."""
    if value is None:
        return "—"
    if value < 0:
        return "-${:,.2f}".format(abs(value))
    return "${:,.2f}".format(value)


def _month_bounds(year, month):
    first = date(year, month, 1)
    last = date(year, month, calendar.monthrange(year, month)[1])
    prev_first = date(year - 1, 12, 1) if month == 1 else date(year, month - 1, 1)
    return first, last, prev_first


def coupon_month_summary(year, month):
    """
    Resumen de los cupones de un mes, sin listarlos uno por uno (pedido del
    usuario, 2026-10-05): "un resumen de la cantidad de cupones que hay en
    ese mes, cuántos se aplicaron a cada EFT y cuáles y cuántos quedaron
    pendientes... debe aclararse cuánta es la cantidad de cupones del mes
    anterior que se aplicó... para así no mezclar los números".

    Cada cupón cuenta según SU fecha: del mes anterior o de este mes (uno
    más viejo solo aparece si un EFT de este mes lo pagó). Está aplicado en
    este mes si alguna línea de un EFT con fecha de este mes lo pagó, y
    pendiente al cierre si ningún EFT con fecha hasta fin de mes lo pagó
    (aunque uno del mes siguiente ya lo haya pagado: pasa al mes siguiente).
    En lo aplicado, el monto es lo que pagó el EFT; en el resto, el Net del
    cupón, y un cupón de batch que sigue en 0 cuenta el saldo del batch una
    sola vez (mismo criterio que get_cupones_with_status).

    Devuelve {cantidad, monto} del mes anterior y de este mes, el detalle
    por EFT y los pendientes al cierre agrupados por fecha del cupón.
    """
    first, last, prev_first = _month_bounds(year, month)
    cupones = {}
    for c in get_cupones_with_status():
        parsed = _parse_cupon_date(c.get("date"))
        if parsed:
            cupones[c["coupon_id"]] = {**c, "day": parsed.date()}

    conn = _connect()
    try:
        lines = conn.execute(
            """
            SELECT ec.coupon, ec.paid_amount, ed.id AS deposit_id, ed.eft_date
            FROM eft_coupons ec JOIN eft_deposits ed ON ed.id = ec.deposit_id
            """
        ).fetchall()
    finally:
        conn.close()
    paid_on = {}  # cupón -> [(fecha del EFT, pagado)]
    for line in lines:
        eft_day = _parse_eft_date(line["eft_date"])
        if line["coupon"] and eft_day:
            paid_on.setdefault(line["coupon"], []).append((eft_day.date(), line["paid_amount"] or 0.0))

    def bucket(coupon_id):
        c = cupones.get(coupon_id)
        if c is None:
            return "unknown"
        if c["day"] < prev_first:
            return "older"
        if c["day"] < first:
            return "prev"
        return "this" if c["day"] <= last else "later"

    def applied_until(coupon_id, until):
        return any(day <= until for day, _ in paid_on.get(coupon_id, ()))

    def paid_this_month(coupon_id):
        return sum(paid for day, paid in paid_on.get(coupon_id, ()) if first <= day <= last)

    def applied_this_month(coupon_id):
        return any(first <= day <= last for day, _ in paid_on.get(coupon_id, ()))

    counted_groups = set()

    def open_amount(coupon_id):
        """Net del cupón; uno de batch sin resolver, el saldo del batch una sola vez."""
        c = cupones[coupon_id]
        net = c.get("net") or 0.0
        if net or c.get("group_remaining") is None:
            return net
        if c.get("reported_group_text") in counted_groups:
            return 0.0
        counted_groups.add(c.get("reported_group_text"))
        return c["group_remaining"]

    def tally(amounts):
        return {"count": len(amounts), "amount": round(sum(amounts), 2)}

    by_bucket = {"prev": [], "this": [], "older": []}
    for coupon_id in cupones:
        b = bucket(coupon_id)
        if b in by_bucket:
            by_bucket[b].append(coupon_id)

    prev_open = [cid for cid in by_bucket["prev"] if not applied_until(cid, first - timedelta(days=1))]
    prev_applied = [cid for cid in prev_open if applied_this_month(cid)]
    prev_still = [cid for cid in prev_open if not applied_until(cid, last)]
    this_applied = [cid for cid in by_bucket["this"] if applied_this_month(cid)]
    this_pending = [cid for cid in by_bucket["this"] if not applied_until(cid, last)]
    older_applied = [cid for cid in by_bucket["older"] if applied_this_month(cid)]

    # Primero los pendientes: el saldo de un batch queda contado ahí (una
    # sola vez) y los totales reusan el mismo monto.
    amounts = {cid: open_amount(cid) for cid in prev_still + this_pending}
    for cid in prev_open + by_bucket["this"]:
        if cid not in amounts:
            amounts[cid] = open_amount(cid)

    pending_by_day = {}
    for cid in prev_still + this_pending:
        entry = pending_by_day.setdefault(cupones[cid]["day"], [0, 0.0])
        entry[0] += 1
        entry[1] += amounts[cid]

    per_eft = []
    for entry in get_month_deposits(year, month):
        dep = entry["deposit"]
        # "unknown": renglón del EFT sin N° de cupón (o con uno que no está
        # cargado en Cupones); cuenta en el total del EFT, aparte.
        groups = {"prev": {}, "this": {}, "older": {}, "unknown": {}}
        for index, line in enumerate(lines):
            if line["deposit_id"] != dep["id"]:
                continue
            b = bucket(line["coupon"]) if line["coupon"] else "unknown"
            key = b if b in groups else "unknown"
            coupon_key = line["coupon"] or f"sin N° {index}"
            groups[key][coupon_key] = groups[key].get(coupon_key, 0.0) + (line["paid_amount"] or 0.0)
        row = {"rcv": dep.get("rcv_number"), "eft_date": _parse_eft_date(dep.get("eft_date"))}
        for key, paid in groups.items():
            row[key] = {"count": len(paid), "amount": round(sum(paid.values()), 2)}
        row["total"] = {
            "count": sum(row[k]["count"] for k in groups),
            "amount": round(sum(row[k]["amount"] for k in groups), 2),
        }
        per_eft.append(row)
    per_eft.sort(key=lambda r: (r["eft_date"] is None, r["eft_date"] or datetime.min))

    return {
        "unknown_applied": {
            "count": sum(r["unknown"]["count"] for r in per_eft),
            "amount": round(sum(r["unknown"]["amount"] for r in per_eft), 2),
        },
        "prev_open": tally([amounts[cid] for cid in prev_open]),
        "prev_applied": tally([paid_this_month(cid) for cid in prev_applied]),
        "prev_still": tally([amounts[cid] for cid in prev_still]),
        "this_total": tally([amounts[cid] for cid in by_bucket["this"]]),
        "this_applied": tally([paid_this_month(cid) for cid in this_applied]),
        "this_pending": tally([amounts[cid] for cid in this_pending]),
        "older_applied": tally([paid_this_month(cid) for cid in older_applied]),
        "pending_by_day": [
            {"day": day, "count": count, "amount": round(amount, 2)}
            for day, (count, amount) in sorted(pending_by_day.items())
        ],
        "per_eft": per_eft,
    }


def cupones_month_view(year, month):
    """
    El historial de Cupones de un mes (pedido del usuario, 2026-10-06:
    "también va a ir por meses... para poder seguir un control de los
    cupones que quedaron pendientes un mes, el límite debería ser el mes que
    se selecciona"). Mismo criterio que coupon_month_summary:

    - "applied": los cupones que pagó un EFT con fecha de este mes, juntos
      por EFT con su total (order_cupones_by_eft); `prev_month` marca los
      que son de un mes anterior.
    - "pending": los cupones de este mes o del anterior que ningún EFT con
      fecha hasta fin de mes pagó (ver _pending_rows), con su total y
      subtotal por mes del cupón; `later` es el EFT que los pagó después.
    - "next_month": los depósitos posteriores al cierre con ventas de este
      mes (el detalle de cupones da el día de venta de cada batch): lo que
      se vendió con tarjeta en el mes y entró como cupón el mes siguiente.
      Cada DDC del depósito va en su renglón con su Gross/Fee/Net de Cupones
      (completados por el EFT que lo pagó), sin decir qué EFT lo aplicó
      (pedido del usuario, 2026-10-06).
    """
    first, last, prev_first = _month_bounds(year, month)
    cupones = get_cupones_flat()
    applied, pending, coupon_info = [], [], {}
    cupon_by_id = {cp["coupon_id"]: cp for cp in cupones}
    for cp in cupones:
        match = cp.get("match")
        eft_day = _parse_eft_date(match.get("eft_date")) if match else None
        eft_day = eft_day.date() if eft_day else None
        cp_day = _parse_cupon_date(cp.get("date"))
        cp_day = cp_day.date() if cp_day else None
        coupon_info[cp["coupon_id"]] = (cp_day, match)
        if eft_day and first <= eft_day <= last:
            cp["prev_month"] = bool(cp_day and cp_day < first)
            applied.append(cp)
        elif cp_day and prev_first <= cp_day <= last and not (eft_day and eft_day <= last):
            cp["later"] = {"rcv": match.get("rcv_number"), "eft_date": eft_day} if eft_day else None
            pending.append(cp)
    applied = order_cupones_by_eft(applied)
    detail_groups = get_detail_groups()
    pending_rows = _pending_rows(pending, cupones, detail_groups)
    pending_totals = _sum_rows(pending_rows)
    pending_by_month = {}
    for row in pending_rows:
        pending_by_month.setdefault((row["day"].year, row["day"].month), []).append(row)

    next_month = []
    for group in detail_groups:
        own = [b for b in group["batches"] if first.isoformat() <= b["batch_date"] <= last.isoformat()]
        if not own:
            continue
        # Fecha del depósito: la de sus DDC en Cupones; sin DDC identificados,
        # el último batch (el depósito sale ese día o uno o dos después).
        days = [coupon_info[c][0] for c in group["coupons"] or [] if c in coupon_info and coupon_info[c][0]]
        deposit = min(days) if days else None
        if (deposit or date.fromisoformat(group["last_date"])) <= last:
            continue
        ddc_rows = []
        for c in group["coupons"] or []:
            cp = cupon_by_id.get(c, {})
            has_amount = any(cp.get(key) for key in ("gross", "fees", "net"))
            ddc_rows.append({
                "coupon": c, "date_display": cp.get("date_display"),
                "gross": cp.get("gross") if has_amount else None,
                "fees": cp.get("fees") if has_amount else None,
                "net": cp.get("net") if has_amount else None,
            })
        next_month.append({
            "coupons": group["coupons"], "deposit": deposit, "ddc_rows": ddc_rows,
            "from": min(b["batch_date"] for b in own), "to": max(b["batch_date"] for b in own),
            "gross": round(sum(b["gross"] or 0.0 for b in own), 2),
            "net": round(sum(b["net"] or 0.0 for b in own), 2),
            "group_gross": group["gross"],
        })
    next_month.sort(key=lambda g: (g["deposit"] or date.fromisoformat(g["to"]), g["from"]))

    return {
        "applied": applied,
        "applied_count": len(applied),
        "applied_net": round(sum(cp.get("net") or 0.0 for cp in applied), 2),
        "eft_count": sum(1 for cp in applied if cp.get("eft_block_total")),
        "pending": pending_rows,
        "pending_count": len(pending),
        "pending_totals": pending_totals,
        "pending_by_month": [
            {"year": y, "month": m, **_sum_rows(rows)} for (y, m), rows in sorted(pending_by_month.items())
        ] if len(pending_by_month) > 1 else [],
        "next_month": next_month,
        "next_month_gross": round(sum(g["gross"] for g in next_month), 2),
        "next_month_totals": {
            key: round(sum(r[key] or 0.0 for g in next_month for r in g["ddc_rows"]), 2)
            for key in ("gross", "fees", "net")
        },
        "next_month_incomplete": any(r["net"] is None for g in next_month for r in g["ddc_rows"]),
    }


def _pending_rows(pending, cupones, detail_groups):
    """
    Los pendientes al cierre como renglones que se pueden sumar (pedido del
    usuario, 2026-10-06: "que te diga cuánto queda para saberlo a simple
    vista y no tener que sumarlo al ojo"). Un DDC con monto propio es un
    renglón; los de un mismo grupo que todavía están en 0 (hasta que un EFT
    los paga) van juntos en uno, con lo que le falta al grupo: el Net de
    `group_remaining` y el Gross/Fee del depósito en el detalle de cupones
    menos lo de los DDC del grupo que ya tienen monto (None si no hay
    detalle). [{"day", "date_display", "coupons", "gross", "fees", "net",
    "laters": [{"coupon", "rcv", "eft_date"}]}]
    """
    detail_by_coupons = {",".join(sorted(g["coupons"])): g for g in detail_groups if g["coupons"]}
    by_group_text = {}
    for cp in cupones:
        if cp.get("reported_group_text"):
            by_group_text.setdefault(cp["reported_group_text"], []).append(cp)

    rows, group_rows = [], {}
    for cp in pending:
        later = cp.get("later")
        laters = [{"coupon": cp["coupon_id"], **later}] if later else []
        own = any(cp.get(key) for key in ("gross", "fees", "net"))
        group_text = cp.get("reported_group_text")
        if own or not group_text or cp.get("group_remaining") is None:
            rows.append({
                "day": _parse_cupon_date(cp.get("date")).date(), "date_display": cp.get("date_display"),
                "coupons": [cp["coupon_id"]], "gross": cp.get("gross") or 0.0, "fees": cp.get("fees") or 0.0,
                "net": cp.get("net") or 0.0, "laters": laters,
            })
            continue
        row = group_rows.get(group_text)
        if row is None:
            members = by_group_text.get(group_text, [])
            resolved = [m for m in members if any(m.get(key) for key in ("gross", "fees", "net"))]
            detail = detail_by_coupons.get(",".join(sorted(m["coupon_id"] for m in members)))
            row = {
                "day": _parse_cupon_date(cp.get("date")).date(), "date_display": cp.get("date_display"),
                "coupons": [], "net": cp["group_remaining"], "laters": [],
                "gross": round(detail["gross"] - sum(m.get("gross") or 0.0 for m in resolved), 2) if detail else None,
                "fees": round(detail["fees"] - sum(m.get("fees") or 0.0 for m in resolved), 2) if detail else None,
            }
            group_rows[group_text] = row
            rows.append(row)
        row["coupons"].append(cp["coupon_id"])
        row["laters"].extend(laters)
    return rows


def _sum_rows(rows):
    """Gross/Fee/Net sumados; `incomplete` si algún grupo no tiene Gross/Fee (sin detalle de cupones)."""
    return {
        "count": sum(len(r["coupons"]) for r in rows),
        "gross": round(sum(r["gross"] or 0.0 for r in rows), 2),
        "fees": round(sum(r["fees"] or 0.0 for r in rows), 2),
        "net": round(sum(r["net"] or 0.0 for r in rows), 2),
        "incomplete": any(r["gross"] is None or r["fees"] is None for r in rows),
    }


def build_eft_pdf_report(year, month, dest_path):
    """
    PDF del módulo nuevo "Reportes" (pedido explícito del usuario,
    2026-09-17): "Lo mismo quiero que hagas con el Reporte de los EFT
    incluyendo todos los datos que se tengan del mes que se selecciono, y
    lo mismo estar incluido en ese PDF los cupones cargados hasta ese
    momento y cuanto acumulan". A diferencia de Chase (que sí se resume
    por Detalle), acá "todos los datos" significa una fila por CADA EFT
    (RCV) cargado ese mes -- no se agrupa nada, se listan todos.

    Secciones (ver `pdf_export.build_multi_section_pdf`):
    1. "EFT del mes" -- un renglón por depósito (RCV/Fecha/Gross/Fees/Net)
       más una fila TOTAL. El período que se imprime en el membrete usa la
       fecha MÍNIMA/MÁXIMA real de los EFT de ese mes (mismo criterio que
       el PDF de Chase) -- nunca asume que el mes esté completo.
    2-4. Cupones, sin listarlos uno por uno (2026-10-05, ver
       coupon_month_summary): resumen del mes anterior y de este mes,
       cupones aplicados en cada EFT y pendientes al cierre por fecha.
    """
    from pdf_export import build_multi_section_pdf

    backfill_grouped_cupones_from_eft()
    autolink_missing_ddc_by_amount()
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

    # Cupones (pedido del usuario, 2026-10-05): ya no se lista cupón por
    # cupón; un resumen del mes que separa lo del mes anterior de lo de
    # este mes, lo aplicado en cada EFT y lo que pasa al mes siguiente (ver
    # coupon_month_summary).
    summary = coupon_month_summary(year, month)
    prev_label = f"{12 if month == 1 else month - 1:02d}/{year - 1 if month == 1 else year}"
    this_label = f"{month:02d}/{year}"
    next_label = f"{1 if month == 12 else month + 1:02d}/{year + 1 if month == 12 else year}"

    def count_amount(item):
        return [str(item["count"]), _fmt_money_pdf(item["amount"])]

    def add_up(items):
        return {"count": sum(i["count"] for i in items), "amount": round(sum(i["amount"] for i in items), 2)}

    # Cada renglón dice de qué mes es el cupón y en qué mes se aplicó: la
    # tabla se imprime centrada, así que no hay sangría que lo aclare.
    resumen_rows = [
        [f"Cupones del {prev_label} sin aplicar al 01/{this_label}"] + count_amount(summary["prev_open"]),
        [f"Del {prev_label}: aplicados en EFT del {this_label}"] + count_amount(summary["prev_applied"]),
        [f"Del {prev_label}: siguen pendientes"] + count_amount(summary["prev_still"]),
        [f"Cupones con fecha del {this_label}"] + count_amount(summary["this_total"]),
        [f"Del {this_label}: aplicados en EFT del {this_label}"] + count_amount(summary["this_applied"]),
        [f"Del {this_label}: pendientes, pasan al {next_label}"] + count_amount(summary["this_pending"]),
    ]
    applied_parts = [summary["prev_applied"], summary["this_applied"]]
    if summary["older_applied"]["count"]:
        resumen_rows.append(
            [f"Anteriores al {prev_label}: aplicados en EFT del {this_label}"] + count_amount(summary["older_applied"])
        )
        applied_parts.append(summary["older_applied"])
    if summary["unknown_applied"]["count"]:
        resumen_rows.append(
            [f"Renglones de EFT del {this_label} sin N° de cupón"] + count_amount(summary["unknown_applied"])
        )
        applied_parts.append(summary["unknown_applied"])
    resumen_rows.append([f"Total aplicado en EFT del {this_label}"] + count_amount(add_up(applied_parts)))

    per_eft = summary["per_eft"]
    per_eft_rows = [
        [
            row["rcv"] or "—",
            row["eft_date"].strftime("%d/%m/%Y") if row["eft_date"] else "—",
            *count_amount(row["prev"]),
            *count_amount(row["this"]),
            *count_amount(row["total"]),
        ]
        for row in per_eft
    ]
    if per_eft_rows:
        per_eft_rows.append(
            [
                "TOTAL", "",
                *count_amount(add_up([r["prev"] for r in per_eft])),
                *count_amount(add_up([r["this"] for r in per_eft])),
                *count_amount(add_up([r["total"] for r in per_eft])),
            ]
        )
    other_count = sum(r["older"]["count"] + r["unknown"]["count"] for r in per_eft)

    pending = summary["pending_by_day"]
    pending_rows = [[d["day"].strftime("%d/%m/%Y"), str(d["count"]), _fmt_money_pdf(d["amount"])] for d in pending]
    if pending_rows:
        pending_rows.append(["TOTAL", *count_amount(add_up(pending))])

    per_eft_heading = "Cupones aplicados en cada EFT"
    pending_heading = "Pendientes al cierre del mes (pasan al mes siguiente)"
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
            "heading": "Resumen de cupones",
            "headers": ["Detalle", "Cupones", "Monto"],
            "rows": resumen_rows,
            "col_widths_mm": [150, 40, 55],
            "bold_last_row": True,
            "footnote": (
                "Cada cupón cuenta según su propia fecha, así lo del mes anterior no se mezcla con lo de este mes. "
                "En los aplicados, el monto es lo que pagó el EFT; en el resto, el Net del cupón."
            ),
        },
        {
            "heading": per_eft_heading,
            "headers": ["RCV", "Fecha", f"Del {prev_label}", "Monto", f"Del {this_label}", "Monto", "Total", "Monto"],
            "rows": per_eft_rows,
            "col_widths_mm": [36, 30, 26, 36, 26, 36, 26, 36],
            "bold_last_row": True,
            "footnote": (
                f"El Total incluye {other_count} renglón(es) de cupones anteriores al {prev_label} o sin N° de cupón "
                "(ver el Resumen)." if other_count else None
            ),
        }
        if per_eft_rows
        else {"heading": per_eft_heading, "note": "No hay EFT cargados este mes."},
        {
            "heading": pending_heading,
            "headers": ["Fecha del cupón", "Cupones", "Monto"],
            "rows": pending_rows,
            "col_widths_mm": [80, 40, 55],
            "bold_last_row": True,
        }
        if pending_rows
        else {"heading": pending_heading, "note": "No quedó ningún cupón del mes anterior ni de este mes sin aplicar."},
    ]
    build_multi_section_pdf(dest_path, title, sections, period_label=period_label, company_header=True)
    return dest_path
