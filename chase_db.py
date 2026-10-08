"""
Persistencia de los movimientos de Chase Bank guardados dentro de la propia
página (ver CLAUDE.md -> "Conversión de Chase Bank a Carga de Datos").
Capa de datos pura, sin ningún conocimiento de CSV/Excel ni de las reglas de
categorización -- chase_rules.py hace la lectura/categorización, este módulo
solo guarda/consulta.

reportes_data/chase.db (gitignored, mismo directorio que reportes_diarios.db
y lottery.db) -- una fila por movimiento bancario real. Cada carga se puede
repetir (el mismo extracto, uno solapado, o el historial completo del banco
de una) sin duplicar nada: la clave natural es fecha + descripción + monto,
y categorizar de nuevo (ej. tras editar una regla) simplemente actualiza el
Detalle ya guardado.
"""

import os
import sqlite3
from datetime import date, datetime, timedelta

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "chase.db")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    # timeout=30 + WAL -- ver reportes_db.py: necesario desde que las cargas
    # en segundo plano (jobs.py) pueden escribir de verdad en paralelo.
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    _ensure_schema(conn)
    return conn


def _ensure_schema(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS chase_transactions (
            posting_date TEXT NOT NULL,
            description TEXT NOT NULL,
            amount REAL NOT NULL,
            balance REAL,
            detalle TEXT,
            type TEXT,
            source_filename TEXT,
            updated_at TEXT,
            detalle_source TEXT,
            supplier_key TEXT,
            supplier_source TEXT,
            occurrence INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY (posting_date, description, amount, occurrence)
        )
        """
    )
    _ensure_columns(conn)
    _migrate_occurrence_pk(conn)
    conn.commit()


def _migrate_occurrence_pk(conn):
    """
    La clave era (fecha, descripción, monto): dos movimientos reales
    idénticos del mismo día (dos MVNT de $2.50, dos UBER iguales) se
    fusionaban y se perdía uno (auditoría 2026-09, 3 perdidos en agosto).
    Ahora la clave suma `occurrence` (0, 1, ... según el orden dentro del
    extracto) -- recargar el mismo extracto sigue siendo idempotente. Se
    reconstruye la tabla una sola vez; las filas ya guardadas quedan con 0.
    """
    existing = {row[1] for row in conn.execute("PRAGMA table_info(chase_transactions)")}
    if "occurrence" in existing:
        return
    cols = ("posting_date, description, amount, balance, detalle, type, source_filename, "
            "updated_at, detalle_source, supplier_key, supplier_source")
    conn.execute("ALTER TABLE chase_transactions RENAME TO chase_transactions_old")
    conn.execute(
        """
        CREATE TABLE chase_transactions (
            posting_date TEXT NOT NULL,
            description TEXT NOT NULL,
            amount REAL NOT NULL,
            balance REAL,
            detalle TEXT,
            type TEXT,
            source_filename TEXT,
            updated_at TEXT,
            detalle_source TEXT,
            supplier_key TEXT,
            supplier_source TEXT,
            occurrence INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY (posting_date, description, amount, occurrence)
        )
        """
    )
    conn.execute(f"INSERT INTO chase_transactions ({cols}) SELECT {cols} FROM chase_transactions_old")
    conn.execute("DROP TABLE chase_transactions_old")


def _ensure_columns(conn):
    """Mini-migrador -- la base real ya existía en disco antes de agregar detalle_source."""
    existing = {row[1] for row in conn.execute("PRAGMA table_info(chase_transactions)")}
    if "detalle_source" not in existing:
        conn.execute("ALTER TABLE chase_transactions ADD COLUMN detalle_source TEXT")
        # Todo lo ya categorizado hasta ahora vino de una regla -- lo sin
        # categorizar queda NULL (todavía nadie lo tocó a mano).
        conn.execute(
            "UPDATE chase_transactions SET detalle_source = 'rule' WHERE detalle IS NOT NULL AND detalle != ''"
        )
    # supplier_key/supplier_source (2026-09-21, pedido explícito del usuario:
    # "los pagos a proveedores se van a mover al proveedor directamente que
    # sale en el asiento") -- vínculo INDEPENDIENTE del Detalle, entre un
    # movimiento de Chase y un proveedor puntual de proveedores_db.py.
    # supplier_key nunca es un Excel/sheet_name, es la clave real que ya usa
    # proveedores_db (mismo criterio que detalle/detalle_source: "rule"
    # cuando lo resolvió proveedores.match_supplier_for_chase_description,
    # "manual" cuando lo eligió el usuario a mano y queda pegado para
    # siempre, ver set_manual_supplier).
    if "supplier_key" not in existing:
        conn.execute("ALTER TABLE chase_transactions ADD COLUMN supplier_key TEXT")
    if "supplier_source" not in existing:
        conn.execute("ALTER TABLE chase_transactions ADD COLUMN supplier_source TEXT")


def _row_to_dict(row):
    return {
        "posting_date": row["posting_date"],
        "description": row["description"],
        "amount": row["amount"],
        "balance": row["balance"],
        "detalle": row["detalle"],
        "detalle_source": row["detalle_source"],
        "type": row["type"],
        "source_filename": row["source_filename"],
        "supplier_key": row["supplier_key"],
        "supplier_source": row["supplier_source"],
    }


def upsert_transactions(rows, source_filename=None):
    """
    Guarda (o recategoriza) una tanda de movimientos ya extraídos por
    chase_rules.extract_chase_transactions -- cada dict trae posting_date
    (date), description, amount, balance, detalle, type. Se puede llamar con
    el extracto completo del banco de una sola vez, o de a poco (un rango
    corto por vez, como se cargaba con el Excel) -- la clave natural
    (fecha+descripción+monto) hace que cualquiera de las dos formas termine
    en el mismo resultado, sin duplicar ni pisar lo ya guardado de otro
    rango.

    Un movimiento ya guardado se recategoriza (Detalle actualizado) si se
    vuelve a cargar el mismo extracto y las reglas cambiaron desde la
    última carga -- EXCEPTO si el usuario ya lo corrigió a mano
    (detalle_source="manual", ver set_manual_detalle) -- esa corrección
    queda pegada para siempre, ninguna carga futura la pisa. El vínculo a un
    proveedor puntual (`row["supplier_key"]`, opcional -- ver
    proveedores.match_supplier_for_chase_description) sigue exactamente el
    mismo criterio de forma INDEPENDIENTE, vía supplier_key/supplier_source.

    Devuelve (inserted, updated).
    """
    conn = _connect()
    try:
        inserted = 0
        updated = 0
        now = datetime.utcnow().isoformat()
        # Clave de cada movimiento del archivo, y el rango de fechas que cubre.
        keyed = []
        seen = {}
        for row in rows:
            posting_date = row["posting_date"]
            if isinstance(posting_date, (date, datetime)):
                posting_date = posting_date.isoformat() if isinstance(posting_date, date) else posting_date.date().isoformat()
            key = (posting_date, row["description"], row["amount"])
            occurrence = seen.get(key, 0)
            seen[key] = occurrence + 1
            keyed.append((row, posting_date, occurrence))
        batch_keys = {(d, r["description"], r["amount"], o) for r, d, o in keyed}
        first_day = min((d for _r, d, _o in keyed), default=None)
        last_day = max((d for _r, d, _o in keyed), default=None)
        renamed = set()
        for row, posting_date, occurrence in keyed:
            cur = conn.execute(
                "SELECT 1 FROM chase_transactions WHERE posting_date = ? AND description = ? AND amount = ? AND occurrence = ?",
                (posting_date, row["description"], row["amount"], occurrence),
            )
            exists = cur.fetchone() is not None
            if not exists:
                stale = _stale_version(conn, posting_date, row["amount"], first_day, last_day, batch_keys, renamed)
                if stale is not None:
                    # El mismo movimiento con la descripción (o la fecha) que
                    # Chase mostraba en una descarga anterior: se le pone la
                    # clave nueva y el upsert de abajo lo actualiza (lo
                    # corregido a mano se conserva).
                    conn.execute(
                        "UPDATE chase_transactions SET posting_date = ?, description = ?, occurrence = ? WHERE rowid = ?",
                        (posting_date, row["description"], occurrence, stale),
                    )
                    renamed.add(stale)
                    exists = True
            conn.execute(
                """
                INSERT INTO chase_transactions
                    (posting_date, description, amount, balance, detalle, type, source_filename, updated_at, detalle_source, supplier_key, supplier_source, occurrence)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(posting_date, description, amount, occurrence) DO UPDATE SET
                    balance = excluded.balance,
                    type = excluded.type,
                    source_filename = excluded.source_filename,
                    updated_at = excluded.updated_at,
                    detalle = CASE WHEN chase_transactions.detalle_source IN ('manual', 'deposito')
                                   THEN chase_transactions.detalle ELSE excluded.detalle END,
                    detalle_source = CASE WHEN chase_transactions.detalle_source IN ('manual', 'deposito')
                                          THEN chase_transactions.detalle_source ELSE excluded.detalle_source END,
                    supplier_key = CASE WHEN chase_transactions.supplier_source = 'manual'
                                        THEN chase_transactions.supplier_key ELSE excluded.supplier_key END,
                    supplier_source = CASE WHEN chase_transactions.supplier_source = 'manual'
                                           THEN chase_transactions.supplier_source ELSE excluded.supplier_source END
                """,
                (
                    posting_date, row["description"], row["amount"], row.get("balance"),
                    row.get("detalle"), row.get("type"), source_filename, now,
                    "rule" if row.get("detalle") else None,
                    row.get("supplier_key"),
                    "rule" if row.get("supplier_key") else None,
                    occurrence,
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


# Cuántos días puede correrse la fecha de un movimiento entre una descarga y
# otra (un pendiente con la fecha de la compra pasa a la fecha en que se asentó).
_STALE_DAYS = 5


def _stale_version(conn, posting_date, amount, first_day, last_day, batch_keys, taken):
    """
    rowid de un movimiento ya guardado que es una versión vieja del que llega
    (pedido del usuario, 2026-10-07: el depósito de $2,057 del 22/09 quedó
    dos veces, "DEPOSIT" de la descarga del 29/09 y "DEPOSIT  ID NUMBER
    293043" de la del 07/10). Chase cambia la descripción de los movimientos
    recientes (y a veces la fecha de un pendiente) entre una descarga y otra.
    Es versión vieja un movimiento guardado del mismo importe, a no más de
    _STALE_DAYS días, dentro del rango de fechas del archivo nuevo y que el
    archivo nuevo NO trae tal cual: si el archivo cubre esa fecha y no lo
    trae, es que Chase lo muestra distinto. El más cercano en fecha.
    """
    if first_day is None:
        return None
    day = date.fromisoformat(posting_date)
    low = max(first_day, (day - timedelta(days=_STALE_DAYS)).isoformat())
    high = min(last_day, (day + timedelta(days=_STALE_DAYS)).isoformat())
    candidates = conn.execute(
        "SELECT rowid AS rowid, posting_date, description, amount, occurrence FROM chase_transactions "
        "WHERE amount = ? AND posting_date BETWEEN ? AND ?",
        (amount, low, high),
    ).fetchall()
    stale = [c for c in candidates
             if c["rowid"] not in taken
             and (c["posting_date"], c["description"], c["amount"], c["occurrence"]) not in batch_keys]
    if not stale:
        return None
    stale.sort(key=lambda c: (abs((date.fromisoformat(c["posting_date"]) - day).days), c["rowid"]))
    return stale[0]["rowid"]


def set_manual_detalle(posting_date, description, amount, detalle):
    """
    Corrección manual de un movimiento puntual -- pedido explícito del
    usuario (2026-09-12): "los datos sin categorizar del chase se puedan
    categorizar". Marca detalle_source="manual" para que una recarga
    posterior del mismo extracto no la pise (ver upsert_transactions).

    Devuelve True si encontró y actualizó el movimiento, False si no existe.
    """
    if isinstance(posting_date, (date, datetime)):
        posting_date = posting_date.isoformat() if isinstance(posting_date, date) else posting_date.date().isoformat()
    detalle = (detalle or "").strip() or None
    # Si lo borra (deja el campo vacío), vuelve a quedar disponible para que
    # una futura carga lo recategorice solo -- "manual" queda reservado para
    # cuando de verdad hay un valor puesto a mano.
    detalle_source = "manual" if detalle else None
    conn = _connect()
    try:
        cur = conn.execute(
            """
            UPDATE chase_transactions SET detalle = ?, detalle_source = ?, updated_at = ?
            WHERE posting_date = ? AND description = ? AND amount = ?
            """,
            (detalle, detalle_source, datetime.utcnow().isoformat(), posting_date, description, amount),
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def set_manual_supplier(posting_date, description, amount, supplier_key):
    """
    Vincula (o desvincula) a mano un movimiento puntual con un proveedor de
    proveedores_db -- pedido explícito del usuario (2026-09-21): un cheque
    sin ninguna descripción útil ("CHECK 1770") nunca va a poder resolverse
    solo por ninguna regla de palabra clave, así que hace falta poder
    asignarlo directo. Igual que set_manual_detalle: queda pegado
    (supplier_source="manual") para que una recarga del mismo extracto, o
    una recategorización retroactiva (ver recategorize_all), no lo pise.
    Dejarlo en blanco lo vuelve a dejar disponible para que una regla lo
    resuelva solo en el futuro.

    Devuelve True si encontró y actualizó el movimiento, False si no existe.
    """
    if isinstance(posting_date, (date, datetime)):
        posting_date = posting_date.isoformat() if isinstance(posting_date, date) else posting_date.date().isoformat()
    supplier_key = (supplier_key or "").strip() or None
    supplier_source = "manual" if supplier_key else None
    conn = _connect()
    try:
        cur = conn.execute(
            """
            UPDATE chase_transactions SET supplier_key = ?, supplier_source = ?, updated_at = ?
            WHERE posting_date = ? AND description = ? AND amount = ?
            """,
            (supplier_key, supplier_source, datetime.utcnow().isoformat(), posting_date, description, amount),
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def get_supplier_transactions(supplier_key):
    """Todos los movimientos (cualquier mes) ya vinculados a un proveedor puntual, más reciente primero."""
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM chase_transactions WHERE supplier_key = ? ORDER BY posting_date DESC",
            (supplier_key,),
        ).fetchall()
        return [_row_to_dict(row) for row in rows]
    finally:
        conn.close()


def get_supplier_payment_totals():
    """{"supplier_key": {"count": N, "total": suma}} -- para la grilla de proveedores guardados."""
    conn = _connect()
    try:
        rows = conn.execute(
            """
            SELECT supplier_key, COUNT(*) AS n, SUM(amount) AS total
            FROM chase_transactions WHERE supplier_key IS NOT NULL GROUP BY supplier_key
            """
        ).fetchall()
        return {row["supplier_key"]: {"count": row["n"], "total": row["total"] or 0.0} for row in rows}
    finally:
        conn.close()


def recategorize_all(categorize_fn, resolve_supplier_fn):
    """
    Recorre TODOS los movimientos ya guardados (cualquier mes, no solo el
    que se esté mirando) y recalcula Detalle y proveedor vinculado contra
    las reglas VIGENTES en este momento -- pedido explícito del usuario
    (2026-09-21): crear, editar o eliminar una regla (de Chase o de pago a
    proveedores) tiene que reflejarse al instante en lo que ya está
    guardado, sin obligar a resubir el extracto del banco de nuevo.

    Nunca toca un valor que el usuario ya fijó a mano (detalle_source o
    supplier_source = "manual") -- los dos se recalculan de forma
    INDEPENDIENTE uno del otro, igual que en upsert_transactions.

    `categorize_fn` recibe (Descripción cruda, monto) -- el monto separa los
    depósitos chicos de Food Truck (monto clavado); `resolve_supplier_fn`
    recibe solo la Descripción. Los dos
    devuelven el Detalle / la supplier_key nuevos (o None) -- se inyectan
    desde afuera (chase_rules.categorize_chase_description /
    proveedores.match_supplier_for_chase_description) para que este módulo
    de datos puro no tenga que importar ninguno de los dos.

    Devuelve (detalle_changed, supplier_changed).
    """
    conn = _connect()
    try:
        rows = conn.execute("SELECT rowid AS _rowid, * FROM chase_transactions").fetchall()
        detalle_changed = 0
        supplier_changed = 0
        now = datetime.utcnow().isoformat()
        for row in rows:
            updates = {}
            if row["detalle_source"] not in ("manual", "deposito"):
                new_detalle = categorize_fn(row["description"], row["amount"])
                if new_detalle != row["detalle"]:
                    updates["detalle"] = new_detalle
                    updates["detalle_source"] = "rule" if new_detalle else None
            if row["supplier_source"] != "manual":
                new_supplier = resolve_supplier_fn(row["description"])
                if new_supplier != row["supplier_key"]:
                    updates["supplier_key"] = new_supplier
                    updates["supplier_source"] = "rule" if new_supplier else None
            if not updates:
                continue
            if "detalle" in updates:
                detalle_changed += 1
            if "supplier_key" in updates:
                supplier_changed += 1
            updates["updated_at"] = now
            set_clause = ", ".join(f"{column} = ?" for column in updates)
            conn.execute(
                f"UPDATE chase_transactions SET {set_clause} WHERE rowid = ?",
                (*updates.values(), row["_rowid"]),
            )
        conn.commit()
        return detalle_changed, supplier_changed
    finally:
        conn.close()


def deposits_on(posting_date):
    """Los depósitos (DEPOSIT, importe positivo) de Chase de un día: [{rowid, amount, detalle, detalle_source}]."""
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT rowid AS rowid, amount, detalle, detalle_source FROM chase_transactions "
            "WHERE posting_date = ? AND amount > 0 AND UPPER(description) LIKE 'DEPOSIT%' ORDER BY rowid",
            (posting_date,),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def deposits_set_by_receipt(detalles):
    """Los depósitos de Chase que categorizó un recibo con alguno de esos Detalle: [{rowid, posting_date, description, amount, detalle}]."""
    conn = _connect()
    try:
        marks = ", ".join("?" for _ in detalles)
        rows = conn.execute(
            "SELECT rowid AS rowid, posting_date, description, amount, detalle FROM chase_transactions "
            f"WHERE detalle_source = 'deposito' AND UPPER(detalle) IN ({marks}) ORDER BY rowid",
            [d.upper() for d in detalles],
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def release_deposit_detalle(rowid, detalle):
    """Un depósito que ya no tiene el recibo que lo categorizó vuelve a la categoría de las reglas (o sin categoría)."""
    conn = _connect()
    try:
        conn.execute(
            "UPDATE chase_transactions SET detalle = ?, detalle_source = ?, updated_at = ? WHERE rowid = ?",
            (detalle, "rule" if detalle else None, datetime.utcnow().isoformat(), rowid),
        )
        conn.commit()
    finally:
        conn.close()


def set_deposit_detalle(rowid, detalle):
    """
    Categoría puesta por un recibo de depósito cargado (pedido del usuario,
    2026-10-07: un recibo de Ice Machine categoriza su depósito en Chase).
    detalle_source="deposito": ni una recarga del extracto ni un cambio de
    reglas la pisan (como "manual"); una corrección a mano sí.
    """
    conn = _connect()
    try:
        conn.execute(
            "UPDATE chase_transactions SET detalle = ?, detalle_source = 'deposito', updated_at = ? WHERE rowid = ?",
            (detalle, datetime.utcnow().isoformat(), rowid),
        )
        conn.commit()
    finally:
        conn.close()


def list_known_details():
    """Detalle ya usados alguna vez (cualquier mes) -- para sugerir en el form de categorización manual."""
    conn = _connect()
    try:
        cur = conn.execute(
            "SELECT DISTINCT detalle FROM chase_transactions WHERE detalle IS NOT NULL AND detalle != '' ORDER BY detalle"
        )
        return [row[0] for row in cur.fetchall()]
    finally:
        conn.close()


def get_month_transactions(year, month):
    """Movimientos de un mes calendario, ordenados por fecha."""
    start = f"{year:04d}-{month:02d}-01"
    end_year, end_month = (year + 1, 1) if month == 12 else (year, month + 1)
    end = f"{end_year:04d}-{end_month:02d}-01"
    conn = _connect()
    try:
        cur = conn.execute(
            """
            SELECT * FROM chase_transactions
            WHERE posting_date >= ? AND posting_date < ?
            ORDER BY posting_date ASC, description ASC
            """,
            (start, end),
        )
        return [_row_to_dict(row) for row in cur.fetchall()]
    finally:
        conn.close()


def get_available_years():
    """
    Años distintos con al menos un movimiento guardado -- pedido explícito
    del usuario (2026-09-17, módulo nuevo "Reportes"): el selector de mes/
    año del reporte de Chase solo debe ofrecer años que de verdad tengan
    datos cargados, no un rango arbitrario/hardcodeado.
    """
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT DISTINCT substr(posting_date, 1, 4) AS y FROM chase_transactions ORDER BY y"
        ).fetchall()
        return [int(row["y"]) for row in rows if row["y"]]
    finally:
        conn.close()


def has_detalle_on_date(fecha, detalle):
    """
    True si existe algún movimiento ya guardado para ese día puntual con ese
    Detalle exacto -- usado por Lottery (ver lottery_db/webapp.py) para
    avisar si la fecha de Chase Bank confirmada de un bloque no tiene, en
    los movimientos ya cargados, ningún pago real de Lottery ese día
    (Detalle "LOTTERY", ver chase_rules.py -- keyword "fla lottery").
    """
    key = fecha.isoformat() if hasattr(fecha, "isoformat") else str(fecha)
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT 1 FROM chase_transactions WHERE posting_date = ? AND detalle = ? LIMIT 1",
            (key, detalle),
        ).fetchone()
        return row is not None
    finally:
        conn.close()


def get_first_posting_date():
    """Fecha del primer movimiento guardado (date), o None: desde dónde está cargado Chase."""
    conn = _connect()
    try:
        row = conn.execute("SELECT MIN(posting_date) AS first FROM chase_transactions").fetchone()
    finally:
        conn.close()
    return date.fromisoformat(row["first"]) if row and row["first"] else None


def get_last_posting_date():
    """Fecha del último movimiento guardado (date), o None: hasta dónde está cargado Chase."""
    conn = _connect()
    try:
        row = conn.execute("SELECT MAX(posting_date) AS last FROM chase_transactions").fetchone()
    finally:
        conn.close()
    return date.fromisoformat(row["last"]) if row and row["last"] else None


def get_uncategorized_count(year, month):
    """Cuántos movimientos del mes quedaron sin ninguna regla que matcheara -- útil para avisar."""
    rows = get_month_transactions(year, month)
    return sum(1 for row in rows if not row["detalle"])


def get_check_transactions():
    """Todos los movimientos de cheque ("CHECK {n}...") guardados, de cualquier mes -- ver cheques_db.py."""
    conn = _connect()
    try:
        rows = conn.execute(
            """
            SELECT posting_date, description, amount, supplier_key FROM chase_transactions
            WHERE UPPER(description) LIKE 'CHECK%' OR UPPER(description) LIKE 'CHEQUE%'
            ORDER BY posting_date
            """
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()
