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
from datetime import date, datetime

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
            PRIMARY KEY (posting_date, description, amount)
        )
        """
    )
    _ensure_columns(conn)
    conn.commit()


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
    queda pegada para siempre, ninguna carga futura la pisa.

    Devuelve (inserted, updated).
    """
    conn = _connect()
    try:
        inserted = 0
        updated = 0
        now = datetime.utcnow().isoformat()
        for row in rows:
            posting_date = row["posting_date"]
            if isinstance(posting_date, (date, datetime)):
                posting_date = posting_date.isoformat() if isinstance(posting_date, date) else posting_date.date().isoformat()
            cur = conn.execute(
                "SELECT 1 FROM chase_transactions WHERE posting_date = ? AND description = ? AND amount = ?",
                (posting_date, row["description"], row["amount"]),
            )
            exists = cur.fetchone() is not None
            conn.execute(
                """
                INSERT INTO chase_transactions
                    (posting_date, description, amount, balance, detalle, type, source_filename, updated_at, detalle_source)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(posting_date, description, amount) DO UPDATE SET
                    balance = excluded.balance,
                    type = excluded.type,
                    source_filename = excluded.source_filename,
                    updated_at = excluded.updated_at,
                    detalle = CASE WHEN chase_transactions.detalle_source = 'manual'
                                   THEN chase_transactions.detalle ELSE excluded.detalle END,
                    detalle_source = CASE WHEN chase_transactions.detalle_source = 'manual'
                                          THEN chase_transactions.detalle_source ELSE excluded.detalle_source END
                """,
                (
                    posting_date, row["description"], row["amount"], row.get("balance"),
                    row.get("detalle"), row.get("type"), source_filename, now,
                    "rule" if row.get("detalle") else None,
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


def get_uncategorized_count(year, month):
    """Cuántos movimientos del mes quedaron sin ninguna regla que matcheara -- útil para avisar."""
    rows = get_month_transactions(year, month)
    return sum(1 for row in rows if not row["detalle"])
