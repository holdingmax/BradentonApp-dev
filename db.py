"""
Puerta única a la base de datos (pase a PostgreSQL, ver PLAN_POSTGRES.md).

Cada módulo *_db.py ya pasado pide su conexión con db.connect(...):
- con la variable de entorno DATABASE_URL (Render), a PostgreSQL;
- sin ella (la PC), al SQLite de siempre en la carpeta de datos.

Mientras dure el pase conviven los dos: solo los módulos ya pasados usan
esta capa, y la PC sigue trabajando con SQLite porque la app no lee el
archivo .env (lo leen las herramientas: la copia de datos, el respaldo y
el comparador, con load_dotenv()).

En Postgres cada base SQLite de hoy es un esquema con su mismo nombre
(jh_facturas, chase, ...) y las tablas las crean las migraciones de Alembic
(carpeta migraciones/), nunca el código de los módulos. Los nombres de tabla
no se repiten entre bases, así que la conexión busca en todos los esquemas
(search_path) y las consultas no llevan el esquema adelante.

Lo que se escribe distinto en cada motor lo resuelve esta capa, para que las
consultas de los módulos sean las mismas en los dos:
- los "?" de SQLite pasan a "%s" (y un "%" literal a "%%");
- las filas se leen por posición y por nombre, como sqlite3.Row;
- "with conn:" confirma o deshace sin cerrar, como en sqlite3;
- conn.close() devuelve la conexión al pool (abrir una por consulta contra
  una base en la red costaría 50-200 ms cada vez);
- un NUMERIC (SUM de enteros, AVG, un literal 0.5) vuelve como int o float,
  nunca como Decimal, igual que en SQLite.
"""

import atexit
import datetime
import functools
import os
import re
import sqlite3
import threading

APP_DIR = os.path.dirname(os.path.abspath(__file__))

# Un esquema por base SQLite de hoy, más "config" (usuarios y reglas, que hoy
# son archivos). Es también el search_path de cada conexión.
SCHEMAS = (
    "config",
    "jh_facturas", "ice_machine", "reporte_mensual", "cheques", "control_cmv", "depositos",
    "horas_trabajo", "jh_mensual", "documents", "fisico", "caja", "gettel_toyota", "lottery",
    "reportes_diarios", "cuenta_kia_toyota", "eft", "chase", "cmv", "proveedores",
)

try:
    import psycopg
    from psycopg.types.numeric import NumericLoader

    Error = (sqlite3.Error, psycopg.Error)
    IntegrityError = (sqlite3.IntegrityError, psycopg.IntegrityError)
except ImportError:  # sin psycopg instalado solo se puede usar SQLite
    psycopg = None
    NumericLoader = object
    Error = (sqlite3.Error,)
    IntegrityError = (sqlite3.IntegrityError,)


def load_dotenv():
    """Lee el .env de la carpeta del código sin pisar variables ya definidas.
    Solo para herramientas (copia, respaldo, comparador): la app no lo lee
    mientras dure el pase, así la PC sigue con SQLite aunque el .env tenga
    la DATABASE_URL de testing."""
    from dotenv import load_dotenv as _load

    _load(os.path.join(APP_DIR, ".env"), override=False)


def database_url():
    return os.environ.get("DATABASE_URL", "").strip()


def uses_postgres():
    return bool(database_url())


def connect(schema, sqlite_path, sqlite_init=None):
    """
    Conexión para el módulo dueño de `schema`. En SQLite es la conexión
    nativa de siempre (sqlite3.Row, timeout 30 s) y `sqlite_init(conn)` crea
    o completa sus tablas; en Postgres las tablas ya las creó Alembic.
    """
    if schema not in SCHEMAS:
        raise ValueError(f"Esquema desconocido: {schema}")
    if uses_postgres():
        pool = _get_pool()
        return _PgConnection(pool, pool.getconn())
    os.makedirs(os.path.dirname(sqlite_path), exist_ok=True)
    conn = sqlite3.connect(sqlite_path, timeout=30)
    conn.row_factory = sqlite3.Row
    if sqlite_init:
        sqlite_init(conn)
    return conn


def upgrade_schema():
    """Aplica las migraciones de Alembic que falten (alembic upgrade head)."""
    from alembic import command
    from alembic.config import Config

    config = Config(os.path.join(APP_DIR, "alembic.ini"))
    config.set_main_option("script_location", os.path.join(APP_DIR, "migraciones"))
    command.upgrade(config, "head")


# --- Traducción de las consultas -------------------------------------------

# Literales, identificadores entre comillas y comentarios se copian tal cual
# (salvo el "%", que psycopg mira en toda la consulta); afuera, "?" es un
# parámetro.
_TOKEN = re.compile(r"'(?:[^']|'')*'|\"(?:[^\"]|\"\")*\"|--[^\n]*|/\*.*?\*/|\?|%", re.S)


@functools.lru_cache(maxsize=2048)
def to_postgres(sql):
    def repl(match):
        text = match.group(0)
        if text == "?":
            return "%s"
        return text.replace("%", "%%")

    return _TOKEN.sub(repl, sql)


def _adapt(value):
    """Lo que sqlite3 convierte solo al guardar: una fecha pasa como texto
    AAAA-MM-DD (y con hora, "AAAA-MM-DD HH:MM:SS"), True/False como 1/0."""
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, datetime.datetime):
        return value.isoformat(" ")
    if isinstance(value, datetime.date):
        return value.isoformat()
    return value


def _prepare(sql, params):
    if params is None or (not isinstance(params, dict) and len(params) == 0):
        return sql, None
    if isinstance(params, dict):
        raise ValueError("Parámetros con nombre (:nombre) todavía no están soportados en Postgres")
    return to_postgres(sql), tuple(_adapt(v) for v in params)


# --- Filas como sqlite3.Row ------------------------------------------------

class Row:
    """Fila que se lee por posición (row[0]) y por nombre (row["amount"],
    sin distinguir mayúsculas, como sqlite3.Row); dict(row) también anda."""

    __slots__ = ("_values", "_names", "_index")

    def __init__(self, values, names, index):
        self._values = tuple(values)
        self._names = names
        self._index = index

    def __getitem__(self, key):
        if isinstance(key, (int, slice)):
            return self._values[key]
        try:
            return self._values[self._index[key.lower()]]
        except KeyError:
            raise IndexError("No item with that key") from None

    def keys(self):
        return list(self._names)

    def __iter__(self):
        return iter(self._values)

    def __len__(self):
        return len(self._values)

    def __eq__(self, other):
        if isinstance(other, Row):
            return self._names == other._names and self._values == other._values
        return NotImplemented

    def __hash__(self):
        return hash((self._names, self._values))

    def __repr__(self):
        return f"Row({dict(zip(self._names, self._values))!r})"


def _row_factory(cursor, hide_rowid=False):
    names = tuple(column.name for column in cursor.description or ())
    keep = [i for i, name in enumerate(names) if not (hide_rowid and name == "rowid")]
    if len(keep) != len(names):
        names = tuple(names[i] for i in keep)
    else:
        keep = None
    index = {}
    for position, name in enumerate(names):
        index.setdefault(name.lower(), position)
    if keep is None:
        return lambda values: Row(values, names, index)
    return lambda values: Row([values[i] for i in keep], names, index)


def _hiding_rowid(cursor):
    return _row_factory(cursor, hide_rowid=True)


# La lista de columnas de un SELECT (lo que va antes del primer FROM).
_SELECT_LIST = re.compile(r"(?is)^\s*select\b(.*?)\bfrom\b")


def _names_rowid(sql):
    """En SQLite `SELECT *` no trae el rowid (solo sale si se lo nombra); en
    Postgres las tablas sin id propio tienen una columna rowid real (la copia
    del de SQLite, para ordenar igual) y `*` la traería. Se oculta salvo que
    la consulta la nombre entre las columnas."""
    match = _SELECT_LIST.match(sql)
    if match is None:
        return True
    # "rowid" o "rowid AS rowid" piden esa salida; "rowid AS _rowid" no (la
    # trae con otro nombre, y el * de Postgres no tiene que sumarla).
    return any(m.group(1) is None or m.group(1).lower() == "rowid" for m in _ROWID_OUTPUT.finditer(match.group(1)))


_ROWID_OUTPUT = re.compile(r"(?i)\browid\b(?:\s+AS\s+(\w+))?")


try:
    from psycopg.adapt import Dumper as _Dumper
except ImportError:
    _Dumper = object


class _UntypedNumberDumper(_Dumper):
    """Números sin tipo: Postgres los toma del tipo de la columna, como hace
    SQLite con su afinidad. Así `WHERE texto = ?` con un número compara como
    texto (en vez de dar error) y en una columna numérica sigue siendo número.
    Un decimal sin parte fraccionaria (1.0) va como entero, que es como lo
    guarda SQLite en una columna INTEGER (Postgres no acepta "1.0" ahí)."""

    oid = 0

    def dump(self, obj):
        if isinstance(obj, float) and not (obj.is_integer() and abs(obj) < 2 ** 63):
            return repr(obj).encode()
        return str(int(obj)).encode()


class _NumberLoader(NumericLoader):
    """NUMERIC como en SQLite: entero si no tiene decimales, float si tiene."""

    def load(self, data):
        value = super().load(data)
        if value.is_finite() and value.as_tuple().exponent >= 0:
            return int(value)
        return float(value)


# --- Conexión y cursor de Postgres con la forma de sqlite3 ------------------

class _PgCursor:
    def __init__(self, cursor):
        self._cursor = cursor

    @property
    def rowcount(self):
        return self._cursor.rowcount

    @property
    def description(self):
        return self._cursor.description

    @property
    def lastrowid(self):
        raise AttributeError("En Postgres no hay lastrowid: usá INSERT ... RETURNING id")

    def fetchone(self):
        return self._cursor.fetchone() if self._cursor.description else None

    def fetchall(self):
        return self._cursor.fetchall() if self._cursor.description else []

    def fetchmany(self, size=1):
        return self._cursor.fetchmany(size) if self._cursor.description else []

    def __iter__(self):
        return iter(self.fetchall())

    def close(self):
        self._cursor.close()


class _PgConnection:
    def __init__(self, pool, conn):
        self._pool = pool
        self._conn = conn

    def _live(self):
        if self._conn is None:
            raise sqlite3.ProgrammingError("Cannot operate on a closed database.")
        return self._conn

    def execute(self, sql, params=None):
        conn = self._live()
        cursor = conn.cursor() if _names_rowid(sql) else conn.cursor(row_factory=_hiding_rowid)
        cursor.execute(*_prepare(sql, params))
        return _PgCursor(cursor)

    def executemany(self, sql, seq_of_params):
        cursor = self._live().cursor()
        cursor.executemany(to_postgres(sql), [tuple(_adapt(v) for v in params) for params in seq_of_params])
        return _PgCursor(cursor)

    def commit(self):
        self._live().commit()

    def rollback(self):
        self._live().rollback()

    @property
    def row_factory(self):
        return Row

    @row_factory.setter
    def row_factory(self, value):
        pass  # las filas ya son Row (como sqlite3.Row)

    def close(self):
        conn, self._conn = self._conn, None
        if conn is None:
            return
        try:
            if conn.info.transaction_status != psycopg.pq.TransactionStatus.IDLE:
                conn.rollback()  # como sqlite3: cerrar sin commit descarta los cambios
        finally:
            self._pool.putconn(conn)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        if exc_type is None:
            self.commit()
        else:
            self.rollback()
        return False

    def __del__(self):
        try:
            self.close()
        except Exception:
            pass


# --- Pool ------------------------------------------------------------------

_pool = None
_pool_lock = threading.Lock()


def _configure(conn):
    conn.adapters.register_loader("numeric", _NumberLoader)
    conn.adapters.register_dumper(int, _UntypedNumberDumper)
    conn.adapters.register_dumper(float, _UntypedNumberDumper)
    conn.autocommit = True
    # pg_catalog al final: así las funciones como las de SQLite que crean las
    # migraciones en public (sum de decimales, round, group_concat) se usan
    # antes que las de Postgres con el mismo nombre.
    conn.execute("SET search_path TO " + ", ".join(SCHEMAS) + ", public, pg_catalog")
    conn.autocommit = False


def _get_pool():
    global _pool
    with _pool_lock:
        if _pool is None:
            if psycopg is None:
                raise RuntimeError("Falta instalar psycopg (requirements.txt) para usar DATABASE_URL")
            from psycopg_pool import ConnectionPool

            _pool = ConnectionPool(
                database_url(),
                min_size=1,
                max_size=int(os.environ.get("BRADENTON_DB_POOL", "8")),
                kwargs={"row_factory": _row_factory, "autocommit": False, "connect_timeout": 15},
                configure=_configure,
                check=ConnectionPool.check_connection,
                name="bradenton",
                open=True,
            )
            atexit.register(_pool.close)
    return _pool
