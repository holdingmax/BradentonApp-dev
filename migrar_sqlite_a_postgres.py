"""
Copia los datos de las bases SQLite a PostgreSQL y verifica que hayan
quedado idénticos (pase a Postgres, punto 7 de PLAN_POSTGRES.md).

    python migrar_sqlite_a_postgres.py jh_facturas [otro_modulo ...] [--origen CARPETA] [--vaciar]

- La base de destino es la de DATABASE_URL (variable de entorno o .env).
- Lee una copia de cada base SQLite (se toma con la API de respaldo de
  SQLite, así sale consistente aunque la app esté abierta); la original no
  se toca.
- Exige que el destino tenga las migraciones al día (alembic upgrade head)
  y que las tablas del módulo estén vacías. --vaciar las vacía antes, y
  solo se acepta contra un Postgres de esta PC (localhost).
- Copia cada módulo en una sola transacción (si algo falla no queda nada a
  medias), respeta las claves foráneas y ajusta los contadores de ids.
- Verifica tabla por tabla: cantidad de filas, suma de cada columna
  numérica y cada fila valor por valor (más un hash para dejar constancia).
  Si algo no coincide, termina con error.
"""

import argparse
import hashlib
import math
import os
import sqlite3
import sys
import tempfile

import app_paths
import db

LOCAL_HOSTS = {"localhost", "127.0.0.1", "::1"}

# Tablas que las migraciones ya cargan con filas iniciales (lo que en SQLite
# se carga al abrir la conexión). Antes de copiar se sacan las que SQLite
# también tiene, así quedan las de SQLite; cualquier otra fila sigue
# frenando la copia.
SEED_TABLES = {"proveedores": {"supplier_settings"}}


def _fail(message):
    print(f"ERROR: {message}")
    sys.exit(1)


def _snapshot(source_path, folder):
    """Copia consistente de una base SQLite (incluye lo que está en el -wal)."""
    target_path = os.path.join(folder, os.path.basename(source_path))
    source = sqlite3.connect(f"file:{source_path}?mode=ro", uri=True)
    target = sqlite3.connect(target_path)
    try:
        source.backup(target)
    finally:
        source.close()
        target.close()
    return target_path


def _sqlite_tables(conn):
    names = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%'")]
    tables = {}
    for name in names:
        info = conn.execute(f'PRAGMA table_info("{name}")').fetchall()
        tables[name] = {
            "columns": [r[1] for r in info],
            "pk": [r[1] for r in sorted((r for r in info if r[5]), key=lambda r: r[5])],
            "deps": {r[2] for r in conn.execute(f'PRAGMA foreign_key_list("{name}")').fetchall()} - {name},
        }
    return tables


def _pg_tables(conn, schema):
    rows = conn.execute(
        "SELECT table_name, column_name FROM information_schema.columns WHERE table_schema = %s ORDER BY table_name, ordinal_position",
        (schema,),
    ).fetchall()
    tables = {}
    for table, column in rows:
        tables.setdefault(table, []).append(column)
    return tables


def _copy_order(tables):
    order, done = [], set()

    def visit(name, path=()):
        if name in done:
            return
        if name in path:
            _fail(f"Claves foráneas circulares en {name}")
        for dep in sorted(tables[name]["deps"]):
            if dep in tables:
                visit(dep, path + (name,))
        done.add(name)
        order.append(name)

    for name in sorted(tables):
        visit(name)
    return order


def _quote(name):
    return '"' + name.replace('"', '""') + '"'


def _normalize(value):
    if isinstance(value, memoryview):
        return bytes(value)
    return value


def _fetch_sorted(conn, sql, pk_positions):
    rows = [tuple(_normalize(v) for v in row) for row in conn.execute(sql).fetchall()]
    rows.sort(key=lambda row: tuple((v is not None, v) for v in (row[i] for i in pk_positions)))
    return rows


def _numeric_sums(rows, columns):
    sums = {}
    for position, column in enumerate(columns):
        values = [row[position] for row in rows if isinstance(row[position], (int, float)) and not isinstance(row[position], bool)]
        if values:
            sums[column] = math.fsum(values)
    return sums


def _rows_hash(rows):
    digest = hashlib.sha256()
    for row in rows:
        digest.update(repr(row).encode("utf-8"))
    return digest.hexdigest()[:16]


def _check_alembic_head(pg):
    from alembic.config import Config
    from alembic.script import ScriptDirectory

    config = Config(os.path.join(app_paths.APP_DIR, "alembic.ini"))
    config.set_main_option("script_location", os.path.join(app_paths.APP_DIR, "migraciones"))
    head = ScriptDirectory.from_config(config).get_current_head()
    try:
        current = pg.execute("SELECT version_num FROM public.alembic_version").fetchone()
    except Exception:
        pg.rollback()
        current = None
    if not current or current[0] != head:
        _fail(f"La base de destino no tiene las migraciones al día (está en {current[0] if current else 'ninguna'}, falta llegar a {head}). Corré: alembic upgrade head")


def migrate_module(pg, schema, sqlite_path, empty_first):
    import psycopg
    from psycopg import sql

    lite = sqlite3.connect(sqlite_path)
    try:
        lite_tables = _sqlite_tables(lite)
        pg_tables = _pg_tables(pg, schema)
        if set(lite_tables) != set(pg_tables):
            _fail(f"{schema}: las tablas no coinciden. Solo en SQLite: {sorted(set(lite_tables) - set(pg_tables))}; "
                  f"solo en Postgres: {sorted(set(pg_tables) - set(lite_tables))}")
        for table, info in lite_tables.items():
            if "rowid" in pg_tables[table] and "rowid" not in info["columns"]:
                info["columns"] = info["columns"] + ["rowid"]
            if set(info["columns"]) != set(pg_tables[table]):
                _fail(f"{schema}.{table}: las columnas no coinciden. Solo en SQLite: {sorted(set(info['columns']) - set(pg_tables[table]))}; "
                      f"solo en Postgres: {sorted(set(pg_tables[table]) - set(info['columns']))}")
            if not info["pk"]:
                _fail(f"{schema}.{table} no tiene clave primaria: no se puede verificar fila por fila")

        with pg.transaction():
            if empty_first:
                names = sql.SQL(", ").join(sql.Identifier(schema, t) for t in lite_tables)
                pg.execute(sql.SQL("TRUNCATE {} RESTART IDENTITY CASCADE").format(names))
            for table in SEED_TABLES.get(schema, ()):
                pk = lite_tables[table]["pk"]
                keys = [list(r) for r in lite.execute(f"SELECT {', '.join(_quote(c) for c in pk)} FROM {_quote(table)}")]
                with pg.cursor() as cur:
                    cur.executemany(
                        sql.SQL("DELETE FROM {} WHERE ({}) = ({})").format(
                            sql.Identifier(schema, table), sql.SQL(", ").join(sql.Identifier(c) for c in pk),
                            sql.SQL(", ").join(sql.Placeholder() for _ in pk)),
                        keys,
                    )
            for table in lite_tables:
                count = pg.execute(sql.SQL("SELECT COUNT(*) FROM {}").format(sql.Identifier(schema, table))).fetchone()[0]
                if count:
                    _fail(f"{schema}.{table} ya tiene {count} filas en Postgres: la copia es solo sobre una base vacía")
            for table in _copy_order(lite_tables):
                columns = lite_tables[table]["columns"]
                copy_sql = sql.SQL("COPY {} ({}) FROM STDIN").format(
                    sql.Identifier(schema, table), sql.SQL(", ").join(sql.Identifier(c) for c in columns)
                )
                select = f"SELECT {', '.join(_quote(c) for c in columns)} FROM {_quote(table)}"
                with pg.cursor().copy(copy_sql) as copy:
                    for row in lite.execute(select):
                        copy.write_row(row)
            # El próximo id sigue desde el último que SQLite ya entregó
            # (sqlite_sequence recuerda también los de filas borradas), así un
            # alta nueva recibe el mismo id que habría recibido en SQLite.
            try:
                used = dict(lite.execute("SELECT name, seq FROM sqlite_sequence").fetchall())
            except sqlite3.OperationalError:
                used = {}
            identities = pg.execute(
                "SELECT table_name, column_name FROM information_schema.columns WHERE table_schema = %s AND is_identity = 'YES'",
                (schema,),
            ).fetchall()
            for table, column in identities:
                pg.execute(
                    sql.SQL("SELECT setval(pg_get_serial_sequence({}, {}), GREATEST(COALESCE((SELECT MAX({}) FROM {}), 0), {}) + 1, false)").format(
                        sql.Literal(f"{schema}.{table}"), sql.Literal(column), sql.Identifier(column), sql.Identifier(schema, table),
                        sql.Literal(int(used.get(table) or 0)),
                    )
                )

        ok = True
        print(f"\n{schema}")
        print(f"  {'tabla':32} {'filas SQLite':>12} {'filas Postgres':>14}  {'hash':16}  resultado")
        for table in sorted(lite_tables):
            columns = lite_tables[table]["columns"]
            pk_positions = [columns.index(c) for c in lite_tables[table]["pk"]]
            select_cols = ", ".join(_quote(c) for c in columns)
            lite_rows = _fetch_sorted(lite, f"SELECT {select_cols} FROM {_quote(table)}", pk_positions)
            pg_rows = _fetch_sorted(pg, sql.SQL("SELECT {} FROM {}").format(
                sql.SQL(", ").join(sql.Identifier(c) for c in columns), sql.Identifier(schema, table)), pk_positions)
            same = lite_rows == pg_rows and _numeric_sums(lite_rows, columns) == _numeric_sums(pg_rows, columns)
            ok = ok and same
            print(f"  {table:32} {len(lite_rows):>12} {len(pg_rows):>14}  {_rows_hash(pg_rows):16}  {'idénticas' if same else 'DISTINTAS'}")
            if not same:
                for lite_row, pg_row in zip(lite_rows, pg_rows):
                    if lite_row != pg_row:
                        print(f"    primera diferencia:\n      SQLite:   {lite_row}\n      Postgres: {pg_row}")
                        break
            for column, total in _numeric_sums(pg_rows, columns).items():
                if column not in lite_tables[table]["pk"]:
                    print(f"      suma {column}: {total:,.2f}")
        return ok
    finally:
        lite.close()


def main():
    parser = argparse.ArgumentParser(description="Copia bases SQLite a PostgreSQL y verifica que queden idénticas.")
    parser.add_argument("modulos", nargs="+", help="nombre de la base sin .db (jh_facturas, chase, ...)")
    parser.add_argument("--origen", default=app_paths.DATA_DIR, help="carpeta con las bases SQLite (por defecto, la de datos de la app)")
    parser.add_argument("--vaciar", action="store_true", help="vacía antes las tablas del módulo (solo contra un Postgres de esta PC)")
    args = parser.parse_args()

    if not db.uses_postgres():
        db.load_dotenv()
    if not db.uses_postgres():
        _fail("Falta DATABASE_URL (variable de entorno o archivo .env).")

    import psycopg

    target = psycopg.conninfo.conninfo_to_dict(db.database_url())
    host = target.get("host") or "localhost"
    print(f"Destino: base {target.get('dbname')} en {host}")
    if args.vaciar and host not in LOCAL_HOSTS:
        _fail("--vaciar solo se acepta contra un Postgres de esta PC, nunca contra testing ni producción.")
    for schema in args.modulos:
        if schema not in db.SCHEMAS:
            _fail(f"Módulo desconocido: {schema}")
        if not os.path.isfile(os.path.join(args.origen, f"{schema}.db")):
            _fail(f"No está {schema}.db en {args.origen}")

    with psycopg.connect(db.database_url(), autocommit=True) as pg, tempfile.TemporaryDirectory(prefix="bradenton_copia_") as folder:
        _check_alembic_head(pg)
        all_ok = True
        for schema in args.modulos:
            snapshot = _snapshot(os.path.join(args.origen, f"{schema}.db"), folder)
            all_ok = migrate_module(pg, schema, snapshot, args.vaciar) and all_ok
    print("\nTodo idéntico." if all_ok else "\nHAY DIFERENCIAS: revisar antes de seguir.")
    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()
