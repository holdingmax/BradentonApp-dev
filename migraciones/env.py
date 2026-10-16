"""
Alembic para la base PostgreSQL de BradentonApp.

La URL es la de DATABASE_URL; si no está en el entorno, se toma del .env de
la carpeta del código. Las migraciones están escritas en SQL (op.execute),
con la misma estructura que las tablas SQLite de cada módulo: la app no usa
el ORM de SQLAlchemy, solo Alembic lo necesita para conectarse.
"""

import os
import sys

from alembic import context
from sqlalchemy import create_engine, pool, text

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import db  # noqa: E402

if not db.uses_postgres():
    db.load_dotenv()
url = db.database_url()
if not url:
    raise SystemExit("Falta DATABASE_URL (variable de entorno o archivo .env): no hay base a la que aplicar las migraciones.")
# SQLAlchemy necesita el nombre del driver (psycopg 3) en la URL.
url = "postgresql+psycopg://" + url.split("://", 1)[1]


def run_migrations_online():
    engine = create_engine(url, poolclass=pool.NullPool)
    with engine.connect() as connection:
        # Candado de Postgres: si dos procesos arrancan juntos (dos instancias
        # en Render, o la app y un "alembic upgrade" a mano), el segundo espera
        # a que el primero termine en lugar de aplicar lo mismo dos veces.
        connection.execute(text("SELECT pg_advisory_lock(20261016)"))
        connection.commit()
        try:
            context.configure(connection=connection, transaction_per_migration=True)
            with context.begin_transaction():
                context.run_migrations()
        finally:
            connection.execute(text("SELECT pg_advisory_unlock(20261016)"))
            connection.commit()


if context.is_offline_mode():
    context.configure(url=url, literal_binds=True)
    with context.begin_transaction():
        context.run_migrations()
else:
    run_migrations_online()
