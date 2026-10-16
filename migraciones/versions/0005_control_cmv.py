"""Reportes de Elistar del Control CMV (control_cmv.db): esquema control_cmv.

Misma estructura que las tablas SQLite de control_cmv.db (con las columnas que les
agregaron los ALTER de cada versión). Fechas como texto y COLLATE "C" para
ordenar igual que SQLite.

Revision ID: 0005_control_cmv
Revises: 0004_cheques
Create Date: 2026-10-16
"""

from alembic import op

revision = "0005_control_cmv"
down_revision = "0004_cheques"
branch_labels = None
depends_on = None


def upgrade():
    op.execute("CREATE SCHEMA IF NOT EXISTS control_cmv")
    op.execute(
        """
        CREATE TABLE control_cmv.elistar_reports (
            year BIGINT,
            month BIGINT,
            kind TEXT COLLATE "C",
            data_json TEXT COLLATE "C" NOT NULL,
            filename TEXT COLLATE "C",
            uploaded_at TEXT COLLATE "C",
            PRIMARY KEY (year, month, kind)
        )
        """
    )


def downgrade():
    op.execute("DROP SCHEMA control_cmv CASCADE")
