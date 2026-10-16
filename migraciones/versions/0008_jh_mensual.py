"""Reportes mensuales de J.H. (jh_mensual.db): esquema jh_mensual.

Misma estructura que las tablas SQLite de jh_mensual.db (con las columnas que les
agregaron los ALTER de cada versión). Fechas como texto y COLLATE "C" para
ordenar igual que SQLite.

Revision ID: 0008_jh_mensual
Revises: 0007_horas_trabajo
Create Date: 2026-10-16
"""

from alembic import op

revision = "0008_jh_mensual"
down_revision = "0007_horas_trabajo"
branch_labels = None
depends_on = None


def upgrade():
    op.execute("CREATE SCHEMA IF NOT EXISTS jh_mensual")
    op.execute(
        """
        CREATE TABLE jh_mensual.jh_monthly_reports (
            kind TEXT COLLATE "C",
            year BIGINT,
            month BIGINT,
            from_date TEXT COLLATE "C",
            to_date TEXT COLLATE "C",
            rows_json TEXT COLLATE "C",
            totals_json TEXT COLLATE "C",
            filename TEXT COLLATE "C",
            updated_at TEXT COLLATE "C",
            PRIMARY KEY (kind, year, month)
        )
        """
    )


def downgrade():
    op.execute("DROP SCHEMA jh_mensual CASCADE")
