"""Facturas varias de J.H. (jh_facturas.db): esquema jh_facturas.

Misma estructura que la tabla SQLite de jh_facturas_db.py. Fechas como
texto AAAA-MM-DD y COLLATE "C" para ordenar igual que SQLite.

Revision ID: 0001_jh_facturas
Revises:
Create Date: 2026-10-16
"""

from alembic import op

revision = "0001_jh_facturas"
down_revision = None
branch_labels = None
depends_on = None


def upgrade():
    op.execute("CREATE SCHEMA IF NOT EXISTS jh_facturas")
    op.execute(
        """
        CREATE TABLE jh_facturas.jh_misc_invoices (
            invoice_number TEXT COLLATE "C" PRIMARY KEY,
            invoice_date TEXT COLLATE "C" NOT NULL,
            due_date TEXT COLLATE "C" NOT NULL,
            amount DOUBLE PRECISION NOT NULL,
            kind TEXT COLLATE "C" NOT NULL,
            description TEXT COLLATE "C",
            detail TEXT COLLATE "C",
            source_filename TEXT COLLATE "C",
            uploaded_at TEXT COLLATE "C"
        )
        """
    )


def downgrade():
    op.execute("DROP SCHEMA jh_facturas CASCADE")
