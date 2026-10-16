"""Payment Summary de la máquina de hielo (ice_machine.db): esquema ice_machine.

Misma estructura que la tabla SQLite de ice_machine_db.py.

Revision ID: 0002_ice_machine
Revises: 0001_jh_facturas
Create Date: 2026-10-16
"""

from alembic import op

revision = "0002_ice_machine"
down_revision = "0001_jh_facturas"
branch_labels = None
depends_on = None


def upgrade():
    op.execute("CREATE SCHEMA IF NOT EXISTS ice_machine")
    op.execute(
        """
        CREATE TABLE ice_machine.payment_summaries (
            summary_no TEXT COLLATE "C" PRIMARY KEY,
            reference TEXT COLLATE "C",
            from_date TEXT COLLATE "C",
            to_date TEXT COLLATE "C" NOT NULL,
            transactions BIGINT,
            gross DOUBLE PRECISION,
            process_fees DOUBLE PRECISION,
            service_fees DOUBLE PRECISION,
            net DOUBLE PRECISION,
            notes TEXT COLLATE "C",
            filename TEXT COLLATE "C",
            uploaded_at TEXT COLLATE "C"
        )
        """
    )
    op.execute("CREATE INDEX idx_payment_summaries_to ON ice_machine.payment_summaries (to_date)")


def downgrade():
    op.execute("DROP SCHEMA ice_machine CASCADE")
