"""Lottery: días, bloques semanales, cierre del mes y reporte mensual (lottery.db): esquema lottery.

Misma estructura que las tablas SQLite de lottery.db (con las columnas que les
agregaron los ALTER de cada versión). Fechas como texto y COLLATE "C" para
ordenar igual que SQLite.

Revision ID: 0013_lottery
Revises: 0012_gettel_toyota
Create Date: 2026-10-16
"""

from alembic import op

revision = "0013_lottery"
down_revision = "0012_gettel_toyota"
branch_labels = None
depends_on = None


def upgrade():
    op.execute("CREATE SCHEMA IF NOT EXISTS lottery")
    op.execute(
        """
        CREATE TABLE lottery.lottery_blocks (
            iso_year BIGINT,
            iso_week BIGINT,
            chase_bank_date TEXT COLLATE "C",
            updated_at TEXT COLLATE "C",
            PRIMARY KEY (iso_year, iso_week)
        )
        """
    )
    op.execute(
        """
        CREATE TABLE lottery.lottery_days (
            date TEXT COLLATE "C",
            online_count BIGINT,
            online_net_sales DOUBLE PRECISION,
            department_source TEXT COLLATE "C",
            department_pdf_filename TEXT COLLATE "C",
            sales DOUBLE PRECISION,
            pagos DOUBLE PRECISION,
            cash_balance DOUBLE PRECISION,
            comis DOUBLE PRECISION,
            prize_free_plays DOUBLE PRECISION,
            skoff_count BIGINT,
            skoff_net_sales DOUBLE PRECISION,
            pays_units DOUBLE PRECISION,
            pays_amount DOUBLE PRECISION,
            skoff_sales_amount DOUBLE PRECISION,
            sales_comm DOUBLE PRECISION,
            sales_report_source TEXT COLLATE "C",
            sales_report_pdf_filename TEXT COLLATE "C",
            updated_at TEXT COLLATE "C",
            PRIMARY KEY (date)
        )
        """
    )
    op.execute(
        """
        CREATE TABLE lottery.lottery_month_closing (
            year BIGINT,
            month BIGINT,
            gastos_admin DOUBLE PRECISION,
            updated_at TEXT COLLATE "C",
            caja_skoff_ajuste DOUBLE PRECISION,
            PRIMARY KEY (year, month)
        )
        """
    )
    op.execute(
        """
        CREATE TABLE lottery.lottery_monthly_reports (
            year BIGINT,
            month BIGINT,
            from_date TEXT COLLATE "C",
            to_date TEXT COLLATE "C",
            values_json TEXT COLLATE "C",
            lines_json TEXT COLLATE "C",
            warnings_json TEXT COLLATE "C",
            filename TEXT COLLATE "C",
            updated_at TEXT COLLATE "C",
            PRIMARY KEY (year, month)
        )
        """
    )


def downgrade():
    op.execute("DROP SCHEMA lottery CASCADE")
