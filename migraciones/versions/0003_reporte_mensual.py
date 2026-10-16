"""Reporte mensual del POS (reporte_mensual.db): esquema reporte_mensual.

Misma estructura que las tablas SQLite de reporte_mensual.db (con las columnas que les
agregaron los ALTER de cada versión). Fechas como texto y COLLATE "C" para
ordenar igual que SQLite.

Revision ID: 0003_reporte_mensual
Revises: 0002_ice_machine
Create Date: 2026-10-16
"""

from alembic import op

revision = "0003_reporte_mensual"
down_revision = "0002_ice_machine"
branch_labels = None
depends_on = None


def upgrade():
    op.execute("CREATE SCHEMA IF NOT EXISTS reporte_mensual")
    op.execute(
        """
        CREATE TABLE reporte_mensual.monthly_report_departments (
            year BIGINT,
            month BIGINT,
            department TEXT COLLATE "C",
            count BIGINT,
            amount DOUBLE PRECISION,
            PRIMARY KEY (year, month, department)
        )
        """
    )
    op.execute(
        """
        CREATE TABLE reporte_mensual.monthly_reports (
            year BIGINT,
            month BIGINT,
            from_date TEXT COLLATE "C",
            to_date TEXT COLLATE "C",
            volume DOUBLE PRECISION,
            sales_fuel DOUBLE PRECISION,
            desc_comb DOUBLE PRECISION,
            non_fuel_total DOUBLE PRECISION,
            desc_otros DOUBLE PRECISION,
            tax_collect DOUBLE PRECISION,
            total_sales DOUBLE PRECISION,
            cash DOUBLE PRECISION,
            local_accounts DOUBLE PRECISION,
            other_amount DOUBLE PRECISION,
            network_revenue DOUBLE PRECISION,
            total_revenue DOUBLE PRECISION,
            credit_terms_json TEXT COLLATE "C",
            printed_department_amount DOUBLE PRECISION,
            printed_department_count BIGINT,
            warnings_json TEXT COLLATE "C",
            filename TEXT COLLATE "C",
            updated_at TEXT COLLATE "C",
            edited_at TEXT COLLATE "C",
            sheet_added_at TEXT COLLATE "C",
            PRIMARY KEY (year, month)
        )
        """
    )


def downgrade():
    op.execute("DROP SCHEMA reporte_mensual CASCADE")
