"""Funciones de SQLite que las consultas de la app usan y Postgres no trae igual.

- round(double precision, integer): Postgres solo redondea a N decimales un
  NUMERIC; SQLite redondea un REAL y devuelve REAL. Esta versión pasa por
  NUMERIC y vuelve a double precision, así la consulta sigue igual
  (eft_db: comparación de batches de cupones). Con importes en centavos da
  lo mismo que SQLite; con 3 decimales o más puede diferir en el último
  (SQLite redondea sobre 17 dígitos del valor binario: 1768.485 → 1768.48,
  acá 1768.49). Probado con 20.000 valores: a 0 decimales, siempre igual.
- group_concat(text): como el de SQLite, une con coma (reportes_db: fuentes
  de los departamentos de un día). Con DISTINCT, Postgres las ordena.

Van en el esquema public, que está al final del search_path de cada conexión.

Revision ID: 0014_funciones_sqlite
Revises: 0013_lottery
Create Date: 2026-10-16
"""

from alembic import op

revision = "0014_funciones_sqlite"
down_revision = "0013_lottery"
branch_labels = None
depends_on = None


def upgrade():
    op.execute(
        """
        CREATE FUNCTION public.round(value double precision, digits integer) RETURNS double precision
        LANGUAGE sql IMMUTABLE PARALLEL SAFE
        AS $$ SELECT round(value::numeric, digits)::double precision $$
        """
    )
    op.execute(
        """
        CREATE FUNCTION public._group_concat_step(acc text, value text) RETURNS text
        LANGUAGE sql IMMUTABLE PARALLEL SAFE
        AS $$ SELECT CASE WHEN value IS NULL THEN acc WHEN acc IS NULL THEN value ELSE acc || ',' || value END $$
        """
    )
    op.execute("CREATE AGGREGATE public.group_concat(text) (SFUNC = public._group_concat_step, STYPE = text)")


def downgrade():
    op.execute("DROP AGGREGATE public.group_concat(text)")
    op.execute("DROP FUNCTION public._group_concat_step(text, text)")
    op.execute("DROP FUNCTION public.round(double precision, integer)")
