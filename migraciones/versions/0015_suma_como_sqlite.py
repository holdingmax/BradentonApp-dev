"""SUM de decimales igual que SQLite.

SQLite (desde 3.43; la PC tiene 3.49) suma los REAL con el método compensado
de Kahan-Babuska-Neumaier: la suma de importes en centavos da el número
justo (3525.73) donde una suma común arrastra el error de cada renglón
(3525.7299999999996). Postgres suma de corrido, y un total que cae en medio
centavo podría redondearse distinto. Este agregado hace lo mismo que el
sumStep de SQLite (mismo paso, misma comparación, resultado suma + error).

Se llama sum y está en public: cada conexión de db.py pone pg_catalog al
final del search_path, así SUM(columna decimal) usa este y SUM de enteros
sigue usando el de Postgres (este solo acepta double precision).

Revision ID: 0015_suma_como_sqlite
Revises: 0014_funciones_sqlite
Create Date: 2026-10-16
"""

from alembic import op

revision = "0015_suma_como_sqlite"
down_revision = "0014_funciones_sqlite"
branch_labels = None
depends_on = None


def upgrade():
    # Estado: {suma, error}. Mismo paso que kahanBabuskaNeumaierStep de SQLite.
    op.execute(
        """
        CREATE FUNCTION public._sum_kbn_step(state double precision[], value double precision)
        RETURNS double precision[]
        LANGUAGE sql IMMUTABLE PARALLEL SAFE
        AS $$
            SELECT CASE
                WHEN value IS NULL THEN state
                WHEN state IS NULL THEN ARRAY[value, 0]::double precision[]
                ELSE ARRAY[
                    state[1] + value,
                    state[2] + CASE WHEN abs(state[1]) > abs(value)
                                    THEN (state[1] - (state[1] + value)) + value
                                    ELSE (value - (state[1] + value)) + state[1] END
                ]::double precision[]
            END
        $$
        """
    )
    op.execute(
        """
        CREATE FUNCTION public._sum_kbn_final(state double precision[]) RETURNS double precision
        LANGUAGE sql IMMUTABLE PARALLEL SAFE
        AS $$ SELECT CASE WHEN state IS NULL THEN NULL ELSE state[1] + state[2] END $$
        """
    )
    op.execute(
        "CREATE AGGREGATE public.sum(double precision) "
        "(SFUNC = public._sum_kbn_step, STYPE = double precision[], FINALFUNC = public._sum_kbn_final)"
    )


def downgrade():
    op.execute("DROP AGGREGATE public.sum(double precision)")
    op.execute("DROP FUNCTION public._sum_kbn_final(double precision[])")
    op.execute("DROP FUNCTION public._sum_kbn_step(double precision[], double precision)")
