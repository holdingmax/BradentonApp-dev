"""
Bloques de 7 días de Lottery (semana ISO lunes-domingo) con Subtotal y
Debito calculados como las fórmulas del Excel de Lottery. Corre contra la
base temporal de conftest.py.
"""

from datetime import date

import pytest

import lottery_db

SEMANA = (2026, 37)  # lunes 07/09 a domingo 13/09/2026


def _cargar_dia(fecha, ventas, departamento):
    campos = dict.fromkeys(lottery_db._SALES_REPORT_FIELDS)
    campos.update(ventas)
    lottery_db.upsert_sales_report_fields(fecha, campos, source="prueba")
    lottery_db.upsert_department_fields(fecha, *departamento, source="prueba")


@pytest.fixture(scope="module")
def semana_cargada():
    _cargar_dia(date(2026, 9, 7), {
        "sales": 1000.0, "pagos": -400.0, "comis": 50.0, "prize_free_plays": 20.0,
        "pays_amount": -300.0, "skoff_sales_amount": 200.0, "sales_comm": 10.0,
    }, (5, 900.0, 2, 150.0))
    _cargar_dia(date(2026, 9, 8), {
        "sales": 500.0, "pagos": -100.0, "comis": 25.0, "prize_free_plays": 0.0,
        "pays_amount": -50.0, "skoff_sales_amount": 100.0, "sales_comm": 5.0,
    }, (3, 450.0, 1, 75.0))
    bloques = lottery_db.build_month_blocks(2026, 9)
    return next(b for b in bloques if (b["iso_year"], b["iso_week"]) == SEMANA)


def test_columnas_del_dia(semana_cargada):
    lunes = semana_cargada["days"][0]
    assert lunes["comis_ratio"] == pytest.approx(0.05)        # J = I/F
    assert lunes["total_comm"] == pytest.approx(70.0)         # M = I+K
    assert lunes["net_total"] == pytest.approx(-90.0)         # U = Q+R+S
    assert lunes["cuenta_final"] == pytest.approx(700.0)      # X = -G-Q
    assert semana_cargada["days"][2]["sales"] is None         # miércoles sin datos


def test_subtotal(semana_cargada):
    sub = semana_cargada["subtotal"]
    assert sub["sales"] == pytest.approx(1500.0)
    assert sub["pagos"] == pytest.approx(-500.0)
    assert sub["online_count"] == 8
    assert sub["online_net_sales"] == pytest.approx(1350.0)
    assert sub["total_comm"] == pytest.approx(95.0)
    assert sub["pays_amount"] == pytest.approx(-350.0)


def test_debito(semana_cargada):
    deb = semana_cargada["debito"]
    assert deb["online_net_sales"] == pytest.approx(-150.0)   # E - F
    assert deb["sales"] == pytest.approx(1105.0)              # F + G + I + K + 10
    assert deb["pays_amount"] == pytest.approx(-35.0)         # Q + R + S
    assert deb["net_debit"] == pytest.approx(1070.0)          # F + Q de la fila Debito


def test_semana_vacia_no_debe_10():
    bloque = next(b for b in lottery_db.build_month_blocks(2026, 3) if b["iso_week"] == 11)
    assert not bloque["has_data"]
    assert bloque["debito"]["sales"] is None
    assert bloque["debito"]["net_debit"] is None


def test_bloques_del_mes_cruzan_el_borde():
    # Agosto 2026 empieza sábado y termina lunes: 6 semanas, del 27/07 al 06/09.
    bloques = lottery_db.build_month_blocks(2026, 8)
    assert [b["iso_week"] for b in bloques] == [31, 32, 33, 34, 35, 36]
    assert bloques[0]["start"] == "2026-07-27"
    assert bloques[-1]["end"] == "2026-09-06"
    assert all(len(b["days"]) == 7 for b in bloques)


def test_fecha_de_pago_sugerida_y_encadenada():
    def sugerida(semana):
        bloque = next(b for b in lottery_db.build_month_blocks(2026, 9) if b["iso_week"] == semana)
        return bloque["suggested_chase_date"]

    # Sin ningún pago confirmado: el miércoles siguiente al bloque.
    assert sugerida(37) == "2026-09-16"
    # Con uno confirmado, los demás se encadenan cada 7 días.
    lottery_db.set_block_chase_date(2026, 37, date(2026, 9, 17))
    try:
        assert sugerida(38) == "2026-09-24"
        assert sugerida(36) == "2026-09-10"
    finally:
        lottery_db.set_block_chase_date(2026, 37, None)
