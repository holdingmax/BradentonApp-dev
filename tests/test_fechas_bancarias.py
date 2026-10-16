"""
Feriados bancarios, día de depósito de J.H. y fechas de Próximos EFT.
Los casos salen de lo medido con datos reales (detalle de cupones de
septiembre 2026, Columbus Day del 12/10/2026): ver control_tarjetas y
eft_pronostico.
"""

from datetime import date

import pytest

from control_tarjetas import bank_holidays, in_store_info, jh_deposit_day
from eft_pronostico import business_days_before, cutoff, is_business_day, next_business_day


def test_feriados_2026():
    assert bank_holidays(2026) == {
        date(2026, 1, 1),    # Año Nuevo
        date(2026, 1, 19),   # Martin Luther King
        date(2026, 2, 16),   # Presidents' Day
        date(2026, 5, 25),   # Memorial Day
        date(2026, 6, 19),   # Juneteenth
        date(2026, 9, 7),    # Labor Day
        date(2026, 10, 12),  # Columbus Day
        date(2026, 11, 11),  # Veterans Day
        date(2026, 11, 26),  # Thanksgiving
        date(2026, 12, 25),  # Navidad
    }


def test_feriado_en_sabado_no_cierra():
    # El 04/07/2026 es sábado: la Fed abre el viernes 03/07.
    assert date(2026, 7, 4) not in bank_holidays(2026)
    assert date(2026, 7, 3) not in bank_holidays(2026)


def test_feriado_en_domingo_pasa_al_lunes():
    # 2022: Juneteenth y Navidad cayeron domingo.
    holidays = bank_holidays(2022)
    assert date(2022, 6, 20) in holidays and date(2022, 6, 19) not in holidays
    assert date(2022, 12, 26) in holidays and date(2022, 12, 25) not in holidays


@pytest.mark.parametrize("venta, deposito", [
    # Semana normal (14-20/09/2026): lunes a miércoles a los 2 días,
    # jueves a sábado juntos el lunes, domingo el martes.
    (date(2026, 9, 14), date(2026, 9, 16)),
    (date(2026, 9, 15), date(2026, 9, 17)),
    (date(2026, 9, 16), date(2026, 9, 18)),
    (date(2026, 9, 17), date(2026, 9, 21)),
    (date(2026, 9, 18), date(2026, 9, 21)),
    (date(2026, 9, 19), date(2026, 9, 21)),
    (date(2026, 9, 20), date(2026, 9, 22)),
    # Labor Day (lunes 07/09): jueves a domingo pasan al martes.
    (date(2026, 9, 3), date(2026, 9, 8)),
    (date(2026, 9, 4), date(2026, 9, 8)),
    (date(2026, 9, 5), date(2026, 9, 8)),
    (date(2026, 9, 6), date(2026, 9, 8)),
    # Columbus Day (lunes 12/10).
    (date(2026, 10, 10), date(2026, 10, 13)),
    # Thanksgiving (jueves 26/11).
    (date(2026, 11, 24), date(2026, 11, 27)),
    # Cambio de año: el 01/01/2027 es viernes y feriado.
    (date(2026, 12, 30), date(2027, 1, 4)),
])
def test_dia_de_deposito_jh(venta, deposito):
    assert jh_deposit_day(venta) == deposito


@pytest.mark.parametrize("batch, esperado", [
    ("9263", True),
    ("0412", True),
    ("8201", True),
    ("slri81d8f", False),
    ("1234567", False),
    ("", False),
    (None, False),
])
def test_batch_del_pos(batch, esperado):
    assert in_store_info(batch) is esperado


def test_dias_habiles():
    assert is_business_day(date(2026, 10, 9))        # viernes
    assert not is_business_day(date(2026, 10, 10))   # sábado
    assert not is_business_day(date(2026, 10, 12))   # Columbus Day
    assert next_business_day(date(2026, 10, 10)) == date(2026, 10, 13)
    assert next_business_day(date(2026, 10, 14)) == date(2026, 10, 14)
    assert business_days_before(date(2026, 10, 13), 1) == date(2026, 10, 9)


@pytest.mark.parametrize("eft, corte", [
    # 2 días hábiles antes del EFT.
    (date(2026, 10, 5), date(2026, 10, 1)),
    (date(2026, 10, 8), date(2026, 10, 6)),
    # El EFT del 13/10/2026 (después de Columbus Day): jueves 08/10.
    (date(2026, 10, 13), date(2026, 10, 8)),
])
def test_corte_del_eft(eft, corte):
    assert cutoff(eft) == corte
