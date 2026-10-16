"""
Motor de reglas de Chase: precedencia (compuestas primero, después la
palabra más larga), grupos con "|", depósitos chicos y las reglas del repo
(las que se siembran la primera vez; las editadas en la página no cambian
estas pruebas).
"""

import pytest

import chase_rules
from chase_rules import _match_rules, _split_small_deposit, categorize_chase_description


def _desc(text):
    return chase_rules.normalize_rule_text(text)


def test_compuesta_gana_aunque_haya_una_mas_larga():
    reglas = [
        {"keyword": "manatee county utilities", "detail": "AGUA"},
        {"keyword": "manatee + tax", "detail": "TANGIBLE TAX"},
    ]
    assert _match_rules(_desc("MANATEE COUNTY UTILITIES TAX"), reglas) == "TANGIBLE TAX"


def test_gana_la_palabra_mas_larga():
    reglas = [
        {"keyword": "kooler", "detail": "PROVEEDORES"},
        {"keyword": "kooler farms llc", "detail": "OTRA"},
    ]
    assert _match_rules(_desc("ACH KOOLER FARMS LLC 123"), reglas) == "OTRA"
    assert _match_rules(_desc("ACH KOOLER ICE"), reglas) == "PROVEEDORES"


def test_grupo_con_alternativas():
    reglas = [{"keyword": "chek + coca|coke|midtown", "detail": "PROVEEDORES"}]
    assert _match_rules(_desc("CHEK 1783 MIDTOWN"), reglas) == "PROVEEDORES"
    assert _match_rules(_desc("CHEK 1784 PEPSI"), reglas) is None


@pytest.mark.parametrize("monto, esperado", [
    (1000.0, "FOOD TRUCK"),       # Food Truck desde ago-2026
    (1030.0, "FOOD TRUCK"),       # Food Truck hasta jul-2026
    (705.0, None),                # chico: a mano (hielo u otro)
    (1030.01, "DEPOSITO"),
    (6500.0, "DEPOSITO"),
    (None, "DEPOSITO"),
])
def test_deposito_chico(monto, esperado):
    assert _split_small_deposit("DEPOSITO", monto) == esperado


def test_deposito_chico_no_toca_otro_detalle():
    assert _split_small_deposit("VENTA ICE", 500.0) == "VENTA ICE"


@pytest.mark.parametrize("descripcion, monto, esperado", [
    ("MANATEE COUNTY TAX COLLECTOR PAYMENT", -1234.56, "TANGIBLE TAX"),
    ("MANATEE COUNTY UTIL BILLPAY", -848.46, "AGUA"),
    ("FPL DIRECT DEBIT ELEC PYMT", -2100.0, "ENERGIA ELECTRICA"),
    ("ORIG CO NAME:CANTALOUPE INC ORIG ID:123", 412.30, "VENTA ICE"),
    ("ORIG CO NAME:FLA LOTTERY ORIG ID:59", -5321.0, "LOTTERY"),
    ("DEPOSIT  ID NUMBER 123456", 1000.0, "FOOD TRUCK"),
    ("DEPOSIT  ID NUMBER 123456", 8450.0, "DEPOSITO"),
    ("ACH SALARY PAYROLL", -900.0, "SUELDO"),
    ("ALGO QUE NO ESTA EN NINGUNA REGLA", -10.0, None),
    ("", 0.0, None),
])
def test_reglas_del_repo(descripcion, monto, esperado):
    assert categorize_chase_description(descripcion, monto) == esperado


def test_comentario_fijo_de_tangible_tax():
    assert chase_rules.detalle_note("TANGIBLE TAX") == "Otros gastos varios"
    assert chase_rules.detalle_note("AGUA") is None
