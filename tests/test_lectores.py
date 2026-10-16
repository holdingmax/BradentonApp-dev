"""
Regresión de los lectores OCR/PDF: la salida de hoy de cada PDF real tiene
que ser igual a la guardada. Sin BRADENTON_TEST_DATA se saltean (ver
tests/lectores.py). Solo estas: `python -m pytest -m lectores`.
"""

import os

import pytest

import lectores

CASOS = lectores.load_cases()

pytestmark = [
    pytest.mark.lectores,
    pytest.mark.skipif(not CASOS, reason=f"sin {lectores.DATA_ENV}: no hay PDFs de prueba"),
]


@pytest.mark.parametrize("caso", CASOS or [None], ids=[lectores.case_id(c) for c in CASOS] or ["sin casos"])
def test_lector(caso):
    if not os.path.exists(lectores.expected_path(caso)):
        pytest.fail("Falta la salida guardada: correr tests/regenerar_esperados.py")
    assert lectores.run_case(caso) == lectores.load_expected(caso)
