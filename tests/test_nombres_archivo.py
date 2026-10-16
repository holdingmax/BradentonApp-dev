"""
Datos que se leen del nombre del archivo subido (por eso nunca se sanitiza
el nombre): día del Close Store y fecha/tipo de los recibos de depósito.
"""

from datetime import date

import pytest

import depositos
from reporte_diario import extract_day_from_filename


@pytest.mark.parametrize("nombre, dia", [
    ("Close Store 01-05.pdf", 1),
    ("Close Store 13-10.pdf", 13),
    ("C:/x/Close Store 31-07.pdf", 31),
])
def test_dia_del_close_store(nombre, dia):
    assert extract_day_from_filename(nombre) == dia


def test_close_store_sin_dia():
    with pytest.raises(ValueError):
        extract_day_from_filename("Close Store.pdf")


def test_fecha_del_recibo_de_deposito():
    assert depositos.filename_date("Transaccion #123 29-09-2026.pdf") == date(2026, 9, 29)
    assert depositos.filename_date("Transaccion #123.pdf") is None


@pytest.mark.parametrize("texto, tipo", [
    ("Food Truck", depositos.FOOD_TRUCK),
    ("Ice Machine", depositos.ICE_MACHINE),
    ("hielo", depositos.ICE_MACHINE),
    ("Vaccumms", "Vaccumms"),
    ("Vacuums", "Vaccumms"),
    ("Invoice", "Invoice"),   # "ice" dentro de otra palabra no es hielo
])
def test_tipo_de_deposito(texto, tipo):
    assert depositos.normalize_kind(texto) == tipo


def test_aclaraciones_del_nombre():
    anotaciones = depositos.filename_annotations("Transaccion #133 (Ice Machine)-134 15-09-2026.pdf")
    assert anotaciones == {133: depositos.ICE_MACHINE}
