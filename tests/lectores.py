"""
Regresión de los lectores de documentos con PDFs reales.

Los PDFs y lo que se espera de cada uno tienen datos del negocio y NO van al
repo: viven en la carpeta que indica BRADENTON_TEST_DATA (en la PC,
C:\\BradentonApp_pruebas):

    casos.json            [{"lector": "proveedores", "archivo": "pdfs/....pdf"}, ...]
                          (opcional "args": argumentos extra del lector)
    pdfs/                 copias de los PDFs del Drive, con su nombre original
    esperado/*.json       la salida guardada de cada caso

`tests/test_lectores.py` compara la salida de hoy contra la guardada.
`tests/regenerar_esperados.py` la vuelve a guardar: correrlo solo después de
revisar que la diferencia es una mejora.
"""

import datetime
import json
import os
import re

from PIL import Image

DATA_ENV = "BRADENTON_TEST_DATA"


def _close_store(path):
    import reporte_diario

    return {
        "store_info": reporte_diario.extract_store_info_for_day(path),
        "departamentos": reporte_diario.extract_department_sales_for_day(path),
        "lottery": reporte_diario.extract_lottery_department_fields_from_pdf(path),
    }


def _lottery_diario(path):
    import reporte_diario

    return reporte_diario.extract_lottery_receipt_fields_from_sales_report(path)


def _lottery_mensual(path):
    import lottery_mensual

    return lottery_mensual.extract_monthly_report(path)


def _depositos(path):
    import depositos

    return depositos.extract_deposits_from_pdf(path, os.path.basename(path))


def _cheques(path):
    import cheques

    return cheques.extract_checks_from_pdf(path)


def _proveedores(path):
    import proveedores

    return proveedores.extract_invoices_from_pdf(path)


def _eft(path):
    import eft_cta_cte

    return eft_cta_cte.extract_eft_data(path)


def _combustible(path):
    import fisico_invoice_parser

    return fisico_invoice_parser.extract_fuel_invoice(path)


def _jh_factura(path):
    import jh_facturas

    return jh_facturas.extract_invoice(path)


def _ice(path):
    import ice_machine

    return ice_machine.read_payment_summary(path)


def _horas(path):
    import horas_trabajo

    return horas_trabajo.extract_hours_report(path)


def _gastos_caja(path, default_year=None):
    import gastos_caja

    return gastos_caja.extract_cash_expenses(path, os.path.basename(path), default_year=default_year)


LECTORES = {
    "close_store": _close_store,
    "lottery_diario": _lottery_diario,
    "lottery_mensual": _lottery_mensual,
    "depositos": _depositos,
    "cheques": _cheques,
    "proveedores": _proveedores,
    "eft": _eft,
    "combustible": _combustible,
    "jh_factura": _jh_factura,
    "ice": _ice,
    "horas": _horas,
    "gastos_caja": _gastos_caja,
}


def data_dir():
    return os.environ.get(DATA_ENV, "").strip()


def load_cases():
    """Los casos de casos.json, o [] si no hay carpeta de datos."""
    base = data_dir()
    if not base:
        return []
    with open(os.path.join(base, "casos.json"), encoding="utf-8") as fh:
        return json.load(fh)


def case_id(case):
    return f"{case['lector']}: {os.path.basename(case['archivo'])}"


def expected_path(case):
    stem = os.path.splitext(os.path.basename(case["archivo"]))[0]
    name = re.sub(r"[^\w.-]+", "_", f"{case['lector']}__{stem}")
    return os.path.join(data_dir(), "esperado", name + ".json")


def _plain(value, base):
    """Pasa la salida de un lector a algo comparable y guardable en JSON."""
    if isinstance(value, dict):
        return {str(k): _plain(v, base) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_plain(v, base) for v in value]
    if isinstance(value, (set, frozenset)):
        return sorted(_plain(v, base) for v in value)
    if isinstance(value, float):
        return round(value, 4)
    if isinstance(value, (datetime.date, datetime.datetime, datetime.time)):
        return value.isoformat()
    if isinstance(value, str):
        return value.replace(base, "<datos>").replace(base.replace("\\", "/"), "<datos>")
    if value is None or isinstance(value, (bool, int)):
        return value
    if isinstance(value, Image.Image):
        # La foto del cheque: su repr trae la dirección de memoria, que cambia en cada corrida.
        return f"<imagen {value.mode} {value.width}x{value.height}>"
    return repr(value)


def run_case(case):
    """La salida del lector para un caso; un error también es una salida (lo que se avisa)."""
    base = data_dir()
    path = os.path.join(base, case["archivo"])
    try:
        result = LECTORES[case["lector"]](path, **case.get("args", {}))
    except Exception as exc:  # el lector rechaza el archivo: se guarda qué dijo
        result = {"error": type(exc).__name__, "mensaje": str(exc)}
    return _plain(result, base)


def load_expected(case):
    with open(expected_path(case), encoding="utf-8") as fh:
        return json.load(fh)


def save_expected(case, output):
    path = expected_path(case)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(output, fh, ensure_ascii=False, indent=1, sort_keys=True)
        fh.write("\n")
