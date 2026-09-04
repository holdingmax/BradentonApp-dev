"""
Helpers compartidos entre los módulos de la sección Controles (solo lectura,
cruzan un Excel ya cerrado de fin de mes contra un reporte externo).

Hoy son dos (Cierre Mensual, Lottery Mensual) y los dos necesitaban leer una
celda de suma mensual que puede ser un literal o una fórmula simple
"=num+num..." -- se comparte acá en vez de duplicarla en cada módulo, mismo
criterio ya usado para ocr_utils.py.
"""

import re


def eval_literal_sum_cell(value, sheet_label):
    """
    Devuelve el número real de una celda -- puede ser un literal (float/int),
    una fórmula "=num+num+num..." (como escribe reporte_diario.py para TC, o
    como el usuario carga a mano un desglose de pagos), o estar vacía
    (None -> 0.0). Nunca evalúa una fórmula con referencias a otras celdas --
    si aparece algo así, falla en vez de arriesgar un resultado incorrecto
    (regla de oro del proyecto: nunca un valor de baja confianza en
    silencio). `sheet_label` (ej. "Store Info", "Lottery") solo se usa para
    que el mensaje de error diga de qué hoja viene el problema.
    """
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str) and value.startswith("="):
        body = value[1:]
        if not re.match(r"^[\d.\-+]+$", body):
            raise ValueError(
                f"No se puede evaluar la fórmula {value!r} de {sheet_label} -- tiene algo "
                "más que números y signos, revisar a mano."
            )
        return sum(float(token) for token in re.findall(r"-?\d+\.?\d*", body))
    raise ValueError(f"Valor inesperado en una celda de {sheet_label}: {value!r}")
