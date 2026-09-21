"""
Módulo "Proyecciones" -- pedido explícito del usuario (2026-09-18): un
tercer apartado (junto a Herramientas y Reportes) con Ventas por
Departamento y Store Info proyectados a fin de mes. Reproduce la lógica
de las filas 37/38/39 (hasta la columna O) de la hoja "Store Info" del
Excel de cierre real que se usaba antes de esta app -- decodificado
leyendo las fórmulas de una planilla real (Cierre_09-26.xlsx) con
openpyxl, no adivinado. El usuario mismo avisó que no entiende del todo
cómo se armaban esas fórmulas ("no las entiendo bien como recolecctaban
la info pero es lo que se usaba") -- por eso este módulo documenta cada
paso con la celda/fórmula real de origen.

Fórmula real (idéntica para Volume/Total Fuel/cada categoría de
departamento/Total Non Fuel -- una proyección "run-rate" a 30 días):
    dias_cargados = cantidad de días del mes con Store Info ya cargado
    proyectado = actual / dias_cargados * (30 - dias_cargados) + actual

Que es álgebraicamente lo mismo que:
    proyectado = actual * 30 / dias_cargados

(la planilla real usa la primera forma; se implementa la segunda, más
simple, con el mismo resultado -- confirmado numéricamente contra
Cierre_09-26.xlsx).

Casos particulares reales (no un patrón parejo -- por eso están
comentados uno por uno):
  - "Total Fuel $" proyectado (F37 en la planilla) se calcula a partir
    del TOTAL FUEL actual (H32: Sales Fuel + Desc. Comb), no de Sales
    Fuel solo (F32) -- confirmado leyendo la fórmula real (`=+H32/...`).
  - El "precio por galón" (G37) es un derivado (Total Fuel $ proyectado /
    Volume proyectado), no algo que se sume o proyecte por separado.
  - La "Utilidad" de cada categoría de departamento proyectada (I39:N39)
    es venta proyectada * margen (I38:N38) -- los mismos 6 márgenes
    estáticos que ya usaba la planilla para la Utilidad ACTUAL (fila 34),
    editables acá (ver reportes_db.get/save_department_margins).
  - El "Total Non Fuel" (O) es un valor cargado directo de Store Info,
    SIEMPRE independiente de la suma de las 6 categorías de Ventas por
    Departamento (discrepancia real y ya conocida de los datos del
    negocio, no un bug de esta app) -- se muestra por separado, nunca se
    fuerza a que coincida.
  - El margen de "Total Non Fuel" proyectado (O38 = O39/O37) es un
    DERIVADO (Utilidad total proyectada / Ventas proyectadas), no un
    input propio -- nunca se edita.

Deliberadamente FUERA de alcance en esta v1 (mismo criterio ya usado en
Reportes/Store Info con otras columnas que dependen de módulos que no
existen todavía): el ajuste de merma de combustible ("Fisico!F9" en la
planilla real, filas 38/39 columnas F/G) depende de una hoja "Fisico"
(inventario físico de combustible) que no existe como módulo en esta
app -- en la planilla real hoy vale 0 (F9 sin cargar) incluso para el
usuario, así que la proyección de Total Fuel $ y el precio por galón que
se muestran acá son la proyección "cruda", sin ese ajuste.
"""

import calendar

import reportes_db
from reporte_diario import DEPARTMENT_GROUPS, group_department_sales

# Márgenes estáticos reales del Excel (filas 33/38, columnas I-N -- mismo
# orden que DEPARTMENT_GROUPS de reporte_diario.py). Quedan acá como
# default de arranque; el valor que de verdad se usa siempre sale de
# reportes_db.get_department_margins (que cae en este default hasta que
# el usuario edite algo).
DEFAULT_MARGINS = {
    "TABACCO": 0.15,
    "SODA": 0.40,
    "BEER/WINE": 0.25,
    "LOTERY/LOTTO": 0.06,
    "Gettel": 1.00,
    "RESTO": 0.40,
}

# Nombres cortos para mostrar en la página -- las etiquetas reales de
# DEPARTMENT_GROUPS son las que usa la base/lógica interna, estas son
# solo cosméticas.
CATEGORY_DISPLAY_LABELS = {
    "TABACCO": "Tabacco",
    "SODA": "Soda",
    "BEER/WINE": "Beer / Wine",
    "LOTERY/LOTTO": "Lottery",
    "Gettel": "Gettel",
    "RESTO": "Resto",
}


def _run_rate(actual, dias_cargados):
    """actual*30/dias_cargados -- None si no hay ningún día cargado todavía (mes recién empezado)."""
    if not dias_cargados:
        return None
    return actual * 30.0 / dias_cargados


def _days_loaded(store_info_rows):
    """
    Días del mes con Store Info ya cargado -- mismo criterio que
    COUNT(columna de fechas) en la planilla real: un día sin ningún dato
    (volume en None) no cuenta, tenga o no fila (ver
    reportes_db.get_month_store_info, que devuelve un renglón por CADA
    día del mes exista o no el dato).
    """
    return sum(1 for row in store_info_rows if row.get("volume") is not None)


def build_projection(year, month):
    """
    Devuelve todo lo que necesita la página de Proyecciones para un mes:
    días cargados/restantes, Store Info (actual + proyectado) y Ventas
    por Departamento (actual + margen + utilidad, actual y proyectado).
    """
    days_in_month = calendar.monthrange(year, month)[1]

    # `_build_store_info_rows` (con los campos calculados de Total Fuel/
    # Total Sales) vive en webapp.py, no en reportes_db.py -- se importa
    # acá adentro de la función (import diferido) para evitar un import
    # circular real (webapp.py ya importa este módulo).
    from webapp import _build_store_info_rows

    store_info_rows = _build_store_info_rows(year, month)
    days_loaded = _days_loaded(store_info_rows)
    days_remaining = max(days_in_month - days_loaded, 0)

    volume_actual = sum(row.get("volume") or 0.0 for row in store_info_rows if row.get("volume") is not None)
    total_fuel_actual = sum(
        row.get("total_fuel") or 0.0 for row in store_info_rows if row.get("total_fuel") is not None
    )
    non_fuel_actual = sum(
        row.get("non_fuel_total") or 0.0 for row in store_info_rows if row.get("non_fuel_total") is not None
    )

    volume_proj = _run_rate(volume_actual, days_loaded)
    total_fuel_proj = _run_rate(total_fuel_actual, days_loaded)
    non_fuel_proj = _run_rate(non_fuel_actual, days_loaded)

    price_per_gallon_actual = (total_fuel_actual / volume_actual) if volume_actual else None
    price_per_gallon_proj = (total_fuel_proj / volume_proj) if (volume_proj and total_fuel_proj is not None) else None

    dept_totals = reportes_db.get_month_department_totals(year, month)
    groups, _unmatched = group_department_sales(dept_totals)
    groups_by_label = {g["label"]: g for g in groups}

    margins = reportes_db.get_department_margins(DEFAULT_MARGINS)

    departments = []
    dept_utilidad_actual_total = 0.0
    dept_utilidad_proj_total = 0.0
    dept_sales_proj_total = 0.0
    for label, _members in DEPARTMENT_GROUPS:
        group = groups_by_label.get(label, {"count": 0, "amount": 0.0})
        actual_amount = group.get("amount") or 0.0
        margin = margins.get(label, DEFAULT_MARGINS.get(label, 0.0))
        proj_amount = _run_rate(actual_amount, days_loaded)
        utilidad_actual = actual_amount * margin
        utilidad_proj = (proj_amount * margin) if proj_amount is not None else None
        dept_utilidad_actual_total += utilidad_actual
        if utilidad_proj is not None:
            dept_utilidad_proj_total += utilidad_proj
        if proj_amount is not None:
            dept_sales_proj_total += proj_amount
        departments.append(
            {
                "label": label,
                "display_label": CATEGORY_DISPLAY_LABELS.get(label, label),
                "count": group.get("count") or 0,
                "amount_actual": round(actual_amount, 2),
                "amount_proj": round(proj_amount, 2) if proj_amount is not None else None,
                "margin": margin,
                "utilidad_actual": round(utilidad_actual, 2),
                "utilidad_proj": round(utilidad_proj, 2) if utilidad_proj is not None else None,
            }
        )

    # Margen proyectado de "Total Non Fuel" -- derivado (O38 = O39/O37),
    # nunca un input propio. Si todavía no hay proyección de departamentos
    # (mes sin ningún día cargado) queda en None -- la plantilla muestra
    # "—".
    non_fuel_margin_proj = (
        (dept_utilidad_proj_total / dept_sales_proj_total) if dept_sales_proj_total else None
    )

    return {
        "year": year,
        "month": month,
        "days_in_month": days_in_month,
        "days_loaded": days_loaded,
        "days_remaining": days_remaining,
        "store_info": {
            "volume_actual": round(volume_actual, 2),
            "volume_proj": round(volume_proj, 2) if volume_proj is not None else None,
            "total_fuel_actual": round(total_fuel_actual, 2),
            "total_fuel_proj": round(total_fuel_proj, 2) if total_fuel_proj is not None else None,
            "price_per_gallon_actual": round(price_per_gallon_actual, 4) if price_per_gallon_actual is not None else None,
            "price_per_gallon_proj": round(price_per_gallon_proj, 4) if price_per_gallon_proj is not None else None,
            "non_fuel_actual": round(non_fuel_actual, 2),
            "non_fuel_proj": round(non_fuel_proj, 2) if non_fuel_proj is not None else None,
            "non_fuel_margin_proj": round(non_fuel_margin_proj, 4) if non_fuel_margin_proj is not None else None,
        },
        "departments": departments,
        "departments_total_actual": round(sum(d["amount_actual"] for d in departments), 2),
        "departments_total_proj": (
            round(dept_sales_proj_total, 2) if any(d["amount_proj"] is not None for d in departments) else None
        ),
        "utilidad_actual_total": round(dept_utilidad_actual_total, 2),
        "utilidad_proj_total": (
            round(dept_utilidad_proj_total, 2) if any(d["utilidad_proj"] is not None for d in departments) else None
        ),
    }
