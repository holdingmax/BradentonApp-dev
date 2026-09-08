"""
Controles: Cierre mensual — primer módulo de la sección Controles.

Cruza el PDF "Store Sales Summary Report" mensual (el mismo tipo de reporte
que Reporte Diario ya lee día a día, acá para todo el mes de una sola vez)
contra el total mensual ya cargado en la hoja "Store Info" del Excel Cierre
-- Total Revenue y Network Revenue. A diferencia del chequeo "dif 1"/"dif 2"
que ya vive como fórmula en la propia hoja (compara una fila contra sí
misma, así que un día salteado, cargado dos veces, o mal tipeado de forma
consistente no lo detecta), este control usa una fuente 100% independiente
-- el reporte oficial del POS para el mes completo.

Un segundo chequeo, `check_department_sales_monthly`, cruza el "Department
Sales Report" del mismo PDF mensual (otra sección del mismo bundle que
Store Sales Summary Report -- el usuario imprime varios reportes del POS
juntos en un solo PDF a fin de mes) contra la hoja CARGA AQUI del Excel de
Ventas ("... Ventas ... ANALISIS.xlsx", un archivo aparte del Excel
Cierre) -- items y $ vendidos por cada departamento del mes.

Es de solo lectura: nunca escribe nada en el Excel ni genera un archivo
para descargar -- el resultado se muestra en pantalla (verde/rojo por
chequeo), según lo que ya se decidió para toda la sección Controles.
"""

import calendar
import os
from datetime import datetime, timedelta

from openpyxl import load_workbook

from controles_utils import eval_literal_sum_cell
from reporte_diario import (
    DATE_SCAN_COLUMN,
    HEADER_ROW,
    HEADER_START_COLUMN,
    STORE_INFO_COL_CASH,
    STORE_INFO_COL_CREDIT,
    STORE_INFO_COL_DESC_COMB,
    STORE_INFO_COL_DESC_OTROS,
    STORE_INFO_COL_LOCAL_ACCOUNTS,
    STORE_INFO_COL_NETWORK_REVENUE,
    STORE_INFO_COL_NON_FUEL,
    STORE_INFO_COL_SALES_FUEL,
    STORE_INFO_COL_TAX_COLLECT,
    STORE_INFO_COL_VOLUME,
    _find_store_info_sheet,
    _get_carga_aqui_sheet,
    _normalize_department_label,
    _resolve_department_columns,
    _store_info_row_for_day,
    _strip_cell,
    build_department_column_map,
    extract_store_info_from_pdf,
    find_row_for_calendar_day,
    parse_elistar_daily_pdf_ocr,
)

TOLERANCE = 0.01

# Other (columna U, 21 -- un lugar después de TC) es la única columna de
# pago que entra en el Total Revenue (W=SUM(S:V)) sin tener su propio
# chequeo individual pedido por el usuario -- reporte_diario.py nunca la
# escribe (no tiene nombre propio ahí), se deriva de STORE_INFO_COL_CREDIT
# para no duplicar el número de columna a mano.
STORE_INFO_COL_OTHER = STORE_INFO_COL_CREDIT + 1


def _eval_literal_sum_cell(value):
    return eval_literal_sum_cell(value, "Store Info")


# Cada entrada agrupa la columna de Store Info a sumar por día con el
# nombre de campo que usa `_extract_store_info_fields` (reporte_diario.py)
# para ese mismo dato en el PDF -- una sola lista maneja tanto la suma
# mensual del Excel como, más abajo, el armado de los 9 chequeos puntuales
# pedidos por el usuario (Volume, Total Fuel Sales, Fuel Discounts, Total
# Non Fuel Sales, Other Discounts, Total Taxes Collected, Cash, TC, Local
# Accounts), sin duplicar la lista de columnas dos veces.
_MONTHLY_SUM_COLUMNS = {
    "volume": STORE_INFO_COL_VOLUME,
    "sales_fuel": STORE_INFO_COL_SALES_FUEL,
    "desc_comb": STORE_INFO_COL_DESC_COMB,
    "non_fuel_total": STORE_INFO_COL_NON_FUEL,
    "desc_otros": STORE_INFO_COL_DESC_OTROS,
    "tax_collect": STORE_INFO_COL_TAX_COLLECT,
    "cash": STORE_INFO_COL_CASH,
    "credit_terms": STORE_INFO_COL_CREDIT,
    "other": STORE_INFO_COL_OTHER,
    "local_accounts": STORE_INFO_COL_LOCAL_ACCOUNTS,
    "network_revenue": STORE_INFO_COL_NETWORK_REVENUE,
}


def _sum_store_info_month(sheet, year, month):
    """
    Suma, para cada día real del mes (year, month), cada columna de
    Store Info listada en `_MONTHLY_SUM_COLUMNS` -- incluye "other" (U),
    que no tiene su propio chequeo pero hace falta para reconstruir Total
    Revenue (Cash+TC+Other+Local Account, columnas S-V -- misma fórmula
    que la propia plantilla usa en W, =SUM(S:V)).

    Nunca lee con data_only=True: eso dependería de que el archivo haya
    sido recalculado y guardado en Excel real antes de subirlo acá -- en
    vez de eso, recalcula cada total en Python a partir de las mismas
    columnas de entrada, así el control funciona apenas se sube el Excel,
    sin pedirle al usuario un paso extra.

    Devuelve (totales, missing_days, periods_seen) -- un día sin fecha real
    cargada en columna A se reporta en missing_days en vez de aportar 0 en
    silencio, para que el usuario sepa que ese día todavía no se cargó (y
    no confundir "no cargado" con "cargado en cero"). `periods_seen` es el
    conjunto de todos los (año, mes) reales encontrados en columna A de
    las filas revisadas -- como esta función ubica cada fila por posición
    (row = day+1 del mes pedido, no por fecha), un Excel Cierre de OTRO
    mes igual tiene una fecha real en cada una de esas filas -- solo que
    de un período distinto. Detectarlo acá permite que
    `check_store_info_monthly` avise que el Excel cargado no pertenece al
    período del PDF (subir esa combinación no está mal -- sirve para
    comparar -- pero hay que avisarlo).
    """
    days_in_month = calendar.monthrange(year, month)[1]
    totals = {field: 0.0 for field in _MONTHLY_SUM_COLUMNS}
    missing_days = []
    periods_seen = set()
    for day in range(1, days_in_month + 1):
        row = _store_info_row_for_day(day)
        date_value = sheet.cell(row=row, column=1).value
        if not isinstance(date_value, datetime):
            missing_days.append(day)
            continue
        periods_seen.add((date_value.year, date_value.month))
        for field, column in _MONTHLY_SUM_COLUMNS.items():
            totals[field] += _eval_literal_sum_cell(sheet.cell(row=row, column=column).value)
    return totals, missing_days, periods_seen


def _rounded_diff(pdf_value, excel_value):
    """round(..., 2), but never -0.00 -- confuses more than it clarifies."""
    diff = round(pdf_value - excel_value, 2)
    return diff if diff != 0 else 0.0


def _build_check(label, pdf_value, excel_value, unit="$", tolerance=TOLERANCE):
    diff = _rounded_diff(pdf_value, excel_value)
    return {
        "label": label,
        "pdf_value": round(pdf_value, 2),
        "excel_value": round(excel_value, 2),
        "diff": diff,
        "ok": abs(diff) <= tolerance,
        "unit": unit,
    }


def check_store_info_monthly(cierre_path, monthly_pdf_path):
    """
    Cruza el PDF "Store Sales Summary Report" mensual (páginas 1-2, "Store
    Sales Summary Report" + "Method of Payment Totals Report") contra la
    hoja Store Info del Excel Cierre para el mismo mes/año (tomado del
    propio período del PDF, "PERIOD FROM: ... TO: ..."). Devuelve un dict
    con un chequeo por cada dato que imprime esas dos páginas -- Volume,
    Total Fuel Sales, Fuel Discounts, Total Non Fuel Sales, Other
    Discounts, Total Taxes Collected, Cash, tarjetas de crédito (TC),
    Local Accounts, Network Revenue y Total Revenue -- listo para mostrar
    en pantalla. Nunca escribe ni descarga nada.
    """
    cierre_path = os.path.abspath(str(cierre_path).strip())
    if not os.path.isfile(cierre_path):
        raise FileNotFoundError(f"Excel Cierre no encontrado: {cierre_path}")

    pdf_fields = extract_store_info_from_pdf(monthly_pdf_path, start_page_index=0)
    # El "PERIOD FROM" que imprime este mismo reporte (Store Sales Summary
    # Report, reusado tal cual del extractor de Reporte Diario) siempre viene
    # un día antes del día real que cubre el reporte -- por el cierre nocturno
    # del negocio, no por un error del POS. Reporte Diario ya corrige esto
    # sumando 1 día antes de decidir en qué fila escribir (ver
    # write_store_info_row en reporte_diario.py) -- acá hace falta la misma
    # corrección antes de decidir qué mes buscar en Store Info, si no, un PDF
    # mensual de agosto (que imprime algo como "PERIOD FROM: Jul 31 ... TO:
    # Aug 31 ...") terminaría comparándose contra julio en vez de agosto.
    # "period_to" no necesita el mismo ajuste -- ya cae en el día real
    # correcto (ver el ejemplo de _parse_period_from_to_line: un reporte
    # "para" el día 19 imprime FROM=día 18, TO=día 19).
    period_from = pdf_fields["from_date"] + timedelta(days=1)
    period_to = pdf_fields["to_date"]

    workbook = load_workbook(cierre_path, data_only=False)
    try:
        sheet = _find_store_info_sheet(workbook)
        totals, missing_days, periods_seen = _sum_store_info_month(sheet, period_from.year, period_from.month)
    finally:
        workbook.close()

    other_periods = sorted(periods_seen - {(period_from.year, period_from.month)})
    period_mismatch = [f"{month:02d}/{year}" for year, month in other_periods]

    pdf_credit_terms = sum(pdf_fields["credit_terms"])
    excel_total_revenue = (
        totals["cash"] + totals["credit_terms"] + totals["other"] + totals["local_accounts"]
    )

    checks = [
        # Cada día carga su propio Volume ya redondeado a centésimas de
        # galón (ver Store Info columna E) -- sumar 28-31 valores así
        # redondeados contra el total que imprime el propio POS a fin de
        # mes puede acumular unos pocos centésimos de diferencia sin que
        # haya ningún día salteado o mal cargado (confirmado contra agosto
        # real: 0.04 gal de diferencia en 40,464.81 gal, un mes sin ningún
        # error real). Tolerancia más ancha solo acá -- el resto de los
        # chequeos son montos en dólares y se quedan en la tolerancia
        # estándar de ±$0.01.
        _build_check("Volume", pdf_fields["volume"], totals["volume"], unit="gal", tolerance=0.15),
        _build_check("Total Fuel Sales", pdf_fields["sales_fuel"], totals["sales_fuel"]),
        _build_check("Fuel Discounts", pdf_fields["desc_comb"], totals["desc_comb"]),
        _build_check("Total Non Fuel Sales", pdf_fields["non_fuel_total"], totals["non_fuel_total"]),
        _build_check("Other Discounts", pdf_fields["desc_otros"], totals["desc_otros"]),
        _build_check("Total Taxes Collected", pdf_fields["tax_collect"], totals["tax_collect"]),
        _build_check("Cash", pdf_fields["cash"], totals["cash"]),
        _build_check("Tarjetas de Crédito (TC)", pdf_credit_terms, totals["credit_terms"]),
        _build_check("Local Accounts", pdf_fields["local_accounts"], totals["local_accounts"]),
        _build_check("Network Revenue", pdf_fields["network_revenue"], totals["network_revenue"]),
        _build_check("Total Revenue", pdf_fields["total_revenue"], excel_total_revenue),
    ]

    return {
        "period_from": period_from,
        "period_to": period_to,
        "missing_days": missing_days,
        "period_mismatch": period_mismatch,
        "checks": checks,
        "all_ok": all(check["ok"] for check in checks) and not missing_days,
    }


def _list_ventas_department_columns(sheet):
    """
    (label, count_col, amount_col) por cada departamento real (no
    protegido) de la fila HEADER_ROW de CARGA AQUI, en orden, sin
    duplicados.

    Reusa build_department_column_map -- la misma resolución de columnas
    que ya usa inject_daily_sales para escribir -- en vez de reinventarla,
    así que un departamento con COUNT/NET SALES invertido entre fila 3 y 4
    (el único caso en que esa función corrige el orden) se lee de la
    columna correcta acá también. Se recorre la fila de encabezados una
    vez más solo para conservar el texto real de cada label (mapping ya lo
    normaliza a minúsculas para poder matchear, no sirve para mostrar).
    """
    column_map, _protected = build_department_column_map(sheet)
    seen_coords = set()
    entries = []
    max_col = max(sheet.max_column, HEADER_START_COLUMN)
    for col in range(HEADER_START_COLUMN, max_col + 1):
        label = _strip_cell(sheet.cell(row=HEADER_ROW, column=col).value)
        if not label:
            continue
        coords = column_map.get(_normalize_department_label(label))
        if coords is None or coords in seen_coords:
            continue
        seen_coords.add(coords)
        entries.append((label, coords[0], coords[1]))
    return entries


def _sum_ventas_departments_month(sheet, year, month, department_columns):
    """
    Suma, para cada día real de (year, month), el par COUNT | NET SALES de
    cada departamento de `department_columns` en CARGA AQUI.

    Mismo criterio que _sum_store_info_month: un día sin fila propia (fecha
    real no encontrada en columna A) se reporta en missing_days en vez de
    aportar 0 en silencio; `periods_seen` detecta un Excel de otro mes.
    Ubica cada fila con find_row_for_calendar_day -- el mismo buscador por
    fecha real que ya usa Reporte Diario, nunca un offset fijo.
    """
    days_in_month = calendar.monthrange(year, month)[1]
    totals = {
        (count_col, amount_col): {"count": 0.0, "amount": 0.0}
        for _label, count_col, amount_col in department_columns
    }
    missing_days = []
    periods_seen = set()
    for day in range(1, days_in_month + 1):
        try:
            row = find_row_for_calendar_day(sheet, day)
        except ValueError:
            missing_days.append(day)
            continue
        date_value = sheet.cell(row=row, column=DATE_SCAN_COLUMN).value
        if isinstance(date_value, datetime):
            periods_seen.add((date_value.year, date_value.month))
        for _label, count_col, amount_col in department_columns:
            key = (count_col, amount_col)
            totals[key]["count"] += eval_literal_sum_cell(
                sheet.cell(row=row, column=count_col).value, "CARGA AQUI"
            )
            totals[key]["amount"] += eval_literal_sum_cell(
                sheet.cell(row=row, column=amount_col).value, "CARGA AQUI"
            )
    return totals, missing_days, periods_seen


def _aggregate_department_records(records, column_map):
    """
    Resuelve cada registro de departamento leído del PDF a su columna real
    de CARGA AQUI con el mismo resolver que ya usa inject_daily_sales para
    escribir (_resolve_department_columns) -- así "HOT DOGS & SANDWICH"
    (como lo imprime el reporte mensual) cae en la misma columna que "HOT
    DOGS" (el nombre real de esa columna en el Excel), en vez de compararse
    como si fueran dos departamentos distintos.

    Un registro que no resuelve a ninguna columna conocida se reporta en
    `unmatched` -- salvo el caso de ruido puro (count=0 y amount=0 a la
    vez), que se descarta directo: la línea "PERIOD FROM: ... TO: ..." de
    la propia página, repetida antes de la tabla real, a veces se cuela
    como si fuera una fila de departamento vacía.
    """
    totals_by_coords = {}
    unmatched = []
    for record in records:
        if record.get("is_total"):
            continue
        if record["count"] == 0 and record["amount"] == 0:
            continue
        coords = _resolve_department_columns(record["department"], column_map)
        if coords is None:
            unmatched.append(record["department"])
            continue
        bucket = totals_by_coords.setdefault(coords, {"count": 0, "amount": 0.0})
        bucket["count"] += record["count"]
        bucket["amount"] += record["amount"]
    return totals_by_coords, unmatched


def check_department_sales_monthly(ventas_path, monthly_pdf_path):
    """
    Cruza el "Department Sales Report" del PDF mensual (otra sección del
    mismo bundle que Store Sales Summary Report, período tomado de su
    propia línea "PERIOD FROM: ... TO: ...") contra la hoja CARGA AQUI del
    Excel de Ventas -- items y $ vendidos por cada departamento del mes.

    Reusa el mismo motor que ya lee este reporte día a día en Reporte
    Diario (parse_elistar_daily_pdf_ocr, build_department_column_map,
    _resolve_department_columns) aplicado a la versión mensual del mismo
    reporte en vez de a un PDF por día. Devuelve un chequeo (items + $) por
    cada departamento real de CARGA AQUI más una fila de total general
    (contra el total impreso por el propio reporte). Nunca escribe ni
    descarga nada.
    """
    ventas_path = os.path.abspath(str(ventas_path).strip())
    if not os.path.isfile(ventas_path):
        raise FileNotFoundError(f"Excel de Ventas no encontrado: {ventas_path}")
    extension = os.path.splitext(ventas_path)[1].lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise ValueError("El Excel de Ventas debe ser .xlsx o .xlsm.")

    records, diagnostics = parse_elistar_daily_pdf_ocr(monthly_pdf_path, start_page_index=0)
    period = diagnostics.get("period")
    if period is None:
        raise ValueError(
            'No se encontró la línea "PERIOD FROM: ... TO: ..." en el Department '
            "Sales Report del PDF."
        )
    # Mismo ajuste de +1 día que check_store_info_monthly, y por el mismo
    # motivo -- es la misma línea "PERIOD FROM" del mismo reporte del POS,
    # solo que leída de otra página del bundle.
    period_from = period["from_date"] + timedelta(days=1)
    period_to = period["to_date"]

    workbook = load_workbook(ventas_path, data_only=False)
    try:
        sheet = _get_carga_aqui_sheet(workbook)
        column_map, _protected = build_department_column_map(sheet)
        department_columns = _list_ventas_department_columns(sheet)
        excel_totals, missing_days, periods_seen = _sum_ventas_departments_month(
            sheet, period_from.year, period_from.month, department_columns
        )
    finally:
        workbook.close()

    pdf_totals, unmatched_departments = _aggregate_department_records(records, column_map)
    unmatched_departments = sorted(set(unmatched_departments))

    other_periods = sorted(periods_seen - {(period_from.year, period_from.month)})
    period_mismatch = [f"{month:02d}/{year}" for year, month in other_periods]

    checks = []
    for label, count_col, amount_col in department_columns:
        coords = (count_col, amount_col)
        excel_bucket = excel_totals.get(coords, {"count": 0.0, "amount": 0.0})
        pdf_bucket = pdf_totals.get(coords, {"count": 0, "amount": 0.0})
        excel_count = int(round(excel_bucket["count"]))
        pdf_count = int(round(pdf_bucket["count"]))
        count_diff = pdf_count - excel_count
        amount_diff = _rounded_diff(pdf_bucket["amount"], excel_bucket["amount"])
        checks.append(
            {
                "label": label,
                "pdf_count": pdf_count,
                "excel_count": excel_count,
                "count_diff": count_diff,
                "count_ok": count_diff == 0,
                "pdf_amount": round(pdf_bucket["amount"], 2),
                "excel_amount": round(excel_bucket["amount"], 2),
                "amount_diff": amount_diff,
                "amount_ok": abs(amount_diff) <= TOLERANCE,
                "ok": count_diff == 0 and abs(amount_diff) <= TOLERANCE,
            }
        )

    printed_totals = diagnostics.get("printed_totals")
    excel_grand_count = sum(check["excel_count"] for check in checks)
    excel_grand_amount = round(sum(check["excel_amount"] for check in checks), 2)
    total_check = None
    if printed_totals is not None:
        total_count_diff = printed_totals["count"] - excel_grand_count
        total_amount_diff = _rounded_diff(printed_totals["amount"], excel_grand_amount)
        total_check = {
            "pdf_count": printed_totals["count"],
            "excel_count": excel_grand_count,
            "count_diff": total_count_diff,
            "count_ok": total_count_diff == 0,
            "pdf_amount": round(printed_totals["amount"], 2),
            "excel_amount": excel_grand_amount,
            "amount_diff": total_amount_diff,
            "amount_ok": abs(total_amount_diff) <= TOLERANCE,
        }

    return {
        "period_from": period_from,
        "period_to": period_to,
        "missing_days": missing_days,
        "period_mismatch": period_mismatch,
        "unmatched_departments": unmatched_departments,
        "checks": checks,
        "total_check": total_check,
        "all_ok": (
            all(check["ok"] for check in checks)
            and not missing_days
            and not unmatched_departments
            and (total_check is None or (total_check["count_ok"] and total_check["amount_ok"]))
        ),
    }
