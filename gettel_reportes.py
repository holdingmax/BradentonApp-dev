"""
Módulo "Reportes -> Gettel" (pedido explícito del usuario, 2026-09-21):
"ese excel que te pase es el que quiero que uses para crear el reporte de
gettel, usando los formatos y colores que tiene" -- reproduce, con
openpyxl, el formato/colores reales de `Gettel formato 08-26 testeo.xlsx`
(decodificado con openpyxl -- fórmulas + colores + fills leídos del
archivo real, nunca adivinados), con 4 hojas por mes (pedido textual del
usuario):

  1. Pendiente MM.YYYY (mes ANTERIOR) -- los días de Gettel/Toyota que
     quedaron sin cobrar del mes anterior.
  2. Gettel-Toyota MM.YYYY (mes actual) -- todos los días de Gettel/Toyota
     cargados este mes (gettel_db.gettel_toyota_days).
  3. Pago Cupones -- los pagos/cupones cargados este mes (gettel_pagos,
     ver gettel_db.get_month_pagos) + el resumen Total del mes/Pendiente
     Mes Anterior/Cobrado/Diferencia/Total a Pagar.
  4. Pendiente MM.YYYY (mes ACTUAL, resultante) -- los días de este mes
     que los pagos recibidos NO llegaron a cubrir. Son, exactamente, los
     que pasan a ser la hoja 1 ("Pendiente Mes Anterior") del reporte del
     MES QUE VIENE -- ver resolve_month, que encadena de un mes al
     siguiente igual que gettel_pagos._resolve_pendiente_anterior.

Pedido textual completo del usuario: "este reporte deberia tener 4 hojas
que muestran en la primera los dias de gettel pendientes del mes
anterior, en la segunda se muestra el mes actual, en la tercera los pagos
que se fueron haciendo y en la ultima los dias que quedaron pendientes
del mes actual (eso se mide en base a los pagos que llegaron durante el
mes y que tanto llegaron a cubrir los dias que se debia el mes anterior y
el mes actual, los dias que no se llegaron a cubrir con los pagos son los
que quedan pendientes para el mes siguiente)".

--- Fórmula real por bloque de días (decodificada de las 3 hojas de
"días" del archivo real -- Pendiente 07.2026 / Gettel-Toyota 08.2026 /
Pendiente Gettel-Toyota 08.2026, las tres con la MISMA fórmula al pie) ---
    VERIFICAR             = suma de cada columna del bloque de días.
    Rebate x Galón (0.02) = -0.02 * (Galones Gettel + Galones Toyota).
    Charge 3%             = 0.03 * (Monto Gettel + Monto Toyota + Rebate).
    TOTAL CUPÓN A COBRAR  = Monto + Rebate + Charge -- celda amarilla
                             (FFFF00), letra roja, negrita (E22/E40
                             reales).

La hoja real usa "Local Account" (columna C -- un monto de banco/POS que
esta app JAMÁS cargó en ningún lado) como base de Monto, en vez de
Gettel+Toyota. Acá se usa Gettel $ + Toyota $ (gettel_db.gettel_toyota_
days) como base -- es la misma que ya usan las celdas ayuda L7/L8 de la
hoja real del mes actual ("Galones"/"Gettel/Toyota $", ambas marcadas
"ACTUALIZAR" en el archivo real), y es el único dato equivalente que el
sistema tiene guardado. AVISO a Alfonso (2026-09-21): el "Total del mes"
que ya usa gettel_pagos.py (el módulo existente "Cuadro de Pagos") suma
SOLO Gettel $ (gettel_db.get_month_gettel_amount), sin Toyota $ ni
Rebate/Charge -- da un monto distinto (menor) al de esta fórmula real.
Ese módulo viejo NO se tocó acá -- se le avisa a Alfonso, no se cambia
nada sin que lo pida.

--- Alocación día a día (para armar la hoja 4) ---
Verificado 1:1 contra el archivo real de ejemplo (pago recibido de
23,707.70 en agosto/2026, con un Pendiente Mes Anterior de 10,300.03 ya
calculado en la hoja "Pendiente 07.2026"): el pago cubre PRIMERO, de más
viejo a más nuevo, los días que quedaron pendientes del mes anterior (ahí
el Rebate/Charge YA está calculado como bloque entero -- si alcanza para
cubrirlo completo no se lo vuelve a partir día por día); lo que sobra se
aplica después a los días del mes actual, ahí SÍ día a día (por eso hace
falta recalcular Rebate/Charge en cada corte posible: son un % del
bloque, no algo que se pueda sumar suelto por día). El primer día (de
cualquiera de los dos bloques) que el pago ya no alcanza a cubrir -- y
todos los que quedan después de ese -- son los que se muestran en la hoja
4 y pasan a ser el Pendiente Mes Anterior del mes que viene. Con los
números reales del archivo de ejemplo esto da EXACTAMENTE los mismos 13
días (19/08 al 31/08) que ya trae ese archivo en su hoja "Pendiente
Gettel-Toyota 08.2026" -- confirmado por script antes de escribir este
módulo.
"""

import gettel_db

REBATE_RATE = 0.02
CHARGE_RATE = 0.03


def _day_amount(day):
    return (day.get("gettel_amount") or 0.0) + (day.get("toyota_amount") or 0.0)


def _day_gallons(day):
    return (day.get("gettel_gallons") or 0.0) + (day.get("toyota_gallons") or 0.0)


def _block_total(days):
    """Rebate/Charge/Total recalculados SOBRE ESTE bloque de días (no es una suma de valores por día -- son % del bloque completo, ver docstring del módulo)."""
    sum_gettel_amt = sum(d.get("gettel_amount") or 0.0 for d in days)
    sum_gettel_gal = sum(d.get("gettel_gallons") or 0.0 for d in days)
    sum_toyota_amt = sum(d.get("toyota_amount") or 0.0 for d in days)
    sum_toyota_gal = sum(d.get("toyota_gallons") or 0.0 for d in days)
    sum_amt = sum_gettel_amt + sum_toyota_amt
    sum_gal = sum_gettel_gal + sum_toyota_gal
    rebate = -REBATE_RATE * sum_gal
    charge = CHARGE_RATE * (sum_amt + rebate)
    total = sum_amt + rebate + charge
    return {
        "sum_gettel_amt": sum_gettel_amt,
        "sum_gettel_gal": sum_gettel_gal,
        "sum_toyota_amt": sum_toyota_amt,
        "sum_toyota_gal": sum_toyota_gal,
        "sum_amt": sum_amt,
        "sum_gal": sum_gal,
        "rebate": rebate,
        "charge": charge,
        "total": total,
    }


def _month_days(year, month):
    rows = gettel_db.get_month_days(year, month)
    return [
        {
            "date": r["date"],
            "gettel_amount": r.get("gettel_amount") or 0.0,
            "gettel_gallons": r.get("gettel_gallons") or 0.0,
            "toyota_amount": r.get("toyota_amount") or 0.0,
            "toyota_gallons": r.get("toyota_gallons") or 0.0,
        }
        for r in rows
    ]


def _covered_prefix(days, budget):
    """
    Cuántos días (desde el más viejo, principio de `days`) alcanza a
    cubrir un pago de `budget` -- la mayor cantidad de días consecutivos
    desde el principio cuyo propio Total Cupón a Cobrar (Rebate/Charge
    recalculado sobre ESE bloque) no supera el presupuesto. Devuelve
    (cantidad_cubierta, total_cubierto).
    """
    covered = 0
    covered_total = 0.0
    for k in range(1, len(days) + 1):
        total = _block_total(days[:k])["total"]
        if total <= budget + 0.005:  # tolerancia de medio centavo por redondeo
            covered, covered_total = k, total
        else:
            break
    return covered, covered_total


_MAX_CHAIN_MONTHS = 240  # 20 años -- tope de seguridad, ver docstring


def resolve_month(year, month, _chain=True, _depth=0):
    """
    Arma el ledger completo del mes -- la data de las 4 hojas del reporte.

    A diferencia de gettel_pagos._resolve_pendiente_anterior (que nunca
    encadena más de UN mes hacia atrás -- ahí es una simplificación
    deliberada porque ese módulo solo maneja un monto suelto, sin días),
    acá SÍ hace falta encadenar TODO lo que haga falta hacia atrás: la
    hoja 1 de un mes tiene que ser exactamente la hoja 4 (con sus días
    reales) del mes anterior, para que el reporte de cada mes siga
    reflejando la cadena completa de lo que nunca se terminó de cobrar.
    La recursión SIEMPRE termina sola apenas encuentra un mes sin ningún
    dato cargado (`has_any_data=False`) o un override manual (pedido
    explícito del usuario -- "ancla" la cadena ahí) -- `_depth`/
    `_MAX_CHAIN_MONTHS` es solo un tope de seguridad ante datos corruptos
    o un bug, nunca se espera llegar a pisarlo en uso normal.
    """
    settings = gettel_db.get_pago_month_settings(year, month)
    override = settings.get("pendiente_anterior_override") if settings else None

    if override is not None:
        # Mismo override manual que ya usa gettel_pagos.py (comparten la
        # tabla gettel_pagos_months) -- acá no hay desglose por día
        # porque es un monto tipeado a mano, no calculado de datos reales.
        # También ancla la cadena acá (no sigue mirando más atrás).
        pendiente_anterior_days = []
        pendiente_anterior_total = float(override)
        pendiente_anterior_source = "manual"
    elif _chain and _depth < _MAX_CHAIN_MONTHS:
        prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
        prev = resolve_month(prev_year, prev_month, _chain=True, _depth=_depth + 1)
        if prev["has_any_data"]:
            pendiente_anterior_days = prev["pendiente_siguiente_days"]
            pendiente_anterior_total = prev["pendiente_siguiente_total"]
            pendiente_anterior_source = "prev_month_computed"
        else:
            pendiente_anterior_days = []
            pendiente_anterior_total = 0.0
            pendiente_anterior_source = "default"
    else:
        pendiente_anterior_days = []
        pendiente_anterior_total = 0.0
        pendiente_anterior_source = "default"

    mes_days = _month_days(year, month)
    mes_breakdown = _block_total(mes_days)

    pagos = gettel_db.get_month_pagos(year, month)
    cobrado = round(sum(p.get("total_cupon") or 0.0 for p in pagos), 2)

    total_a_pagar = pendiente_anterior_total + mes_breakdown["total"]
    diferencia = cobrado - total_a_pagar

    # Alocación de lo cobrado este mes, más viejo primero: primero lo
    # pendiente del mes anterior, después los días del mes actual (ver
    # docstring del módulo).
    remaining_budget = cobrado
    pendiente_no_desglosable = 0.0

    if pendiente_anterior_days:
        covered_prev, covered_prev_total = _covered_prefix(pendiente_anterior_days, remaining_budget)
        remaining_budget -= covered_prev_total
        uncovered_prev_days = pendiente_anterior_days[covered_prev:]
    else:
        uncovered_prev_days = []
        if remaining_budget >= pendiente_anterior_total:
            remaining_budget -= pendiente_anterior_total
        else:
            # Pendiente Mes Anterior ingresado a mano (sin desglose por
            # día) y el pago del mes no llegó a cubrirlo entero -- lo que
            # falta queda como un monto suelto, sin días propios que
            # mostrar (ver el note en el sheet builder).
            pendiente_no_desglosable = pendiente_anterior_total - remaining_budget
            remaining_budget = 0.0

    if uncovered_prev_days or pendiente_no_desglosable:
        # No alcanzó a cubrirse todo el mes anterior -- nada del mes
        # actual se considera cubierto todavía (se cubre de más viejo a
        # más nuevo, sin saltear bloques).
        covered_curr = 0
    else:
        covered_curr, _covered_curr_total = _covered_prefix(mes_days, max(remaining_budget, 0.0))

    pendiente_siguiente_days = uncovered_prev_days + mes_days[covered_curr:]
    pendiente_siguiente_breakdown = _block_total(pendiente_siguiente_days)
    pendiente_siguiente_total = round(pendiente_siguiente_breakdown["total"] + pendiente_no_desglosable, 2)

    has_any_data = bool(mes_days) or bool(pagos) or bool(pendiente_anterior_days) or bool(pendiente_anterior_total)

    return {
        "year": year,
        "month": month,
        "has_any_data": has_any_data,
        "pendiente_anterior_days": pendiente_anterior_days,
        "pendiente_anterior_total": round(pendiente_anterior_total, 2),
        "pendiente_anterior_source": pendiente_anterior_source,
        "mes_days": mes_days,
        "mes_breakdown": mes_breakdown,
        "pagos": pagos,
        "cobrado": cobrado,
        "total_a_pagar": round(total_a_pagar, 2),
        "diferencia": round(diferencia, 2),
        "pendiente_siguiente_days": pendiente_siguiente_days,
        "pendiente_siguiente_breakdown": pendiente_siguiente_breakdown,
        "pendiente_siguiente_total": pendiente_siguiente_total,
        "pendiente_no_desglosable": round(pendiente_no_desglosable, 2),
    }


def get_available_years():
    """Años con algo cargado -- días de Gettel/Toyota, pagos, o algún override -- mismo criterio que gettel_pagos.get_available_years, más los años de gettel_toyota_days."""
    import gettel_pagos

    years = set(gettel_db.get_toyota_days_years()) | set(gettel_pagos.get_available_years())
    return sorted(years)


def _fmt_money(value):
    return "—" if value is None else "{:,.2f}".format(value)


# ---------------------------------------------------------------------------
# Export a Excel -- 4 hojas, mismos colores/formatos que el archivo real
# (Gettel formato 08-26 testeo.xlsx, decodificado con openpyxl).
# ---------------------------------------------------------------------------

def _write_days_sheet(workbook, sheet_title, month_label, days, total_value, note=None):
    """
    Una hoja de "días" (Pendiente mes anterior / Gettel-Toyota mes actual
    / Pendiente mes actual resultante) -- mismas 3 hojas del archivo real,
    misma fórmula VERIFICAR/Rebate x Galón/Charge 3%/TOTAL CUPÓN A COBRAR
    al pie (ver docstring del módulo). Columnas Local Account/DIF del
    archivo real NO se replican -- el sistema nunca cargó ese dato (ver
    aviso en el docstring); en su lugar se agrega "Total $ (Gettel +
    Toyota)" por día, que es la base real que sí usa este reporte.

    `total_value` es SIEMPRE el número que se muestra en la celda final
    TOTAL CUPÓN A COBRAR (autoridad del caller, ver resolve_month) -- no
    necesariamente igual a Rebate/Charge recalculado sobre `days` cuando
    hay un Pendiente Mes Anterior ingresado a mano sin desglose (ver
    `note` en ese caso).
    """
    import openpyxl
    from openpyxl.styles import Alignment as XlAlignment, Border as XlBorder, Font as XlFont, PatternFill, Side as XlSide

    HEADER_GRAY = PatternFill("solid", fgColor="FFD9D9D9")
    TOTAL_YELLOW = PatternFill("solid", fgColor="FFFFFF00")
    MONEY_FMT = '"$" #,##0.00'
    GAL_FMT = "#,##0.000"
    THIN_SIDE = XlSide(style="thin", color="FF000000")
    THIN_BORDER = XlBorder(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
    CENTER = XlAlignment(horizontal="center", vertical="center")
    RIGHT = XlAlignment(horizontal="right", vertical="center")
    RED_BOLD = XlFont(bold=True, color="FFFF0000")

    sheet = workbook.create_sheet(title=sheet_title[:31])

    title_cell = sheet.cell(row=1, column=1, value=f"CONTROL GETTEL/TOYOTA - {month_label}")
    title_cell.font = XlFont(name="Segoe UI", bold=True, size=14)
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    sheet.row_dimensions[1].height = 21

    headers = ["Fecha", "Gettel $", "Gettel Gal", "Toyota $", "Toyota Gal", "Total $ (Gettel + Toyota)"]
    for col, text in enumerate(headers, start=1):
        cell = sheet.cell(row=3, column=col, value=text)
        cell.font = XlFont(bold=True)
        cell.fill = HEADER_GRAY
        cell.alignment = XlAlignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER

    row = 4
    for day in days:
        date_cell = sheet.cell(row=row, column=1, value=day["date"])
        date_cell.number_format = "mm/dd/yyyy"
        g_amt = sheet.cell(row=row, column=2, value=day["gettel_amount"])
        g_gal = sheet.cell(row=row, column=3, value=day["gettel_gallons"])
        t_amt = sheet.cell(row=row, column=4, value=day["toyota_amount"])
        t_gal = sheet.cell(row=row, column=5, value=day["toyota_gallons"])
        tot = sheet.cell(row=row, column=6, value=day["gettel_amount"] + day["toyota_amount"])
        g_amt.number_format = MONEY_FMT
        t_amt.number_format = MONEY_FMT
        tot.number_format = MONEY_FMT
        g_gal.number_format = GAL_FMT
        t_gal.number_format = GAL_FMT
        for col in range(1, 7):
            cell = sheet.cell(row=row, column=col)
            cell.border = THIN_BORDER
            cell.alignment = CENTER
        row += 1

    if not days:
        empty_cell = sheet.cell(row=row, column=1, value="Sin días cargados en este bloque.")
        empty_cell.font = XlFont(italic=True)
        row += 1

    breakdown = _block_total(days)
    verificar_row = row + 1
    sheet.cell(row=verificar_row, column=1, value="VERIFICAR").font = XlFont(italic=True)
    v_vals = [
        None,
        breakdown["sum_gettel_amt"],
        breakdown["sum_gettel_gal"],
        breakdown["sum_toyota_amt"],
        breakdown["sum_toyota_gal"],
        breakdown["sum_amt"],
    ]
    for col in range(2, 7):
        cell = sheet.cell(row=verificar_row, column=col, value=v_vals[col - 1])
        cell.font = XlFont(bold=True)
        cell.number_format = GAL_FMT if col in (3, 5) else MONEY_FMT

    rebate_row = verificar_row + 2
    sheet.merge_cells(start_row=rebate_row, start_column=1, end_row=rebate_row, end_column=5)
    label = sheet.cell(row=rebate_row, column=1, value=f"Rebate x Galón ({REBATE_RATE:g})")
    label.font = XlFont(bold=True)
    label.alignment = RIGHT
    value_cell = sheet.cell(row=rebate_row, column=6, value=breakdown["rebate"])
    value_cell.font = XlFont(bold=True)
    value_cell.number_format = MONEY_FMT

    charge_row = rebate_row + 2
    sheet.merge_cells(start_row=charge_row, start_column=1, end_row=charge_row, end_column=5)
    label = sheet.cell(row=charge_row, column=1, value=f"Charge {CHARGE_RATE*100:g}%")
    label.font = XlFont(bold=True)
    label.alignment = RIGHT
    value_cell = sheet.cell(row=charge_row, column=6, value=breakdown["charge"])
    value_cell.font = XlFont(bold=True)
    value_cell.number_format = MONEY_FMT

    total_row = charge_row + 2
    sheet.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=5)
    label = sheet.cell(row=total_row, column=1, value="TOTAL CUPÓN A COBRAR")
    label.font = XlFont(bold=True)
    label.alignment = RIGHT
    total_cell = sheet.cell(row=total_row, column=6, value=total_value)
    total_cell.font = RED_BOLD
    total_cell.fill = TOTAL_YELLOW
    total_cell.number_format = MONEY_FMT
    total_cell.border = THIN_BORDER

    if note:
        note_row = total_row + 1
        note_cell = sheet.cell(row=note_row, column=1, value=note)
        note_cell.font = XlFont(italic=True, size=9)
        sheet.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=6)

    for col_letter, width in zip("ABCDEF", (13, 13, 11, 13, 11, 20)):
        sheet.column_dimensions[col_letter].width = width

    return sheet


def _write_pago_cupones_sheet(workbook, report, year, month, prev_year, prev_month, prev_label):
    """
    Hoja 3 -- "Pago Cupones", mismos pagos/colores/celdas unificadas que
    ya usa gettel_pagos.build_pagos_export_workbook (misma tabla real
    "Pago Cupones" decodificada), pero con "Total del mes" corregido a la
    fórmula real completa (Gettel+Toyota+Rebate+Charge -- ver
    report["mes_breakdown"]) en vez del monto parcial que sigue usando el
    módulo viejo gettel_pagos.py (ver aviso en el docstring del módulo).
    """
    import gettel_pagos  # solo para reusar grouped_pagos -- no toca ni llama a build_month_report
    from openpyxl.styles import Alignment as XlAlignment, Border as XlBorder, Font as XlFont, PatternFill, Side as XlSide

    HEADER_GRAY = PatternFill("solid", fgColor="FFD9D9D9")
    DATA_GREEN = PatternFill("solid", fgColor="FF92D050")
    MONEY_FMT = '"$" #,##0.00'
    THIN_SIDE = XlSide(style="thin", color="FF000000")
    THIN_BORDER = XlBorder(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
    CENTER = XlAlignment(horizontal="center", vertical="center")
    CENTER_WRAP = XlAlignment(horizontal="center", vertical="center", wrap_text=True)

    sheet = workbook.create_sheet(title="Pago Cupones")

    title_cell = sheet.cell(row=1, column=1, value=f"CONTROL - {month:02d}/{year} (CUPONES QUE ENVIA RICK A FINAL DE MES)")
    title_cell.font = XlFont(name="Segoe UI", bold=True, size=14)
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    sheet.row_dimensions[1].height = 21

    headers = ["Fecha", "Pago N°", "Transc N°", "Total Cupón"]
    for col, text in enumerate(headers, start=1):
        cell = sheet.cell(row=3, column=col, value=text)
        cell.font = XlFont(bold=True)
        cell.fill = HEADER_GRAY
        cell.alignment = CENTER_WRAP
        cell.border = THIN_BORDER

    f3 = sheet.cell(row=3, column=6, value="Total")
    f3.font = XlFont(bold=True)
    f3.fill = HEADER_GRAY
    f3.alignment = CENTER
    f3.border = THIN_BORDER

    sheet.merge_cells(start_row=4, start_column=6, end_row=5, end_column=6)
    f4 = sheet.cell(row=4, column=6, value=report["mes_breakdown"]["total"])
    f4.font = XlFont(bold=True)
    f4.number_format = MONEY_FMT
    f4.alignment = CENTER
    f4.border = THIN_BORDER
    sheet.merge_cells(start_row=4, start_column=7, end_row=5, end_column=7)
    g4 = sheet.cell(row=4, column=7, value=f"Total del mes de Gettel+Toyota ({month:02d}/{year})")
    g4.font = XlFont(bold=True)
    g4.alignment = CENTER_WRAP
    g4.border = THIN_BORDER

    for col in (6, 7):
        sheet.merge_cells(start_row=7, start_column=col, end_row=8, end_column=col)
    f7 = sheet.cell(row=7, column=6, value=report["pendiente_anterior_total"])
    f7.font = XlFont(bold=True)
    f7.number_format = MONEY_FMT
    f7.fill = DATA_GREEN
    f7.alignment = CENTER
    f7.border = THIN_BORDER
    g7 = sheet.cell(row=7, column=7)
    g7.fill = DATA_GREEN
    g7.border = THIN_BORDER
    sheet.merge_cells(start_row=7, start_column=8, end_row=7, end_column=10)
    h7 = sheet.cell(row=7, column=8, value=f"(DIAS PENDIENTES DEL MES ANTERIOR -- {prev_label})")
    h7.alignment = XlAlignment(horizontal="left", vertical="center")

    row_cursor = 4
    for group in gettel_pagos.grouped_pagos(report["pagos"]):
        start_row = row_cursor
        for pago in group["rows"]:
            sheet.cell(row=row_cursor, column=3, value=pago.get("transc_n") or "").alignment = CENTER
            sheet.cell(row=row_cursor, column=3).border = THIN_BORDER
            amount_cell = sheet.cell(row=row_cursor, column=4, value=pago.get("total_cupon"))
            amount_cell.number_format = MONEY_FMT
            amount_cell.alignment = CENTER
            amount_cell.border = THIN_BORDER
            for col in (1, 2):
                sheet.cell(row=row_cursor, column=col).border = THIN_BORDER
            row_cursor += 1
        end_row = row_cursor - 1
        fecha_value = group["rows"][0].get("fecha")
        pago_n_value = group["rows"][0].get("pago_n")
        if end_row > start_row:
            sheet.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
            sheet.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
        fecha_cell = sheet.cell(row=start_row, column=1, value=fecha_value)
        fecha_cell.alignment = CENTER
        fecha_cell.font = XlFont(bold=True)
        pago_cell = sheet.cell(row=start_row, column=2, value=pago_n_value)
        pago_cell.alignment = CENTER
        pago_cell.font = XlFont(bold=True)

    totals_row = max(row_cursor, 32)
    c_label = sheet.cell(row=totals_row, column=3, value="COBRADO")
    c_label.font = XlFont(bold=True)
    c_label.alignment = XlAlignment(horizontal="right")
    d_total = sheet.cell(row=totals_row, column=4, value=report["cobrado"])
    d_total.font = XlFont(bold=True)
    d_total.number_format = MONEY_FMT
    d_total.fill = HEADER_GRAY
    d_total.alignment = CENTER
    d_total.border = THIN_BORDER
    e_dif = sheet.cell(row=totals_row, column=5, value=report["diferencia"])
    e_dif.font = XlFont(bold=True)
    e_dif.number_format = MONEY_FMT
    e_dif.alignment = CENTER
    f_pagar = sheet.cell(row=totals_row, column=6, value=report["total_a_pagar"])
    f_pagar.font = XlFont(bold=True)
    f_pagar.number_format = MONEY_FMT
    f_pagar.fill = HEADER_GRAY
    f_pagar.alignment = CENTER
    f_pagar.border = THIN_BORDER

    note_row = totals_row + 1
    note_label = sheet.cell(
        row=note_row, column=4,
        value="a acreditarse en 48 hs" if report["diferencia"] >= 0 else "pasa a Pendiente del mes siguiente",
    )
    note_label.font = XlFont(bold=True, italic=True)
    note_label.alignment = XlAlignment(horizontal="right")

    for col_letter, width in zip("ABCDEFGHIJ", (12, 9, 12, 13, 3, 13, 22, 24, 8, 8)):
        sheet.column_dimensions[col_letter].width = width

    return sheet


def build_gettel_reportes_workbook(report, year, month, dest_path):
    """
    Arma las 4 hojas del reporte, en el orden pedido por el usuario
    (Pendiente mes anterior / Mes actual / Pago Cupones / Pendiente mes
    actual resultante), y guarda el Excel NUEVO en `dest_path` -- nunca
    toca ningún archivo real, mismo criterio que el resto de los exports
    de la app.
    """
    import openpyxl

    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)

    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)

    prev_label = f"{prev_month:02d}.{prev_year}"
    _write_days_sheet(
        workbook,
        f"Pendiente {prev_label}",
        prev_label,
        report["pendiente_anterior_days"],
        report["pendiente_anterior_total"],
        note=(
            "Pendiente ingresado a mano (sin desglose por día disponible) -- ver Cuadro de Pagos de Gettel."
            if report["pendiente_anterior_source"] == "manual"
            else None
        ),
    )

    mes_label = f"{month:02d}.{year}"
    _write_days_sheet(
        workbook,
        f"Gettel-Toyota {mes_label}",
        mes_label,
        report["mes_days"],
        report["mes_breakdown"]["total"],
    )

    _write_pago_cupones_sheet(workbook, report, year, month, prev_year, prev_month, prev_label)

    pendiente_note = None
    if report["pendiente_no_desglosable"]:
        pendiente_note = (
            f"Incluye ${report['pendiente_no_desglosable']:,.2f} sin desglose por día "
            "(arrastrado de un Pendiente Mes Anterior ingresado a mano que el pago del mes no llegó a cubrir)."
        )
    _write_days_sheet(
        workbook,
        f"Pendiente {mes_label}",
        mes_label,
        report["pendiente_siguiente_days"],
        report["pendiente_siguiente_total"],
        note=pendiente_note,
    )

    workbook.save(dest_path)
    return dest_path


# ---------------------------------------------------------------------------
# Export a PDF -- pedido explícito del usuario (2026-09-22): "el reporte de
# gettel deberia ser un PDF, no un excel, el excel de gettel deberia ir en
# su apartado de la barra lateral dentro de cuadro del mes" -- reemplaza al
# Excel como el reporte de "Reportes -> Gettel" (el Excel de arriba sigue
# existiendo, ahora servido desde /carga-datos/gettel/historial, ver
# webapp.py: reportes_gettel_excel / carga_datos_gettel_historial). Mismas
# 4 secciones que las 4 hojas del Excel, mismo `resolve_month`, sin
# reproducir ningún color/fórmula de Excel -- un PDF resumido, mismo
# criterio que Chase/EFT/Caja/Proveedores en este mismo módulo "Reportes".
# ---------------------------------------------------------------------------

def _fmt_date_ddmmyyyy_gettel(value):
    """`value` como lo guarda gettel_db (texto "YYYY-MM-DD")."""
    if not value:
        return "—"
    parts = str(value).split("-")
    if len(parts) != 3:
        return str(value)
    year, month, day = parts
    return f"{day}/{month}/{year}"


def _fmt_money_pdf_gettel(value):
    """Mismo criterio de signo que chase_rules._fmt_money_pdf -- "-$" antes del monto, nunca "$-"."""
    if value is None:
        return "—"
    if value < 0:
        return "-${:,.2f}".format(abs(value))
    return "${:,.2f}".format(value)


def _gettel_days_section(heading, days, breakdown, total_value, note=None):
    """
    Una sección de días (mismo contenido que una hoja "Pendiente"/"Gettel-
    Toyota" del Excel, sin colores) -- Fecha/Gettel $/Gettel Gal/Toyota $/
    Toyota Gal por día, más una fila TOTAL con la fórmula real (Rebate x
    Galón + Charge 3%) ya calculada por `_block_total`/`resolve_month`.
    """
    if not days:
        return {
            "heading": heading,
            "note": note or f"Sin días -- Total Cupón a Cobrar: {_fmt_money_pdf_gettel(total_value)}.",
        }
    rows = []
    for day in days:
        rows.append(
            [
                _fmt_date_ddmmyyyy_gettel(day.get("date")),
                _fmt_money_pdf_gettel(day.get("gettel_amount")),
                "{:,.2f}".format(day.get("gettel_gallons") or 0.0),
                _fmt_money_pdf_gettel(day.get("toyota_amount")),
                "{:,.2f}".format(day.get("toyota_gallons") or 0.0),
            ]
        )
    rows.append(
        [
            "Rebate x Galón (0.02)",
            "",
            "",
            "",
            _fmt_money_pdf_gettel(breakdown["rebate"]),
        ]
    )
    rows.append(["Charge 3%", "", "", "", _fmt_money_pdf_gettel(breakdown["charge"])])
    rows.append(["TOTAL CUPÓN A COBRAR", "", "", "", _fmt_money_pdf_gettel(total_value)])
    section = {
        "heading": heading,
        "headers": ["Fecha", "Gettel $", "Gettel Gal", "Toyota $", "Toyota Gal"],
        "rows": rows,
        "col_widths_mm": [50, 50, 50, 50, 50],
        "bold_last_row": True,
    }
    if note:
        section["heading"] = f"{heading} — {note}"
    return section


def build_gettel_pdf_report(report, year, month, dest_path):
    """
    PDF del módulo "Reportes -> Gettel" -- mismas 4 secciones que las 4
    hojas del Excel (`build_gettel_reportes_workbook`, ver el docstring del
    módulo para la fórmula real de cada bloque), en el mismo orden pedido
    por el usuario. `report` es el dict que devuelve `resolve_month`.
    """
    from pdf_export import build_multi_section_pdf

    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    prev_label = f"{prev_month:02d}.{prev_year}"
    mes_label = f"{month:02d}.{year}"

    pagos_rows = []
    for pago in report["pagos"]:
        pagos_rows.append(
            [
                _fmt_date_ddmmyyyy_gettel(pago.get("fecha")),
                str(pago.get("pago_n") or "—"),
                pago.get("transc_n") or "—",
                pago.get("empresa") or "—",
                _fmt_money_pdf_gettel(pago.get("total_cupon")),
            ]
        )
    if pagos_rows:
        pagos_rows.append(["TOTAL", "", "", "", _fmt_money_pdf_gettel(report["cobrado"])])

    resumen_rows = [
        [f"Total del mes ({mes_label})", _fmt_money_pdf_gettel(report["mes_breakdown"]["total"])],
        [f"Pendiente Mes Anterior ({prev_label})", _fmt_money_pdf_gettel(report["pendiente_anterior_total"])],
        ["Total a Pagar", _fmt_money_pdf_gettel(report["total_a_pagar"])],
        ["Cobrado (pagos cargados este mes)", _fmt_money_pdf_gettel(report["cobrado"])],
        ["Diferencia (Cobrado − Total a Pagar)", _fmt_money_pdf_gettel(report["diferencia"])],
    ]

    title = f"Gettel / Toyota — {mes_label}"
    sections = [
        _gettel_days_section(
            f"Pendiente {prev_label}",
            report["pendiente_anterior_days"],
            _block_total(report["pendiente_anterior_days"]),
            report["pendiente_anterior_total"],
            note=(
                "ingresado a mano, sin desglose por día -- ver Cuadro de Pagos de Gettel"
                if report["pendiente_anterior_source"] == "manual"
                else None
            ),
        ),
        _gettel_days_section(
            f"Gettel-Toyota {mes_label}",
            report["mes_days"],
            report["mes_breakdown"],
            report["mes_breakdown"]["total"],
        ),
        {
            "heading": "Pago Cupones",
            "headers": ["Fecha", "Pago N°", "N° Transacción", "Empresa", "Total Cupón"],
            "rows": pagos_rows,
            "col_widths_mm": [45, 35, 55, 45, 50],
            "bold_last_row": True,
        }
        if pagos_rows
        else {"heading": "Pago Cupones", "note": "Sin pagos cargados este mes."},
        {
            "heading": "Resumen",
            "headers": ["Detalle", "Total"],
            "rows": resumen_rows,
            "col_widths_mm": [140, 80],
            "bold_last_row": False,
        },
        _gettel_days_section(
            f"Pendiente {mes_label} (resultante -- pasa al mes que viene)",
            report["pendiente_siguiente_days"],
            report["pendiente_siguiente_breakdown"],
            report["pendiente_siguiente_total"],
            note=(
                f"Incluye {_fmt_money_pdf_gettel(report['pendiente_no_desglosable'])} sin desglose por día."
                if report["pendiente_no_desglosable"]
                else None
            ),
        ),
    ]
    build_multi_section_pdf(dest_path, title, sections, period_label=None, company_header=True)
    return dest_path
