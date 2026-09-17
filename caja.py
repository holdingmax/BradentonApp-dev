"""
Módulo Caja -- Carga de Datos: cruza en el momento lo que ya guardaron Chase
Bank, Lottery y Reporte Diario (Store Info), más los Gastos cargados a mano
acá mismo (ver caja_db.py) -- sin subir ni escribir ningún Excel, mismo
criterio que un módulo de Controles.

Pedido explícito del usuario (2026-09-12, cuarta tanda): "caja no es algo
que tendria que estar en el modulo de herramientas ya que no se le tiene
que cargar ningun PDF o excel para completar" -- el viejo módulo de
Herramientas (`/caja`, escribía las columnas K/N/S/T del Excel Cierre real a
partir de 3 Excel subidos: Cierre + Chase + Lottery) se eliminó del todo.
Quedan acá, sin relación con eso, las piezas de solo lectura que sigue
usando Controles → Caja (`controles_caja.py`) para leer la hoja CAJA real de
un Excel Cierre recién subido y validarla contra los Mayores de Chase/Caja
-- `_get_caja_sheet`/`_iter_caja_dates`/las constantes CAJA_COL_*.
"""

import calendar
from datetime import date, datetime, timedelta

import caja_db
import chase_db
import lottery_db
import reportes_db

CAJA_SHEET_NAME = "CAJA"
CAJA_DATA_START_ROW = 4
CAJA_COL_DATE = 1  # A — Fecha del día de negocio (la que se usa para matchear)
CAJA_COL_CHASE_DEPOSITS = 11  # K
CAJA_COL_EXPENSES_CASH = 13  # M — gastos pagados con caja (cada celda con un comentario del proveedor/persona pagada)
CAJA_COL_FOOD_ICE = 19  # S
CAJA_COL_FOOD_ICE_LABEL = 20  # T — de qué se trata el importe de S (Food Truck / ICE MACHINE)

CHASE_DETALLE_DEPOSITO = "DEPOSITO"
# Depósito de la máquina/casino Gettel: cuenta como depósito normal en K.
CHASE_DETALLE_GETTEL = "DEPOSITO GETTEL"

# Nota: el Detalle real de un depósito físico de hielo es "DEPOSITO VENTA ICE"
# (Type DEPOSIT) — "VENTA ICE" a secas es la venta reportada por la máquina vía
# ACH_CREDIT/MISC_CREDIT, que NO es un depósito físico y no cuenta acá.
CHASE_DETALLE_FOOD_TRUCK = "FOOD TRUCK"
CHASE_DETALLE_ICE = "DEPOSITO VENTA ICE"
FOOD_ICE_LABELS = {
    CHASE_DETALLE_FOOD_TRUCK: "Food Truck",
    CHASE_DETALLE_ICE: "ICE MACHINE",
}
# Orden fijo de presentación cuando un mismo día combina ambos.
FOOD_ICE_LABEL_ORDER = (CHASE_DETALLE_ICE, CHASE_DETALLE_FOOD_TRUCK)


def _normalize_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def _get_caja_sheet(workbook):
    for name in workbook.sheetnames:
        if name.strip().lower() == CAJA_SHEET_NAME.lower():
            return workbook[name]
    raise ValueError(
        f'Hoja "{CAJA_SHEET_NAME}" no encontrada. Disponibles: {", ".join(workbook.sheetnames)}'
    )


def _iter_caja_dates(sheet):
    """Fechas (columna A) de CAJA desde la fila 4 hasta que se acaben los datos del mes."""
    row = CAJA_DATA_START_ROW
    while True:
        date_key = _normalize_date(sheet.cell(row=row, column=CAJA_COL_DATE).value)
        if date_key is None:
            return
        yield row, date_key
        row += 1


def _format_food_ice_label(detalle_keys):
    labels = [FOOD_ICE_LABELS[key] for key in FOOD_ICE_LABEL_ORDER if key in detalle_keys]
    return ", ".join(labels)


# ---------------------------------------------------------------------------
# Caja -- Carga de Datos (2026-09-12, ampliado más tarde el mismo día): el
# mismo cruce K/S/T/N de arriba, pero leído directo de chase_db/lottery_db
# en vez de un Excel -- pedido explícito del usuario: "el modulo de caja...
# se completaria automaticamente con los datos que haya guardado en el
# chase de tal mes". Ampliado a pedido explícito del usuario para replicar
# TODA la hoja CAJA real, no solo K/N/S/T -- se investigó la hoja real
# (Cierre 08-26.xlsx) con openpyxl antes de escribir nada acá, confirmando
# columna por columna qué es fórmula/qué se pisa a mano:
#   E Total SALES = Store Info!R, F Cash = Store Info!S,
#   G TC = Store Info!T (suma de "credit_terms"), H Other = Store Info!U,
#   J Total Revenue = Store Info!W (ya guardado tal cual, no hace falta
#   recalcularlo -- Store Info!W es la MISMA suma que J calcularía),
#   K CHASE = Depósitos, L OUT CASH y M EXPENSES CASH = a mano (nada más
#   los tiene), N lottery = Cuenta Final, O DIF EFECT = F-K+H-L-M-N (fórmula
#   real confirmada contra el archivo), P Saldo = Saldo(día anterior)+O.
# Solo L/M (a mano) y el Saldo Inicial/Final de cada mes necesitan guardado
# propio (caja_db.py) -- todo lo demás se recalcula en el momento, sin
# escribir nada, mismo espíritu que los módulos de Controles.
# ---------------------------------------------------------------------------

def _collect_chase_amounts_from_db(year, month):
    """Mismo criterio de detección que _collect_chase_amounts, pero sobre las filas ya guardadas en chase_db (columna Detalle) en vez de un Excel."""
    deposits_by_date = {}
    food_ice_by_date = {}
    food_ice_labels_by_date = {}
    gettel_dates = set()

    for tx in chase_db.get_month_transactions(year, month):
        detalle_norm = (tx.get("detalle") or "").strip().upper()
        if not detalle_norm:
            continue
        try:
            date_key = datetime.strptime(tx["posting_date"], "%Y-%m-%d").date()
        except (TypeError, ValueError):
            continue
        amount = tx.get("amount")
        if amount is None:
            continue

        if detalle_norm == CHASE_DETALLE_DEPOSITO:
            deposits_by_date[date_key] = deposits_by_date.get(date_key, 0.0) + float(amount)
        elif detalle_norm == CHASE_DETALLE_GETTEL:
            deposits_by_date[date_key] = deposits_by_date.get(date_key, 0.0) + float(amount)
            gettel_dates.add(date_key)
        elif detalle_norm in FOOD_ICE_LABELS:
            food_ice_by_date[date_key] = food_ice_by_date.get(date_key, 0.0) + float(amount)
            food_ice_labels_by_date.setdefault(date_key, set()).add(detalle_norm)

    return deposits_by_date, food_ice_by_date, food_ice_labels_by_date, gettel_dates


def _resolve_opening_balance(year, month, _chain=True):
    """
    El Saldo Inicial de un mes es, por default, el Saldo Final del mes
    anterior -- pedido explícito del usuario (2026-09-12): "el saldo
    inicial de la caja de un mes deberia ser el saldo final de la caja del
    mes anterior". Si el usuario ya fijó un Saldo Inicial a mano para ESTE
    mes (caja_db.set_month_opening_balance), ese manda SIEMPRE -- este
    chequeo corre primero sin importar `_chain`, así que un override manual
    nunca se pisa, ni siquiera cuando esta función se llama para resolver
    el mes ANTERIOR de otro mes (ver más abajo).

    Si no hay override para este mes, y `_chain` es True, se encadena del
    mes anterior: su propio ajuste manual de Saldo Final si lo tiene, o si
    no, su Saldo corrido calculado -- recorriendo su propio reporte una
    sola vez con `_chain=False` (nunca se recursa un tercer mes hacia
    atrás; alcanza con un solo salto). Si el mes anterior tampoco tiene
    ningún dato, default 0.0 -- queda flaggeado (opening_source="default")
    para que la pantalla pueda pedirle al usuario que lo cargue a mano.
    """
    settings = caja_db.get_month_settings(year, month)
    if settings and settings.get("opening_balance") is not None:
        return settings["opening_balance"], "manual"
    if not _chain:
        return 0.0, "default"

    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    prev_settings = caja_db.get_month_settings(prev_year, prev_month)
    if prev_settings and prev_settings.get("closing_balance_override") is not None:
        return prev_settings["closing_balance_override"], "prev_month_override"

    prev_report = build_month_report_from_db(prev_year, prev_month, _recursion_guard=False)
    if prev_report["rows"]:
        return prev_report["computed_closing_balance"], "prev_month_computed"
    return 0.0, "default"


_EMPTY_TOTALS_FIELDS = (
    "total_sales", "cash", "tc", "other_amount", "total_revenue",
    "deposit", "food_ice", "cuenta_final", "expenses_cash", "dif_efect",
)


def _empty_future_month_report(year, month):
    """
    Un mes que todavía no empezó no tiene ningún dato posible -- pedido
    explícito del usuario (2026-09-12): "los meses siguientes al actual
    deberian aparecer con todos los datos en 0 obviamente". Se corta acá,
    sin ni siquiera intentar encadenar un Saldo Inicial -- proyectar un
    saldo hacia un mes que no ocurrió todavía sería confuso, no informativo.
    """
    days_in_month = calendar.monthrange(year, month)[1]
    rows = [
        {
            "date": date(year, month, day).isoformat(),
            "day": day,
            "total_sales": 0.0, "cash": 0.0, "tc": 0.0, "other_amount": 0.0, "total_revenue": 0.0,
            "deposit": 0.0, "is_gettel": False, "food_ice": 0.0, "food_ice_label": "",
            "expenses_cash": 0.0, "cuenta_final": 0.0, "dif_efect": 0.0, "saldo": 0.0,
        }
        for day in range(1, days_in_month + 1)
    ]
    return {
        "rows": rows,
        "totals": {field: 0.0 for field in _EMPTY_TOTALS_FIELDS},
        "total_lottery": None,
        "opening_balance": 0.0,
        "opening_source": "future",
        "computed_closing_balance": 0.0,
        "closing_balance_override": None,
        "effective_closing_balance": 0.0,
    }


def build_month_report_from_db(year, month, _recursion_guard=True):
    """
    Un renglón por día del mes, mismas columnas que la hoja CAJA real (ver
    el bloque de comentarios de arriba) -- calculado en el momento contra
    chase_db/lottery_db/reportes_db/caja_db, sin necesitar ningún Excel ni
    subir nada nuevo acá (Chase/Lottery/Reporte Diario ya se cargan por sus
    propios módulos; Gastos se cargan a mano acá mismo, ver
    caja_db.set_day_expenses). OUT CASH (columna L del Excel real) se sacó
    del todo 2026-09-12 -- pedido explícito del usuario, "no se usa".
    """
    if (year, month) > (date.today().year, date.today().month):
        return _empty_future_month_report(year, month)

    deposits_by_date, food_ice_by_date, food_ice_labels_by_date, gettel_dates = _collect_chase_amounts_from_db(
        year, month
    )
    store_info_by_date = {row["date"]: row for row in reportes_db.get_month_store_info(year, month)}
    expenses_by_date = caja_db.get_month_expenses(year, month)

    days_in_month = calendar.monthrange(year, month)[1]

    opening_balance, opening_source = _resolve_opening_balance(year, month, _chain=_recursion_guard)
    running_saldo = opening_balance

    rows = []
    totals = {field: 0.0 for field in _EMPTY_TOTALS_FIELDS}
    any_lottery = False

    for day in range(1, days_in_month + 1):
        d = date(year, month, day)
        key = d.isoformat()

        deposit = deposits_by_date.get(d)
        food_ice = food_ice_by_date.get(d)
        food_ice_label = _format_food_ice_label(food_ice_labels_by_date.get(d, set()))

        lottery_day = lottery_db.get_day(d)
        cuenta_final = None
        if lottery_day is not None:
            cuenta_final = lottery_db.decorate_day(lottery_day).get("cuenta_final")
        if cuenta_final is not None:
            any_lottery = True

        info = store_info_by_date.get(key) or {}
        total_sales = info.get("total_sales")
        cash = info.get("cash")
        tc = round(sum(info.get("credit_terms") or []), 2) if info else None
        other_amount = info.get("other_amount")
        total_revenue = info.get("total_revenue")
        local_accounts = info.get("local_accounts")

        expenses_cash = expenses_by_date.get(key)

        # DIF EFECT =+F-K+H-M-N (fórmula real de la hoja CAJA, confirmada
        # contra el archivo -- sin el término L/OUT CASH, sacado del todo)
        # -- una celda en blanco vale 0 en la aritmética de Excel, así que
        # acá también: un día sin nada cargado todavía da DIF EFECT 0 y el
        # Saldo sigue igual al del día anterior, ni más ni menos, igual que
        # pasaría abriendo la hoja real con esas celdas vacías.
        dif_efect = round(
            (cash or 0.0) - (deposit or 0.0) + (other_amount or 0.0)
            - (expenses_cash or 0.0) - (cuenta_final or 0.0),
            2,
        )
        saldo_previous = running_saldo
        running_saldo = round(running_saldo + dif_efect, 2)

        # Desgloses para el cuadro flotante "qué valores usaron para llegar
        # a ese resultado" -- pedido explícito del usuario (2026-09-19):
        # "cuando hagas click en una celda en la que haya una formula de
        # suma que te muestre... que es lo que suma", mismo mecanismo ya
        # usado en Store Info (ver reporte_store_info_historial.html).
        total_revenue_breakdown = [
            ("Cash", cash),
            ("Tarjeta/Crédito", tc),
            ("Other", other_amount),
            ("Local Acc.", local_accounts),
        ]
        dif_efect_breakdown = [
            ("Cash", cash or 0.0),
            ("Depósitos (resta)", -(deposit or 0.0)),
            ("Other", other_amount or 0.0),
            ("Gastos (resta)", -(expenses_cash or 0.0)),
            ("Lottery Cuenta Final (resta)", -(cuenta_final or 0.0)),
        ]
        saldo_breakdown = [
            ("Saldo día anterior", saldo_previous),
            ("Dif Efect", dif_efect),
        ]

        totals["total_sales"] += total_sales or 0.0
        totals["cash"] += cash or 0.0
        totals["tc"] += tc or 0.0
        totals["other_amount"] += other_amount or 0.0
        totals["total_revenue"] += total_revenue or 0.0
        totals["deposit"] += deposit or 0.0
        totals["food_ice"] += food_ice or 0.0
        totals["cuenta_final"] += cuenta_final or 0.0
        totals["expenses_cash"] += expenses_cash or 0.0
        totals["dif_efect"] += dif_efect

        rows.append(
            {
                "date": key,
                "day": day,
                "total_sales": round(total_sales, 2) if total_sales is not None else None,
                "cash": round(cash, 2) if cash is not None else None,
                "tc": tc,
                "other_amount": round(other_amount, 2) if other_amount is not None else None,
                "total_revenue": round(total_revenue, 2) if total_revenue is not None else None,
                "deposit": round(deposit, 2) if deposit is not None else None,
                "is_gettel": d in gettel_dates,
                "food_ice": round(food_ice, 2) if food_ice is not None else None,
                "food_ice_label": food_ice_label,
                "expenses_cash": expenses_cash,
                "cuenta_final": round(cuenta_final, 2) if cuenta_final is not None else None,
                "dif_efect": dif_efect,
                "saldo": running_saldo,
                "total_revenue_breakdown": total_revenue_breakdown,
                "dif_efect_breakdown": dif_efect_breakdown,
                "saldo_breakdown": saldo_breakdown,
            }
        )

    settings = caja_db.get_month_settings(year, month)
    closing_override = settings.get("closing_balance_override") if settings else None
    computed_closing = rows[-1]["saldo"] if rows else round(opening_balance, 2)

    return {
        "rows": rows,
        "totals": {k: round(v, 2) for k, v in totals.items()},
        "total_lottery": round(totals["cuenta_final"], 2) if any_lottery else None,
        "opening_balance": round(opening_balance, 2),
        "opening_source": opening_source,
        "computed_closing_balance": computed_closing,
        "closing_balance_override": round(closing_override, 2) if closing_override is not None else None,
        "effective_closing_balance": round(
            closing_override if closing_override is not None else computed_closing, 2
        ),
    }


def build_caja_export_workbook(report, year, month, dest_path):
    """
    Excel NUEVO (no toca ningún archivo real) con la Caja del mes -- mismo
    criterio y mismo método que reporte_diario.build_store_info_export_
    workbook. Pedido explícito del usuario (2026-09-19): "hagamos algo
    igual del exportar la caja y que quede con el formato que tenia en el
    cierre" -- y, tras la primera versión, mandó su propia copia real
    limpiada (`CAJA 08-2026 WEB.xlsx`) con dos rondas de correcciones:

    (a) Las columnas "hs"/TC/Local Account de la hoja real están OCULTAS
    (`column_dimensions[letra].hidden = True`) -- Alfonso nunca las ve, no
    hace falta replicarlas (no se muestran en pantalla tampoco). Las DOS
    columnas de Fecha (desde/hasta) sí están siempre visibles -- van las
    dos, sin ninguna columna "hs" al lado.
    (b) El "espacio entre columnas" real no es un padding cualquiera --
    son DOS columnas en blanco genuinas (Q ancho ~11.9, R ancho ~6.6, más
    angosta) entre el cuadro principal (hasta Saldo/P) y la nota suelta de
    Food Truck/Ice (S/T, "FONDO FIJO").
    (c) La primera fila de datos real (P3) trae el saldo de arranque del
    mes pintado de azul (FF0070C0), con la etiqueta "Caja al INICIO" al
    lado -- se agrega acá como su propia fila antes del día 1.
    (d) El orden real es Depósitos(K)/Gastos(M)/Lottery(N) -- Gastos antes
    que Lottery, no al revés (así se corrigió también en la pantalla).
    (e) Dif Efect/Total Revenue/Saldo se exportan siempre como VALOR, NUNCA
    como fórmula de Excel (a diferencia del archivo real, que sí las tiene
    como fórmulas) -- pedido explícito del usuario, para que el número que
    ve acá sea siempre el mismo que ya calculó la página, sin depender de
    que Excel recalcule nada.
    """
    import openpyxl
    from openpyxl.styles import Alignment as XlAlignment, Border as XlBorder, Font as XlFont, PatternFill, Side as XlSide

    HEADER_GRAY = PatternFill("solid", fgColor="FFD9D9D9")
    HEADER_ORANGE = PatternFill("solid", fgColor="FFFFC000")
    DATA_GREEN_LIGHT = PatternFill("solid", fgColor="FFA9D18E")
    DATA_GREEN = PatternFill("solid", fgColor="FF92D050")
    DATA_YELLOW = PatternFill("solid", fgColor="FFFFFF00")
    DATA_BLUEGRAY = PatternFill("solid", fgColor="FFADB9CA")
    DATA_GRAY_LIGHT = PatternFill("solid", fgColor="FFF2F2F2")
    OPENING_BLUE = PatternFill("solid", fgColor="FF0070C0")
    MONEY_FMT = '"$" #,##0.00'
    PLAIN_FMT = '#,##0.00_ ;[Red]\\-#,##0.00\\ '
    THIN_SIDE = XlSide(style="thin", color="FF000000")
    MEDIUM_SIDE = XlSide(style="medium", color="FF000000")
    THIN_BORDER = XlBorder(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
    HEADER_BORDER = XlBorder(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=MEDIUM_SIDE)

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "CAJA"

    sheet.cell(row=1, column=1, value=f"BGS - {month:02d}/{year} - Caja")
    sheet["A1"].font = XlFont(name="Segoe UI", bold=True, size=14, underline="single")
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
    sheet.row_dimensions[1].height = 21

    # (texto, fill, fmt) -- columna por columna, en el mismo orden que ya
    # muestra carga_datos_caja_historial.html (Gastos antes que Lottery).
    # Las columnas 12/13 quedan vacías a propósito (el mismo par de
    # columnas Q/R en blanco del archivo real).
    header_spec = [
        ("Fecha", HEADER_GRAY, None),
        ("Fecha", HEADER_GRAY, None),
        ("Total Sales", HEADER_GRAY, MONEY_FMT),
        ("Cash", HEADER_GRAY, MONEY_FMT),
        ("Other", HEADER_GRAY, MONEY_FMT),
        ("Total Revenue", HEADER_GRAY, MONEY_FMT),
        ("Depósitos", HEADER_ORANGE, PLAIN_FMT),
        ("Gastos", HEADER_ORANGE, PLAIN_FMT),
        ("Lottery\nCuenta Final", HEADER_ORANGE, PLAIN_FMT),
        ("Dif Efect", HEADER_ORANGE, PLAIN_FMT),
        ("Saldo", HEADER_ORANGE, PLAIN_FMT),
        (None, None, None),
        (None, None, None),
        ("Food Truck/Ice", None, MONEY_FMT),
    ]
    data_fill_by_col = {
        3: DATA_GREEN_LIGHT,
        4: DATA_GREEN,
        5: DATA_YELLOW,
        6: DATA_BLUEGRAY,
        7: DATA_YELLOW,
        9: HEADER_ORANGE,
        10: DATA_YELLOW,
        11: DATA_GRAY_LIGHT,
    }
    SALDO_COL = 11

    for col, (text, fill, _fmt) in enumerate(header_spec, start=1):
        cell = sheet.cell(row=2, column=col, value=text)
        if text is None:
            continue
        cell.font = XlFont(bold=True)
        if fill is not None:
            cell.fill = fill
        cell.alignment = XlAlignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = HEADER_BORDER
    sheet.row_dimensions[2].height = 45.75
    sheet.freeze_panes = "A4"

    # Fila de arranque -- el saldo inicial del mes, con su etiqueta al lado
    # (mismo lugar que P3/Q3 en el archivo real) -- pedido explícito del
    # usuario (2026-09-19): el azul cubre TODA la fila, de A a K (no solo
    # la celda de Saldo), aunque el resto quede vacío.
    r = 3
    for col in range(1, SALDO_COL + 1):
        cell = sheet.cell(row=r, column=col)
        cell.fill = OPENING_BLUE
        cell.border = THIN_BORDER
    opening_cell = sheet.cell(row=r, column=SALDO_COL, value=report.get("opening_balance"))
    opening_cell.font = XlFont(bold=True, color="FFFFFFFF")
    opening_cell.number_format = PLAIN_FMT
    label_cell = sheet.cell(row=r, column=SALDO_COL + 1, value="Caja al INICIO")
    label_cell.font = XlFont(bold=True, italic=True)

    # Columnas con el número resaltado en negrita -- pedido explícito del
    # usuario (2026-09-19), confirmado contra el archivo real: Total
    # Sales/Total Revenue/Depósitos/Lottery/Dif Efect/Saldo van en
    # negrita; Other y Gastos quedan sin resaltar. Cash además va en
    # cursiva (igual que la etiqueta "Caja al INICIO" de al lado).
    BOLD_COLS = {3, 6, 7, 9, 10, 11}  # Total Sales, Total Revenue, Depósitos, Lottery, Dif Efect, Saldo
    BOLD_ITALIC_COLS = {4}  # Cash

    for row in report["rows"]:
        r = sheet.max_row + 1
        business_date = datetime.strptime(row["date"], "%Y-%m-%d") if row.get("date") else None
        values = {
            1: business_date,
            2: business_date + timedelta(days=1) if business_date else None,
            3: row.get("total_sales"),
            4: row.get("cash"),
            5: row.get("other_amount"),
            6: row.get("total_revenue"),
            7: row.get("deposit"),
            8: row.get("expenses_cash"),
            9: row.get("cuenta_final"),
            10: row.get("dif_efect"),
            11: row.get("saldo"),
            14: row.get("food_ice"),
        }
        for col, value in values.items():
            cell = sheet.cell(row=r, column=col, value=value)
            cell.border = THIN_BORDER
            if col in (1, 2):
                cell.number_format = "mm-dd-yy"
                cell.alignment = XlAlignment(horizontal="center")
            else:
                cell.number_format = header_spec[col - 1][2] or MONEY_FMT
            if col in BOLD_ITALIC_COLS:
                cell.font = XlFont(bold=True, italic=True)
            elif col in BOLD_COLS:
                cell.font = XlFont(bold=True)
            fill = data_fill_by_col.get(col)
            if fill is not None:
                cell.fill = fill

    # Fila "Total del mes" -- mismo criterio que la fila 35 real (SUM por
    # columna, con el mismo color de bloque que sus datos, en negrita) --
    # Saldo no se suma (es un acumulado corrido, se muestra el saldo final
    # efectivo del mes, igual que ya hace la pantalla).
    r = sheet.max_row + 1
    totals = report["totals"]
    total_values = {
        1: "Total del mes",
        3: totals.get("total_sales"),
        4: totals.get("cash"),
        5: totals.get("other_amount"),
        6: totals.get("total_revenue"),
        7: totals.get("deposit"),
        8: totals.get("expenses_cash"),
        9: report.get("total_lottery"),
        10: totals.get("dif_efect"),
        11: report.get("effective_closing_balance"),
        14: totals.get("food_ice"),
    }
    for col, value in total_values.items():
        cell = sheet.cell(row=r, column=col, value=value)
        cell.border = THIN_BORDER
        cell.font = XlFont(bold=True)
        if col > 1:
            cell.number_format = header_spec[col - 1][2] or MONEY_FMT
        fill = data_fill_by_col.get(col, HEADER_GRAY if col == 1 else None)
        if fill is not None:
            cell.fill = fill

    # Anchos confirmados 1:1 contra la copia real (CAJA 08-2026 WEB.xlsx,
    # columnas visibles A/C/E/F/H/J/K/M/N/O/P mapeadas a las nuestras en
    # el mismo orden) -- las dos últimas antes de Food Truck/Ice son el
    # mismo par de columnas vacías Q/R (11.9 y 6.6, la segunda más
    # angosta) que separan el cuadro principal de esa nota en el archivo
    # real.
    for col_letter, width in zip(
        "ABCDEFGHIJKLMN",
        (10.3, 11.4, 12.9, 12.7, 8.6, 12.9, 11.9, 9.7, 10.4, 12.7, 11.6, 11.9, 6.6, 12.3),
    ):
        sheet.column_dimensions[col_letter].width = width

    workbook.save(dest_path)
    return dest_path


def _fmt_money_pdf(value):
    return "—" if value is None else "{:,.2f}".format(value)


def build_caja_export_pdf(report, year, month, dest_path):
    """
    PDF (líneas/bordes + colores, sin fórmulas) con la Caja del mes --
    pedido explícito del usuario (2026-09-19), alternativa liviana al
    Excel (build_caja_export_workbook) -- mismas columnas que ya se ven
    en /carga-datos/caja (sin las dos columnas de fecha ni las columnas
    vacías de espacio, que solo tienen sentido dentro del Excel real).
    El Saldo Inicial va como nota arriba de la tabla, no como fila propia.
    Colores idénticos a los del export a Excel (mismos hex, ver
    build_caja_export_workbook) -- pedido explícito del usuario tras ver
    la primera versión sin color.
    """
    from pdf_export import build_simple_table_pdf

    GRAY, ORANGE, YELLOW = "#D9D9D9", "#FFC000", "#FFFF00"
    GREEN_LIGHT, GREEN, BLUEGRAY, GRAY_LIGHT = "#A9D18E", "#92D050", "#ADB9CA", "#F2F2F2"
    header_fill_by_col = {0: GRAY, 1: GRAY, 2: GRAY, 3: GRAY, 4: GRAY, 5: ORANGE, 6: ORANGE, 7: ORANGE, 8: ORANGE, 9: ORANGE}
    data_fill_by_col = {1: GREEN_LIGHT, 2: GREEN, 3: YELLOW, 4: BLUEGRAY, 5: YELLOW, 7: ORANGE, 8: YELLOW, 9: GRAY_LIGHT}

    headers = [
        "Día", "Total Sales", "Cash", "Other", "Total Revenue",
        "Depósitos", "Gastos", "Lottery\nCta. Final", "Dif Efect", "Saldo", "Food Truck/Ice",
    ]
    table_rows = [
        [
            "{:02d}".format(row["day"]),
            _fmt_money_pdf(row.get("total_sales")),
            _fmt_money_pdf(row.get("cash")),
            _fmt_money_pdf(row.get("other_amount")),
            _fmt_money_pdf(row.get("total_revenue")),
            _fmt_money_pdf(row.get("deposit")),
            _fmt_money_pdf(row.get("expenses_cash")),
            _fmt_money_pdf(row.get("cuenta_final")),
            _fmt_money_pdf(row.get("dif_efect")),
            _fmt_money_pdf(row.get("saldo")),
            _fmt_money_pdf(row.get("food_ice")),
        ]
        for row in report["rows"]
    ]
    totals = report["totals"]
    table_rows.append(
        [
            "Total",
            _fmt_money_pdf(totals.get("total_sales")),
            _fmt_money_pdf(totals.get("cash")),
            _fmt_money_pdf(totals.get("other_amount")),
            _fmt_money_pdf(totals.get("total_revenue")),
            _fmt_money_pdf(totals.get("deposit")),
            _fmt_money_pdf(totals.get("expenses_cash")),
            _fmt_money_pdf(report.get("total_lottery")),
            _fmt_money_pdf(totals.get("dif_efect")),
            _fmt_money_pdf(report.get("effective_closing_balance")),
            _fmt_money_pdf(totals.get("food_ice")),
        ]
    )
    title = f"Caja — {month:02d}/{year} — Saldo Inicial: ${_fmt_money_pdf(report.get('opening_balance'))}"
    col_widths_mm = [16, 29, 26, 23, 29, 26, 26, 29, 26, 29, 29]
    return build_simple_table_pdf(
        dest_path,
        title,
        headers,
        table_rows,
        col_widths_mm,
        header_fill_by_col=header_fill_by_col,
        data_fill_by_col=data_fill_by_col,
        bold_last_row=True,
    )



def get_available_years():
    """
    Años con datos de Caja -- pedido explícito del usuario (2026-09-17,
    módulo "Reportes"): a diferencia de Chase/Lottery/Store Info, Caja no
    guarda casi nada por su cuenta (se arma cruzando esas 3 fuentes +
    gastos a mano, ver build_month_report_from_db más arriba), así que no
    hay una sola tabla de la que sacar "qué años tienen Caja". Se unen los
    años de las 4 fuentes reales -- si CUALQUIERA de ellas tiene algo ese
    año, Caja puede tener un reporte con algo cargado.
    """
    years = set(chase_db.get_available_years())
    years |= set(lottery_db.get_available_years())
    years |= set(reportes_db.get_store_info_years())
    years |= set(caja_db.get_expense_years())
    return sorted(years)


def build_caja_pdf_resumen(report, year, month, dest_path):
    """
    PDF resumido de Caja para el módulo "Reportes" -- pedido explícito
    del usuario (2026-09-17, sesión siguiente): "agrega tambien el de
    caja" (a los otros 3 reportes de Reportes, todos ya resumidos con
    Detalle/Total). A diferencia de build_caja_export_pdf (una fila por
    CADA día del mes -- el que sigue usando el botón "Exportar PDF" ya
    existente de /carga-datos/caja, sin tocar), acá se muestra solo la
    fila de totales de ese mismo reporte (`report["totals"]`, ya validada
    y usada tal cual en el PDF completo) en formato Detalle/Total, mismo
    estilo que los otros 3 reportes de este módulo.

    Saldo Inicial/Saldo Final NO salen de sumar nada (son un saldo
    corrido, no un monto del día) -- se muestran tal cual ya los expone
    el reporte (`opening_balance`/`effective_closing_balance`), mismo
    criterio que ya usa build_caja_export_pdf en el título.
    """
    from pdf_export import build_simple_table_pdf

    totals = report["totals"]
    table_rows = [
        ["Saldo Inicial", _fmt_money_pdf(report.get("opening_balance"))],
        ["Total Sales", _fmt_money_pdf(totals.get("total_sales"))],
        ["Cash", _fmt_money_pdf(totals.get("cash"))],
        ["Tarjeta/Crédito", _fmt_money_pdf(totals.get("tc"))],
        ["Other", _fmt_money_pdf(totals.get("other_amount"))],
        ["Total Revenue", _fmt_money_pdf(totals.get("total_revenue"))],
        ["Depósitos", _fmt_money_pdf(totals.get("deposit"))],
        ["Gastos", _fmt_money_pdf(totals.get("expenses_cash"))],
        ["Lottery Cta. Final", _fmt_money_pdf(report.get("total_lottery"))],
        ["Dif Efectivo (mes)", _fmt_money_pdf(totals.get("dif_efect"))],
        ["Food Truck/Ice", _fmt_money_pdf(totals.get("food_ice"))],
        ["Saldo Final", _fmt_money_pdf(report.get("effective_closing_balance"))],
    ]

    return build_simple_table_pdf(
        dest_path,
        f"Caja — Resumen — {month:02d}/{year}",
        ["Detalle", "Total"],
        table_rows,
        col_widths_mm=[110, 80],
        bold_last_row=True,
        company_header=True,
    )
