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
from datetime import date, datetime

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
        running_saldo = round(running_saldo + dif_efect, 2)

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
