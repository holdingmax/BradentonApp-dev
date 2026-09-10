"""
Controles: CAJA -- sexto módulo de la sección Controles.

Cruza la hoja CAJA del Excel Cierre contra dos Mayores contables (mismo
formato de Libro Mayor exportado del sistema contable que ya usan
Cupones/Mercadería: línea "Cuenta: ... Emisor:...", línea
"Desde: ... Hasta: ...", encabezado "Asiento/Cpbte. | Fecha | Descripcion
| Detalle | Debito | Credito | Saldo"):

- Depósitos Ice Machine / Food Truck (columna S de CAJA, separados por la
  etiqueta de la columna T) contra el Mayor de Chase (cuenta 1.01.02.00 -
  CHASE BANK) -- se compara el total del mes de cada uno contra la suma de
  los depósitos (Débito) del Mayor cuya Descripción menciona la palabra
  "ICE"/"TRUCK" (pedido explícito del usuario 2026-09-10: el Mayor real a
  veces anota "(ICE MACHINE)" en el texto del depósito bancario).
- Columna K (CHASE, depósitos normales de ventas) contra el Mayor de Caja
  (cuenta 1.01.01.00 - CAJA) -- los depósitos en ese Mayor siempre tienen
  la palabra "DEPOSIT" en su Descripción (regla dada por el usuario); se
  excluyen los que además mencionan ICE/TRUCK, para no contarlos dos veces
  contra el chequeo de arriba.
- Columna M (EXPENSES CASH, gastos pagados en efectivo) contra el mismo
  Mayor de Caja -- cada celda de M tiene un comentario de Excel con el/los
  nombre(s) de a quién se le pagó, usado para buscar la fila del Mayor
  correspondiente entre las que no son depósito ni están en la lista de
  exclusión de abajo.

Filas del Mayor de Caja que no son depósito ni gasto real (dadas
explícitamente por el usuario): "LIQUIDACION CIERRE RECAUDACION",
"LIQUIDACION CIERRE RECAUDACION COMISIONES", "LIQUIDACION CIERRE LOTTERY",
"AJUSTE", "PROVISION CAJA POR LIQ PENDIENTE DE LOTTERY".

Tolerancia $0 en los 4 chequeos de totales (pedido explícito del usuario
2026-09-10: "los depositos deberian dar igual en caja que en la caja del
cierre al igual que los gastos, ice machine y food truck").

Es de solo lectura: nunca escribe nada ni genera un archivo para
descargar -- el resultado se muestra en pantalla, según lo ya decidido
para toda la sección Controles.
"""

import os
import re
from datetime import date, datetime

import pandas as pd
from openpyxl import load_workbook

from caja import (
    CAJA_COL_CHASE_DEPOSITS,
    CAJA_COL_EXPENSES_CASH,
    CAJA_COL_FOOD_ICE,
    CAJA_COL_FOOD_ICE_LABEL,
    _get_caja_sheet,
    _iter_caja_dates,
)

TOLERANCE = 0.0

_PERIOD_LINE_RE = re.compile(r"desde:\s*([\d/-]+)\s*hasta:\s*([\d/-]+)", re.IGNORECASE)
_MAYOR_DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y")

# "DEPOSIT..." es el más común, pero un depósito de terminal de tarjeta a
# veces se anota solo como "TRANSACCION #..." sin la palabra DEPOSIT --
# confirmado contra datos reales (ver validación en CLAUDE.md): sin incluir
# también "TRANSACCION" acá, esas filas se contaban como gasto en vez de
# depósito y el total de columna K no cerraba contra el Mayor de Caja.
_DEPOSIT_RE = re.compile(r"deposit|transaccion", re.IGNORECASE)
_TRUCK_RE = re.compile(r"\btruck\b", re.IGNORECASE)
_ICE_RE = re.compile(r"\bice\b", re.IGNORECASE)

# Filas del Mayor de Caja que no son ni depósito ni gasto real -- liquidaciones
# de Lottery/Recaudación y ajustes contables que pasan por la misma cuenta sin
# ser un movimiento de caja del día a día (dadas explícitamente por el usuario).
EXCLUDED_CAJA_MAYOR_PATTERNS = (
    "liquidacion cierre recaudacion",
    "liquidacion cierre lottery",
    "ajuste",
    "provision caja por liq pendiente de lottery",
)

# Cuando Ice Machine y Food Truck llegan el mismo día, el usuario tipea la
# suma a mano en la celda de S (ej. "650+1030", como fórmula real de Excel o
# como texto plano) en vez de dejar un solo número ya sumado -- confirmado
# explícitamente 2026-09-10. Food Truck siempre cobra un monto EXACTO y fijo
# ($1.030 hasta julio-2026, $1.000 desde agosto-2026 en adelante), así que
# el sumando que coincide con uno de estos valores es Food Truck y el otro
# es Ice Machine -- ver _split_combined_food_ice_amount.
KNOWN_FOOD_TRUCK_AMOUNTS = {1000.0, 1030.0}

_SUM_LITERAL_RE = re.compile(r"^\s*=?\s*(-?\d+(?:\.\d+)?(?:\s*\+\s*-?\d+(?:\.\d+)?)+)\s*$")


def _parse_mayor_date(text):
    text = text.strip()
    for fmt in _MAYOR_DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _coerce_date(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        return _parse_mayor_date(value)
    return None


def _read_mayor_rows(path):
    """
    Lee un Mayor contable genérico (.xls/.xlsx) -- misma estructura que ya
    usan Cupones/Mercadería (ver controles_cupones._read_mayor_ledger /
    controles_mercaderia._read_mercaderia_mayor). Busca la línea
    "Desde: ... Hasta: ..." y la fila de encabezado real por su propio
    contenido ("Fecha"/"Debito"/"Credito", nunca una fila fija).

    Devuelve (period_from, period_to, rows) -- rows es una lista de dicts
    {fecha, descripcion, debito, credito} por cada fila con fecha real (la
    fila "SALDO INICIAL" y la de totales al pie no tienen fecha, quedan
    afuera solas).
    """
    df = pd.read_excel(path, header=None)

    period_from = period_to = None
    header_row_index = None
    col_index = {}

    for idx in range(len(df)):
        cells = ["" if pd.isna(v) else str(v).strip() for v in df.iloc[idx]]
        if period_from is None:
            joined = " ".join(c for c in cells if c)
            match = _PERIOD_LINE_RE.search(joined)
            if match:
                period_from = _parse_mayor_date(match.group(1))
                period_to = _parse_mayor_date(match.group(2))
        if header_row_index is None:
            lowered = [c.lower() for c in cells]
            if "fecha" in lowered and "debito" in lowered and "credito" in lowered:
                header_row_index = idx
                col_index = {
                    "fecha": lowered.index("fecha"),
                    "descripcion": lowered.index("descripcion") if "descripcion" in lowered else None,
                    "debito": lowered.index("debito"),
                    "credito": lowered.index("credito"),
                }

    if header_row_index is None:
        raise ValueError('No se encontró la fila de encabezado ("Fecha"/"Debito"/"Credito") en el Mayor.')
    if period_from is None or period_to is None:
        raise ValueError('No se encontró la línea "Desde: ... Hasta: ..." en el Mayor.')

    rows = []
    for idx in range(header_row_index + 1, len(df)):
        fecha = _coerce_date(df.iat[idx, col_index["fecha"]])
        if fecha is None:
            continue
        descripcion = ""
        if col_index["descripcion"] is not None:
            raw = df.iat[idx, col_index["descripcion"]]
            descripcion = "" if pd.isna(raw) else str(raw).strip()
        debito = df.iat[idx, col_index["debito"]]
        credito = df.iat[idx, col_index["credito"]]
        rows.append(
            {
                "fecha": fecha,
                "descripcion": descripcion,
                "debito": float(debito) if pd.notna(debito) else 0.0,
                "credito": float(credito) if pd.notna(credito) else 0.0,
            }
        )

    return period_from, period_to, rows


def _is_excluded_mayor_row(descripcion):
    lowered = descripcion.lower()
    return any(pattern in lowered for pattern in EXCLUDED_CAJA_MAYOR_PATTERNS)


def _parse_sum_literal(raw_value):
    """
    Si `raw_value` (el contenido crudo de la celda, leído con
    data_only=False) es una suma tipeada a mano tipo "650+1030" -- con o
    sin el "=" de fórmula real de Excel al principio, el usuario usa las
    dos formas -- devuelve la lista de sumandos como float. Si es un solo
    número, u otra cosa, devuelve None.
    """
    if not isinstance(raw_value, str):
        return None
    match = _SUM_LITERAL_RE.match(raw_value)
    if not match:
        return None
    return [float(part.strip()) for part in match.group(1).split("+")]


def _split_combined_food_ice_amount(addends):
    """
    Dado un día donde S combina Food Truck + Ice Machine en una sola celda
    (ej. [650.0, 1030.0]), separa cuál es cuál usando que Food Truck
    siempre cobra un monto exacto conocido (KNOWN_FOOD_TRUCK_AMOUNTS).
    Devuelve (truck_amount, ice_amount), o None si no se pudo separar con
    certeza (ni uno solo de los sumandos coincide con un monto conocido de
    Food Truck, o hay más de dos sumandos).
    """
    if len(addends) != 2:
        return None
    truck_candidates = [a for a in addends if a in KNOWN_FOOD_TRUCK_AMOUNTS]
    if len(truck_candidates) != 1:
        return None
    truck_amount = truck_candidates[0]
    remaining = list(addends)
    remaining.remove(truck_amount)
    return truck_amount, remaining[0]


def _extract_vendor_lines_from_comment(comment):
    """
    Excel antepone el autor del comentario como primera línea de su propio
    texto (ej. "usuario:\nGordon Food" -- confirmado contra datos reales,
    no es un campo de metadata separado que openpyxl ya limpie solo). Se
    descarta esa primera línea cuando termina en ":" y se devuelve el
    resto como una lista de nombres -- un mismo día puede combinar más de
    un pago en una sola celda de M (visto en datos reales: "Pago Sueldo
    Lauren" / "Gordon Food Service" / "CMS" los tres en un mismo comentario).
    """
    if comment is None or not comment.text:
        return []
    lines = [line.strip() for line in comment.text.splitlines()]
    lines = [line for line in lines if line]
    if not lines:
        return []
    if lines[0].endswith(":"):
        lines = lines[1:]
    return lines


def _name_found_in_descripcion(name, descripcion):
    tokens = [t for t in re.split(r"\W+", name.lower()) if len(t) > 2]
    if not tokens:
        return False
    lowered = descripcion.lower()
    return all(token in lowered for token in tokens)


def _build_check(label, caja_value, mayor_value, unit="$"):
    caja_value = round(caja_value, 2)
    mayor_value = round(mayor_value, 2)
    diff = round(caja_value - mayor_value, 2)
    return {
        "label": label,
        "caja_value": caja_value,
        "mayor_value": mayor_value,
        "diff": diff,
        "ok": abs(diff) <= TOLERANCE,
        "unit": unit,
    }


def _format_date(value):
    if value is None:
        return None
    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y")
    return str(value)


def check_caja_mayores(cierre_path, mayor_chase_path, mayor_caja_path):
    """
    Cruza la hoja CAJA del Excel Cierre contra los Mayores contables de
    Chase (1.01.02.00) y Caja (1.01.01.00). Devuelve un dict listo para el
    template, con 4 chequeos de totales del mes (Ice Machine, Food Truck,
    columna K, columna M) y el detalle de gastos de columna M. Nunca
    escribe ni descarga nada.
    """
    cierre_path = os.path.abspath(str(cierre_path).strip())
    if not os.path.isfile(cierre_path):
        raise FileNotFoundError(f"Excel Cierre no encontrado: {cierre_path}")

    workbook = load_workbook(cierre_path, data_only=True)
    workbook_raw = load_workbook(cierre_path, data_only=False)
    try:
        sheet = _get_caja_sheet(workbook)
        sheet_raw = _get_caja_sheet(workbook_raw)
        caja_rows = list(_iter_caja_dates(sheet))

        if not caja_rows:
            raise ValueError('No se encontraron fechas en la columna A de la hoja CAJA.')

        period_year = caja_rows[0][1].year
        period_month = caja_rows[0][1].month

        k_total = 0.0
        truck_total = 0.0
        ice_total = 0.0
        ambiguous_food_ice_days = []
        m_entries = []

        for row, date_key in caja_rows:
            k_value = sheet.cell(row=row, column=CAJA_COL_CHASE_DEPOSITS).value
            if isinstance(k_value, (int, float)):
                k_total += float(k_value)

            s_value = sheet.cell(row=row, column=CAJA_COL_FOOD_ICE).value
            t_label = sheet.cell(row=row, column=CAJA_COL_FOOD_ICE_LABEL).value
            if isinstance(s_value, (int, float)) and s_value:
                label = (t_label or "").strip().lower()
                has_truck = "truck" in label
                has_ice = "ice" in label
                if has_truck and has_ice:
                    # La etiqueta ya avisa que este día combina los dos --
                    # recién acá hace falta ir a buscar el contenido crudo
                    # de la celda (fórmula/texto tipo "650+1030") para
                    # separar cuál sumando es cuál (ver KNOWN_FOOD_TRUCK_AMOUNTS).
                    s_raw = sheet_raw.cell(row=row, column=CAJA_COL_FOOD_ICE).value
                    addends = _parse_sum_literal(s_raw)
                    split = _split_combined_food_ice_amount(addends) if addends is not None else None
                    if split is not None:
                        truck_amount, ice_amount = split
                        truck_total += truck_amount
                        ice_total += ice_amount
                    else:
                        ambiguous_food_ice_days.append(date_key)
                elif has_truck:
                    # Puede ser un solo pago o la suma de más de uno del
                    # mismo tipo (ej. dos Food Truck el mismo día,
                    # "=1030+1030") -- la etiqueta ya dice que es solo Food
                    # Truck, así que el total completo va ahí sin necesidad
                    # de separar nada.
                    truck_total += float(s_value)
                elif has_ice:
                    ice_total += float(s_value)

            m_cell = sheet.cell(row=row, column=CAJA_COL_EXPENSES_CASH)
            m_value = m_cell.value
            if isinstance(m_value, (int, float)) and m_value:
                vendor_names = _extract_vendor_lines_from_comment(m_cell.comment)
                m_entries.append(
                    {"fecha": date_key, "amount": round(float(m_value), 2), "vendor_names": vendor_names}
                )
    finally:
        workbook.close()
        workbook_raw.close()

    chase_period_from, chase_period_to, chase_rows = _read_mayor_rows(mayor_chase_path)
    caja_period_from, caja_period_to, caja_mayor_rows = _read_mayor_rows(mayor_caja_path)

    period_warnings = []
    if (chase_period_from.year, chase_period_from.month) != (period_year, period_month):
        period_warnings.append(
            f"El Mayor de Chase cubre {chase_period_from:%d/%m/%Y} al {chase_period_to:%d/%m/%Y}, "
            f"distinto al mes de esta hoja CAJA ({period_month:02d}/{period_year})."
        )
    if (caja_period_from.year, caja_period_from.month) != (period_year, period_month):
        period_warnings.append(
            f"El Mayor de Caja cubre {caja_period_from:%d/%m/%Y} al {caja_period_to:%d/%m/%Y}, "
            f"distinto al mes de esta hoja CAJA ({period_month:02d}/{period_year})."
        )

    mayor_truck_total = sum(r["debito"] for r in chase_rows if _TRUCK_RE.search(r["descripcion"]))
    mayor_ice_total = sum(r["debito"] for r in chase_rows if _ICE_RE.search(r["descripcion"]))

    expense_rows = [
        r
        for r in caja_mayor_rows
        if r["credito"] > 0
        and not _DEPOSIT_RE.search(r["descripcion"])
        and not _is_excluded_mayor_row(r["descripcion"])
    ]
    mayor_k_total = sum(
        r["credito"]
        for r in caja_mayor_rows
        if _DEPOSIT_RE.search(r["descripcion"]) and not (_ICE_RE.search(r["descripcion"]) or _TRUCK_RE.search(r["descripcion"]))
    )
    mayor_expenses_total = sum(r["credito"] for r in expense_rows)

    m_total = sum(entry["amount"] for entry in m_entries)

    # Matching informativo (no bloquea el resultado del chequeo de totales,
    # que se basa en la suma agregada de arriba): para cada nombre extraído
    # del comentario de una celda M, marcar si aparece en alguna fila de
    # gasto del Mayor todavía no usada por otro nombre de este mismo lote.
    remaining_expense_rows = list(expense_rows)
    expense_detail = []
    for entry in m_entries:
        row_matches = []
        for name in entry["vendor_names"] or [None]:
            found_row = None
            if name:
                for idx, r in enumerate(remaining_expense_rows):
                    if _name_found_in_descripcion(name, r["descripcion"]):
                        found_row = remaining_expense_rows.pop(idx)
                        break
            row_matches.append({"name": name, "found": found_row is not None})
        expense_detail.append(
            {
                # Formateada a string acá mismo (no un objeto date) -- el
                # resultado entero de esta función viaja por flask.session
                # (ver webapp.py: control_caja) para que un F5 en la página
                # no deje ver de nuevo la comparación anterior, y la sesión
                # de Flask solo sabe serializar tipos JSON simples.
                "fecha": _format_date(entry["fecha"]),
                "amount": entry["amount"],
                "names": row_matches,
            }
        )

    checks = [
        _build_check("Depósitos Ice Machine", ice_total, mayor_ice_total),
        _build_check("Depósitos Food Truck", truck_total, mayor_truck_total),
        _build_check("Depósitos Chase", k_total, mayor_k_total),
        _build_check("Gastos en efectivo", m_total, mayor_expenses_total),
    ]

    return {
        "period_year": period_year,
        "period_month": period_month,
        "checks": checks,
        "expense_detail": expense_detail,
        "ambiguous_food_ice_days": [_format_date(d) for d in ambiguous_food_ice_days],
        "period_warnings": period_warnings,
        "all_ok": all(check["ok"] for check in checks) and not period_warnings,
    }
