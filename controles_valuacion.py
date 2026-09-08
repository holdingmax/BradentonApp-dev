"""
Controles: Valuación de Existencia Final -- quinto módulo de la sección Controles.

Cruza la Existencia Final de mercadería calculada "según contabilidad"
(Existencia Inicial + Compras - CMV) contra la valuación física "según
auditoría" del reporte Chevron Category Cost Report, y completa las dos
celdas del Excel BGS que hoy el usuario pega a mano cada mes.

Pautas de negocio confirmadas con el usuario (2026-09-08), a partir de la
hoja RESUMEN real del Excel BGS de julio-2026:
- **Existencia Inicial**: el valor de la fila "SALDO INICIAL" del Mayor de
  Mercadería en C-Store, antes de cualquier asiento con fecha (mismo
  helper que ya usa `controles_mercaderia._read_mercaderia_mayor`).
- **Compras**: el neto Débito - Crédito del Mayor para el período --
  exactamente el mismo `mayor_total` que ya calcula el control de
  Mercadería (confirmado numéricamente contra la fórmula real de N18).
- **CMV**: la celda L10 de la hoja RESUMEN del Excel BGS (`=+E31`, ya
  armada por el propio libro a partir de las 24 hojas de departamento) --
  nunca se recalcula desde cero, se lee su valor cacheado y se la deja
  como referencia viva en la fórmula que se escribe.
- **Valuación de auditoría**: la fila "TOTAL STORE:" del reporte Chevron
  Category Cost Report, último valor numérico de esa fila (columna
  "NET COST $" -- se ubica por posición relativa a la fila, no por
  columna fija, porque el reporte tiene dos sub-tablas con distinto
  ancho de columnas antes de esa fila).
- **Margen razonable**: 9%-13% (confirmado por el usuario, "así me dijo mi
  compañera") -- por fuera de ese rango se marca como advertencia, no dentro
  significa revisar.
- **El ajuste ad-hoc que a veces aparece en N18** (ej. "-4000" visto en el
  N18 real de julio-2026) es puntual de ese mes, no un término general de
  la fórmula -- confirmado con el usuario 2026-09-08 ("ajuste puntual de
  julio, no se repite"). Este control SIEMPRE escribe la fórmula limpia
  (EI + Compras - CMV, sin ningún ajuste extra); si un mes puntual necesita
  un ajuste manual como ese, el usuario lo agrega a mano en el Excel ya
  descargado.

A diferencia de los otros 4 controles (100% de solo lectura), este
**completa** una copia del Excel BGS (nunca el original -- siempre sobre
la copia ya guardada en el workspace temporal de la request) con:
  1. N16 (RESUMEN): la valuación de auditoría de Chevron.
  2. N18 (RESUMEN): la fórmula `=EI+Compras-L10` (limpia, sin ajustes).
  3. P{fila}: la misma valuación de auditoría, en la fila de la tabla
     histórica (una por mes calendario) cuya columna Q dice
     "EF-Valor Costo {MM}-{YYYY}" para el mes del período -- la fila ya
     existe en la plantilla (12 filas por año), no se crea ninguna.
Pedido explícito del usuario (2026-09-08): "quiero que si complete el
excel, pero que el archivo aparezca abajo del reporte para que el usuario
lo pueda descargar por su cuenta" -- el resultado se muestra en pantalla
como el resto de Controles, con un link de descarga aparte para la copia
ya completada.
"""

import os
import re
import tempfile
from datetime import datetime

from openpyxl import load_workbook

from controles_mercaderia import _format_date, _read_mercaderia_mayor

TOLERANCE_MIN = 0.09
TOLERANCE_MAX = 0.13

RESUMEN_SHEET = "RESUMEN"
CELL_CMV = "L10"
CELL_EF_CONTABILIDAD = "N18"
CELL_EF_AUDITORIA = "N16"
CELL_AUDITORIA_LABEL = "O16"
COL_TABLE_LABEL = 17  # Q
COL_TABLE_CONTABILIDAD = 15  # O
COL_TABLE_AUDITORIA = 16  # P

_LABEL_DATE_RE = re.compile(r"\d{2}/\d{2}/\d{4}")

_CHEVRON_INVENTORY_DATE_RE = re.compile(r"current inventory date:\s*([\d/-]+)", re.IGNORECASE)
_TABLE_LABEL_RE = re.compile(r"ef-valor costo\s+(\d{1,2})-(\d{4})", re.IGNORECASE)


def _create_temp_workbook_path():
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="valuacion_")
    os.close(fd)
    return temp_path


def _read_chevron_valuation(path):
    """
    Lee el Chevron Category Cost Report (.xlsx) -- de solo lectura.

    Ubica la fila "TOTAL STORE:" (columna B) y toma el último valor
    numérico de esa fila como la valuación de auditoría (columna
    "NET COST $", siempre la última con datos en esa fila puntual, sin
    depender de en qué columna exacta cae -- las dos sub-tablas del
    reporte tienen distinto ancho de columnas antes de esta fila).

    También lee "Current inventory date:" para poder cruzarlo contra el
    período del Mayor y avisar si son de meses distintos.

    Devuelve (valuacion, inventory_date).
    """
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook[workbook.sheetnames[0]]
        valuacion = None
        inventory_date = None
        for row in sheet.iter_rows():
            values = [cell.value for cell in row]
            if inventory_date is None and isinstance(values[0], str):
                match = _CHEVRON_INVENTORY_DATE_RE.search(values[0])
                if match:
                    try:
                        inventory_date = datetime.strptime(match.group(1).strip(), "%Y-%m-%d").date()
                    except ValueError:
                        inventory_date = None
            label = values[1] if len(values) > 1 else None
            if isinstance(label, str) and label.strip().lower() == "total store:":
                numeric_values = [v for v in values if isinstance(v, (int, float))]
                if numeric_values:
                    valuacion = round(float(numeric_values[-1]), 2)
    finally:
        workbook.close()

    if valuacion is None:
        raise ValueError('No se encontró la fila "TOTAL STORE:" en el reporte Chevron.')
    return valuacion, inventory_date


def _find_valuation_table_row(sheet, month, year):
    """Busca en la columna Q la fila "EF-Valor Costo {MM}-{YYYY}" del mes pedido."""
    for row in range(1, sheet.max_row + 1):
        value = sheet.cell(row=row, column=COL_TABLE_LABEL).value
        if not isinstance(value, str):
            continue
        match = _TABLE_LABEL_RE.search(value)
        if match and int(match.group(1)) == month and int(match.group(2)) == year:
            return row
    return None


def check_and_complete_valuation(mayor_path, bgs_path, chevron_path):
    """
    Calcula la Existencia Final "según contabilidad" y la cruza contra la
    valuación "según auditoría" de Chevron, y completa una copia del Excel
    BGS con ambos valores (ver docstring del módulo). Nunca toca los
    archivos originales -- siempre lee de las copias ya guardadas en el
    workspace temporal de la request, y guarda el resultado en un archivo
    nuevo aparte.
    """
    mayor_path = os.path.abspath(str(mayor_path).strip())
    if not os.path.isfile(mayor_path):
        raise FileNotFoundError(f"Mayor de Mercadería en C-Store no encontrado: {mayor_path}")

    bgs_path = os.path.abspath(str(bgs_path).strip())
    if not os.path.isfile(bgs_path):
        raise FileNotFoundError(f"Excel BGS no encontrado: {bgs_path}")
    if os.path.splitext(bgs_path)[1].lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("El Excel BGS debe ser .xlsx o .xlsm.")

    chevron_path = os.path.abspath(str(chevron_path).strip())
    if not os.path.isfile(chevron_path):
        raise FileNotFoundError(f"Chevron Category Cost Report no encontrado: {chevron_path}")

    period_from, period_to, mayor_debito, mayor_credito, _invoice_count, saldo_inicial = (
        _read_mercaderia_mayor(mayor_path)
    )
    if saldo_inicial is None:
        raise ValueError('No se encontró la fila "SALDO INICIAL" en el Mayor de Mercadería.')
    compras = round(mayor_debito - mayor_credito, 2)

    ef_auditoria, chevron_inventory_date = _read_chevron_valuation(chevron_path)
    if chevron_inventory_date is not None and (
        chevron_inventory_date.month != period_to.month or chevron_inventory_date.year != period_to.year
    ):
        raise ValueError(
            f"El Mayor es del período {_format_date(period_to)} pero el reporte Chevron "
            f"es del {_format_date(chevron_inventory_date)} -- verificá que sean del mismo mes."
        )

    workbook_cached = load_workbook(bgs_path, data_only=True)
    try:
        if RESUMEN_SHEET not in workbook_cached.sheetnames:
            raise ValueError(f'El Excel BGS no tiene una hoja "{RESUMEN_SHEET}".')
        cmv = workbook_cached[RESUMEN_SHEET][CELL_CMV].value
        if not isinstance(cmv, (int, float)):
            raise ValueError(
                f"No se pudo leer un valor numérico de CMV en {RESUMEN_SHEET}!{CELL_CMV} -- "
                "abrí el Excel BGS en Excel real y guardalo una vez para que la fórmula "
                "quede con un valor calculado."
            )
        cmv = round(float(cmv), 2)
    finally:
        workbook_cached.close()

    ef_contabilidad = round(saldo_inicial + compras - cmv, 2)
    diff_pct = round(ef_contabilidad / ef_auditoria - 1, 4) if ef_auditoria else None
    ok = diff_pct is not None and TOLERANCE_MIN <= abs(diff_pct) <= TOLERANCE_MAX

    workbook = load_workbook(bgs_path, data_only=False)
    try:
        sheet = workbook[RESUMEN_SHEET]
        target_row = _find_valuation_table_row(sheet, period_to.month, period_to.year)
        sheet[CELL_EF_AUDITORIA] = ef_auditoria
        sheet[CELL_EF_CONTABILIDAD] = f"={saldo_inicial:.2f}+{compras:.2f}-{CELL_CMV}"
        if target_row is not None:
            # Pedido explícito del usuario (2026-09-08): O{fila} se pega
            # como VALOR (ef_contabilidad ya calculado), no como fórmula
            # "=+N18" -- la fila queda congelada con el número real de
            # este mes en vez de seguir apuntando a N18, que el mes que
            # viene ya va a tener otra cuenta.
            sheet.cell(row=target_row, column=COL_TABLE_CONTABILIDAD, value=ef_contabilidad)
            sheet.cell(row=target_row, column=COL_TABLE_AUDITORIA, value=ef_auditoria)

        label_cell = sheet[CELL_AUDITORIA_LABEL]
        if isinstance(label_cell.value, str):
            new_date = period_to.strftime("%d/%m/%Y")
            label_cell.value = _LABEL_DATE_RE.sub(new_date, label_cell.value, count=1)

        # openpyxl no evalúa fórmulas -- sin esto, Excel puede seguir
        # mostrando el valor cacheado viejo de N18/O{fila}/M{fila}/N{fila}
        # (todas dependen de N18) hasta que el usuario fuerce un recálculo
        # a mano. `fullCalcOnLoad` le pide a Excel que recalcule TODO al
        # abrir el archivo, sin depender de ningún valor cacheado.
        workbook.calculation.fullCalcOnLoad = True

        temp_path = _create_temp_workbook_path()
        workbook.save(temp_path)
    finally:
        workbook.close()

    return {
        "period_from": period_from,
        "period_to": period_to,
        "period_from_display": _format_date(period_from),
        "period_to_display": _format_date(period_to),
        "saldo_inicial": saldo_inicial,
        "compras": compras,
        "cmv": cmv,
        "ef_contabilidad": ef_contabilidad,
        "ef_auditoria": ef_auditoria,
        "diff_pct": diff_pct,
        "ok": ok,
        "target_row_found": target_row is not None,
        "download_path": temp_path,
        "download_filename": os.path.basename(bgs_path),
    }
