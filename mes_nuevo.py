"""
Mes Nuevo -- automatiza preparar la carpeta del mes siguiente a partir de la
carpeta de cierre del mes actual (el mismo trabajo manual que Alfonso hacía
con la carpeta "Lienzo en blanco", validado archivo por archivo, hoja por
hoja, en la sesión del 2026-09-04 -- este módulo es esa misma lógica ya
puesta en código, no una reinvención).

Entrada: un .zip con la carpeta completa del mes que se está por cerrar (con
la misma estructura que usa siempre Alfonso -- ver CLAUDE.md, "Módulo Mes
Nuevo"). Salida: un .zip con la misma carpeta ya preparada para el mes
siguiente -- fechas corridas, hojas rotadas, archivos/carpetas renombrados.

Nunca escribe nada fuera de un directorio de trabajo temporal propio -- el
.zip que sube el usuario nunca se toca, se extrae a una copia aparte.
"""

import calendar
import os
import re
import shutil
import tempfile
import zipfile
from copy import copy
from datetime import datetime, timedelta

from contextlib import contextmanager

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.cell_range import CellRange

MESES = [
    "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

# El nombre de la carpeta raíz siempre trae "MM NombreMes - AA" (ej. "07 Julio
# - 26") -- es la única fuente que se usa para saber de qué mes es la carpeta
# subida; el resto de los archivos/hojas usan formatos de fecha distintos
# entre sí (MM-AA, MM.AAAA, nombre de mes largo, etc.) y no son confiables
# como fuente única.
_FOLDER_NAME_RE = re.compile(r"^(\d{1,2})\s+[A-Za-zÁÉÍÓÚÑáéíóúñ]+\s*-\s*(\d{2,4})")


_SUM_RANGE_RE = re.compile(r"^(=SUM\()([A-Z]+)(\d+):([A-Z]+)(\d+)(\))$", re.IGNORECASE)

# Set de referencia de "Invoice y EFT JH Williams" (el mismo que trae el
# Lienzo en blanco) -- pedido explícito del usuario (2026-09-04): esta
# carpeta tiene que terminar SIEMPRE con esta misma cantidad de archivos y
# estos mismos títulos, listos para que el usuario los pise con el PDF real
# del mes nuevo -- cualquier PDF real que se haya acumulado durante el mes
# que se cierra (más facturas de las que caben acá) se descarta, nunca se
# suma como archivo extra.
_EFT_INVOICE_REFERENCE_NAMES = {
    "invoice n° oil invoice n°1.pdf",
    "invoice n° oil invoice n°2.pdf",
    "invoice n° oil invoice n°3.pdf",
    "invoice n° oil invoice n°4.pdf",
    "invoice n° oil invoice n°5.pdf",
    "invoice nº chevron networkfee.pdf",
    "invoice nº gift gas cards.pdf",
    "invoice nº mako netw fee.pdf",
    "invoice nº vpn fee.pdf",
}

# Un mes ya vivido casi siempre trae estas mismas facturas fijas mensuales
# con el N° de factura REAL intercalado en el nombre (ej. "Invoice
# N°211258 Oil Invoice N°1.pdf") en vez del nombre de referencia limpio --
# eso NO las hace "facturas de más": son la misma categoría fija de
# siempre. Se reconocen por el texto de después del N° (el label) y se
# renombran de vuelta al nombre de referencia limpio, listas para que el
# usuario pise el contenido el mes que viene (bug real reportado por el
# usuario 2026-09-04: antes se borraban como si fueran facturas de más).
_EFT_INVOICE_NUMBERED_RE = re.compile(r"^invoice\s+n[°º]\s*\d*\s+(.+)\.pdf$", re.IGNORECASE)
_EFT_INVOICE_LABEL_TO_REFERENCE = {
    "oil invoice n°1": "Invoice N° Oil Invoice N°1.pdf",
    "oil invoice n°2": "Invoice N° Oil Invoice N°2.pdf",
    "oil invoice n°3": "Invoice N° Oil Invoice N°3.pdf",
    "oil invoice n°4": "Invoice N° Oil Invoice N°4.pdf",
    "oil invoice n°5": "Invoice N° Oil Invoice N°5.pdf",
    "chevron networkfee": "Invoice Nº Chevron Networkfee.pdf",
    "gift gas cards": "Invoice Nº Gift Gas Cards.pdf",
    "mako netw fee": "Invoice Nº Mako Netw Fee.pdf",
    "vpn fee": "Invoice Nº VPN Fee.pdf",
}
_EFT_DATE_RE = re.compile(r"\(\d{2}\.\d{2}\.\d{4}\)")


class MesNuevoError(ValueError):
    """Error esperable de datos/formato -- se muestra al usuario tal cual, sin traceback."""


@contextmanager
def _isolate_step(summary, label):
    """
    Aísla un paso del proceso -- pedido explícito del usuario (2026-09-04):
    cada mes real va a tener variaciones (un archivo que no está, una hoja
    con una estructura distinta a la esperada, lo que sea), y un problema
    puntual en UN archivo/carpeta nunca debería cancelar TODO el proceso --
    mismo criterio ya usado en cada módulo de Herramientas (una factura que
    falla no tira abajo el lote entero). Cualquier excepción acá queda
    como aviso en summary["warnings"] y el resto de los pasos sigue.
    """
    try:
        yield
    except Exception as exc:
        summary["warnings"].append(
            f"{label}: {exc} -- se salteó este paso puntual, el resto del proceso siguió."
        )


def _parse_root_folder_month(folder_name):
    match = _FOLDER_NAME_RE.match(folder_name.strip())
    if not match:
        raise MesNuevoError(
            f'No se pudo leer el mes del nombre de la carpeta "{folder_name}" -- '
            'se espera algo como "07 Julio - 26".'
        )
    month = int(match.group(1))
    year_part = match.group(2)
    year = 2000 + int(year_part) if len(year_part) == 2 else int(year_part)
    if not (1 <= month <= 12):
        raise MesNuevoError(f'Mes inválido ({month}) en el nombre de la carpeta "{folder_name}".')
    return year, month


def _next_month(year, month):
    if month == 12:
        return year + 1, 1
    return year, month + 1


def _mm_yy(year, month):
    return f"{month:02d}-{year % 100:02d}"


def _mm_yyyy(year, month):
    return f"{month:02d}.{year:04d}"


def _mm_dash_yyyy(year, month):
    return f"{month:02d}-{year:04d}"


def _mm_slash_yyyy(year, month):
    return f"{month:02d}/{year:04d}"


def _mmdd_yyyy(year, month, day):
    return f"{day:02d}.{month:02d}.{year:04d}"


def _find_workbook(root, *name_fragments_and_ext):
    """
    Busca, recursivamente bajo root, el único archivo cuyo nombre contiene
    TODOS los fragmentos dados (case-insensitive) y termina con la extensión
    indicada (último elemento). Falla limpio si no encuentra exactamente uno
    -- nunca adivina entre varios candidatos.
    """
    *fragments, ext = name_fragments_and_ext
    matches = []
    for dirpath, _dirs, files in os.walk(root):
        for name in files:
            if name.startswith("~$"):
                # Archivo de bloqueo temporal de Office (queda al lado del
                # real mientras ese Excel está abierto) -- caso real
                # encontrado 2026-09-04: "~$CHASE 08-2026.xlsx" matcheaba
                # igual que el archivo real y hacía fallar la búsqueda por
                # "encontré 2, esperaba 1".
                continue
            lower = name.lower()
            if lower.endswith(ext.lower()) and all(frag.lower() in lower for frag in fragments):
                matches.append(os.path.join(dirpath, name))
    if len(matches) == 1:
        return matches[0]
    if not matches:
        return None
    raise MesNuevoError(
        f"Encontré {len(matches)} archivos que matchean {fragments} (.{ext}) -- "
        "esperaba exactamente uno, revisar la carpeta a mano."
    )


def _rename_in_place(path, new_name):
    if path is None:
        return None
    new_path = os.path.join(os.path.dirname(path), new_name)
    if os.path.abspath(path) != os.path.abspath(new_path):
        os.rename(path, new_path)
    return new_path


def _clear_cell(cell):
    if cell.__class__.__name__ == "MergedCell":
        return
    cell.value = None


def _resize_day_block(sheet, first_row, old_days, new_days):
    """
    Ajusta un bloque de "una fila por día del mes" (CAJA, PROYECTADO, CARGA
    AQUI, Reply to Report c-Store) de old_days a new_days filas, arrancando
    en first_row.

    Si new_days > old_days: inserta filas nuevas copiando estilo/fórmulas de
    la última fila real existente, traduciendo las fórmulas (relativas y
    entre hojas) a la nueva posición con openpyxl.formula.translate --
    _reformulate_rows_below (proveedores.py) no alcanza acá porque esas
    fórmulas SÍ tienen que correrse relativas a su propia fila, no
    recalcularse contra "la fila real más cercana arriba".

    Si new_days < old_days: borra las filas sobrantes del final del bloque.

    Devuelve (new_last_row, skipped_refs) -- skipped_refs es la lista de
    coordenadas (ej. "F38") cuya fórmula referencia una fila fija por
    ENCIMA del bloque (saldo inicial, encabezado) y por eso no se tradujo
    sola; el llamador decide si avisarle al usuario que las revise a mano.

    Cualquier fila que haya DEBAJO del bloque (totales, ratios, lo que sea
    -- sin asumir cuántas ni cuáles columnas) queda físicamente corrida por
    insert_rows/delete_rows pero con el TEXTO de sus fórmulas intacto (igual
    que en cualquier otro insert/delete de openpyxl) -- se retraduce cada
    una acá mismo, como si se hubiera movido de su fila vieja a la nueva,
    que es exactamente lo que pasó.
    """
    old_last_row = first_row + old_days - 1
    new_last_row = first_row + new_days - 1
    shift = new_days - old_days
    if shift == 0:
        return new_last_row, []

    # insert_rows/delete_rows NO actualizan sheet.merged_cells -- cualquier
    # celda combinada que caiga debajo del bloque de días (ej. una etiqueta
    # de dos columnas como "Saldo Caja" en E:F) queda apuntando a la fila
    # VIEJA después del corrimiento. Si no se corrige, esa celda pasa a
    # combinarse sobre el contenido de OTRA fila -- y al guardar el libro,
    # Excel/openpyxl descartan el valor de la celda no-ancla de esa
    # combinación (bug real encontrado 2026-09-04: "=+P3" en CAJA F38
    # desaparecía después de un resize, solo se notaba al reabrir el
    # archivo guardado). Se desarma cualquier combinación que caiga debajo
    # del bloque ANTES de mover filas, y se rearma en su nueva posición
    # después.
    trailing_merges = [mcr.coord for mcr in list(sheet.merged_cells.ranges) if mcr.min_row > old_last_row]
    for coord in trailing_merges:
        sheet.unmerge_cells(coord)

    if shift > 0:
        template_row = old_last_row
        insert_at = old_last_row + 1
        sheet.insert_rows(insert_at, amount=shift)
        for offset in range(shift):
            new_row = insert_at + offset
            for col in range(1, sheet.max_column + 1):
                src_cell = sheet.cell(row=template_row, column=col)
                if src_cell.__class__.__name__ == "MergedCell":
                    continue
                dst_cell = sheet.cell(row=new_row, column=col)
                dst_cell.font = copy(src_cell.font)
                dst_cell.border = copy(src_cell.border)
                dst_cell.alignment = copy(src_cell.alignment)
                dst_cell.number_format = src_cell.number_format
                if isinstance(src_cell.value, str) and src_cell.value.startswith("="):
                    dst_cell.value = Translator(
                        src_cell.value, origin=src_cell.coordinate
                    ).translate_formula(dst_cell.coordinate)
                else:
                    dst_cell.value = src_cell.value
        trailing_start = insert_at + shift
    else:
        n_remove = -shift
        sheet.delete_rows(new_last_row + 1, amount=n_remove)
        trailing_start = new_last_row + 1

    for coord in trailing_merges:
        shifted = CellRange(coord)
        shifted.shift(row_shift=shift)
        sheet.merge_cells(start_row=shifted.min_row, start_column=shifted.min_col,
                           end_row=shifted.max_row, end_column=shifted.max_col)

    skipped_refs = []
    for row in range(trailing_start, sheet.max_row + 1):
        old_row_number = row - shift
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row, column=col)
            if cell.__class__.__name__ == "MergedCell":
                continue
            if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                continue
            sum_match = _SUM_RANGE_RE.match(cell.value)
            if sum_match and int(sum_match.group(3)) <= first_row:
                # =SUM(E4:E34) -- el límite de ARRIBA (E4) es un ancla fija
                # al inicio del bloque de días, nunca se mueve; solo el de
                # ABAJO corre con el último día real. Traducir la fórmula
                # entera acá (como se hace abajo) desplazaría los dos
                # límites por igual y perdería el ancla.
                cell.value = (
                    f"{sum_match.group(1)}{sum_match.group(2)}{sum_match.group(3)}:"
                    f"{sum_match.group(4)}{new_last_row}{sum_match.group(6)}"
                )
                continue
            # Cualquier otra referencia de fila (ej. "=+P3", un saldo inicial
            # fijo bien por ENCIMA del bloque de días) no se toca -- solo se
            # trasladan las fórmulas donde CADA fila referenciada cae en o
            # después de first_row, o sea que plausiblemente se mueve junto
            # con el bloque (ej. "=+F35+H35" apuntando a la fila de totales
            # pegada abajo). Una referencia por encima del bloque es un
            # ancla fija (saldo inicial, encabezado) -- forzarla a correr iría
            # en contra de lo que esa celda representa.
            referenced_rows = [int(r) for r in re.findall(r"\$?[A-Z]+\$?(\d+)", cell.value)]
            if referenced_rows and min(referenced_rows) < first_row:
                skipped_refs.append(cell.coordinate)
                continue
            # Nunca Translator acá a propósito: Translator jamás corre una
            # referencia absoluta ($D$34), porque en Excel eso significa
            # "fija para siempre" -- pero acá, dentro de este mismo bloque
            # que se está corriendo entero, "$D$34" casi siempre es una
            # referencia a OTRA fila del propio bloque (ej. el "Promedio"
            # de PROYECTADO dividiendo por el total un par de filas más
            # arriba) que tiene que correr junto con todo lo demás -- no un
            # ancla real (si lo fuera, ya se hubiera salteado arriba, antes
            # de llegar acá). Se corren TODAS las referencias de fila por
            # igual, tengan "$" o no (bug real reportado por el usuario
            # 2026-09-04: las fórmulas del pie de PROYECTADO -- Promedio,
            # Proyectado en la columna I -- se iban rompiendo solas en cada
            # resize).
            cell.value = re.sub(
                r"(\$?[A-Z]+\$?)(\d+)",
                lambda m, _shift=shift: f"{m.group(1)}{int(m.group(2)) + _shift}",
                cell.value,
            )

    return new_last_row, skipped_refs


def _set_day_dates(sheet, first_row, days, year, month, columns):
    """columns: lista de (col_index, day_offset) -- ej. [(1, 0), (3, 1)] para A=dia, C=dia+1."""
    for i in range(days):
        row = first_row + i
        day = i + 1
        base_date = datetime(year, month, day)
        for col, offset in columns:
            value = base_date + timedelta(days=offset) if offset else base_date
            sheet.cell(row=row, column=col, value=value)


def _prepare_proyectado(path, year, month, days, summary):
    wb = load_workbook(path, data_only=False)
    sheet = wb.worksheets[0]
    # fila 3 = dia 1 -- old_days se mide escaneando fechas reales, NUNCA con
    # sheet.max_row (ese incluye ademas el cuadro de totales/Dias
    # Mes/Promedio que sigue despues del bloque de dias, y confundir uno con
    # otro borra ese cuadro entero en vez de solo las filas de dias).
    old_days = 0
    for row in range(3, sheet.max_row + 1):
        if isinstance(sheet.cell(row=row, column=1).value, datetime):
            old_days = row - 2
        else:
            break
    sheet.title = f"{MESES[month].upper()} {year % 100:02d}"
    if days != old_days:
        new_last, skipped = _resize_day_block(sheet, 3, old_days, days)
        if skipped:
            summary.setdefault("assumptions", []).append(
                f"PROYECTADO: {len(skipped)} fórmula(s) que ya apuntaban a una fila fija de "
                "la cabecera (por encima del bloque de días) se dejaron exactamente igual a "
                "propósito -- es lo correcto, no hace falta tocar nada: " + ", ".join(skipped)
            )
        # Las 3 fórmulas "Promedio" (I13/I19/I25, ej. "=+$D$34/23") están
        # FUERA del rango que _resize_day_block ya corrige solo (viven en
        # el cuadro de la derecha, filas 11-25, no en el bloque de días ni
        # en su pie) -- pero SÍ referencian con "$" la fila de TOTAL, que
        # sí se mueve cuando cambia la cantidad de días del mes. Si no se
        # actualiza esta referencia, la fórmula termina apuntando a la
        # fila vieja (bug real reportado por el usuario 2026-09-04,
        # "fórmulas dispersas en la columna I" rompiéndose solas). El "/23"
        # (días transcurridos) NO se toca -- eso es manual, ver más abajo.
        new_totals_row = new_last + 1
        for promedio_row in (13, 19, 25):
            promedio_cell = sheet.cell(row=promedio_row, column=9)
            if isinstance(promedio_cell.value, str) and promedio_cell.value.startswith("="):
                promedio_cell.value = re.sub(
                    r"\$([A-Z]+)\$\d+", rf"$\1${new_totals_row}", promedio_cell.value
                )
    _set_day_dates(sheet, 3, days, year, month, [(1, 0)])
    for i in range(days):
        row = 3 + i
        sheet.cell(row=row, column=2, value=0)
        sheet.cell(row=row, column=3, value=0)
    # El cuadro de la derecha (Total Comb./C-Store/Total) tiene 3 celdas
    # "Dias Mes" -- esas sí son la cantidad real de días del mes, se
    # actualizan siempre. "Promedio" es OTRA cosa (bug real reportado por
    # el usuario 2026-09-04, "no debes cambiar la formula que estan en las
    # celdas"): su fórmula (ej. "=+$D$34/23") divide por la cantidad de
    # DÍAS YA TRANSCURRIDOS a la fecha de la última actualización manual
    # -- no por el total de días del mes -- y el propio archivo trae una
    # nota fija (H7: "SIEMPRE SE DEBE IR CAMBIANDO LA CANTIDAD DE DIAS
    # TRANSCURRIDOS...") aclarando que el usuario la recalcula a mano. Se
    # deja completamente intacta.
    for dias_mes_row in (11, 17, 23):
        sheet.cell(row=dias_mes_row, column=9, value=days)
    wb.save(path)
    wb.close()


def _prepare_trucks(path, year, month):
    wb = load_workbook(path, data_only=False)
    sheet = wb.active
    sheet["A1"] = f"CONTROL TRUCKS - {_mm_slash_yyyy(year, month)}"
    _clear_cell(sheet.cell(row=3, column=1))
    _clear_cell(sheet.cell(row=4, column=1))
    wb.save(path)
    wb.close()


def _keep_only_last_sheet(path, summary):
    """
    El Excel de Horas de Trabajo (BDT. HOURS...) acumula una hoja por
    semana del mes que se cierra -- pedido explícito del usuario
    (2026-09-07): al pasar al mes nuevo hay que borrarlas todas menos
    una, de preferencia la ÚLTIMA hoja del archivo (normalmente la
    semana más reciente/en curso, la que sirve de base para el mes que
    arranca). Si el archivo ya tiene una sola hoja, no hay nada que
    hacer.
    """
    wb = load_workbook(path, data_only=False)
    if len(wb.sheetnames) <= 1:
        wb.close()
        return
    keep_name = wb.sheetnames[-1]
    removed = [name for name in wb.sheetnames if name != keep_name]
    for name in removed:
        wb.remove(wb[name])
    wb.save(path)
    wb.close()
    summary["assumptions"].append(
        f"Horas Trabajo C-Store: se borraron {len(removed)} hoja(s) del mes que se cierra "
        f'("{", ".join(removed)}"), se dejó solo "{keep_name}" (la última del archivo).'
    )


def _prepare_cierre(path, year, month, days, summary):
    wb = load_workbook(path, data_only=False)

    # Cada hoja de este libro queda aislada en su propio _isolate_step --
    # pedido explícito del usuario 2026-09-04 aplicado también DENTRO de un
    # mismo archivo: un problema puntual en una hoja (ej. no encontrar la
    # fila "Ventas" en Store Info) no debe impedir que las demás hojas se
    # procesen igual, ni perder el guardado final del resto.
    with _isolate_step(summary, "CAJA"):
        caja = wb["CAJA"]
        caja["A1"] = f"BGS - {_mm_slash_yyyy(year, month)} - Cash transactions - TO REVIEW"
        old_days_caja = 0
        for row in range(4, caja.max_row + 1):
            if isinstance(caja.cell(row=row, column=1).value, datetime):
                old_days_caja = row - 3
            else:
                break
        if days != old_days_caja:
            _new_last, skipped = _resize_day_block(caja, 4, old_days_caja, days)
            if skipped:
                summary.setdefault("assumptions", []).append(
                    "CAJA: " + f"{len(skipped)} fórmula(s) que ya apuntaban a una fila fija de la "
                    "cabecera (ej. el saldo inicial) se dejaron exactamente igual a propósito -- "
                    "es lo correcto, no hace falta tocar nada: " + ", ".join(skipped)
                )
        for i in range(days):
            row = 4 + i
            day = i + 1
            new_a = datetime(year, month, day)
            caja.cell(row=row, column=1, value=new_a)
            caja.cell(row=row, column=3, value=new_a + timedelta(days=1))
            # K (Depósitos Chase), M (EXPENSES CASH), N (Cuenta Final
            # Lottery), S/T (Food Truck / Ice Machine + su etiqueta) son
            # datos del mes que se cierra -- tienen que quedar limpios para
            # que el módulo Caja los cargue de cero el mes que viene (bug
            # real reportado por el usuario 2026-09-04). K y N vuelven a su
            # default de plantilla (0, igual que trae el Lienzo en blanco);
            # M/S/T quedan en blanco (el Lienzo en blanco tampoco los trae
            # precargados).
            caja.cell(row=row, column=11, value=0)  # K
            _clear_cell(caja.cell(row=row, column=13))  # M
            caja.cell(row=row, column=14, value=0)  # N
            _clear_cell(caja.cell(row=row, column=19))  # S
            _clear_cell(caja.cell(row=row, column=20))  # T

    # ---- Store Info: la llena "Reporte Diario" día a día -- acá solo se
    # vacía el bloque de días (fechas incluidas, las pone esa otra
    # herramienta a medida que subís cada PDF) -- bug real reportado por el
    # usuario 2026-09-04, no cambiaba nada, ni las fechas.
    with _isolate_step(summary, "Store Info"):
        si = wb["Store Info"]
        totals_row_si = None
        for row in range(2, si.max_row + 1):
            if si.cell(row=row, column=3).value == "Ventas":
                totals_row_si = row
                break
        if totals_row_si is None:
            summary.setdefault("warnings", []).append(
                "Store Info: no encontré la fila 'Ventas' -- no se pudo vaciar el bloque de "
                "días, revisar a mano."
            )
        else:
            old_days_si = totals_row_si - 2
            if days != old_days_si:
                _new_last, skipped = _resize_day_block(si, 2, old_days_si, days)
                if skipped:
                    summary.setdefault("assumptions", []).append(
                        "Store Info: " + f"{len(skipped)} fórmula(s) que ya apuntaban a una fila "
                        "fija de la cabecera se dejaron exactamente igual a propósito -- es lo "
                        "correcto, no hace falta tocar nada: " + ", ".join(skipped)
                    )
            for i in range(days):
                row = 2 + i
                # Columnas de dato crudo (las llena Reporte Diario) -- A-D
                # fecha/hora, E-G volumen/ventas comb., I-Q ventas por
                # rubro, S-T cash/TC. H,R,W,Y,Z,AA,AB son fórmulas por fila
                # (se resizean solas arriba) -- nunca se tocan.
                for col in (1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 19, 20):
                    _clear_cell(si.cell(row=row, column=col))
                si.cell(row=row, column=21, value=0)  # U
                si.cell(row=row, column=22, value=0)  # V
                si.cell(row=row, column=24, value=0)  # X
            summary.setdefault("assumptions", []).append(
                "Store Info: se vació el bloque de días (incluidas las fechas) -- lo vuelve a "
                "cargar el módulo Reporte Diario, día a día, a medida que subís cada PDF."
            )

    # ---- Fisico: facturas de combustible del mes -- se vacían, el
    # inventario inicial teórico (fila 3) NO se toca (no hay forma de
    # calcular sola cuál debería ser el arrastre sin el resultado real del
    # inventario físico) -- bug real reportado por el usuario 2026-09-04.
    with _isolate_step(summary, "Fisico"):
        fisico = wb["Fisico"]
        totals_row_fisico = None
        for row in range(4, fisico.max_row + 1):
            if fisico.cell(row=row, column=3).value == "Total Compras":
                totals_row_fisico = row
                break
        if totals_row_fisico is None:
            summary.setdefault("warnings", []).append(
                "Fisico: no encontré la fila 'Total Compras' -- no se pudo vaciar, revisar a mano."
            )
        else:
            for row in range(4, totals_row_fisico):
                for col in (1, 2, 3, 4, 5):
                    _clear_cell(fisico.cell(row=row, column=col))
            summary.setdefault("assumptions", []).append(
                "Fisico: se vaciaron las filas de facturas de combustible del mes que se cierra. "
                "La fila 3 ('Inventario Inicial Teórico') NO se tocó -- avisame cómo debería "
                "arrastrarse de un mes al otro si querés que también se actualice sola."
            )

    # ---- Rotacion Pendiente / Gettel-Toyota / Pendiente Gettel-Toyota ----
    with _isolate_step(summary, "Rotación Gettel-Toyota"):
        prev_year, prev_month = (year, month - 1) if month > 1 else (year - 1, 12)

        pendiente_candidates = [n for n in wb.sheetnames if n.lower().startswith("pendiente ") and "gettel" not in n.lower()]
        draft_candidates = [n for n in wb.sheetnames if n.lower().startswith("pendiente gettel-toyota")]
        gettel_candidates = [n for n in wb.sheetnames if n.lower().startswith("gettel-toyota")]
        if len(pendiente_candidates) == 1 and len(draft_candidates) == 1 and len(gettel_candidates) == 1:
            old_pendiente_name = pendiente_candidates[0]
            gettel_draft_name = draft_candidates[0]
            gt_old_name = gettel_candidates[0]

            new_pendiente_name = f"Pendiente {_mm_yyyy(prev_year, prev_month)}"
            # El draft se reutiliza para preparar la cola del mes ACTUAL
            # (year/month) -- corrección 2026-09-04, pedido explícito del
            # usuario: el nombre tiene que ser el MISMO mes que se está
            # armando (son los días que quedan pendientes DE ESE mes), no
            # uno de más adelante -- la suposición anterior ("un mes por
            # delante") era incorrecta.
            new_draft_name = f"Pendiente Gettel-Toyota {month:02d}.{year:04d}"
            gt_new_name = f"Gettel-Toyota {_mm_yyyy(year, month)}"

            pendiente_sheet = wb[old_pendiente_name]
            draft_sheet = wb[gettel_draft_name]

            for coord in [mcr.coord for mcr in list(pendiente_sheet.merged_cells.ranges)]:
                pendiente_sheet.unmerge_cells(coord)
            for row in range(1, pendiente_sheet.max_row + 1):
                for col in range(1, pendiente_sheet.max_column + 1):
                    _clear_cell(pendiente_sheet.cell(row=row, column=col))
            for row in range(1, draft_sheet.max_row + 1):
                for col in range(1, draft_sheet.max_column + 1):
                    src_cell = draft_sheet.cell(row=row, column=col)
                    if src_cell.__class__.__name__ == "MergedCell":
                        continue
                    dst_cell = pendiente_sheet.cell(row=row, column=col)
                    dst_cell.value = src_cell.value
                    dst_cell.number_format = src_cell.number_format
            for mcr in draft_sheet.merged_cells.ranges:
                pendiente_sheet.merge_cells(range_string=str(mcr.coord))
            pendiente_sheet.title = new_pendiente_name
            pendiente_sheet["A1"] = f"CONTROL GETTEL/TOYOTA- {_mm_slash_yyyy(prev_year, prev_month)}"

            draft_sheet.title = new_draft_name
            draft_sheet["A1"] = f"CONTROL GETTEL/TOYOTA- {_mm_slash_yyyy(year, month)}"
            last_row_with_data = 0
            for row in range(3, draft_sheet.max_row + 1):
                v = draft_sheet.cell(row=row, column=1).value
                if v not in (None, "", "VERIFICAR"):
                    last_row_with_data = row
            for row in range(3, last_row_with_data + 1):
                for col in range(1, 10):
                    _clear_cell(draft_sheet.cell(row=row, column=col))
            summary.setdefault("assumptions", []).append(
                f'"{gettel_draft_name}" -> "{new_draft_name}": filas de datos vaciadas '
                "(queda lista para cargar la cola del mes a mano)."
            )

            gt = wb[gt_old_name]
            gt.title = gt_new_name
            # El titulo real trae ademas una aclaracion fija ("LOCAL ACCOUNT DEL
            # ARCHIVO DE CIERRE") que no es una fecha -- se mantiene igual, solo
            # cambia la fecha.
            gt["A1"] = f"CONTROL GETTEL/TOYOTA - {_mm_slash_yyyy(year, month)} (LOCAL ACCOUNT DEL ARCHIVO DE CIERRE)"
            # El bloque de dias se resize igual que CAJA/PROYECTADO (nunca un
            # numero fijo de filas -- un "18" hardcodeado acá era en realidad
            # cuanto habia llegado a llenarse a mano la muestra de referencia
            # en el momento en que se armó, no una regla real; con eso, los
            # dias 19+ del mes viejo quedaban sin tocar -- bug real reportado
            # por el usuario 2026-09-04, incluida la fecha que "volvia" a
            # agosto en esas filas nunca actualizadas).
            old_days_gt = 0
            for row in range(3, gt.max_row + 1):
                if isinstance(gt.cell(row=row, column=1).value, datetime):
                    old_days_gt = row - 2
                else:
                    break
            if days != old_days_gt:
                _new_last, skipped = _resize_day_block(gt, 3, old_days_gt, days)
                if skipped:
                    summary.setdefault("assumptions", []).append(
                        "Gettel-Toyota: " + f"{len(skipped)} fórmula(s) que ya apuntaban a una fila "
                        "fija de la cabecera se dejaron exactamente igual a propósito -- es lo "
                        "correcto, no hace falta tocar nada: " + ", ".join(skipped)
                    )
            for i in range(days):
                row = 3 + i
                day = i + 1
                new_a = datetime(year, month, day)
                gt.cell(row=row, column=1, value=new_a)
                gt.cell(row=row, column=2, value=new_a + timedelta(days=1))
                # Solo se limpian los datos crudos (Gettel/Gallon/Toyota/Gallon)
                # -- la columna I (DIF) es una fórmula calculada (=+C-E-G) que
                # tiene que seguir viva, nunca se borra a mano.
                for col in (5, 6, 7, 8):
                    _clear_cell(gt.cell(row=row, column=col))

            rename_map = {gt_old_name: gt_new_name, old_pendiente_name: new_pendiente_name,
                          gettel_draft_name: new_draft_name}
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.__class__.__name__ == "MergedCell":
                            continue
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            new_formula = cell.value
                            for old_name, new_name in rename_map.items():
                                if f"'{old_name}'" in new_formula:
                                    new_formula = new_formula.replace(f"'{old_name}'", f"'{new_name}'")
                            if new_formula != cell.value:
                                cell.value = new_formula
        else:
            summary.setdefault("warnings", []).append(
                "No encontré exactamente una hoja Pendiente / Gettel-Toyota / Pendiente "
                "Gettel-Toyota -- esa rotación no se aplicó, revisar a mano."
            )

    with _isolate_step(summary, "PAGO Cupones"):
        if "PAGO Cupones" in wb.sheetnames:
            pc = wb["PAGO Cupones"]
            pc["A1"] = f"CONTROL - {_mm_slash_yyyy(year, month)} (CUPONES QUE ENVIA RICK A FINAL DE MES)"
            # Filas de datos reales: columna C (Transc N°) es un numero, o esta
            # vacia (continuacion). Para en la primera fila con texto ahi (la
            # fila de totales "COBRADO") -- nunca asume un largo fijo de filas.
            row = 4
            while row <= pc.max_row:
                c_val = pc.cell(row=row, column=3).value
                if c_val is not None and not isinstance(c_val, (int, float)):
                    break
                for col in range(1, 6):
                    _clear_cell(pc.cell(row=row, column=col))
                row += 1

    wb.save(path)
    wb.close()


def _prepare_bgs(path, year, month):
    wb = load_workbook(path, data_only=False)
    if "RESUMEN" in wb.sheetnames:
        resumen = wb["RESUMEN"]
        last_day = calendar.monthrange(year, month)[1]
        resumen["A2"] = f"P & L Report 01 - {last_day} {MESES[month]} {year}"
        resumen["B3"] = datetime(year, month, 1)
    updated = 0
    for name in wb.sheetnames:
        if name in ("RESUMEN", "CONTROL", "COSTO.TODOS"):
            continue
        sheet = wb[name]
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.__class__.__name__ == "MergedCell":
                    continue
                if isinstance(cell.value, str):
                    match = re.match(r"^(TOTAL )(\d{2}-\d{4})$", cell.value.strip(), re.IGNORECASE)
                    if match:
                        cell.value = f"{match.group(1)}{month:02d}-{year:04d}"
                        updated += 1
    wb.save(path)
    wb.close()
    return updated


def _prepare_ventas(path, year, month, days, summary):
    wb = load_workbook(path, data_only=False)
    if "CARGA AQUI" not in wb.sheetnames:
        raise MesNuevoError('El archivo de Ventas no tiene una hoja "CARGA AQUI".')
    ca = wb["CARGA AQUI"]
    ca.cell(row=1, column=2, value=datetime(year, month, 1))
    old_days = 0
    for row in range(5, ca.max_row + 1):
        if isinstance(ca.cell(row=row, column=1).value, datetime):
            old_days = row - 4
        else:
            break
    if days != old_days:
        _new_last, skipped = _resize_day_block(ca, 5, old_days, days)
        if skipped:
            summary.setdefault("assumptions", []).append(
                "CARGA AQUI: " + f"{len(skipped)} fórmula(s) que ya apuntaban a una fila fija "
                "de la cabecera se dejaron exactamente igual a propósito -- es lo correcto, "
                "no hace falta tocar nada: " + ", ".join(skipped)
            )
    bp_col = column_index_from_string("BP")
    bs_col = column_index_from_string("BS")
    dept_first_col = column_index_from_string("C")
    dept_last_col = column_index_from_string("BN")
    for i in range(days):
        row = 5 + i
        day = i + 1
        new_a = datetime(year, month, day)
        ca.cell(row=row, column=1, value=new_a)
        ca.cell(row=row, column=2, value=new_a + timedelta(days=1))
        # C-BN son los pares COUNT/NET SALES por departamento -- los llena
        # Reporte Diario día a día, igual que Store Info; el Lienzo en
        # blanco los trae en 0 (nunca en blanco), así que se resetea a eso,
        # no a None (bug real encontrado en este análisis 2026-09-04, sin
        # reporte previo del usuario -- CARGA AQUI tenía exactamente el
        # mismo problema que Store Info: no se vaciaba nada, ni siquiera
        # las fechas). Y/Z ("VARIOS/BOLSA") son la excepción -- son
        # fórmulas que suman el resto de las columnas de esta misma fila,
        # no un departamento real con dato propio -- nunca se pisan.
        y_col = column_index_from_string("Y")
        z_col = column_index_from_string("Z")
        for col in range(dept_first_col, dept_last_col + 1):
            if col in (y_col, z_col):
                continue
            ca.cell(row=row, column=col, value=0)
        _clear_cell(ca.cell(row=row, column=bp_col))
        _clear_cell(ca.cell(row=row, column=bs_col))
    summary.setdefault("assumptions", []).append(
        "CARGA AQUI: se resetearon a 0 las columnas de conteo/venta por departamento (C a BN), "
        "y se vaciaron BP y BS (los totales que cargás vos a mano para chequear contra el POS) "
        "-- los vuelve a cargar Reporte Diario / los cargás vos de nuevo."
    )

    if "Resumen Venta" in wb.sheetnames:
        wb["Resumen Venta"].cell(row=1, column=1, value=f"C-STORE - BRADENTON - VENTAS {_mm_slash_yyyy(year, month)}")

    if "Reply to Report c-Store" in wb.sheetnames:
        rr = wb["Reply to Report c-Store"]
        rr.cell(row=1, column=2, value=datetime(year, month, 1))
        old_days_rr = 0
        for row in range(5, rr.max_row + 1):
            if isinstance(rr.cell(row=row, column=1).value, datetime):
                old_days_rr = row - 4
            else:
                break
        if days != old_days_rr:
            _new_last, skipped = _resize_day_block(rr, 5, old_days_rr, days)
            if skipped:
                summary.setdefault("assumptions", []).append(
                    "Reply to Report c-Store: " + f"{len(skipped)} fórmula(s) que ya apuntaban "
                    "a una fila fija de la cabecera se dejaron exactamente igual a propósito "
                    "-- es lo correcto, no hace falta tocar nada: " + ", ".join(skipped)
                )
        for i in range(days):
            row = 5 + i
            day = i + 1
            new_a = datetime(year, month, day)
            rr.cell(row=row, column=1, value=new_a)
            rr.cell(row=row, column=2, value=new_a + timedelta(days=1))

    wb.save(path)
    wb.close()


def prepare_next_month(upload_zip_path):
    """
    Función pública: recibe la ruta de un .zip subido por el usuario (la
    carpeta de cierre del mes actual) y devuelve (output_zip_path, summary).

    summary = {
        "source_period": "Julio 2026", "target_period": "Agosto 2026",
        "warnings": [...],    # archivos/hojas que no se pudieron procesar
        "assumptions": [...], # decisiones que conviene que el usuario revise
    }
    """
    summary = {"warnings": [], "assumptions": []}

    workdir = tempfile.mkdtemp(prefix="mes_nuevo_")
    extract_dir = os.path.join(workdir, "extract")
    os.makedirs(extract_dir, exist_ok=True)
    with zipfile.ZipFile(upload_zip_path) as zf:
        zf.extractall(extract_dir)

    entries = [e for e in os.listdir(extract_dir) if not e.startswith("__MACOSX")]
    if len(entries) == 1 and os.path.isdir(os.path.join(extract_dir, entries[0])):
        root = os.path.join(extract_dir, entries[0])
        root_name = entries[0]
    else:
        root = extract_dir
        root_name = os.path.basename(upload_zip_path).rsplit(".", 1)[0]

    year, month = _parse_root_folder_month(root_name)
    days = calendar.monthrange(year, month)[1]
    next_year, next_month = _next_month(year, month)
    next_days = calendar.monthrange(next_year, next_month)[1]
    summary["source_period"] = f"{MESES[month]} {year}"
    summary["target_period"] = f"{MESES[next_month]} {next_year}"

    def maybe_rename_top(fragments_ext, new_name):
        p = _find_workbook(root, *fragments_ext)
        if p is None:
            summary["warnings"].append(f"No encontré el archivo esperado ({' '.join(fragments_ext[:-1])}).")
            return None
        return _rename_in_place(p, new_name)

    # ---- Archivos de solo-renombrar (sin cambios de contenido) ----
    # Cada paso queda aislado con _isolate_step (pedido explícito del
    # usuario 2026-09-04): cada mes real trae variaciones -- un archivo que
    # no está, una hoja con una estructura distinta a la esperada -- y un
    # problema puntual en UNA parte no debe cancelar TODO el proceso, mismo
    # criterio ya usado en cada módulo de Herramientas (una factura que
    # falla no tira abajo el lote entero).
    chase_path = None
    tc_eft_final = None
    bce_path = None

    with _isolate_step(summary, "CHASE"):
        chase_path = maybe_rename_top(("chase", "xlsx"), f"CHASE {_mm_dash_yyyy(next_year, next_month)}.xlsx")

    with _isolate_step(summary, "Aplicacion TC y EFT"):
        tc_eft = _find_workbook(root, "aplicacion", "tc", "eft", "xlsx")
        if tc_eft:
            tc_eft_final = tc_eft
            m = re.match(r"^(.*\bal\s+)\(?\d{2}\.\d{2}\.\d{4}\)?(.*)$", os.path.basename(tc_eft))
            if m:
                new_name = f"{m.group(1)}{_mmdd_yyyy(next_year, next_month, calendar.monthrange(next_year, next_month)[1])}{m.group(2)}"
                tc_eft_final = _rename_in_place(tc_eft, new_name)
            else:
                summary["warnings"].append("No pude leer el patrón de fecha de 'Aplicacion TC y EFT' -- revisar el nombre a mano.")

    with _isolate_step(summary, "Bce Brandenton"):
        bce = _find_workbook(root, "bce", "xlsx")
        if bce:
            last_day = calendar.monthrange(next_year, next_month)[1]
            bce_path = _rename_in_place(bce, f"Bce Brandenton {last_day:02d}-{next_month:02d}-{next_year % 100:02d}.xlsx")
            summary["assumptions"].append("Bce Brandenton: solo se renombró el archivo, el contenido no se toca (pendiente de definir).")

    with _isolate_step(summary, "Horas Trabajo C-Store"):
        horas = _find_workbook(root, "hours", "xlsx")
        if horas:
            horas = _rename_in_place(horas, f"BDT. HOURS {_mm_yyyy(next_year, next_month)}.xlsx")
            # Pedido explícito del usuario 2026-09-07: este Excel acumula
            # una hoja por semana del mes que se cierra -- hay que borrarlas
            # todas menos una (de preferencia la última del archivo) antes
            # de pasarlo al mes nuevo, para que no arrastre el historial.
            _keep_only_last_sheet(horas, summary)
        else:
            horas_legacy = _find_workbook(root, "hours", "xls")
            if horas_legacy:
                horas = horas_legacy
                summary["warnings"].append(
                    "El Excel de Horas de Trabajo (BDT. HOURS...) sigue en formato .xls viejo -- "
                    "abrilo en Excel y guardalo como .xlsx antes de subir el .zip, no lo pude "
                    "renombrar ni borrarle las hojas del mes anterior en este formato."
                )
        # La carpeta solo debe tener ese Excel -- el Lienzo en blanco no
        # trae ningún PDF ahí. Un mes ya vivido acumula un PDF de pago
        # semanal por empleado por semana; esos son del mes que se cierra
        # y no deben pasar al mes nuevo (bug real reportado por el usuario
        # 2026-09-04).
        horas_dir = os.path.dirname(horas) if horas else None
        if horas_dir is None:
            for dirpath, dirs, _files in os.walk(root):
                for d in dirs:
                    if "horas" in d.lower() and "trabajo" in d.lower():
                        horas_dir = os.path.join(dirpath, d)
                        break
                if horas_dir:
                    break
        if horas_dir:
            removed_horas = 0
            for name in os.listdir(horas_dir):
                full = os.path.join(horas_dir, name)
                if not os.path.isfile(full):
                    continue
                if name.lower() == "desktop.ini" or name.startswith("~$") or name.lower().endswith((".xls", ".xlsx")):
                    continue
                os.remove(full)
                removed_horas += 1
            if removed_horas:
                summary["assumptions"].append(
                    f"Horas Trabajo C-Store: se borraron {removed_horas} PDF(s) del mes que se "
                    "cierra (esa carpeta solo debe tener el Excel de horas)."
                )

    # ---- PROYECTADO ----
    new_p = None
    with _isolate_step(summary, "PROYECTADO"):
        proyectado = _find_workbook(root, "proyectado", "xlsx")
        if proyectado:
            new_p = _rename_in_place(proyectado, f"PROYECTADO BRADENTON {_mm_yy(next_year, next_month)}.xlsx")
            _prepare_proyectado(new_p, next_year, next_month, next_days, summary)
        else:
            summary["warnings"].append("No encontré el archivo PROYECTADO.")

    # ---- Trucks ----
    with _isolate_step(summary, "Trucks"):
        trucks = _find_workbook(root, "trucks", "xlsx")
        if trucks:
            new_t = _rename_in_place(trucks, f"Trucks-{MESES[next_month]}.xlsx")
            _prepare_trucks(new_t, next_year, next_month)
            # La carpeta solo debe tener ese Excel -- el Lienzo en blanco no
            # trae ningún PDF ahí. Un mes ya vivido acumula un comprobante
            # real por transacción de Food Truck; esos son del mes que se
            # cierra y no deben pasar al mes nuevo (bug real reportado por
            # el usuario 2026-09-04).
            trucks_dir = os.path.dirname(new_t)
            removed_trucks = 0
            for name in os.listdir(trucks_dir):
                full = os.path.join(trucks_dir, name)
                if not os.path.isfile(full):
                    continue
                if name.lower() == "desktop.ini" or name.startswith("~$") or name.lower().endswith(".xlsx"):
                    continue
                os.remove(full)
                removed_trucks += 1
            if removed_trucks:
                summary["assumptions"].append(
                    f"Trucks: se borraron {removed_trucks} comprobante(s) del mes que se cierra "
                    "(esa carpeta solo debe tener el Excel de Trucks)."
                )

    # ---- Carpeta Gettel-Toyota (dentro de Controles Ice, Trucks, Gettel) ----
    # Encontrado en este mismo análisis 2026-09-04 (mismo tipo de bug que
    # Trucks, sin reporte previo del usuario): esta carpeta solo debe tener
    # "gettel hoja de Rick.xlsx" y las subcarpetas Gettel/Toyota vacías --
    # un mes ya vivido acumula ahí los reportes de Gettel/Toyota y los
    # comprobantes de pago (Pagos1, Pagos2...) del mes que se cierra.
    with _isolate_step(summary, "Carpeta Gettel-Toyota"):
        gt_folder = None
        for dirpath, dirs, _files in os.walk(root):
            for d in dirs:
                if d.lower() == "gettel-toyota":
                    gt_folder = os.path.join(dirpath, d)
                    break
            if gt_folder:
                break
        if gt_folder:
            removed_gt_folder = 0
            for name in os.listdir(gt_folder):
                full = os.path.join(gt_folder, name)
                if os.path.isfile(full):
                    if name.lower() == "desktop.ini" or name.startswith("~$") or name.lower().endswith(".xlsx"):
                        continue
                    os.remove(full)
                    removed_gt_folder += 1
                elif os.path.isdir(full) and name.lower() in ("gettel", "toyota"):
                    for sub_name in os.listdir(full):
                        sub_full = os.path.join(full, sub_name)
                        if not os.path.isfile(sub_full):
                            continue
                        if sub_name.lower() == "desktop.ini" or sub_name.startswith("~$"):
                            continue
                        os.remove(sub_full)
                        removed_gt_folder += 1
            if removed_gt_folder:
                summary["assumptions"].append(
                    f"Carpeta Gettel-Toyota: se borraron {removed_gt_folder} reporte(s)/comprobante(s) "
                    "del mes que se cierra (esa carpeta solo debe tener el Excel de Rick, las "
                    "subcarpetas Gettel/Toyota quedan vacías)."
                )

    # ---- Cierre ----
    new_c = None
    with _isolate_step(summary, "Cierre"):
        cierre = _find_workbook(root, "cierre", "xlsx")
        if cierre:
            last_day = calendar.monthrange(next_year, next_month)[1]
            new_c = _rename_in_place(cierre, f"Cierre {next_month:02d}-{next_year % 100:02d}.xlsx")
            _prepare_cierre(new_c, next_year, next_month, next_days, summary)
        else:
            summary["warnings"].append("No encontré el archivo Cierre -- se omitió toda esa parte.")

    # ---- Archivos sueltos en la raíz ----
    # La raíz del Lienzo en blanco SOLO tiene los 5 Excel de arriba (CHASE,
    # Aplicacion TC y EFT, Bce Brandenton, PROYECTADO, Cierre) -- nada más.
    # Un mes ya vivido acumula ahí archivos reales de más (ej. un "Arqueo"
    # de caja) que no deberían pasar al mes nuevo -- bug real reportado por
    # el usuario 2026-09-04 ("Arqueo Rick.pdf" quedaba colgado en la raíz).
    with _isolate_step(summary, "Archivos sueltos en la raíz"):
        expected_root_names = {
            os.path.basename(p) for p in (chase_path, tc_eft_final, bce_path, new_p, new_c) if p
        }
        removed_root = 0
        for name in os.listdir(root):
            full = os.path.join(root, name)
            if not os.path.isfile(full):
                continue
            if name.lower() == "desktop.ini" or name.startswith("~$"):
                continue
            if name in expected_root_names:
                continue
            os.remove(full)
            removed_root += 1
        if removed_root:
            summary["assumptions"].append(
                f"Raíz de la carpeta: se borraron {removed_root} archivo(s) reales de más "
                "(ej. un Arqueo de caja) que no forman parte de los 5 Excel fijos del mes."
            )

    # ---- Carpeta Stock ----
    stock_dir = None
    with _isolate_step(summary, "Carpeta Stock"):
        found_stock = None
        for dirpath, dirs, _files in os.walk(root):
            for d in dirs:
                if d.lower().startswith("stock"):
                    found_stock = os.path.join(dirpath, d)
                    break
            if found_stock:
                break
        if found_stock:
            new_stock_dir = os.path.join(os.path.dirname(found_stock), f"Stock {next_month:02d}-{next_year % 100:02d}")
            if os.path.abspath(found_stock) != os.path.abspath(new_stock_dir):
                os.rename(found_stock, new_stock_dir)
            stock_dir = new_stock_dir
        else:
            summary["warnings"].append("No encontré ninguna carpeta Stock.")

    if stock_dir:
        with _isolate_step(summary, "BGS. CMV"):
            bgs = _find_workbook(stock_dir, "bgs", "cmv", "xlsx")
            if bgs:
                last_day = calendar.monthrange(next_year, next_month)[1]
                new_bgs = _rename_in_place(
                    bgs,
                    f"BGS. CMV {_mmdd_yyyy(next_year, next_month, 1)} al {_mmdd_yyyy(next_year, next_month, last_day)} TODO EL MES.xlsx",
                )
                _prepare_bgs(new_bgs, next_year, next_month)
            else:
                summary["warnings"].append("No encontré el archivo BGS. CMV dentro de Stock.")

        with _isolate_step(summary, "Ventas ANALISIS"):
            ventas = _find_workbook(stock_dir, "ventas", "analisis", "xlsx")
            if ventas:
                new_v = _rename_in_place(ventas, f"Bradenton. Analisis C-Store. Ventas {_mm_yyyy(next_year, next_month)} ANALISIS.xlsx")
                _prepare_ventas(new_v, next_year, next_month, next_days, summary)
            else:
                ventas_xls = _find_workbook(stock_dir, "ventas", "analisis", "xls")
                if ventas_xls:
                    summary["warnings"].append(
                        "El archivo de Ventas ANALISIS sigue en formato .xls viejo -- "
                        "abrilo en Excel y guardalo como .xlsx antes de subir el .zip, "
                        "no lo pude editar en este formato."
                    )
                else:
                    summary["warnings"].append("No encontré el archivo de Ventas ANALISIS dentro de Stock.")

    # ---- Carpeta Invoice y EFT JH Williams: dejar solo el set de referencia ----
    with _isolate_step(summary, "Invoice y EFT JH Williams"):
        eft_dir = None
        for dirpath, dirs, _files in os.walk(root):
            for d in dirs:
                if "invoice" in d.lower() and "eft" in d.lower():
                    eft_dir = os.path.join(dirpath, d)
                    break
            if eft_dir:
                break
        if eft_dir:
            removed_extra = 0
            eft_candidates = []
            numbered_candidates = {}
            for name in sorted(os.listdir(eft_dir)):
                if name.lower() == "desktop.ini" or name.startswith("~$"):
                    continue  # se limpian aparte al final, no son "facturas de más"
                lower = name.lower()
                if lower in _EFT_INVOICE_REFERENCE_NAMES:
                    continue  # factura de referencia fija -- se deja tal cual, sin fecha que actualizar
                if "eft" in lower:
                    eft_candidates.append(name)  # categoria EFT -- se colapsa a una sola mas abajo
                    continue
                numbered_match = _EFT_INVOICE_NUMBERED_RE.match(name)
                label = numbered_match.group(1).lower() if numbered_match else None
                if label in _EFT_INVOICE_LABEL_TO_REFERENCE:
                    numbered_candidates.setdefault(label, []).append(name)
                    continue
                # Ni la referencia fija, ni EFT, ni una factura numerada
                # conocida -- es justo lo que no debería quedar (ej. una
                # factura de un tipo nuevo, sin categoría fija) -- se borra,
                # igual que en Depósitos.
                os.remove(os.path.join(eft_dir, name))
                removed_extra += 1

            if eft_candidates:
                keep = eft_candidates[0]
                for extra in eft_candidates[1:]:
                    os.remove(os.path.join(eft_dir, extra))
                    removed_extra += 1
                new_name = _EFT_DATE_RE.sub(f"(00.{next_month:02d}.{next_year:04d})", keep)
                if new_name != keep:
                    os.rename(os.path.join(eft_dir, keep), os.path.join(eft_dir, new_name))

            for label, names in numbered_candidates.items():
                keep = names[0]
                for extra in names[1:]:
                    os.remove(os.path.join(eft_dir, extra))
                    removed_extra += 1
                reference_name = _EFT_INVOICE_LABEL_TO_REFERENCE[label]
                target_path = os.path.join(eft_dir, reference_name)
                if os.path.exists(target_path):
                    # La versión de referencia (sin número) ya estaba
                    # presente en el mismo lote -- esta ya no hace falta.
                    os.remove(os.path.join(eft_dir, keep))
                    removed_extra += 1
                else:
                    os.rename(os.path.join(eft_dir, keep), target_path)

            if removed_extra:
                summary["assumptions"].append(
                    f"Invoice y EFT JH Williams: se borraron {removed_extra} PDF(s) reales de más "
                    "(quedó el mismo set de referencia que trae el Lienzo en blanco, listo para "
                    "pisar)."
                )
        else:
            summary["warnings"].append("No encontré la carpeta Invoice y EFT JH Williams.")

    # ---- Carpeta Depositos: renombrar carpeta + PDFs (dia->00, mes/anio actualizado) ----
    with _isolate_step(summary, "Depósitos"):
        dep_dir = None
        for dirpath, dirs, _files in os.walk(root):
            for d in dirs:
                if d.lower().startswith("deposito"):
                    dep_dir = os.path.join(dirpath, d)
                    break
            if dep_dir:
                break
        if dep_dir:
            new_dep_dir = os.path.join(os.path.dirname(dep_dir), f"Depositos {MESES[next_month]}")
            if os.path.abspath(dep_dir) != os.path.abspath(new_dep_dir):
                os.rename(dep_dir, new_dep_dir)
            dep_dir = new_dep_dir
            date_suffix_re = re.compile(r"\(\d{2}-\d{2}-\d{4}\)\.pdf$", re.IGNORECASE)
            replacement = f"00-{next_month:02d}-{next_year:04d}.pdf"
            # La carpeta tiene que quedar igual que en el "Lienzo en blanco":
            # UN solo comprobante de referencia por categoría (normal, Food
            # Truck, Ice Machine), con el día en "00". Un mes ya vivido
            # acumula un PDF real por cada depósito que pasó de verdad --
            # esos son justo los que no deberían quedar (pedido explícito del
            # usuario 2026-09-04): se borran todos menos uno por categoría, y
            # ese uno se renombra al formato de referencia.
            food_truck_re = re.compile(r"\(food truck\)", re.IGNORECASE)
            ice_machine_re = re.compile(r"\(ice machine\)", re.IGNORECASE)
            groups = {"food_truck": [], "ice_machine": [], "regular": []}
            for name in sorted(os.listdir(dep_dir)):
                if not date_suffix_re.search(name):
                    continue
                if food_truck_re.search(name):
                    groups["food_truck"].append(name)
                elif ice_machine_re.search(name):
                    groups["ice_machine"].append(name)
                else:
                    groups["regular"].append(name)

            removed = 0
            for names in groups.values():
                if not names:
                    continue
                keep, extras = names[0], names[1:]
                for extra in extras:
                    os.remove(os.path.join(dep_dir, extra))
                    removed += 1
                new_name = date_suffix_re.sub(replacement, keep)
                if new_name != keep:
                    os.rename(os.path.join(dep_dir, keep), os.path.join(dep_dir, new_name))
            if removed:
                summary["assumptions"].append(
                    f"Depósitos: se borraron {removed} comprobante(s) real(es) de más (quedó uno "
                    "solo por tipo, como en el Lienzo en blanco)."
                )

    # ---- Carpeta Reportes Diarios: queda vacía, como en el Lienzo en blanco ----
    # Un mes ya vivido acumula ahí los 28-31 "Close Store" del mes + el
    # Resumen Ventas; esos no deben pasar al mes nuevo (bug real reportado
    # por el usuario 2026-09-04).
    with _isolate_step(summary, "Reportes Diarios"):
        reportes_dir = None
        for dirpath, dirs, _files in os.walk(root):
            for d in dirs:
                if "reportes" in d.lower() and "diario" in d.lower():
                    reportes_dir = os.path.join(dirpath, d)
                    break
            if reportes_dir:
                break
        if reportes_dir:
            removed_reportes = 0
            for name in os.listdir(reportes_dir):
                full = os.path.join(reportes_dir, name)
                if not os.path.isfile(full):
                    continue
                if name.lower() == "desktop.ini" or name.startswith("~$"):
                    continue
                os.remove(full)
                removed_reportes += 1
            if removed_reportes:
                summary["assumptions"].append(
                    f"Reportes Diarios: se borraron {removed_reportes} PDF(s) del mes que se "
                    "cierra (esa carpeta queda vacía, como en el Lienzo en blanco)."
                )

    # ---- Carpeta Gastos del Mes: queda vacía, como en el Lienzo en blanco ----
    # Un mes ya vivido acumula ahí los comprobantes de gastos reales de ese
    # mes; esos no deben pasar al mes nuevo (bug real reportado por el
    # usuario 2026-09-04, misma carpeta seguía con PDFs de agosto).
    with _isolate_step(summary, "Gastos del Mes"):
        gastos_dir = None
        for dirpath, dirs, _files in os.walk(root):
            for d in dirs:
                if "gastos" in d.lower() and "mes" in d.lower():
                    gastos_dir = os.path.join(dirpath, d)
                    break
            if gastos_dir:
                break
        if gastos_dir:
            removed_gastos = 0
            for name in os.listdir(gastos_dir):
                full = os.path.join(gastos_dir, name)
                if not os.path.isfile(full):
                    continue
                if name.lower() == "desktop.ini" or name.startswith("~$"):
                    continue
                os.remove(full)
                removed_gastos += 1
            if removed_gastos:
                summary["assumptions"].append(
                    f"Gastos del Mes: se borraron {removed_gastos} comprobante(s) del mes que se "
                    "cierra (esa carpeta queda vacía, como en el Lienzo en blanco)."
                )

    # ---- Carpeta Controles Ice, Trucks, Gettel/Trucks: ya se resolvio arriba (rename generico) ----
    # Ice, Gettel-Toyota (Rick), Reportes Diarios, FPL: sin cambios.

    # ---- Comprimir resultado ----
    output_root_name = f"{next_month:02d} {MESES[next_month]} - {next_year % 100:02d}"
    final_root = os.path.join(workdir, output_root_name)
    os.rename(root, final_root)

    # limpiar desktop.ini y archivos de bloqueo de Office (~$...) heredados
    for dirpath, _dirs, files in os.walk(final_root):
        for name in files:
            if name.lower() == "desktop.ini" or name.startswith("~$"):
                os.remove(os.path.join(dirpath, name))

    # El contenido va a la RAÍZ del zip, sin envolverlo en una carpeta
    # "output_root_name/" -- si se envuelve, "Extraer todo..." de Windows
    # (que siempre crea su propia carpeta con el nombre del zip) termina
    # duplicando el nivel: "Descargas\09 Septiembre - 26\09 Septiembre -
    # 26\..." (bug real reportado por el usuario 2026-09-04, confirmado
    # contra su carpeta de Descargas real). Sin el envoltorio, extraer el
    # zip deja el contenido en un solo nivel, listo para mover tal cual.
    output_zip_path = os.path.join(workdir, f"{output_root_name}.zip")
    with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for dirpath, _dirs, files in os.walk(final_root):
            rel = os.path.relpath(dirpath, final_root)
            arcdir = "" if rel == "." else rel.replace(os.sep, "/")
            if not files:
                # Un zip no guarda carpetas vacías solas -- sin esto,
                # carpetas que deben quedar vacías a propósito (Reportes
                # Diarios, Gastos del Mes, FPL) directamente desaparecen del
                # resultado en vez de existir vacías. La raíz misma nunca
                # necesita este marcador (siempre existe implícita).
                if arcdir:
                    zf.writestr(arcdir + "/", "")
                continue
            for name in files:
                full = os.path.join(dirpath, name)
                arcname = f"{arcdir}/{name}" if arcdir else name
                zf.write(full, arcname)

    return output_zip_path, summary


# ---------------------------------------------------------------------------
# Mes Nuevo -- Lottery (2026-09-05)
#
# El archivo "LOTTERY. Analisis MM.YYYY.xlsx" no funciona por "día 1 a día N
# del mes" como Book keeping -- funciona por BLOQUES de 7 días (7 filas de
# día + 1 fila de subtotal + 1 fila "Debito"/Chase Bank, 9 filas por bloque,
# arrancando en la fila 4), y esos bloques cruzan el límite del mes sin
# ningún problema (el bloque de fin de mes casi siempre tiene días del mes
# viejo y del nuevo mezclados). Reglas explícitas del usuario (2026-09-05):
#   1. El ÚLTIMO bloque de la hoja actual -- tenga los 7 días completos o
#      no -- se traslada TAL CUAL (mismos datos, misma fecha) como PRIMER
#      bloque del mes que viene, porque son los días que todavía quedan
#      pendientes de asentar/cobrar.
#   2. Los demás bloques (los que ya estaban completos con datos del mes
#      que se cierra) se "reciclan": se redatan continuos a partir de
#      donde termina el bloque trasladado, con los datos en 0 pero las
#      fórmulas intactas.
#   3. Si reciclar esos bloques no alcanza para cubrir todos los días del
#      mes que viene, se crean bloques nuevos (mismas fórmulas) hasta
#      cubrirlo -- el último bloque de cualquier mes vuelve a meterse unos
#      días en el mes siguiente, a propósito, repitiendo el ciclo.
# La sección de "LIQUIDACION CIERRE LOTTERY" al pie de la hoja (fórmulas
# manuales que referencian los bloques de ESE mes puntual) no se traslada
# -- es del cierre del mes viejo, no aplica al nuevo.
# ---------------------------------------------------------------------------

_LOTTERY_SHEET_NAME_RE = re.compile(r"^(\d{2})\.(\d{4})$")
_LOTTERY_BLOCK_ROWS = 9
_LOTTERY_FIRST_ROW = 4
# D, E, F, G, H, I, K, N, O, P, Q, R, S -- datos crudos por día (los carga
# el módulo Lottery de Herramientas, día a día).
_LOTTERY_DATA_COLS = (4, 5, 6, 7, 8, 9, 11, 14, 15, 16, 17, 18, 19)
# D, E, F, G, I, K, M, P, Q, R, S -- las que SÍ suma la fila de subtotal
# (H y N/O quedan afuera a propósito, así viene la plantilla real).
_LOTTERY_SUBTOTAL_SUM_COLS = (4, 5, 6, 7, 9, 11, 13, 16, 17, 18, 19)


def _lottery_write_day_row(sheet, row, date_a, date_b, data):
    sheet.cell(row=row, column=1, value=date_a)
    sheet.cell(row=row, column=2, value=date_b)
    for col in _LOTTERY_DATA_COLS:
        sheet.cell(row=row, column=col, value=data.get(col, 0))
    sheet.cell(row=row, column=10, value=f"=+I{row}/F{row}")
    sheet.cell(row=row, column=12, value=f"=+K{row}/F{row}")
    sheet.cell(row=row, column=13, value=f"=+I{row}+K{row}")
    sheet.cell(row=row, column=20, value=f"=+S{row}/R{row}")
    sheet.cell(row=row, column=21, value=f"=+Q{row}+R{row}+S{row}")
    sheet.cell(row=row, column=24, value=f"=-G{row}-Q{row}")


def _lottery_write_subtotal_row(sheet, row, day_first, day_last):
    sheet.cell(row=row, column=3, value=0)
    for col in _LOTTERY_SUBTOTAL_SUM_COLS:
        letter = get_column_letter(col)
        sheet.cell(row=row, column=col, value=f"=SUM({letter}{day_first}:{letter}{day_last})")
    sheet.cell(row=row, column=12, value=f"=+K{row}/F{row}")


def _lottery_write_debito_row(sheet, row, subtotal_row, chase_text=None):
    # V (=+F+Q, no confundir con la columna X "CUENTA FINAL" que lee
    # caja.py -- son dos columnas distintas de la misma fila) siempre lleva
    # su fórmula -- bug real reportado por el usuario 2026-09-06: quedaba
    # en blanco en los bloques nuevos/reciclados, en vez de traer la
    # fórmula que ya tenían los demás bloques (E/F/Q también se escriben
    # siempre, reconciliado o no; V tiene que seguir el mismo criterio).
    # W ("Chase Bank\n{fecha}") se calcula afuera (ver _parse_chase_bank_date/
    # _render_chase_bank_text) y se pasa ya armado en chase_text -- acá solo
    # se escribe si vino algo (puede venir None si no se pudo interpretar la
    # fecha real del bloque trasladado).
    sheet.cell(row=row, column=5, value=f"=+E{subtotal_row}-F{subtotal_row}")
    sheet.cell(row=row, column=6, value=f"=+F{subtotal_row}+G{subtotal_row}+I{subtotal_row}+K{subtotal_row}+10")
    sheet.cell(row=row, column=17, value=f"=+Q{subtotal_row}+R{subtotal_row}+S{subtotal_row}")
    sheet.cell(row=row, column=21, value="Debito")
    sheet.cell(row=row, column=22, value=f"=+F{row}+Q{row}")
    if chase_text:
        sheet.cell(row=row, column=23, value=chase_text)


# La fecha de "Chase Bank" (columna W de la fila Debito) es el día en que
# ese monto semanal de Lottery llega al banco -- pedido explícito del
# usuario (2026-09-07): pasa cada 7 días exactos, a partir de la última
# fecha real ya cargada en el bloque trasladado (bloque 1). Nunca se
# reconstruye el texto entero a mano -- se toma el texto real que el
# usuario ya tipeó ("Chase Bank" + salto de línea + fecha) y solo se le
# reemplaza la fecha, preservando separador/ancho/orden día-mes tal cual
# estaban.
_CHASE_BANK_DATE_RE = re.compile(r"(\d{1,2})([/.\-])(\d{1,2})([/.\-])(\d{2,4})")


def _parse_chase_bank_date(text, reference_date):
    """
    Extrae la fecha real del texto de "Chase Bank" -- nunca asume un orden
    día/mes fijo, prueba las dos lecturas posibles del texto real y elige
    la que cae más cerca de reference_date (el último día real del bloque
    trasladado) para desambiguar cuando ambas son fechas válidas. Devuelve
    (fecha, match, day_is_first) o (None, None, None) si no hay texto o no
    se pudo interpretar con confianza.
    """
    if not text:
        return None, None, None
    match = _CHASE_BANK_DATE_RE.search(str(text))
    if not match:
        return None, None, None
    g1, sep1, g2, sep2, g3 = match.groups()
    if sep1 != sep2:
        return None, None, None
    year = int(g3) if len(g3) == 4 else 2000 + int(g3)

    candidates = []
    try:
        candidates.append((datetime(year, int(g2), int(g1)), True))  # g1=día, g2=mes
    except ValueError:
        pass
    try:
        candidates.append((datetime(year, int(g1), int(g2)), False))  # g1=mes, g2=día
    except ValueError:
        pass
    if not candidates:
        return None, None, None
    if len(candidates) == 1:
        date, day_is_first = candidates[0]
    else:
        date, day_is_first = min(candidates, key=lambda c: abs((c[0] - reference_date).days))
    return date, match, day_is_first


def _render_chase_bank_text(template_text, match, day_is_first, new_date):
    """
    Reemplaza solo la fecha dentro del texto real ya tipeado por el
    usuario (ver _parse_chase_bank_date), preservando el resto tal cual --
    el label "Chase Bank", el salto de línea, el separador, y el ancho de
    cada número (01 vs 1, año de 2 vs 4 dígitos).
    """
    g1, sep1, g2, sep2, g3 = match.groups()
    day_str, month_str = (g1, g2) if day_is_first else (g2, g1)
    new_day = f"{new_date.day:02d}" if len(day_str) == 2 else str(new_date.day)
    new_month = f"{new_date.month:02d}" if len(month_str) == 2 else str(new_date.month)
    new_year = str(new_date.year) if len(g3) == 4 else f"{new_date.year % 100:02d}"
    first, second = (new_day, new_month) if day_is_first else (new_month, new_day)
    new_date_text = f"{first}{sep1}{second}{sep2}{new_year}"
    return template_text[: match.start()] + new_date_text + template_text[match.end() :]


def _lottery_merge_block(sheet, block_start):
    subtotal_row = block_start + 7
    debito_row = block_start + 8
    sheet.merge_cells(start_row=subtotal_row, start_column=1, end_row=debito_row, end_column=1)
    sheet.merge_cells(start_row=subtotal_row, start_column=2, end_row=debito_row, end_column=2)
    sheet.merge_cells(start_row=debito_row, start_column=6, end_row=debito_row, end_column=11)
    sheet.merge_cells(start_row=debito_row, start_column=17, end_row=debito_row, end_column=19)


def _lottery_capture_style(sheet, block_start):
    # Incluye "fill" (color de fondo) -- bug real reportado por el usuario
    # 2026-09-05: sin esto, los bloques nuevos/reciclados perdían todo el
    # color de la plantilla (M/U/V/X en cada día, F/G/M/P/Q/R/S/V en el
    # subtotal, E/F/Q/V en el debito).
    style_rows = []
    for offset in range(_LOTTERY_BLOCK_ROWS):
        row = block_start + offset
        row_styles = {}
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row, column=col)
            if cell.__class__.__name__ == "MergedCell":
                continue
            row_styles[col] = (
                copy(cell.font), copy(cell.border), copy(cell.alignment),
                cell.number_format, copy(cell.fill),
            )
        style_rows.append(row_styles)
    return style_rows


def _lottery_apply_style(sheet, block_start, style_rows):
    for offset in range(_LOTTERY_BLOCK_ROWS):
        row = block_start + offset
        for col, (font, border, alignment, number_format, fill) in style_rows[offset].items():
            cell = sheet.cell(row=row, column=col)
            if cell.__class__.__name__ == "MergedCell":
                continue
            cell.font = copy(font)
            cell.border = copy(border)
            cell.alignment = copy(alignment)
            cell.number_format = number_format
            cell.fill = copy(fill)


def _lottery_cell_has_content(cell):
    """Una celda "cuenta" para detectar dónde termina la sección de
    liquidación tanto si tiene un valor como si solo tiene relleno de color
    (la "cajita" visual puede terminar en una fila pintada sin ningún dato) --
    bug real encontrado en esta auditoría 2026-09-06: la detección vieja
    solo miraba `.value`, así que una fila final sin valor pero coloreada
    quedaba afuera de `footer_end_old` y `sheet.delete_rows` (que sí borra
    hasta el final real de la hoja) se la comía sin que _lottery_capture_footer
    llegara a guardarla antes -- se perdía para siempre."""
    if cell.value is not None:
        return True
    fill = getattr(cell, "fill", None)
    return bool(fill and fill.patternType is not None)


def _lottery_capture_footer(sheet, start_row, end_row):
    """
    Captura tal cual (valor crudo -- fórmula o literal -- y estilo completo
    incluido el color de fondo) la sección que vive debajo del último
    bloque -- pedido explícito del usuario 2026-09-05: esa sección
    ("LIQUIDACION CIERRE LOTTERY", cierre manual de ESE mes puntual) tiene
    que pasar al mes que viene con el mismo formato y las mismas
    referencias a los bloques semanales (esas se revisan a mano). Las
    referencias que apuntan DENTRO de la propia sección si se corrigen,
    ver _lottery_paste_footer.
    """
    # Captura TODAS las celdas, tengan valor o no -- bug real reportado por
    # el usuario 2026-09-06: varias celdas de esta sección están pintadas
    # (amarillo/verde, formando la "cajita" visual de la liquidación) pero
    # sin ningún valor adentro -- saltearlas porque estaban vacías perdía
    # el color y dejaba la sección con huecos blancos, pareciendo rota.
    rows = []
    for row in range(start_row, end_row + 1):
        row_cells = {}
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row, column=col)
            if cell.__class__.__name__ == "MergedCell":
                continue
            row_cells[col] = (
                cell.value, copy(cell.font), copy(cell.border), copy(cell.alignment),
                cell.number_format, copy(cell.fill),
            )
        rows.append(row_cells)
    merges = []
    for mcr in sheet.merged_cells.ranges:
        if mcr.min_row >= start_row:
            merges.append((mcr.min_row - start_row, mcr.min_col, mcr.max_row - start_row, mcr.max_col))
    return rows, merges


def _lottery_paste_footer(sheet, new_start_row, footer_rows, footer_merges, footer_start_old, footer_end_old):
    """
    Pega la sección tal cual, PERO corrige las referencias que apuntan
    DENTRO de la propia sección (ej. "=+F58-E58", la fila de arriba menos
    esta misma) -- como la sección entera se corrió de lugar, esas
    referencias quedaban rotas (#¡VALOR!/#¡REF!, bug real reportado por el
    usuario 2026-09-05: "falta corregir la exactitud"). Las referencias a
    los bloques semanales (fuera de este rango) NO se tocan -- esas sí
    quedan tal cual, a propósito, para que el usuario las revise a mano.
    """
    delta = new_start_row - footer_start_old
    row_ref_re = re.compile(r"(\$?[A-Z]+\$?)(\d+)")

    def shift_internal_refs(formula):
        def repl(m):
            row_num = int(m.group(2))
            if footer_start_old <= row_num <= footer_end_old:
                return f"{m.group(1)}{row_num + delta}"
            return m.group(0)
        return row_ref_re.sub(repl, formula)

    for offset, row_cells in enumerate(footer_rows):
        row = new_start_row + offset
        for col, (value, font, border, alignment, number_format, fill) in row_cells.items():
            if isinstance(value, str) and value.startswith("="):
                value = shift_internal_refs(value)
            cell = sheet.cell(row=row, column=col)
            cell.value = value  # nunca sheet.cell(..., value=value) -- ese
            # atajo de openpyxl NO escribe nada si value es None (ver
            # _clear_cell más arriba en este archivo) -- acá sí hace falta
            # poder escribir None de verdad, para celdas pintadas sin dato.
            cell.font = copy(font)
            cell.border = copy(border)
            cell.alignment = copy(alignment)
            cell.number_format = number_format
            cell.fill = copy(fill)
    for row_offset_min, col_min, row_offset_max, col_max in footer_merges:
        sheet.merge_cells(
            start_row=new_start_row + row_offset_min, start_column=col_min,
            end_row=new_start_row + row_offset_max, end_column=col_max,
        )


def prepare_next_month_lottery(upload_path):
    """
    Función pública del sub-módulo Mes Nuevo -- Lottery: recibe la ruta de
    un .xlsx "LOTTERY. Analisis MM.YYYY.xlsx" (el que sube el usuario, nunca
    se toca) y devuelve (output_path, summary) con el .xlsx ya preparado
    para el mes que viene, guardado en un temporal propio.
    """
    summary = {"warnings": [], "assumptions": []}
    wb = load_workbook(upload_path, data_only=False)

    sheet = None
    for name in wb.sheetnames:
        if _LOTTERY_SHEET_NAME_RE.match(name.strip()):
            sheet = wb[name]
            break
    if sheet is None:
        raise MesNuevoError(
            'No encontré una hoja con nombre "MM.YYYY" (ej. "08.2026") en el archivo de Lottery.'
        )
    match = _LOTTERY_SHEET_NAME_RE.match(sheet.title.strip())
    month, year = int(match.group(1)), int(match.group(2))
    next_year, next_month = _next_month(year, month)
    summary["source_period"] = f"{MESES[month]} {year}"
    summary["target_period"] = f"{MESES[next_month]} {next_year}"

    # ---- Contar los bloques existentes (7 filas de día + subtotal + debito) ----
    total_blocks = 0
    while isinstance(
        sheet.cell(row=_LOTTERY_FIRST_ROW + _LOTTERY_BLOCK_ROWS * total_blocks, column=1).value,
        datetime,
    ):
        total_blocks += 1
    if total_blocks == 0:
        raise MesNuevoError("No encontré ningún bloque semanal (7 días) a partir de la fila 4.")

    last_block_start = _LOTTERY_FIRST_ROW + _LOTTERY_BLOCK_ROWS * (total_blocks - 1)

    # ---- Capturar el último bloque (se traslada tal cual) ----
    carried_dates = []
    carried_data = []
    for i in range(7):
        row = last_block_start + i
        carried_dates.append((sheet.cell(row=row, column=1).value, sheet.cell(row=row, column=2).value))
        carried_data.append({col: sheet.cell(row=row, column=col).value for col in _LOTTERY_DATA_COLS})
    debito_row_old = last_block_start + 8
    carried_chase_text = sheet.cell(row=debito_row_old, column=23).value

    next_date = carried_dates[-1][1]  # columna B del 7mo día del bloque trasladado

    # ---- Fecha de "Chase Bank" (columna W) -- se repite cada 7 días exactos
    # a partir de la del bloque trasladado (pedido explícito del usuario
    # 2026-09-07) -- ver _parse_chase_bank_date/_render_chase_bank_text.
    chase_date, chase_match, chase_day_first = _parse_chase_bank_date(carried_chase_text, next_date)
    if carried_chase_text and chase_date is None:
        summary["warnings"].append(
            'No pude interpretar la fecha del texto de "Chase Bank" (columna W) del último '
            "bloque -- los bloques nuevos quedaron sin esa fecha, completala a mano."
        )

    # ---- Plantilla de estilo (del primer bloque -- estructuralmente igual a todos) ----
    style_rows = _lottery_capture_style(sheet, _LOTTERY_FIRST_ROW)

    # ---- Capturar tal cual la sección debajo del último bloque
    # ("LIQUIDACION CIERRE LOTTERY") -- pedido explícito del usuario
    # 2026-09-05: pasa igual al mes que viene, con el mismo formato,
    # como está en el mes anterior -- no se reconstruye ni se actualiza.
    footer_start_old = _LOTTERY_FIRST_ROW + _LOTTERY_BLOCK_ROWS * total_blocks
    footer_end_old = footer_start_old - 1
    for row in range(footer_start_old, sheet.max_row + 1):
        if any(_lottery_cell_has_content(sheet.cell(row=row, column=col)) for col in range(1, sheet.max_column + 1)):
            footer_end_old = row
    footer_rows, footer_merges = (
        _lottery_capture_footer(sheet, footer_start_old, footer_end_old)
        if footer_end_old >= footer_start_old
        else ([], [])
    )

    # ---- Vaciar todo desde la fila 4: los bloques viejos Y la sección de
    # liquidación (ya capturada arriba, se vuelve a pegar tal cual más abajo) ----
    for coord in [mcr.coord for mcr in list(sheet.merged_cells.ranges) if mcr.min_row >= _LOTTERY_FIRST_ROW]:
        sheet.unmerge_cells(coord)
    total_old_rows = sheet.max_row - _LOTTERY_FIRST_ROW + 1
    if total_old_rows > 0:
        sheet.delete_rows(_LOTTERY_FIRST_ROW, amount=total_old_rows)

    # ---- Calcular cuántos bloques hacen falta para cubrir el mes que viene ----
    last_day_target = datetime(next_year, next_month, calendar.monthrange(next_year, next_month)[1])
    if next_date > last_day_target:
        extra_blocks = 0
    else:
        days_to_cover = (last_day_target - next_date).days + 1
        extra_blocks = -(-days_to_cover // 7)  # división hacia arriba
    total_new_blocks = 1 + extra_blocks

    sheet.insert_rows(_LOTTERY_FIRST_ROW, amount=_LOTTERY_BLOCK_ROWS * total_new_blocks)

    # ---- Bloque 1: el trasladado, tal cual ----
    block_start = _LOTTERY_FIRST_ROW
    for i in range(7):
        date_a, date_b = carried_dates[i]
        _lottery_write_day_row(sheet, block_start + i, date_a, date_b, carried_data[i])
    _lottery_write_subtotal_row(sheet, block_start + 7, block_start, block_start + 6)
    _lottery_write_debito_row(sheet, block_start + 8, block_start + 7, chase_text=carried_chase_text)
    _lottery_apply_style(sheet, block_start, style_rows)
    _lottery_merge_block(sheet, block_start)

    # ---- Bloques siguientes: reciclados/nuevos, fechas continuas, datos en 0 ----
    cursor = next_date
    for b in range(1, total_new_blocks):
        block_start = _LOTTERY_FIRST_ROW + _LOTTERY_BLOCK_ROWS * b
        for i in range(7):
            date_a = cursor
            date_b = cursor + timedelta(days=1)
            _lottery_write_day_row(sheet, block_start + i, date_a, date_b, {})
            cursor = date_b
        _lottery_write_subtotal_row(sheet, block_start + 7, block_start, block_start + 6)
        if chase_date is not None:
            block_chase_text = _render_chase_bank_text(
                carried_chase_text, chase_match, chase_day_first, chase_date + timedelta(days=7 * b)
            )
        else:
            block_chase_text = None
        _lottery_write_debito_row(sheet, block_start + 8, block_start + 7, chase_text=block_chase_text)
        _lottery_apply_style(sheet, block_start, style_rows)
        _lottery_merge_block(sheet, block_start)

    # ---- Pegar tal cual la sección de liquidación (capturada arriba),
    # justo debajo del último bloque nuevo -- mismo formato, sin actualizar
    # ninguna referencia, igual que pedido explícito del usuario 2026-09-05.
    if footer_rows:
        footer_start_new = _LOTTERY_FIRST_ROW + _LOTTERY_BLOCK_ROWS * total_new_blocks
        _lottery_paste_footer(
            sheet, footer_start_new, footer_rows, footer_merges,
            footer_start_old, footer_end_old,
        )

    sheet.title = f"{next_month:02d}.{next_year:04d}"
    chase_note = (
        ' La fecha de "Chase Bank" (columna W) se calculó sola cada 7 días a partir de la '
        "del bloque trasladado en los bloques nuevos."
        if chase_date is not None
        else ""
    )
    summary["assumptions"].append(
        f"Se trasladó el último bloque de {MESES[month]} tal cual (con los datos que ya tenía) "
        f"como primer bloque de {MESES[next_month]}, y se agregaron {total_new_blocks - 1} "
        "bloque(s) más en 0 para cubrir el resto del mes. La sección de liquidación al pie se "
        "copió tal cual estaba en el mes anterior (mismo formato y fórmulas, sin actualizar "
        "ninguna referencia) -- revisala a mano antes de usarla, las fórmulas todavía apuntan a "
        f"los bloques del mes viejo.{chase_note}"
    )

    workdir = tempfile.mkdtemp(prefix="mes_nuevo_lottery_")
    output_name = f"LOTTERY. Analisis {next_month:02d}.{next_year:04d}.xlsx"
    output_path = os.path.join(workdir, output_name)
    wb.save(output_path)
    wb.close()
    return output_path, summary
