"""
Herramientas: Balance Mensual -- nueva herramienta para completar el Excel
"Bce ..." (balance mensual), que tiene muchas hojas independientes,
alimentada cada una por el Mayor de una cuenta contable distinta.

Tres formas de completar una hoja, según la cuenta:

- **Reemplazo simple** (SIMPLE_REPLACE_ACCOUNTS) -- Adm Fees, Seguros, Com
  y Gtos Bcarios, Com Lottery, Alq Food Truck, Sale tax: la hoja ES,
  literalmente, una copia pegada del Mayor de esa cuenta (título,
  encabezado, SALDO INICIAL, un renglón por asiento, fila de totales) --
  se reemplaza por completo, sin ningún procesamiento extra.
- **Reemplazo con lógica propia** (SPECIAL_REPLACE_ACCOUNTS) -- 2.J.H.
  WILLIAMS OIL y Caja: también se reemplaza toda la hoja, pero además se
  completa/pinta contenido derivado del propio Mayor (ver
  `_build_jhw_sheet`/`_build_caja_sheet` y sus docstrings).
- **Acumulado anual** (APPEND_ACCOUNTS) -- Otros Gtos y Reparaciones: la
  hoja NUNCA se reemplaza -- cada mes se pega un bloque nuevo (título +
  encabezado + SALDO INICIAL + asientos + totales) unas filas debajo del
  último bloque ya cargado, dejando el resto de la hoja (los meses
  anteriores del año) completamente intacto. Se limpia una vez al año a
  mano (no automatizado acá).

Reemplazo/append quirúrgico a nivel de XML, no re-serializar todo el libro
--------------------------------------------------------------------------
El archivo real (confirmado leyendo "Bce Brandenton 31-07-26.xlsx", solo
lectura) tiene gráficos, una tabla de consulta/conexión de datos vieja
(de 2023, ligada a la PC de otra persona, ya sin uso) y comentarios de
celda en otras hojas -- abrir el libro entero con openpyxl y volver a
guardarlo (el patrón que usa el resto de este proyecto) le degrada el
estilo a los gráficos (pierde el "chart style" moderno) y cambia el
formato interno de los comentarios, aunque nunca se toquen esas hojas
directamente -- es una limitación conocida de openpyxl al reserializar un
libro, no un problema de este archivo en particular. Como este es un
archivo mensual real que probablemente se comparte/presenta (hay una
hoja "Informe Power Point"), y esta herramienta se va a correr todos los
meses, cualquier degradación se notaría y acumularía con el tiempo.

Por eso, en vez de abrir con openpyxl y volver a guardar todo el libro,
esta herramienta edita el .xlsx como lo que es (un .zip con partes XML):
- Reemplazo: solo se reescribe el contenido (`<sheetData>`/
  `<mergeCells>`/`<dimension>`) del `sheetN.xml` de esa hoja puntual.
- Append: se insertan filas NUEVAS justo antes de `</sheetData>` y merges
  nuevos antes de `</mergeCells>` -- ni una sola fila/celda ya existente
  se toca.
Cada otra parte del archivo (estilos, gráficos, comentarios, conexiones,
configuración de impresión, dibujos, y hasta las demás hojas) se copia
byte a byte sin abrir, así que no hay ningún riesgo de degradarlas. El
texto nuevo se escribe como "inline string" (`t="inlineStr"`) en vez de
agregarse a `sharedStrings.xml`, para no tener que tocar esa tabla
compartida tampoco. Los estilos (`s="..."`) de cada celda nueva se toman
de la propia hoja ANTES de tocarla (título/encabezado/SALDO INICIAL/fila
de datos/fila de totales -- o, para Otros Gtos/Reparaciones, del ÚLTIMO
bloque mensual ya cargado), así siempre son índices que ya existían en
`styles.xml` -- nunca hace falta agregar un estilo nuevo.

Openpyxl SÍ se usa, pero solo para LEER (nunca para guardar) contenido
que es mucho más cómodo de resolver así que a mano en XML crudo (texto
con referencia a `sharedStrings.xml`, ubicar una fila por su propio
contenido) -- abrir un archivo sin llamar a `.save()` no tiene ningún
riesgo de degradación, la escritura real siempre pasa por la cirugía de
zip/XML de arriba.
"""

import os
import re
import tempfile
import zipfile
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_COLUMNS = ["A", "B", "C", "D", "E", "F", "G", "H"]
_TEXT_COLUMNS = {"B", "C", "D", "E"}

# Cuenta contable (código exacto tal cual aparece en la línea "Cuenta: ...
# Emisor:...") -> nombre de hoja en el Excel de Balance. El nombre de hoja
# se usa tal cual (incluye el espacio final real de "Sale tax ").
SIMPLE_REPLACE_ACCOUNTS = {
    "5.09.00.00": "Adm Fees",
    "5.08.00.00": "Seguros",
    "5.01.00.00": "Com y Gtos Bcarios",
    "4.10.00.00": "Com Lottery",
    "4.06.00.00": "Alq Food Truck",
    "2.02.01.00": "Sale tax ",
}

# Mismo reemplazo de hoja completa que arriba, pero además con contenido
# derivado (fórmulas/pintado) que no viene del Mayor en sí.
SPECIAL_REPLACE_ACCOUNTS = {
    "2.01.01.00": "2.J.H. WILLIAMS OIL",
    "1.01.01.00": "Caja",
}

# Estas NUNCA se reemplazan -- cada mes se pega un bloque nuevo debajo del
# último ya cargado.
APPEND_ACCOUNTS = {
    "5.11.00.00": "Otros Gtos",
    "5.10.00.00": "Reparaciones",
}

# Filas entre la fila de totales del último bloque y el título del bloque
# nuevo, por cuenta -- confirmado contra el archivo real: Otros Gtos deja
# 2 filas en blanco, Reparaciones deja 4 (ahí además va la plantilla sin
# completar de GASTOS PUNTUALES/HABITUALES, ver `_append_reparaciones`).
APPEND_GAP_ROWS = {
    "5.11.00.00": 2,
    "5.10.00.00": 4,
}

# Cuentas que todavía no tienen ninguna lógica implementada -- si se sube
# un Mayor de una de estas, se avisa en vez de arriesgar un reemplazo que
# rompería lo que ya tiene esa hoja. Vacío por ahora (las 10 cuentas ya
# habladas con el usuario están todas implementadas) -- se va a ir
# llenando de nuevo a medida que aparezcan más hojas del Balance sin
# explicar todavía.
KNOWN_UNSUPPORTED_ACCOUNTS = {}

_CUENTA_LINE_RE = re.compile(r"^Cuenta:\s*([\d.]+)\s*-\s*(.+?)\s*Emisor:", re.IGNORECASE)
_DESDE_HASTA_RE = re.compile(r"desde:\s*([\d/-]+)\s*hasta:\s*([\d/-]+)", re.IGNORECASE)

_SHEET_DATA_RE = re.compile(r"<sheetData>.*?</sheetData>", re.S)
_MERGE_CELLS_BLOCK_RE = re.compile(r'<mergeCells count="(\d+)">(.*?)</mergeCells>', re.S)
_DIMENSION_RE = re.compile(r'<dimension ref="[^"]*"/>')
_ROW_CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
_TOTALS_MERGE_RE = re.compile(r'<mergeCell ref="A(\d+):D(\d+)"/>')
_ROW_RE_TEMPLATE = r'<row r="{row}"[^>]*>(.*?)</row>'

_MAYOR_DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y")


class MayorRow:
    __slots__ = ("asiento", "fecha", "cliente", "descripcion", "detalle", "debito", "credito", "saldo")

    def __init__(self, asiento, fecha, cliente, descripcion, detalle, debito, credito, saldo):
        self.asiento = asiento
        self.fecha = fecha
        self.cliente = cliente
        self.descripcion = descripcion
        self.detalle = detalle
        self.debito = debito
        self.credito = credito
        self.saldo = saldo


def _clean_text(value):
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    text = str(value).strip()
    return text or None


def _parse_mayor_date(text):
    text = text.strip()
    for fmt in _MAYOR_DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _read_mayor_full(path):
    """
    Lee un Mayor exportado completo (.xls o .xlsx): línea "Cuenta: ...",
    línea "Desde: ... Hasta: ...", fila de encabezado (por contenido, no
    posición fija) y todas las filas reales debajo (SALDO INICIAL +
    asientos), hasta la primera fila totalmente vacía.

    Devuelve un dict con account_code, account_name, title_line1 (texto
    tal cual de la línea "Cuenta: ..."), title_line2 (texto tal cual de
    "Desde: ... Hasta: ..."), period_from/period_to (fechas ya
    parseadas), y rows (lista de MayorRow, en el mismo orden del archivo
    -- la primera siempre debería ser SALDO INICIAL).
    """
    df = pd.read_excel(path, header=None)

    title_line1 = None
    account_code = None
    account_name = None
    title_line2 = None
    period_from = period_to = None
    header_row_index = None
    col_index = {}

    for idx in range(len(df)):
        cells = ["" if pd.isna(v) else str(v).strip() for v in df.iloc[idx]]
        joined = " ".join(c for c in cells if c)
        if title_line1 is None and joined.lower().startswith("cuenta:"):
            match = _CUENTA_LINE_RE.search(joined)
            if match:
                title_line1 = joined
                account_code = match.group(1).strip()
                account_name = match.group(2).strip()
        if title_line2 is None:
            match = _DESDE_HASTA_RE.search(joined)
            if match:
                title_line2 = joined
                period_from = _parse_mayor_date(match.group(1))
                period_to = _parse_mayor_date(match.group(2))
        if header_row_index is None:
            lowered = [c.lower() for c in cells]
            if "fecha" in lowered and "debito" in lowered and "credito" in lowered:
                header_row_index = idx

                def _find(name, lowered=lowered):
                    for i, c in enumerate(lowered):
                        if name in c:
                            return i
                    return None

                col_index = {
                    "asiento": _find("asiento"),
                    "fecha": _find("fecha"),
                    "cliente": _find("cliente"),
                    "descripcion": _find("descripcion"),
                    "detalle": _find("detalle"),
                    "debito": _find("debito"),
                    "credito": _find("credito"),
                    "saldo": _find("saldo"),
                }

    if title_line1 is None or account_code is None:
        raise ValueError('No se encontró la línea "Cuenta: ..." en el Mayor.')
    if title_line2 is None:
        raise ValueError('No se encontró la línea "Desde: ... Hasta: ..." en el Mayor.')
    if header_row_index is None:
        raise ValueError('No se encontró la fila de encabezado ("Fecha"/"Debito"/"Credito") en el Mayor.')

    def _cell(row_idx, key):
        col = col_index.get(key)
        if col is None:
            return None
        return _clean_text(df.iat[row_idx, col])

    def _num(row_idx, key):
        col = col_index.get(key)
        if col is None:
            return None
        value = df.iat[row_idx, col]
        if pd.isna(value):
            return None
        return float(value)

    rows = []
    for idx in range(header_row_index + 1, len(df)):
        fecha = _cell(idx, "fecha")
        descripcion = _cell(idx, "descripcion")
        debito = _num(idx, "debito")
        credito = _num(idx, "credito")
        saldo = _num(idx, "saldo")
        if not fecha and not descripcion and debito is None and credito is None and saldo is None:
            break
        rows.append(MayorRow(
            asiento=_cell(idx, "asiento"),
            fecha=fecha,
            cliente=_cell(idx, "cliente"),
            descripcion=descripcion,
            detalle=_cell(idx, "detalle"),
            debito=debito,
            credito=credito,
            saldo=saldo,
        ))

    if not rows:
        raise ValueError("El Mayor no tiene ninguna fila de datos debajo del encabezado.")
    first_desc = (rows[0].descripcion or "").strip().lower()
    if "saldo inicial" not in first_desc:
        raise ValueError('La primera fila del Mayor debería ser "SALDO INICIAL" y no lo es -- revisar el archivo.')

    return {
        "account_code": account_code,
        "account_name": account_name,
        "title_line1": title_line1,
        "title_line2": title_line2,
        "period_from": period_from,
        "period_to": period_to,
        "rows": rows,
    }


def _map_sheet_paths(zf):
    """Nombre de hoja -> ruta del sheetN.xml dentro del .zip, vía workbook.xml + sus rels."""
    import xml.etree.ElementTree as ET

    wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {}
    for rel in rels_xml.findall(f"{{{_PKG_REL_NS}}}Relationship"):
        rel_targets[rel.attrib["Id"]] = rel.attrib["Target"]

    mapping = {}
    for sheet in wb_xml.findall(f"{{{_MAIN_NS}}}sheets/{{{_MAIN_NS}}}sheet"):
        name = sheet.attrib["name"]
        rid = sheet.attrib.get(f"{{{_REL_NS}}}id")
        target = rel_targets.get(rid) if rid else None
        if target and target.startswith("worksheets/"):
            mapping[name] = "xl/" + target
    return mapping


def _col_to_index(col):
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _extract_row_styles(xml_text, row_number):
    """Dict {columna: style_id} de una fila existente, leyendo el atributo s="" de cada celda."""
    match = re.search(_ROW_RE_TEMPLATE.format(row=row_number), xml_text, re.S)
    if not match:
        return {}
    styles = {}
    for cell_match in _ROW_CELL_RE.finditer(match.group(1)):
        col_letter = cell_match.group(1)
        attrs = cell_match.group(3)
        style_match = re.search(r's="(\d+)"', attrs)
        styles[col_letter] = style_match.group(1) if style_match else None
    return styles


def _find_totals_rows(xml_text):
    """
    Ubica TODAS las filas de totales de una hoja por su forma distintiva
    -- el merge "A{n}:D{n}" (mismo número a ambos lados) -- confirmando
    que la fila tenga contenido real en F o G (algunas hojas, ej. Caja,
    tienen además una fila "espaciadora" con el mismo tipo de merge pero
    sin ningún valor, que si no quedaría tomada por error). Devuelve la
    lista de números de fila encontrados, ordenada ascendente.
    """
    candidates = sorted({int(m.group(1)) for m in _TOTALS_MERGE_RE.finditer(xml_text) if m.group(1) == m.group(2)})
    result = []
    for row_num in candidates:
        row_match = re.search(_ROW_RE_TEMPLATE.format(row=row_num), xml_text, re.S)
        if not row_match:
            continue
        f_match = re.search(r'<c r="F\d+"[^>]*>(.*?)</c>', row_match.group(1), re.S)
        g_match = re.search(r'<c r="G\d+"[^>]*>(.*?)</c>', row_match.group(1), re.S)
        if (f_match and f_match.group(1).strip()) or (g_match and g_match.group(1).strip()):
            result.append(row_num)
    return result


def _find_single_totals_row(xml_text, sheet_name):
    """La única fila de totales de una hoja de reemplazo simple/especial (un solo Mayor pegado)."""
    rows = _find_totals_rows(xml_text)
    if not rows:
        return None
    if len(rows) > 1:
        raise ValueError(
            f'La hoja "{sheet_name}" tiene más de una fila de totales reconocible -- revisar a mano.'
        )
    return rows[0]


def _load_style_templates(xml_text, sheet_name, title_row=1, data_row=None, totals_row=None):
    """
    Captura los estilos por columna de título/encabezado/SALDO INICIAL/
    primer asiento/totales. `data_row` por default es la primera fila de
    asiento (title_row+4) -- pero eso asume que esa fila tiene el estilo
    "plano" de siempre, cierto en todas las hojas ya validadas salvo
    Caja, donde la primera fila real (julio-2026) resulta ser una de las
    coloreadas (ver `_find_plain_row`) -- por eso `_build_caja_sheet` pasa
    acá, en cambio, una fila sin pintar encontrada a propósito. `totals_row`
    hace falta pasarlo explícito en una hoja acumulativa (Otros Gtos/
    Reparaciones), que tiene más de un bloque mensual y por lo tanto más
    de una fila de totales -- ahí no alcanza con buscarla sola.
    """
    header_row = title_row + 2
    saldo_row = title_row + 3
    data_row = data_row if data_row is not None else title_row + 4

    row1 = _extract_row_styles(xml_text, title_row)
    row2 = _extract_row_styles(xml_text, title_row + 1)
    row3 = _extract_row_styles(xml_text, header_row)
    row4 = _extract_row_styles(xml_text, saldo_row)
    row5 = _extract_row_styles(xml_text, data_row)

    last_row_num = totals_row if totals_row is not None else _find_single_totals_row(xml_text, sheet_name)
    if last_row_num is None:
        raise ValueError(
            f'La hoja "{sheet_name}" no tiene una fila de totales reconocible (un merge "A{{n}}:D{{n}}") '
            "-- revisar a mano."
        )
    last_row = _extract_row_styles(xml_text, last_row_num)

    # No todas las hojas escriben una celda explícita (con o sin estilo)
    # para cada una de las 8 columnas -- Excel puede omitir del todo una
    # celda en blanco con estilo por default (confirmado en "Sale tax ",
    # que tiene varias columnas de la fila 4/5 sin ninguna celda propia).
    # Alcanza con que la fila exista y tenga AL MENOS una celda.
    checks = [
        ("título (fila 1)", row1), ("título (fila 2)", row2), ("encabezado", row3),
        ("SALDO INICIAL", row4), ("primer asiento", row5), ("totales", last_row),
    ]
    for label, styles in checks:
        if not styles:
            raise ValueError(
                f'La hoja "{sheet_name}" no tiene la forma esperada (no se encontró ninguna '
                f"celda en {label}) -- revisar a mano."
            )

    return {"title1": row1, "title2": row2, "header": row3, "saldo_inicial": row4, "data": row5, "totals": last_row}


def _xml_escape(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _format_number(value):
    native = float(value)
    if native.is_integer():
        return str(int(native))
    return repr(native)


def _cell_xml(col, row, style_id, value=None, is_text=False):
    ref = f"{col}{row}"
    s_attr = f' s="{style_id}"' if style_id is not None else ""
    if value is None or (is_text and value == ""):
        return f'<c r="{ref}"{s_attr}/>'
    if is_text:
        escaped = _xml_escape(str(value))
        return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{escaped}</t></is></c>'
    return f'<c r="{ref}"{s_attr}><v>{_format_number(value)}</v></c>'


def _formula_cell_xml(col, row, style_id, formula, cached_value):
    ref = f"{col}{row}"
    s_attr = f' s="{style_id}"' if style_id is not None else ""
    cached = _format_number(cached_value)
    return f'<c r="{ref}"{s_attr}><f>{_xml_escape(formula)}</f><v>{cached}</v></c>'


def _row_xml_from_map(row_number, cell_map, height=None):
    """cell_map: {columna: cell_xml}. Calcula spans a partir de las columnas realmente presentes."""
    if not cell_map:
        return ""
    cols_sorted = sorted(cell_map.keys(), key=_col_to_index)
    min_idx = _col_to_index(cols_sorted[0])
    max_idx = _col_to_index(cols_sorted[-1])
    height_attr = f' ht="{height}" customHeight="1"' if height else ""
    cells_xml = "".join(cell_map[c] for c in cols_sorted)
    return f'<row r="{row_number}" spans="{min_idx}:{max_idx}"{height_attr} x14ac:dyDescent="0.25">{cells_xml}</row>'


def _rows_map_to_xml(rows_map):
    return "".join(
        _row_xml_from_map(r, rows_map[r].get("cells"), height=rows_map[r].get("height"))
        for r in sorted(rows_map.keys())
    )


def _build_base_rows(mayor, styles, start_row=1):
    """
    Arma las filas A-H comunes a toda hoja de Mayor (título x2, encabezado,
    SALDO INICIAL, un renglón por asiento, fila de totales con
    `=SUM(...)`), como un dict {fila: {"cells": {columna: cell_xml},
    "height": alto|None}} -- empezando en `start_row` (1 para un
    reemplazo completo, o la fila calculada para pegar un mes nuevo debajo
    del último ya existente en una hoja acumulativa).

    Devuelve (rows_map, meta) -- meta tiene title1_row/title2_row/
    header_row/saldo_row/first_data_row/last_data_row/totals_row/
    total_debito/total_credito/transactions (la lista de MayorRow sin la
    de SALDO INICIAL).
    """
    rows = mayor["rows"]
    saldo_inicial = rows[0]
    transactions = rows[1:]
    if not transactions:
        raise ValueError(
            f'El Mayor de "{mayor["account_name"]}" no tiene ningún movimiento en el período '
            "(solo SALDO INICIAL) -- revisalo a mano antes de cargarlo."
        )

    rows_map = {}

    title1_row = start_row
    title2_row = start_row + 1
    header_row = start_row + 2
    saldo_row = start_row + 3
    first_data_row = start_row + 4

    title1_style = styles["title1"]
    r1_cells = {col: _cell_xml(col, title1_row, title1_style.get(col),
                                mayor["title_line1"] if col == "A" else None, is_text=True) for col in _COLUMNS}
    rows_map[title1_row] = {"cells": r1_cells, "height": "15.75"}

    title2_style = styles["title2"]
    r2_cells = {col: _cell_xml(col, title2_row, title2_style.get(col),
                                mayor["title_line2"] if col == "A" else None, is_text=True) for col in _COLUMNS}
    rows_map[title2_row] = {"cells": r2_cells, "height": "15.75"}

    header_style = styles["header"]
    header_labels = {
        "A": "Asiento/Cpbte.", "B": "Fecha", "C": "Cliente/Proveedor", "D": "Descripcion",
        "E": "Detalle", "F": "Debito", "G": "Credito", "H": "Saldo",
    }
    r3_cells = {col: _cell_xml(col, header_row, header_style.get(col), header_labels[col], is_text=True)
                for col in _COLUMNS}
    rows_map[header_row] = {"cells": r3_cells, "height": None}

    saldo_style = styles["saldo_inicial"]
    r4_cells = {}
    for col in _COLUMNS:
        if col == "D":
            r4_cells[col] = _cell_xml(col, saldo_row, saldo_style.get(col),
                                       saldo_inicial.descripcion or "SALDO INICIAL", is_text=True)
        elif col == "H":
            r4_cells[col] = _cell_xml(col, saldo_row, saldo_style.get(col), saldo_inicial.saldo)
        else:
            r4_cells[col] = _cell_xml(col, saldo_row, saldo_style.get(col))
    rows_map[saldo_row] = {"cells": r4_cells, "height": None}

    data_style = styles["data"]
    row_number = first_data_row
    for entry in transactions:
        values = {
            "A": entry.asiento, "B": entry.fecha, "C": entry.cliente, "D": entry.descripcion,
            "E": entry.detalle, "F": entry.debito, "G": entry.credito, "H": entry.saldo,
        }
        cells = {col: _cell_xml(col, row_number, data_style.get(col), values[col], is_text=(col in _TEXT_COLUMNS))
                 for col in _COLUMNS}
        rows_map[row_number] = {"cells": cells, "height": None}
        row_number += 1

    last_data_row = row_number - 1
    totals_row = row_number
    total_debito = round(sum(r.debito or 0.0 for r in transactions), 2)
    total_credito = round(sum(r.credito or 0.0 for r in transactions), 2)

    totals_style = styles["totals"]
    totals_cells = {}
    for col in _COLUMNS:
        style_id = totals_style.get(col)
        if col == "F":
            totals_cells[col] = _formula_cell_xml(col, totals_row, style_id,
                                                   f"SUM(F{first_data_row}:F{last_data_row})", total_debito)
        elif col == "G":
            totals_cells[col] = _formula_cell_xml(col, totals_row, style_id,
                                                   f"SUM(G{first_data_row}:G{last_data_row})", total_credito)
        else:
            totals_cells[col] = _cell_xml(col, totals_row, style_id)
    rows_map[totals_row] = {"cells": totals_cells, "height": None}

    meta = {
        "title1_row": title1_row, "title2_row": title2_row, "header_row": header_row, "saldo_row": saldo_row,
        "first_data_row": first_data_row, "last_data_row": last_data_row, "totals_row": totals_row,
        "total_debito": total_debito, "total_credito": total_credito, "transactions": transactions,
    }
    return rows_map, meta


def _recolor_row(rows_map, row_number, entry, style_map):
    """Reescribe las 8 celdas A-H de una fila de transacción ya construida, con otro estilo por columna."""
    values = {
        "A": entry.asiento, "B": entry.fecha, "C": entry.cliente, "D": entry.descripcion,
        "E": entry.detalle, "F": entry.debito, "G": entry.credito, "H": entry.saldo,
    }
    cells = {col: _cell_xml(col, row_number, style_map.get(col), values[col], is_text=(col in _TEXT_COLUMNS))
             for col in _COLUMNS}
    rows_map[row_number]["cells"] = cells


def _replace_sheet_content(original_xml, sheet_data_inner_xml, merge_refs, dimension_ref):
    sheet_data_xml = "<sheetData>" + sheet_data_inner_xml + "</sheetData>"
    merge_cells_xml = (
        f'<mergeCells count="{len(merge_refs)}">'
        + "".join(f'<mergeCell ref="{ref}"/>' for ref in merge_refs)
        + "</mergeCells>"
    )
    new_xml = _DIMENSION_RE.sub(lambda _m: f'<dimension ref="{dimension_ref}"/>', original_xml, count=1)
    new_xml = _SHEET_DATA_RE.sub(lambda _m: sheet_data_xml, new_xml, count=1)
    if re.search(r"<mergeCells[^>]*>.*?</mergeCells>", new_xml, re.S):
        new_xml = re.sub(r"<mergeCells[^>]*>.*?</mergeCells>", lambda _m: merge_cells_xml, new_xml, count=1, flags=re.S)
    else:
        new_xml = new_xml.replace(sheet_data_xml, sheet_data_xml + merge_cells_xml, 1)
    return new_xml


# ---------------------------------------------------------------------------
# Reemplazo simple
# ---------------------------------------------------------------------------

def _build_simple_sheet(original_xml, sheet_name, mayor):
    styles = _load_style_templates(original_xml, sheet_name)
    rows_map, meta = _build_base_rows(mayor, styles, start_row=1)
    merge_refs = [f"A{meta['title1_row']}:H{meta['title1_row']}", f"A{meta['title2_row']}:H{meta['title2_row']}",
                  f"A{meta['totals_row']}:D{meta['totals_row']}"]
    dimension_ref = f"A1:H{meta['totals_row']}"
    return _replace_sheet_content(original_xml, _rows_map_to_xml(rows_map), merge_refs, dimension_ref)


# ---------------------------------------------------------------------------
# 2.J.H. WILLIAMS OIL -- saldo final + composición de facturas pendientes
# ---------------------------------------------------------------------------

# Un asiento de factura siempre trae su número real de 6 dígitos en algún
# lado del texto (ej. "INVOICE NRO.201808 ..." / "INVOICE NRO 1 - 199806"
# / "... - INVOICE 208063") -- confirmado contra julio-2026, donde
# coincide exacto con el número que después aparece en la cancelación.
_INVOICE_NUMBER_RE = re.compile(r"\b(\d{6})\b")
_CANCELACION_RE = re.compile(r"cancelacion\s+invoice\s+([\d\-]+)", re.IGNORECASE)
_MONTH_YEAR_RE = re.compile(r"(\d{2})-(\d{4})")


def _compute_jhw_pending(transactions):
    """
    Determina qué facturas (Credito, Descripcion con "INVOICE") del mes
    quedaron sin cancelar -- una factura está cancelada si su número de 6
    dígitos aparece en el texto de una fila posterior "... CANCELACION
    INVOICE {numero}[-{numero}...]" (confirmado contra julio-2026: pagos
    que cancelan varias facturas a la vez las separan con "-"). Devuelve
    una lista de (índice de la transacción, número de factura) en el
    mismo orden en que aparecen.
    """
    cancelled_numbers = set()
    for entry in transactions:
        desc = entry.descripcion or ""
        match = _CANCELACION_RE.search(desc)
        if match:
            cancelled_numbers.update(re.findall(r"\d{6}", match.group(1)))

    pending = []
    for idx, entry in enumerate(transactions):
        desc = entry.descripcion or ""
        if "cancelacion" in desc.lower():
            continue
        if not entry.credito or entry.credito <= 0:
            continue
        if "invoice" not in desc.lower():
            continue
        numbers = _INVOICE_NUMBER_RE.findall(desc)
        if len(numbers) != 1:
            # No se puede identificar el número de forma confiable -- se
            # deja afuera de la composición a propósito (el chequeo de
            # J8 deja de dar 0 y avisa que hace falta revisar a mano, en
            # vez de arriesgar sumar/perder una factura por una mala
            # lectura).
            continue
        if numbers[0] not in cancelled_numbers:
            pending.append((idx, numbers[0]))
    return pending


def _read_cell_text(workbook, sheet_name, row, col_letter):
    value = workbook[sheet_name].cell(row=row, column=_col_to_index(col_letter)).value
    return "" if value is None else str(value)


def _build_jhw_sheet(workbook, source_zip, sheet_paths, original_xml, mayor, caja_yellow_style):
    """
    Además del reemplazo A-H de siempre, completa la composición del
    saldo final (columnas J/K/L, filas "SALDO INICIAL"/"primer asiento"/
    "segundo asiento") y pinta de amarillo el Credito de cada factura
    pendiente -- pedido y confirmado por el usuario (2026-09-10):

    - J{saldo_row} (misma fila que SALDO INICIAL): etiqueta "COMPOSICION
      SALDO J.H WILLIAMS {mes-año actual}" (se toma el texto que ya
      estaba y solo se le reemplaza la fecha, igual que ya hace
      controles_valuacion.py con su propio rótulo de mes).
    - J{saldo_row+1}: `=+H{última fila}` (el saldo final real) + K con la
      etiqueta "SALDO A PAGAR" que ya traía la hoja.
    - Si ese saldo da 0, no se completa nada más (J{saldo_row+2} en
      blanco). Si no da 0: J{saldo_row+2} suma uno por uno el Credito de
      cada factura pendiente (`=+G16+G17+...`), K{saldo_row+2} dice
      "Pendientes {n1} - {n2}...", y J{saldo_row+5} (misma fila que ya
      usaba el archivo real para este chequeo) hace
      `=+J{saldo_row+1}-J{saldo_row+2}-J{saldo_row+3}` -- debería dar 0,
      confirmando que el saldo final son solo esas facturas y ninguna
      otra cosa.
    - Las celdas G de las facturas pendientes se pintan de amarillo,
      reusando el mismo estilo (fill + formato de moneda) que ya usa la
      hoja Caja para sus propios depósitos -- los índices de estilo son
      globales al libro, así que un estilo de Caja es válido acá también.
    """
    sheet_name = "2.J.H. WILLIAMS OIL"
    styles = _load_style_templates(original_xml, sheet_name)
    rows_map, meta = _build_base_rows(mayor, styles, start_row=1)

    saldo_row = meta["saldo_row"]
    final_row = saldo_row + 1
    pending_row = saldo_row + 2
    reserved_row = saldo_row + 3
    check_row = saldo_row + 4

    if check_row >= meta["totals_row"]:
        # El panel de composición (filas SALDO INICIAL a SALDO INICIAL+4)
        # se superpone con las primeras filas de asiento reales -- si el
        # mes tiene muy pocos movimientos, esas filas ni siquiera existen
        # todavía (o ya es la fila de totales) -- mejor avisar que
        # arriesgar pisar el total o crashear.
        raise ValueError(
            f'"{sheet_name}" tiene muy pocos movimientos este mes para completar el panel de composición '
            "del saldo (necesita al menos 4 asientos) -- completalo a mano."
        )

    label_style = _extract_row_styles(original_xml, saldo_row)
    final_style = _extract_row_styles(original_xml, final_row)
    pending_style = _extract_row_styles(original_xml, pending_row)
    check_style = _extract_row_styles(original_xml, check_row)
    for label, extracted in (("SALDO INICIAL", label_style), ("primer asiento", final_style),
                             ("segundo asiento", pending_style), ("chequeo", check_style)):
        if not extracted.get("J"):
            raise ValueError(
                f'No se encontró el panel de composición del saldo (columna J, fila de {label}) en '
                f'"{sheet_name}" -- revisar a mano.'
            )

    composicion_text = _read_cell_text(workbook, sheet_name, saldo_row, "J")
    saldo_a_pagar_text = _read_cell_text(workbook, sheet_name, final_row, "K")
    period_from = mayor["period_from"]
    composicion_text = _MONTH_YEAR_RE.sub(f"{period_from.month:02d}-{period_from.year}", composicion_text, count=1)

    rows_map[saldo_row]["cells"]["J"] = _cell_xml("J", saldo_row, label_style.get("J"), composicion_text, is_text=True)
    rows_map[saldo_row]["cells"]["K"] = _cell_xml("K", saldo_row, label_style.get("K"))
    rows_map[saldo_row]["cells"]["L"] = _cell_xml("L", saldo_row, label_style.get("L"))

    final_balance = round((meta["transactions"][-1].saldo or 0.0), 2)
    rows_map[final_row]["cells"]["J"] = _formula_cell_xml(
        "J", final_row, final_style.get("J"), f"+H{meta['last_data_row']}", final_balance
    )
    rows_map[final_row]["cells"]["K"] = _cell_xml("K", final_row, final_style.get("K"), saldo_a_pagar_text, is_text=True)
    rows_map[final_row]["cells"]["L"] = _cell_xml("L", final_row, final_style.get("L"))

    merge_refs = [f"A{meta['title1_row']}:H{meta['title1_row']}", f"A{meta['title2_row']}:H{meta['title2_row']}",
                  f"A{meta['totals_row']}:D{meta['totals_row']}",
                  f"J{saldo_row}:L{saldo_row}", f"K{final_row}:L{final_row}"]

    if final_balance != 0:
        transactions = meta["transactions"]
        pending = _compute_jhw_pending(transactions)
        if not pending:
            raise ValueError(
                f'El saldo final de "{sheet_name}" no da 0 pero no se pudo identificar ninguna factura '
                "pendiente por su número -- revisar a mano."
            )
        first_data_row = meta["first_data_row"]
        pending_rows = [first_data_row + idx for idx, _num in pending]
        formula = "+" + "+".join(f"G{r}" for r in pending_rows)
        pending_sum = round(sum(transactions[idx].credito for idx, _num in pending), 2)
        rows_map[pending_row]["cells"]["J"] = _formula_cell_xml(
            "J", pending_row, pending_style.get("J"), formula, pending_sum
        )
        pending_text = "Pendientes " + " - ".join(num for _idx, num in pending)
        rows_map[pending_row]["cells"]["K"] = _cell_xml(
            "K", pending_row, pending_style.get("K"), pending_text, is_text=True
        )
        rows_map[pending_row]["cells"]["L"] = _cell_xml("L", pending_row, pending_style.get("L"))
        merge_refs.append(f"K{pending_row}:L{pending_row}")

        rows_map[check_row]["cells"]["J"] = _formula_cell_xml(
            "J", check_row, check_style.get("J"), f"+J{final_row}-J{pending_row}-J{reserved_row}", 0.0
        )

        data_style = styles["data"]
        for idx, _num in pending:
            entry = transactions[idx]
            row_number = first_data_row + idx
            style_map = dict(data_style)
            style_map["G"] = caja_yellow_style
            _recolor_row(rows_map, row_number, entry, style_map)

    merge_refs.append(f"K{reserved_row}:L{reserved_row}")

    dimension_ref = f"A1:L{meta['totals_row']}"
    return _replace_sheet_content(original_xml, _rows_map_to_xml(rows_map), merge_refs, dimension_ref)


# ---------------------------------------------------------------------------
# Caja -- depósitos (amarillo) y gastos pagados en efectivo (celeste)
# ---------------------------------------------------------------------------

# Mismas reglas ya confirmadas y validadas por el usuario para el control
# de solo lectura equivalente (ver controles_caja.py) -- se reusan tal
# cual para no divergir de un criterio ya probado contra datos reales.
_DEPOSIT_RE = re.compile(r"deposit|transaccion", re.IGNORECASE)
_EXCLUDED_CAJA_PATTERNS = (
    "liquidacion cierre recaudacion",
    "liquidacion cierre lottery",
    "ajuste",
    "provision caja por liq pendiente de lottery",
)
_ANULACION_RE = re.compile(r"anulacion\s+asiento\s+(\d+)", re.IGNORECASE)

_CAJA_FOOTER_OFFSETS = {"ajuste": 2, "depositos": 4, "gastos": 6, "liquidacion": 8, "provision": 10}


def _normalize_asiento(value):
    if value is None:
        return None
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text or None


def _find_cancelled_asientos(transactions):
    """
    Números de Asiento que participan de un par "ANULACION ASIENTO N" --
    tanto la anulación en sí como el asiento original que cancela -- para
    excluirlos de depósitos/gastos (un movimiento anulado en el mismo mes
    no debe sumarse a ningún lado).
    """
    cancelled = set()
    for entry in transactions:
        desc = entry.descripcion or ""
        match = _ANULACION_RE.search(desc)
        if match:
            cancelled.add(match.group(1))
            asiento = _normalize_asiento(entry.asiento)
            if asiento:
                cancelled.add(asiento)
    return cancelled


def _categorize_caja_row(entry, cancelled_asientos):
    """Devuelve "deposito", "gasto", o None (no se pinta ni se suma)."""
    if _normalize_asiento(entry.asiento) in cancelled_asientos:
        return None
    if not entry.credito or entry.credito <= 0:
        return None
    desc = (entry.descripcion or "").lower()
    if any(pattern in desc for pattern in _EXCLUDED_CAJA_PATTERNS):
        return None
    if _DEPOSIT_RE.search(desc):
        return "deposito"
    return "gasto"


def _caja_rows_from_sheet(workbook, sheet_name, first_data_row, last_data_row):
    """Reconstruye la lista de MayorRow de la hoja Caja YA existente (antes de reemplazarla), para poder categorizar y encontrar una fila de ejemplo de cada color."""
    ws = workbook[sheet_name]
    rows = []
    for row in range(first_data_row, last_data_row + 1):
        rows.append(MayorRow(
            asiento=ws.cell(row=row, column=1).value,
            fecha=ws.cell(row=row, column=2).value,
            cliente=ws.cell(row=row, column=3).value,
            descripcion=ws.cell(row=row, column=4).value,
            detalle=ws.cell(row=row, column=5).value,
            debito=ws.cell(row=row, column=6).value,
            credito=ws.cell(row=row, column=7).value,
            saldo=ws.cell(row=row, column=8).value,
        ))
    return rows


def _find_caja_reference_styles(workbook, caja_xml, sheet_name="Caja"):
    """
    Busca, en el contenido ACTUAL de Caja (antes de reemplazarla), una
    fila de ejemplo ya categorizada como "depósito" y otra como "gasto"
    por el mismo criterio que se va a aplicar de acá en más, para reusar
    sus estilos (fill + formato de moneda) -- evita inventar un estilo
    nuevo, ya existen en `styles.xml` porque el propio archivo ya los usa.
    """
    old_totals_row = _find_single_totals_row(caja_xml, sheet_name)
    if old_totals_row is None:
        raise ValueError('No se encontró la fila de totales actual de "Caja" -- revisar a mano.')
    rows = _caja_rows_from_sheet(workbook, sheet_name, 5, old_totals_row - 1)
    cancelled = _find_cancelled_asientos(rows)

    deposito_row = gasto_row = None
    for idx, entry in enumerate(rows):
        category = _categorize_caja_row(entry, cancelled)
        row_number = 5 + idx
        if category == "deposito" and deposito_row is None:
            deposito_row = row_number
        elif category == "gasto" and gasto_row is None:
            gasto_row = row_number
        if deposito_row and gasto_row:
            break

    if deposito_row is None or gasto_row is None:
        raise ValueError(
            'No se encontró, en el contenido actual de "Caja", ningún ejemplo de fila de depósito/gasto '
            "para copiar el estilo -- revisar a mano."
        )
    return _extract_row_styles(caja_xml, deposito_row), _extract_row_styles(caja_xml, gasto_row)


def _read_caja_footer_labels(workbook, sheet_name, totals_row):
    return {
        key: _read_cell_text(workbook, sheet_name, totals_row + offset, "D")
        for key, offset in _CAJA_FOOTER_OFFSETS.items()
    }


def _find_plain_row(workbook, sheet_name, start_row, end_row):
    """
    Primera fila sin ningún relleno de color en ese rango -- usada como
    plantilla de estilo "plano" para las filas de asiento. No se puede
    asumir que la primera fila de datos sirve (en Caja, julio-2026, esa
    fila resulta ser justo una de las coloreadas -- ver
    `_build_caja_sheet`).
    """
    ws = workbook[sheet_name]
    for row in range(start_row, end_row + 1):
        fill = ws.cell(row=row, column=1).fill
        if fill is None or fill.patternType is None:
            return row
    return None


def _build_caja_sheet(workbook, source_zip, sheet_paths, original_xml, mayor):
    """
    Además del reemplazo A-H de siempre: identifica, en el Mayor recién
    subido, qué asientos son "depósito" (Descripcion con "DEPOSIT"/
    "TRANSACCION") y cuáles son "gasto pagado por Caja" (Credito > 0, no
    depósito, no una liquidación/ajuste de fin de mes -- mismas reglas ya
    confirmadas y usadas en controles_caja.py), los pinta de amarillo/
    celeste respectivamente (A a H) y completa DEPOSITOS EFECTIVO EN
    CHASE / GASTOS PAGADOS X CAJA con la suma. Un asiento que se anula a
    sí mismo en el mes ("ANULACION ASIENTO N") no se pinta ni se suma en
    ningún lado (pedido explícito del usuario). AJUSTE DE SALDOS /
    LIQUIDACION CIERRE LOTTERY / PROVISION CAJA quedan con su misma
    etiqueta de siempre pero el valor en blanco -- eso lo sigue
    completando el usuario a mano, no está automatizado.
    """
    sheet_name = "Caja"
    old_totals_row = _find_single_totals_row(original_xml, sheet_name)
    plain_row = _find_plain_row(workbook, sheet_name, 5, old_totals_row - 1)
    if plain_row is None:
        raise ValueError(
            'No se encontró ninguna fila sin pintar en "Caja" para usar de plantilla de estilo -- revisar a mano.'
        )
    styles = _load_style_templates(original_xml, sheet_name, data_row=plain_row)
    footer_labels = _read_caja_footer_labels(workbook, sheet_name, old_totals_row)
    spacer_style = _extract_row_styles(original_xml, old_totals_row + 1)
    footer_styles = {key: _extract_row_styles(original_xml, old_totals_row + offset)
                     for key, offset in _CAJA_FOOTER_OFFSETS.items()}
    check_styles = _extract_row_styles(original_xml, old_totals_row + 12)
    for key, extracted in footer_styles.items():
        if not extracted:
            raise ValueError(f'No se encontró el pie de página de Caja (bloque "{key}") -- revisar a mano.')
    if not spacer_style or not check_styles:
        raise ValueError('No se encontró el pie de página de Caja completo -- revisar a mano.')

    deposito_style, gasto_style = _find_caja_reference_styles(workbook, original_xml, sheet_name)

    rows_map, meta = _build_base_rows(mayor, styles, start_row=1)
    transactions = meta["transactions"]
    first_data_row = meta["first_data_row"]

    cancelled = _find_cancelled_asientos(transactions)
    deposito_total = 0.0
    gasto_total = 0.0
    for idx, entry in enumerate(transactions):
        category = _categorize_caja_row(entry, cancelled)
        row_number = first_data_row + idx
        if category == "deposito":
            _recolor_row(rows_map, row_number, entry, deposito_style)
            deposito_total += entry.credito
        elif category == "gasto":
            _recolor_row(rows_map, row_number, entry, gasto_style)
            gasto_total += entry.credito
    deposito_total = round(deposito_total, 2)
    gasto_total = round(gasto_total, 2)

    totals_row = meta["totals_row"]
    merge_refs = [f"A{meta['title1_row']}:H{meta['title1_row']}", f"A{meta['title2_row']}:H{meta['title2_row']}",
                  f"A{totals_row}:D{totals_row}"]

    spacer_row = totals_row + 1
    rows_map[spacer_row] = {
        "cells": {col: _cell_xml(col, spacer_row, spacer_style.get(col)) for col in _COLUMNS}, "height": None,
    }
    merge_refs.append(f"A{spacer_row}:D{spacer_row}")

    computed_values = {"depositos": deposito_total, "gastos": gasto_total}
    for key, offset in _CAJA_FOOTER_OFFSETS.items():
        row = totals_row + offset
        block_styles = footer_styles[key]
        cells = {"D": _cell_xml("D", row, block_styles.get("D"), footer_labels[key], is_text=True)}
        value = computed_values.get(key)
        cells["F"] = _cell_xml("F", row, block_styles.get("F"), value)
        for extra_col in ("G", "H"):
            extra_style = block_styles.get(extra_col)
            if extra_style is not None:
                cells[extra_col] = _cell_xml(extra_col, row, extra_style)
        rows_map[row] = {"cells": cells, "height": None}

    check_row = totals_row + 12
    check_formula = (
        f"+F{totals_row + 2}+F{totals_row + 4}+F{totals_row + 6}+F{totals_row + 8}+F{totals_row + 10}"
        f"-G{totals_row}"
    )
    check_cached = round(gasto_total + deposito_total - meta["total_credito"], 2)
    rows_map[check_row] = {
        "cells": {
            "E": _cell_xml("E", check_row, check_styles.get("E")),
            "F": _formula_cell_xml("F", check_row, check_styles.get("F"), check_formula, check_cached),
            "G": _cell_xml("G", check_row, check_styles.get("G")),
        },
        "height": None,
    }

    dimension_ref = f"A1:H{check_row}"
    return _replace_sheet_content(original_xml, _rows_map_to_xml(rows_map), merge_refs, dimension_ref)


# ---------------------------------------------------------------------------
# Otros Gtos / Reparaciones -- acumulado anual, nunca se reemplaza
# ---------------------------------------------------------------------------

def _find_last_block_title_row(workbook, sheet_name):
    """Fila del ÚLTIMO título "Cuenta: ..." ya cargado en una hoja acumulativa."""
    ws = workbook[sheet_name]
    last_row = None
    for row in range(1, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value
        if isinstance(value, str) and value.strip().lower().startswith("cuenta:"):
            last_row = row
    if last_row is None:
        raise ValueError(f'No se encontró ningún bloque "Cuenta: ..." ya cargado en "{sheet_name}".')
    return last_row


def _build_reparaciones_footer(rows_map, meta, styles_source_xml, last_block_totals_row):
    """
    Agrega, sin autocompletar la suma (pedido explícito del usuario: "no
    hace falta que autocompletes esa suma ya que sería mucho trabajo"),
    la plantilla en blanco de GASTOS PUNTUALES/GASTOS HABITUALES debajo
    de la fila de totales del bloque nuevo -- mismo formato/diseño (F con
    el resaltado propio de cada categoría, G con la etiqueta) que ya usan
    los meses anteriores.
    """
    totals_row = meta["totals_row"]
    puntuales_row = totals_row + 2
    habituales_row = totals_row + 3
    puntuales_style = _extract_row_styles(styles_source_xml, last_block_totals_row + 2)
    habituales_style = _extract_row_styles(styles_source_xml, last_block_totals_row + 3)
    if not puntuales_style.get("F") or not habituales_style.get("F"):
        raise ValueError('No se encontró la plantilla "GASTOS PUNTUALES"/"GASTOS HABITUALES" -- revisar a mano.')

    rows_map[puntuales_row] = {
        "cells": {
            "F": _cell_xml("F", puntuales_row, puntuales_style.get("F")),
            "G": _cell_xml("G", puntuales_row, puntuales_style.get("G"), "GASTOS PUNTUALES", is_text=True),
        },
        "height": None,
    }
    rows_map[habituales_row] = {
        "cells": {
            "F": _cell_xml("F", habituales_row, habituales_style.get("F")),
            "G": _cell_xml("G", habituales_row, habituales_style.get("G"), "GASTOS HABITUALES", is_text=True),
        },
        "height": None,
    }
    return habituales_row


def _splice_append_xml(original_xml, new_rows_xml, new_merge_refs, new_dimension_ref):
    new_xml = original_xml.replace("</sheetData>", new_rows_xml + "</sheetData>", 1)

    def _merge_repl(match):
        existing_count = int(match.group(1))
        existing_cells = match.group(2)
        added = "".join(f'<mergeCell ref="{ref}"/>' for ref in new_merge_refs)
        return f'<mergeCells count="{existing_count + len(new_merge_refs)}">{existing_cells}{added}</mergeCells>'

    new_xml = _MERGE_CELLS_BLOCK_RE.sub(_merge_repl, new_xml, count=1)
    new_xml = _DIMENSION_RE.sub(lambda _m: f'<dimension ref="{new_dimension_ref}"/>', new_xml, count=1)
    return new_xml


def _build_append_sheet(workbook, original_xml, sheet_name, mayor):
    """
    A diferencia de un reemplazo, acá puede haber varios bloques
    mensuales ya cargados (una fila de totales por mes) -- se ubica el
    ÚLTIMO (el de mayor número de fila) y el bloque nuevo se pega
    `APPEND_GAP_ROWS[cuenta]` filas más abajo, sin tocar ni una sola fila
    de los meses anteriores.
    """
    account_code = mayor["account_code"]
    gap_rows = APPEND_GAP_ROWS[account_code]
    last_title_row = _find_last_block_title_row(workbook, sheet_name)
    candidates = _find_totals_rows(original_xml)
    if not candidates:
        raise ValueError(f'No se encontró ninguna fila de totales ya cargada en "{sheet_name}".')
    last_totals_row = candidates[-1]

    new_start_row = last_totals_row + gap_rows + 1
    styles = _load_style_templates(original_xml, sheet_name, title_row=last_title_row, totals_row=last_totals_row)
    rows_map, meta = _build_base_rows(mayor, styles, start_row=new_start_row)

    merge_refs = [f"A{meta['title1_row']}:H{meta['title1_row']}", f"A{meta['title2_row']}:H{meta['title2_row']}",
                  f"A{meta['totals_row']}:D{meta['totals_row']}"]
    dimension_extra_row = meta["totals_row"]

    if account_code == "5.10.00.00":  # Reparaciones -- plantilla sin autocompletar (ver docstring)
        dimension_extra_row = _build_reparaciones_footer(rows_map, meta, original_xml, last_totals_row)

    new_rows_xml = _rows_map_to_xml(rows_map)
    dimension_ref = f"A1:H{dimension_extra_row}"
    return _splice_append_xml(original_xml, new_rows_xml, merge_refs, dimension_ref)


# ---------------------------------------------------------------------------
# Orquestador
# ---------------------------------------------------------------------------

def replace_mayor_sheets(balance_path, mayor_paths):
    """
    Actualiza, en el Excel de Balance, cada hoja cuyo Mayor se subió --
    reemplazo simple/especial o append, según la cuenta (ver
    SIMPLE_REPLACE_ACCOUNTS/SPECIAL_REPLACE_ACCOUNTS/APPEND_ACCOUNTS).

    mayor_paths: lista de rutas a archivos de Mayor (.xls/.xlsx), uno por
    cuenta -- se identifica automáticamente la cuenta de cada uno por su
    propia línea "Cuenta: ...", no hace falta indicar a mano a qué hoja
    va cada archivo.

    Nunca toca ninguna otra hoja ni ninguna otra parte del archivo
    (estilos, gráficos, comentarios, conexiones, configuración de
    impresión, dibujos) -- ver el docstring del módulo para el porqué.

    Devuelve (temp_path, summary). Si ningún Mayor matchea una cuenta
    soportada, temp_path es None. summary tiene "updated" (nombres de
    hoja actualizados), "unsupported" (cuentas conocidas pero sin lógica
    todavía) y "unmatched" (Mayores cuya cuenta no se reconoce).
    """
    balance_path = os.path.abspath(str(balance_path).strip())
    if not os.path.isfile(balance_path):
        raise FileNotFoundError(f"Excel de Balance no encontrado: {balance_path}")

    simple_batch = []
    special_batch = []
    append_batch = []
    unsupported = []
    unmatched = []
    for mayor_path in mayor_paths:
        mayor_path = os.path.abspath(str(mayor_path).strip())
        data = _read_mayor_full(mayor_path)
        code = data["account_code"]
        if code in SIMPLE_REPLACE_ACCOUNTS:
            data["sheet_name"] = SIMPLE_REPLACE_ACCOUNTS[code]
            simple_batch.append(data)
        elif code in SPECIAL_REPLACE_ACCOUNTS:
            data["sheet_name"] = SPECIAL_REPLACE_ACCOUNTS[code]
            special_batch.append(data)
        elif code in APPEND_ACCOUNTS:
            data["sheet_name"] = APPEND_ACCOUNTS[code]
            append_batch.append(data)
        elif code in KNOWN_UNSUPPORTED_ACCOUNTS:
            unsupported.append(KNOWN_UNSUPPORTED_ACCOUNTS[code])
        else:
            unmatched.append(data.get("account_name") or code)

    all_parsed = simple_batch + special_batch + append_batch
    if not all_parsed:
        return None, {"updated": [], "unsupported": unsupported, "unmatched": unmatched}

    updated_sheets = []
    # Solo se abre con openpyxl (nunca se guarda) si hace falta leer texto
    # existente (2.J.H. WILLIAMS OIL/Caja/hojas acumulativas) -- las
    # cuentas de reemplazo simple no lo necesitan.
    needs_workbook = bool(special_batch or append_batch)
    workbook = load_workbook(balance_path, data_only=False) if needs_workbook else None
    try:
        with zipfile.ZipFile(balance_path, "r") as source_zip:
            sheet_paths = _map_sheet_paths(source_zip)
            replacements = {}

            # 2.J.H. WILLIAMS OIL pinta sus facturas pendientes con el
            # mismo estilo (fill amarillo + formato de moneda) que ya usa
            # Caja para sus propios depósitos -- los índices de estilo
            # son globales al libro, así que se puede leer directo del
            # contenido ACTUAL de Caja (esté o no Caja en este mismo
            # lote; si también se está reemplazando, el estilo no cambia,
            # solo cambian los valores).
            caja_yellow_style = None
            if any(d["sheet_name"] == "2.J.H. WILLIAMS OIL" for d in special_batch):
                caja_path = sheet_paths.get("Caja")
                if caja_path is None:
                    raise ValueError(
                        'No se encontró la hoja "Caja" en el Balance -- hace falta para copiar el estilo '
                        "de amarillo de sus depósitos."
                    )
                caja_current_xml = source_zip.read(caja_path).decode("utf-8")
                deposito_style, _gasto_style = _find_caja_reference_styles(workbook, caja_current_xml, "Caja")
                caja_yellow_style = deposito_style.get("G")

            for data in simple_batch:
                sheet_name = data["sheet_name"]
                xml_path = sheet_paths.get(sheet_name)
                if xml_path is None:
                    unmatched.append(data.get("account_name") or data["account_code"])
                    continue
                original_xml = source_zip.read(xml_path).decode("utf-8")
                replacements[xml_path] = _build_simple_sheet(original_xml, sheet_name, data).encode("utf-8")
                updated_sheets.append(sheet_name.strip())

            for data in special_batch:
                sheet_name = data["sheet_name"]
                xml_path = sheet_paths.get(sheet_name)
                if xml_path is None:
                    unmatched.append(data.get("account_name") or data["account_code"])
                    continue
                original_xml = source_zip.read(xml_path).decode("utf-8")
                if sheet_name == "Caja":
                    new_xml = _build_caja_sheet(workbook, source_zip, sheet_paths, original_xml, data)
                else:  # 2.J.H. WILLIAMS OIL
                    new_xml = _build_jhw_sheet(workbook, source_zip, sheet_paths, original_xml, data, caja_yellow_style)
                replacements[xml_path] = new_xml.encode("utf-8")
                updated_sheets.append(sheet_name.strip())

            for data in append_batch:
                sheet_name = data["sheet_name"]
                xml_path = sheet_paths.get(sheet_name)
                if xml_path is None:
                    unmatched.append(data.get("account_name") or data["account_code"])
                    continue
                original_xml = source_zip.read(xml_path).decode("utf-8")
                replacements[xml_path] = _build_append_sheet(workbook, original_xml, sheet_name, data).encode("utf-8")
                updated_sheets.append(sheet_name.strip())

            fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
            os.close(fd)
            try:
                with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as out_zip:
                    for item in source_zip.infolist():
                        payload = replacements.get(item.filename)
                        if payload is None:
                            payload = source_zip.read(item.filename)
                        out_zip.writestr(item, payload)
            except Exception:
                os.remove(temp_path)
                raise
    finally:
        if workbook is not None:
            workbook.close()

    summary = {"updated": updated_sheets, "unsupported": unsupported, "unmatched": unmatched}
    return temp_path, summary
