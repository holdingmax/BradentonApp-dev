"""
Generador de PDF básico y reusable para exportar tablas -- pedido
explícito del usuario (2026-09-19): "solo necesitaria que salieran los
datos limpios en un PDF basico, no hacen falta los colores, solo con
poner lineas y bordes se entenderia". Sin dependencias de sistema
operativo -- reportlab es puro Python, corre igual en Windows y en el
Dockerfile de Render (ver CLAUDE.md, "Cross-platform by default").

A diferencia de los export a Excel (build_store_info_export_workbook/
build_caja_export_workbook, que replican colores/bordes/anchos exactos de
la hoja real del Cierre), este PDF es deliberadamente más simple: mismas
columnas que ya se ven en pantalla, sin colores, solo líneas de grilla.

Membrete (logo + nombre de la empresa) -- agregado 2026-09-17, pedido
explícito del usuario para el módulo nuevo "Reportes" (Chase/EFT/Reporte
Diario/Lottery): "Incluye en esos PDF el Logo de Brandeton y al lado el
nombre oficial de la empresa con un formato de texto decente, todo bien
encuadrado y con los bordes que se necesiten para que se vean prolijos".
`build_simple_table_pdf` ganó dos parámetros OPCIONALES (`company_header`/
`period_label`) para esto -- default `False`/`None`, así que los tres
exports que ya usaban esta función (Lottery/Store Info/Caja) siguen
saliendo pixel-a-pixel iguales que antes, sin membrete, a menos que un
caller nuevo lo pida.
"""

import os

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, legal
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

COMPANY_NAME = "Bradenton Gas Station USA LLC"
_LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "logo.png")


def _letterhead_table(available_width):
    """
    Membrete compartido por todos los PDF de "Reportes": logo a la
    izquierda + nombre oficial de la empresa al lado, en una sola fila con
    su propio recuadro -- para que quede "bien encuadrado" y se distinga
    claramente del resto del documento (pedido explícito del usuario).
    Devuelve una lista de flowables (la tabla + un espacio abajo) o una
    lista vacía si el logo no está disponible (nunca rompe el PDF por
    esto -- lo demás se genera igual, con o sin logo).
    """
    styles = getSampleStyleSheet()
    name_style = ParagraphStyle(
        "CompanyName",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
    )
    if os.path.isfile(_LOGO_PATH):
        logo_w = 20 * mm
        logo_h = logo_w * 809 / 914  # aspect ratio real del PNG (914x809)
        logo = Image(_LOGO_PATH, width=logo_w, height=logo_h)
        header = Table(
            [[logo, Paragraph(COMPANY_NAME, name_style)]],
            colWidths=[logo_w + 6 * mm, available_width - logo_w - 6 * mm],
    )
    else:
        header = Table([[Paragraph(COMPANY_NAME, name_style)]], colWidths=[available_width])
    header.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.8, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (0, 0), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return [header, Spacer(1, 12)]


def build_simple_table_pdf(
    dest_path,
    title,
    headers,
    rows,
    col_widths_mm=None,
    header_fill_by_col=None,
    data_fill_by_col=None,
    bold_last_row=False,
    font_size=8.5,
    cell_padding=5,
    company_header=False,
    period_label=None,
    footer_note=None,
):
    """
    `headers`: una fila de encabezado (texto), en negrita.
    `rows`: lista de listas ya como texto/número final para mostrar --
    esta función no calcula nada, solo dibuja lo que le pasen.
    `col_widths_mm` opcional -- si no se pasa, reparte el ancho disponible
    en partes iguales entre todas las columnas.

    `header_fill_by_col`/`data_fill_by_col` opcionales -- pedido explícito
    del usuario (2026-09-19): "si queriamos ponerle colores" -- dict
    {índice de columna (0-based): color hex "#RRGGBB"}, mismos colores que
    ya usan los export a Excel (build_store_info_export_workbook/
    build_caja_export_workbook) para que ambos formatos se vean iguales.
    `header_fill_by_col` pinta la fila de encabezado; `data_fill_by_col`
    pinta esa columna en todas las filas de datos (nunca el encabezado).

    `bold_last_row` -- pedido explícito del usuario (2026-09-19): "en los
    totales de los PDF... esten marcados en negrita" -- cuando `rows`
    termina en una fila de totales, poné esto en True para que salga en
    negrita igual que el encabezado.

    `font_size`/`cell_padding` opcionales -- por default quedan igual que
    siempre (8.5pt / 5pt), para no afectar a los PDF ya ajustados (Store
    Info/Caja). Un caller con muchas columnas (ej. Lottery) puede pasar un
    valor más grande si lo necesita -- pedido explícito del usuario
    (2026-09-16): "se ve medio apretado".

    `company_header`/`period_label` -- opcionales, default `False`/`None`
    (sin cambios para los callers existentes). `company_header=True` agrega
    el membrete (logo + nombre de la empresa, ver `_letterhead_table`)
    arriba del título; `period_label` agrega una línea debajo del título
    con el rango de fechas real que cubre el reporte (ej. "Período:
    01/08/2026 al 20/08/2026") -- pedido explícito del usuario para los PDF
    de "Reportes": "hay que aclarar hasta que dia llega el reporte".

    `footer_note` -- opcional, default `None` (sin cambios para los
    callers existentes). Un párrafo de texto libre DEBAJO de la tabla, para
    avisar algo que la tabla sola no puede explicar (ej. "N bloques sin
    fecha de Chase Bank confirmada, el total no los incluye" en el resumen
    de Lottery) -- nunca reemplaza la tabla, solo la complementa.
    """
    doc = SimpleDocTemplate(
        dest_path,
        pagesize=landscape(legal),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=title,
    )
    styles = getSampleStyleSheet()
    elements = []
    available_width = doc.width
    if company_header:
        elements.extend(_letterhead_table(available_width))
    elements.append(Paragraph(title, styles["Heading2"]))
    if period_label:
        elements.append(Paragraph(period_label, styles["Normal"]))
    elements.append(Spacer(1, 10))

    if col_widths_mm is None:
        col_widths = [available_width / len(headers)] * len(headers)
    else:
        col_widths = [w * mm for w in col_widths_mm]

    table = Table([headers] + rows, colWidths=col_widths, repeatRows=1)
    # Tamaño general subido -- pedido explícito del usuario (2026-09-19):
    # "quiero que a todo lo hagas un poco mas grande de tamano".
    style_commands = [
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), cell_padding),
        ("BOTTOMPADDING", (0, 0), (-1, -1), cell_padding),
    ]
    last_row = len(rows)
    for col, hex_color in (header_fill_by_col or {}).items():
        style_commands.append(("BACKGROUND", (col, 0), (col, 0), colors.HexColor(hex_color)))
    for col, hex_color in (data_fill_by_col or {}).items():
        if last_row >= 1:
            style_commands.append(("BACKGROUND", (col, 1), (col, last_row), colors.HexColor(hex_color)))
    if bold_last_row and last_row >= 1:
        style_commands.append(("FONTNAME", (0, last_row), (-1, last_row), "Helvetica-Bold"))
    table.setStyle(TableStyle(style_commands))
    elements.append(table)
    if footer_note:
        elements.append(Spacer(1, 8))
        elements.append(Paragraph(footer_note, styles["Normal"]))
    doc.build(elements)
    return dest_path


def _section_table(headers, rows, available_width, col_widths_mm, font_size, cell_padding, bold_last_row):
    if col_widths_mm is None:
        col_widths = [available_width / len(headers)] * len(headers)
    else:
        col_widths = [w * mm for w in col_widths_mm]
    table = Table([headers] + rows, colWidths=col_widths, repeatRows=1)
    style_commands = [
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), cell_padding),
        ("BOTTOMPADDING", (0, 0), (-1, -1), cell_padding),
    ]
    last_row = len(rows)
    if bold_last_row and last_row >= 1:
        style_commands.append(("FONTNAME", (0, last_row), (-1, last_row), "Helvetica-Bold"))
    table.setStyle(TableStyle(style_commands))
    return table


def build_multi_section_pdf(dest_path, title, sections, period_label=None, company_header=True):
    """
    PDF con VARIAS tablas (o notas) en el mismo documento, cada una con su
    propio subtítulo -- pedido explícito del usuario para el Reporte de
    EFT (2026-09-17): el mismo PDF tiene que traer los EFT del mes Y,
    debajo, un resumen de los cupones acumulados -- dos secciones, no una
    sola tabla como `build_simple_table_pdf`. Mismo membrete/título/período
    que esa función (ver ahí el detalle) -- acá `sections` es una lista de
    bloques, cada uno un dict con:
      - `heading` (texto, subtítulo en negrita de esa sección)
      - `headers`/`rows`/`col_widths_mm`/`bold_last_row` -- una tabla, con
        el mismo significado que en `build_simple_table_pdf` (col_widths_mm
        opcional, reparte el ancho en partes iguales si falta)
      - `note` -- alternativa a `headers`/`rows`: un párrafo de texto libre
        en vez de una tabla completa, para un resumen corto (ej. "Cupones
        cargados: 1157 -- Acumulado: $123,456.78") sin la formalidad de
        una tabla de una sola fila.
    Landscape/legal, mismos márgenes que `build_simple_table_pdf` -- pensado
    para que ambas funciones generen documentos que se vean "de la misma
    familia".
    """
    doc = SimpleDocTemplate(
        dest_path,
        pagesize=landscape(legal),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=title,
    )
    styles = getSampleStyleSheet()
    elements = []
    available_width = doc.width
    if company_header:
        elements.extend(_letterhead_table(available_width))
    elements.append(Paragraph(title, styles["Heading2"]))
    if period_label:
        elements.append(Paragraph(period_label, styles["Normal"]))
    elements.append(Spacer(1, 10))

    for section in sections:
        heading = section.get("heading")
        if heading:
            elements.append(Paragraph(heading, styles["Heading3"]))
            elements.append(Spacer(1, 4))
        if section.get("note"):
            elements.append(Paragraph(section["note"], styles["Normal"]))
        else:
            elements.append(
                _section_table(
                    section["headers"],
                    section["rows"],
                    available_width,
                    section.get("col_widths_mm"),
                    section.get("font_size", 8.5),
                    section.get("cell_padding", 5),
                    section.get("bold_last_row", False),
                )
            )
        elements.append(Spacer(1, 14))

    doc.build(elements)
    return dest_path
