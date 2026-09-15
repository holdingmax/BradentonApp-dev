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
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, legal
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


def build_simple_table_pdf(
    dest_path,
    title,
    headers,
    rows,
    col_widths_mm=None,
    header_fill_by_col=None,
    data_fill_by_col=None,
    bold_last_row=False,
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
    elements = [Paragraph(title, styles["Heading2"]), Spacer(1, 10)]

    available_width = doc.width
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
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
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
    doc.build(elements)
    return dest_path
