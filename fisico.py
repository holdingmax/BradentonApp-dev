"""
Módulo "Físico" -- reconciliación mensual de combustible: Inventario
Teórico (lo que DEBERÍA quedar en los tanques según compras y ventas)
contra la lectura física real (lo que un gauge/medición mide de verdad).
Reproduce la lógica real de la hoja "Fisico" del Excel de cierre --
decodificada con openpyxl leyendo un ejemplo real (`hoja_fisico.xlsx`),
no adivinada.

Fórmula real por mes (columnas Gal/$'):
    Inventario Inicial Teórico   = Inventario Final Teórico del mes anterior
                                    (encadenado, mismo criterio que
                                    caja._resolve_opening_balance -- un
                                    override manual para ESTE mes manda
                                    siempre; si no hay, se encadena UN
                                    solo mes atrás, nunca más)
    Total Compras                = Inventario Inicial + SUM(facturas del mes)
    Precio promedio ($/gal)      = Total Compras $ / Total Compras Gal
    Ventas (gal)                 = -Volume del mes (Store Info, mismo dato
                                    que ya usa Proyecciones) -- negativo
                                    porque es una salida de combustible
    Ventas ($)                   = Ventas (gal) * Precio promedio
                                    (valuado al COSTO de compra, no al
                                    precio de venta -- así lo hace la
                                    hoja real: `=+D12*F10`)
    Inventario Final Teórico     = Total Compras + Ventas (gal y $)
    Inventario Final Real        = lectura física de los tanques, SIEMPRE
                                    un dato externo cargado a mano (ver
                                    fisico_db.set_month_real_ending) --
                                    "pendiente" hasta que se cargue (la
                                    hoja real la toma recién el día 1 del
                                    mes siguiente, así que un mes recién
                                    cerrado normalmente todavía no la
                                    tiene)
    Diferencia (gal)             = Inventario Final Teórico - Inventario
                                    Final Real (positivo = falta
                                    combustible/merma; negativo = sobra)
    Diferencia ($)                = Diferencia (gal) * Precio promedio
    Diferencia (%)                 = Diferencia (gal) / Inventario Final
                                    Teórico (gal)

Cada factura puede traer el galonaje partido en dos productos que se
suman (ej. `=7600+1199` en la hoja real, típicamente Regular + Diesel) --
`fisico_db.add_invoice` ya recibe un solo número de galones, así que esa
suma (si aplica) se hace en el formulario de carga (ver carga_datos_
combustible en webapp.py), replicando el mismo criterio.
"""

import calendar
from datetime import date, datetime

import fisico_db
import reportes_db


def _month_volume_actual(year, month):
    """Suma de Volume (galones vendidos) del mes, directo de Store Info -- mismo dato que proyecciones.build_projection."""
    rows = reportes_db.get_month_store_info(year, month)
    return sum(row.get("volume") or 0.0 for row in rows if row.get("volume") is not None)


def _resolve_initial(year, month, _chain=True):
    """
    (gallons, amount, source) del Inventario Inicial Teórico -- mismo
    patrón que caja._resolve_opening_balance: un override manual de ESTE
    mes manda siempre; si no hay y `_chain` es True, se encadena UN mes
    atrás (su Inventario Final Teórico, calculado); si tampoco hay nada
    ahí, 0.0/0.0 (mes realmente inicial, sin nada cargado antes).
    """
    settings = fisico_db.get_month_settings(year, month)
    if settings and settings.get("initial_gallons_override") is not None:
        return settings["initial_gallons_override"], settings.get("initial_amount_override") or 0.0, "manual"
    if not _chain:
        return 0.0, 0.0, "default"

    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    prev_report = build_month_report(prev_year, prev_month, _chain=False)
    if prev_report["has_any_data"]:
        return (
            prev_report["final_teorico_gallons"],
            prev_report["final_teorico_amount"],
            "prev_month_computed",
        )
    return 0.0, 0.0, "default"


def build_month_report(year, month, _chain=True):
    invoices = fisico_db.get_month_invoices(year, month)
    settings = fisico_db.get_month_settings(year, month)

    initial_gallons, initial_amount, initial_source = _resolve_initial(year, month, _chain=_chain)

    invoices_gallons = round(sum(inv["gallons"] or 0.0 for inv in invoices), 2)
    invoices_amount = round(sum(inv["amount"] or 0.0 for inv in invoices), 2)

    total_gallons = round(initial_gallons + invoices_gallons, 2)
    total_amount = round(initial_amount + invoices_amount, 2)
    avg_price = (total_amount / total_gallons) if total_gallons else None

    ventas_gallons = round(-_month_volume_actual(year, month), 2)
    ventas_amount = round(ventas_gallons * avg_price, 2) if avg_price is not None else None

    final_teorico_gallons = round(total_gallons + ventas_gallons, 2)
    final_teorico_amount = (
        round(total_amount + ventas_amount, 2) if ventas_amount is not None else total_amount
    )

    real_gallons = settings.get("real_ending_gallons") if settings else None
    real_reading_date = settings.get("real_reading_date") if settings else None

    if real_gallons is not None:
        diferencia_gallons = round(final_teorico_gallons - real_gallons, 2)
        diferencia_amount = round(diferencia_gallons * avg_price, 2) if avg_price is not None else None
        diferencia_pct = (
            (diferencia_gallons / final_teorico_gallons) if final_teorico_gallons else None
        )
    else:
        diferencia_gallons = None
        diferencia_amount = None
        diferencia_pct = None

    has_any_data = bool(invoices) or (settings is not None) or initial_gallons or initial_amount

    return {
        "year": year,
        "month": month,
        "has_any_data": has_any_data,
        "invoices": invoices,
        "initial_gallons": initial_gallons,
        "initial_amount": round(initial_amount, 2),
        "initial_source": initial_source,
        "invoices_gallons": invoices_gallons,
        "invoices_amount": invoices_amount,
        "total_gallons": total_gallons,
        "total_amount": total_amount,
        "avg_price": round(avg_price, 2) if avg_price is not None else None,
        "ventas_gallons": ventas_gallons,
        "ventas_amount": ventas_amount,
        "final_teorico_gallons": final_teorico_gallons,
        "final_teorico_amount": final_teorico_amount,
        "real_gallons": real_gallons,
        "real_reading_date": real_reading_date,
        "diferencia_gallons": diferencia_gallons,
        "diferencia_amount": diferencia_amount,
        "diferencia_pct": diferencia_pct,
    }


def get_available_years():
    years = set(fisico_db.get_invoice_years()) | set(fisico_db.get_settings_years())
    return sorted(years)


# ---------------------------------------------------------------------------
# Exportar a Excel y PDF -- pedido explícito del usuario (2026-09-22):
# "quiero que haya un boton para exportar a excel y pdf como tienen los
# demas modulos" -- mismo botón "Exportar" (popover Excel/PDF) que ya usa
# Lottery, con el mismo resumen que ya muestra /fisico en pantalla. A
# diferencia de Caja/Store Info/Lottery, Físico no replica ninguna hoja de
# un Excel real (es un módulo nuevo, sin plantilla original que calcar) --
# el Excel usa el mismo estilo simple ya usado en otros reportes derivados
# (ver gettel_reportes._write_days_sheet) en vez de intentar clonar colores
# de un archivo que no existe.
# ---------------------------------------------------------------------------

def _fmt_gal(value):
    return f"{value:,.2f}" if value is not None else "—"


def _fmt_money(value):
    return f"${value:,.2f}" if value is not None else "—"


def _fmt_date_ddmmyyyy(value):
    """`value` como lo guarda fisico_db (texto "YYYY-MM-DD")."""
    if not value:
        return "—"
    parts = str(value).split("-")
    if len(parts) != 3:
        return str(value)
    year, month, day = parts
    return f"{day}/{month}/{year}"


def _parse_date_for_excel(value):
    """Convierte "YYYY-MM-DD" a un date real para que la celda de Excel quede como fecha, no como texto."""
    if not value:
        return None
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").date()
    except ValueError:
        return str(value)


def _resumen_rows(report):
    invoices = report["invoices"]
    initial_label = "Inventario Inicial" + (
        " (manual)" if report["initial_source"] == "manual"
        else " (encadenado)" if report["initial_source"] == "prev_month_computed"
        else " (sin datos)"
    )
    return [
        (initial_label, report["initial_gallons"], report["initial_amount"]),
        (f"+ Compras del mes ({len(invoices)} factura{'s' if len(invoices) != 1 else ''})",
         report["invoices_gallons"], report["invoices_amount"]),
        ("= Total Compras", report["total_gallons"], report["total_amount"]),
        ("Ventas del mes", report["ventas_gallons"], report["ventas_amount"]),
        ("Inventario Final Teórico", report["final_teorico_gallons"], report["final_teorico_amount"]),
    ]


def build_fisico_export_workbook(report, year, month, dest_path):
    """Excel NUEVO (no toca ningún archivo real) con el Inventario Teórico + facturas del mes."""
    import openpyxl
    from openpyxl.styles import Alignment as XlAlignment, Border as XlBorder, Font as XlFont, PatternFill, Side as XlSide

    HEADER_GRAY = PatternFill("solid", fgColor="FFD9D9D9")
    MONEY_FMT = '"$" #,##0.00'
    GAL_FMT = "#,##0.00"
    THIN_SIDE = XlSide(style="thin", color="FF000000")
    THIN_BORDER = XlBorder(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
    CENTER = XlAlignment(horizontal="center", vertical="center", wrap_text=True)

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Fisico"

    title_cell = sheet.cell(row=1, column=1, value=f"BGS - Físico - {month:02d}/{year}")
    title_cell.font = XlFont(name="Segoe UI", bold=True, size=14)
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    sheet.row_dimensions[1].height = 21

    headers = ["Detalle", "Galones", "Monto $"]
    for col, text in enumerate(headers, start=1):
        cell = sheet.cell(row=3, column=col, value=text)
        cell.font = XlFont(bold=True)
        cell.fill = HEADER_GRAY
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    row = 4
    for label, gallons, amount in _resumen_rows(report):
        label_cell = sheet.cell(row=row, column=1, value=label)
        gal_cell = sheet.cell(row=row, column=2, value=gallons)
        amt_cell = sheet.cell(row=row, column=3, value=amount)
        gal_cell.number_format = GAL_FMT
        amt_cell.number_format = MONEY_FMT
        for col in range(1, 4):
            sheet.cell(row=row, column=col).border = THIN_BORDER
        row += 1

    price_cell = sheet.cell(row=row, column=1, value="Precio Promedio / Gal")
    price_val_cell = sheet.cell(row=row, column=3, value=report["avg_price"])
    price_val_cell.number_format = '"$" #,##0.00'
    for col in range(1, 4):
        sheet.cell(row=row, column=col).border = THIN_BORDER
        sheet.cell(row=row, column=col).font = XlFont(bold=True)
    row += 2

    sheet.column_dimensions["A"].width = 34
    sheet.column_dimensions["B"].width = 14
    sheet.column_dimensions["C"].width = 16

    inv_title_row = row
    inv_title = sheet.cell(row=inv_title_row, column=1, value=f"Facturas de {month:02d}/{year}")
    inv_title.font = XlFont(bold=True, size=12)
    row += 1

    inv_headers = ["Fecha Factura", "Fecha Vencimiento", "N° Factura", "Galones", "Monto $", "Precio $/Gal"]
    for col, text in enumerate(inv_headers, start=1):
        cell = sheet.cell(row=row, column=col, value=text)
        cell.font = XlFont(bold=True)
        cell.fill = HEADER_GRAY
        cell.alignment = CENTER
        cell.border = THIN_BORDER
    row += 1

    for inv in report["invoices"]:
        gallons = inv["gallons"] or 0.0
        amount = inv["amount"] or 0.0
        values = [
            _parse_date_for_excel(inv["invoice_date"]), _parse_date_for_excel(inv["due_date"]) or "—",
            inv["invoice_number"] or "—",
            gallons, amount, (amount / gallons) if gallons else None,
        ]
        for col, value in enumerate(values, start=1):
            cell = sheet.cell(row=row, column=col, value=value)
            cell.border = THIN_BORDER
            if col in (1, 2) and isinstance(value, date):
                cell.number_format = "dd/mm/yyyy"
            elif col == 4:
                cell.number_format = GAL_FMT
            elif col in (5, 6):
                cell.number_format = MONEY_FMT if col == 5 else '"$" #,##0.00'
        row += 1

    for extra_col in range(4, 7):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(extra_col)].width = 16

    workbook.save(dest_path)
    return dest_path


def build_fisico_pdf_report(report, year, month, dest_path):
    """Versión PDF del export de arriba -- mismo resumen + tabla de facturas, sin colores (ver pdf_export.py)."""
    from pdf_export import build_multi_section_pdf

    resumen_rows = [
        [label, _fmt_gal(gallons), _fmt_money(amount)]
        for label, gallons, amount in _resumen_rows(report)
    ]
    resumen_rows.append(["Precio Promedio / Gal", "", f"${report['avg_price']:,.2f}" if report["avg_price"] is not None else "—"])

    invoices = report["invoices"]
    invoice_rows = []
    for inv in invoices:
        gallons = inv["gallons"] or 0.0
        amount = inv["amount"] or 0.0
        price = (amount / gallons) if gallons else None
        invoice_rows.append([
            _fmt_date_ddmmyyyy(inv["invoice_date"]), _fmt_date_ddmmyyyy(inv["due_date"]), inv["invoice_number"] or "—",
            _fmt_gal(gallons), _fmt_money(amount),
            f"${price:,.2f}" if price is not None else "—",
        ])

    sections = [
        {
            "heading": "Resumen del mes",
            "headers": ["Detalle", "Galones", "Monto $"],
            "rows": resumen_rows,
            "col_widths_mm": [110, 55, 55],
            "bold_last_row": True,
        },
    ]
    if invoice_rows:
        sections.append({
            "heading": f"Facturas de {month:02d}/{year}",
            "headers": ["Fecha Factura", "Fecha Vencimiento", "N° Factura", "Galones", "Monto $", "Precio $/Gal"],
            "rows": invoice_rows,
        })
    else:
        sections.append({"heading": f"Facturas de {month:02d}/{year}", "note": "Sin facturas cargadas este mes."})

    build_multi_section_pdf(dest_path, f"Físico — {month:02d}.{year}", sections, period_label=None, company_header=True)
    return dest_path
