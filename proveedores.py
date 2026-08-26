"""
Proveedores — carga de facturas de compra (PDF) al libro de cuentas
corrientes por proveedor ("Bradenton. Cta Cte Proveedores.xlsx"), una hoja
por proveedor.

Cada proveedor tiene su propia forma de facturar (texto digital vs.
escaneado, una o varias páginas, foto de cheque incluida, etc.), así que
la extracción vive en SUPPLIER_REGISTRY: un extractor por proveedor,
detectado automáticamente por el contenido del PDF. Se agregan proveedores
nuevos ahí, sin tocar la lógica de inserción de filas.
"""

import os
import re
import sys
import tempfile
from copy import copy
from datetime import datetime

try:
    import pdfplumber
except ImportError:  # pragma: no cover - environment guard
    pdfplumber = None  # type: ignore[assignment]

try:
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill

    OPENPYXL_AVAILABLE = True
except ImportError:  # pragma: no cover - environment guard
    load_workbook = None  # type: ignore[assignment,misc]
    PatternFill = None  # type: ignore[assignment,misc]
    OPENPYXL_AVAILABLE = False

# Columnas de cada hoja de proveedor (1-based): DATE|COMPROB|N|Amount(DEBE)|
# Amount(HABER)|BALANCE|DETAIL. Las filas de factura solo escriben
# DATE/COMPROB/N/DEBE — HABER queda para los pagos (fase futura, via Chase).
INVOICE_ROW_COLUMNS = (1, 2, 3, 4, 6, 7)
COL_DATE, COL_COMPROB, COL_NUMERO, COL_DEBE, COL_HABER, COL_BALANCE, COL_DETALLE = range(1, 8)

# El color de la columna N (COMPROB) alterna mes a mes; se detecta el último
# usado y se alterna, nunca se hardcodea cuál mes es cuál color.
GREEN_FILL = "FF92D050"
YELLOW_FILL = "FFFFFF00"
MONTH_FILL_COLORS = (GREEN_FILL, YELLOW_FILL)


def _ensure_pdfplumber():
    if pdfplumber is None:
        raise ImportError(
            "Proveedores requiere pdfplumber. Instale con: pip install pdfplumber"
        )


def _create_temp_workbook_path():
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="proveedores_")
    os.close(fd)
    return temp_path


def _launch_temp_workbook(temp_path):
    abs_path = os.path.abspath(temp_path)
    if sys.platform == "win32":
        os.startfile(abs_path)
    elif sys.platform == "darwin":
        os.system(f'open "{abs_path}"')
    else:
        os.system(f'xdg-open "{abs_path}"')


# ---- Extractores por proveedor ----

def _extract_ht_hackney_invoice(pdf_path):
    """
    Las facturas de H.T. Hackney vienen en varias páginas; el N° de
    invoice y la fecha están en el encabezado de la primera página, y el
    total final real (con fuel surcharge, delivery charge e impuestos) solo
    aparece en la última página.
    """
    _ensure_pdfplumber()
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() or ""
        last_text = pdf.pages[-1].extract_text() or ""

    invoice_match = re.search(r"Invoice #:\s*(\d+)", first_text)
    date_match = re.search(r"Invoice Date:\s*(\d{1,2}/\d{1,2}/\d{2,4})", first_text)
    totals = re.findall(r"Total:\s*([\d,]+\.\d{2})", last_text)

    if not (invoice_match and date_match and totals):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total "
            "del PDF de H.T. Hackney."
        )

    invoice_no = int(invoice_match.group(1))
    invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%y")
    amount = float(totals[-1].replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


SUPPLIER_REGISTRY = {
    "ht_hackney": {
        "label": "H.T. Hackney",
        "sheet_name": "HT Hackney",
        "detect": lambda text: "h.t. hackney" in text.lower(),
        "extract": _extract_ht_hackney_invoice,
    },
}


def _detect_supplier(pdf_path):
    _ensure_pdfplumber()
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""
    for key, config in SUPPLIER_REGISTRY.items():
        if config["detect"](text):
            return key
    supported = ", ".join(cfg["label"] for cfg in SUPPLIER_REGISTRY.values())
    raise ValueError(
        f"{os.path.basename(pdf_path)}: proveedor no reconocido todavía. "
        f"Proveedores soportados por ahora: {supported}."
    )


def _get_supplier_sheet(workbook, sheet_name):
    target = sheet_name.strip().lower()
    for name in workbook.sheetnames:
        if name.strip().lower() == target:
            return workbook[name]
    raise ValueError(
        f'Hoja "{sheet_name}" no encontrada. Disponibles: {", ".join(workbook.sheetnames)}'
    )


def _find_last_real_row(sheet):
    """Última fila con una fecha real cargada en la columna A."""
    last_real_row = None
    for row in range(1, sheet.max_row + 1):
        if sheet.cell(row=row, column=COL_DATE).value not in (None, ""):
            last_real_row = row
    return last_real_row


def _find_last_month_color(sheet, from_row):
    for row in range(from_row, 0, -1):
        fill = sheet.cell(row=row, column=COL_NUMERO).fill
        if fill and fill.patternType == "solid" and fill.fgColor and fill.fgColor.rgb in MONTH_FILL_COLORS:
            return fill.fgColor.rgb
    return None


def _existing_invoice_numbers(sheet):
    numbers = set()
    for row in range(1, sheet.max_row + 1):
        value = sheet.cell(row=row, column=COL_NUMERO).value
        if isinstance(value, (int, float)):
            numbers.add(int(value))
    return numbers


def _append_invoice_row(sheet, style_ref_row, last_row, last_color, last_date, invoice):
    """
    Inserta una fila de factura justo después de last_row: si cambia el mes
    respecto a last_date, salta una fila (el separador en blanco ya
    existente entre meses) y alterna el color de la columna N; si es el
    mismo mes, la agrega justo debajo con el mismo color.

    El estilo (fuente/borde/alineación/formato) se copia de style_ref_row
    -- la última factura real ya cargada -- en vez de confiar en lo que
    traiga la fila destino, que puede tener overrides viejos inconsistentes.

    Devuelve (fila_nueva, color_usado) para encadenar la siguiente factura.
    """
    new_date = invoice["date"]
    month_changed = (new_date.year, new_date.month) != (last_date.year, last_date.month)
    target_row = last_row + 2 if month_changed else last_row + 1

    if month_changed:
        new_color = GREEN_FILL if last_color == YELLOW_FILL else YELLOW_FILL
    else:
        new_color = last_color or GREEN_FILL

    for col in INVOICE_ROW_COLUMNS:
        ref_cell = sheet.cell(row=style_ref_row, column=col)
        new_cell = sheet.cell(row=target_row, column=col)
        new_cell.font = copy(ref_cell.font)
        new_cell.border = copy(ref_cell.border)
        new_cell.alignment = copy(ref_cell.alignment)
        new_cell.number_format = ref_cell.number_format

    sheet.cell(row=target_row, column=COL_DATE, value=new_date)
    sheet.cell(row=target_row, column=COL_COMPROB, value="invoice")
    sheet.cell(row=target_row, column=COL_NUMERO, value=invoice["invoice_no"])
    sheet.cell(row=target_row, column=COL_DEBE, value=invoice["amount"])
    sheet.cell(
        row=target_row,
        column=COL_BALANCE,
        value=f"=+{sheet.cell(row=target_row - 1, column=COL_BALANCE).coordinate}"
        f"+{sheet.cell(row=target_row, column=COL_DEBE).coordinate}"
        f"-{sheet.cell(row=target_row, column=COL_HABER).coordinate}",
    )
    sheet.cell(row=target_row, column=COL_DETALLE).value = None

    sheet.cell(row=target_row, column=COL_NUMERO).fill = PatternFill(
        start_color=new_color, end_color=new_color, fill_type="solid"
    )

    return target_row, new_color


def append_supplier_invoices(ledger_path, pdf_paths):
    """
    Lee cada PDF, detecta su proveedor y agrega una fila "invoice" a la
    hoja correspondiente del libro -- en orden de fecha, encadenada a la
    última fila real que ya tenga esa hoja. Un mismo invoice ya cargado
    (por N° de factura) se omite. Guarda una copia temporal y la abre; el
    usuario la revisa y hace "Guardar como" como en el resto de los
    módulos.
    """
    if not OPENPYXL_AVAILABLE:
        raise ImportError("Proveedores requiere openpyxl. Instale con: pip install openpyxl")

    ledger_path = os.path.abspath(ledger_path)
    if not os.path.isfile(ledger_path):
        raise FileNotFoundError(f"Excel no encontrado: {ledger_path}")

    by_supplier = {}
    for pdf_path in pdf_paths:
        supplier_key = _detect_supplier(pdf_path)
        invoice = SUPPLIER_REGISTRY[supplier_key]["extract"](pdf_path)
        invoice["filename"] = os.path.basename(pdf_path)
        by_supplier.setdefault(supplier_key, []).append(invoice)

    workbook = load_workbook(ledger_path, data_only=False)

    batch_results = []
    total_appended = 0
    for supplier_key, invoices in by_supplier.items():
        config = SUPPLIER_REGISTRY[supplier_key]
        sheet = _get_supplier_sheet(workbook, config["sheet_name"])
        invoices.sort(key=lambda inv: inv["date"])

        existing_numbers = _existing_invoice_numbers(sheet)
        last_row = _find_last_real_row(sheet)
        last_date = sheet.cell(row=last_row, column=COL_DATE).value
        last_color = _find_last_month_color(sheet, last_row)
        style_ref_row = last_row

        appended = 0
        duplicates_skipped = []
        for invoice in invoices:
            if invoice["invoice_no"] in existing_numbers:
                duplicates_skipped.append(invoice["filename"])
                continue
            last_row, last_color = _append_invoice_row(
                sheet, style_ref_row, last_row, last_color, last_date, invoice
            )
            last_date = invoice["date"]
            existing_numbers.add(invoice["invoice_no"])
            appended += 1

        total_appended += appended
        batch_results.append(
            {
                "supplier": config["label"],
                "sheet_name": config["sheet_name"],
                "invoices_appended": appended,
                "duplicates_skipped": duplicates_skipped,
            }
        )

    temp_path = _create_temp_workbook_path()
    workbook.save(os.path.abspath(temp_path))
    workbook.close()

    _launch_temp_workbook(temp_path)

    summary = {
        "files_processed": len(pdf_paths),
        "invoices_appended": total_appended,
        "batch_results": batch_results,
    }
    return temp_path, summary
