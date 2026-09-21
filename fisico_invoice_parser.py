"""
Lector de facturas de compra de combustible en PDF -- pedido explícito de
Alfonso (2026-09-21), subiendo 5 facturas reales del proveedor de
combustible (formato con tabla de producto "Qty/Item/Description/
Contract/Unit Price w/o Tax/Unit Price with Tax/Total Price Including
Tax", una tabla "FREIGHT SUMMARY" y un total final "TOTAL INVOICE AMOUNT
DUE") para que el módulo Combustible/Físico se actualice solo en vez de
cargar todo a mano.

Son PDF de texto real (no escaneados/fotografiados) -- decodificado
leyendo las 5 facturas reales con pdfplumber: `page.extract_tables()` ya
da filas limpias (sin falsos saltos de línea a mitad de celda) para la
tabla de producto y la de FREIGHT SUMMARY, así que no hace falta regex
frágil sobre texto corrido para esas dos partes. El nombre del
proveedor (logo "J.H.WILLIAMS OIL COMPANY INC.") es una imagen, NO texto
seleccionable -- confirmado con pdfplumber -- así que la detección de
que "esto es una factura de combustible de este formato" es estructural
(existencia de la tabla de producto + FREIGHT SUMMARY + TOTAL INVOICE
AMOUNT DUE), no por nombre de proveedor.

Qué se extrae de cada factura y por qué (pedido explícito del usuario,
2026-09-21):
  - "la cantidad de compra" -> el galonaje de FREIGHT-FUELS (columna Qty
    de FREIGHT SUMMARY), NO la columna Qty de la tabla de producto --
    son números distintos (ej. factura real SI-211258: 7,455.0 galones
    facturados de 87 octanos, pero 7,600 en su fila FREIGHT-FUELS)
    porque uno es galonaje neto de facturación y el otro el bruto que el
    camión trae de verdad -- lo que se anota en la hoja física real del
    cierre (confirmado por el usuario: "ahi dice la cantidad de
    FREIGHT-FUELS que se trae del camion que es lo que se anota en la
    hoja de fisico"). Coincide EXACTO con el ejemplo real ya
    documentado en fisico.py (`=7600+1199`, de una carga manual anterior
    de esta misma factura SI-211258) -- confirma que la interpretación
    es la correcta, no es una suposición nueva.
  - "el precio del combustible con y sin tax" -> Unit Price w/o Tax /
    Unit Price with Tax de la tabla de producto, por grado (87/93
    octanos), tal cual figuran -- se guardan en `fuel_invoice_lines`,
    a título informativo/histórico (Físico sigue usando el agregado
    gallons/amount, sin split por grado, ver fisico_db.py).
  - "el total de lo que salio la factura que esta abajo de todo" ->
    TOTAL INVOICE AMOUNT DUE (la línea "PLEASE PAY THIS AMOUNT", con
    flete e impuestos incluidos), NO el "INVOICE TOTAL" de arriba de la
    tabla de producto (que es solo el costo del combustible sin flete/
    impuestos) -- es el costo real pagado, lo que se guarda como
    `amount` (mismo campo que ya usaba la carga manual).

Cada línea de producto se empareja POSICIONALMENTE con su fila
FREIGHT-FUELS correspondiente (mismo orden en que aparecen en cada
tabla -- no hay ningún campo compartido tipo código de producto en
FREIGHT SUMMARY para cruzar de otra forma). Validado en las 5 facturas
reales: la proporción FREIGHT-FUELS/facturación siempre cae entre 1.01 y
1.03 -- se rechaza cualquier factura donde esa proporción se salga de un
rango razonable, señal de que el emparejamiento por posición no es
confiable para ese PDF puntual.

Regla de oro de OCR del proyecto: si algo no se puede leer o no cuadra
con confianza, se lanza ValueError con un mensaje claro -- nunca se
guarda un valor de baja confianza en silencio. Cada factura del lote se
aísla (ver `carga_datos_combustible_subir_pdf` en webapp.py) -- una
factura rota nunca tira abajo el resto.
"""

import os
import re
from datetime import datetime

try:
    import pdfplumber
except ImportError:  # pragma: no cover - environment guard
    pdfplumber = None  # type: ignore[assignment]

try:
    from pdfplumber.utils.exceptions import MalformedPDFException, PdfminerException
except ImportError:  # pragma: no cover - environment guard
    MalformedPDFException = PdfminerException = Exception  # type: ignore[assignment,misc]

from ocr_utils import ensure_pdfplumber

# Mismo criterio de aislamiento que proveedores.py/proveedores_dynamic_extractors.py,
# más MalformedPDFException/PdfminerException -- un PDF genuinamente corrupto (no
# solo "de otro formato") hace que pdfplumber.open() tire esto, no ValueError/
# OSError; confirmado probando con un archivo .pdf basura real durante esta sesión.
PDF_READ_EXCEPTIONS = (
    ValueError, TypeError, AttributeError, RuntimeError, OSError,
    MalformedPDFException, PdfminerException,
)

_ITEM_HEADER_FIRST_CELL = "Qty"
_FREIGHT_TABLE_TITLE = "FREIGHT SUMMARY"
_FREIGHT_FUELS_CODE = "FREIGHT-FUELS"

_INVOICE_NO_RE = re.compile(r"Invoice No:\s*(\S+)")
_INVOICE_DATE_RE = re.compile(r"Invoice Date:\s*(\d{1,2}/\d{1,2}/\d{2,4})")
_DUE_DATE_RE = re.compile(r"Invoice Due Date:\s*(\d{1,2}/\d{1,2}/\d{2,4})")
_BOL_RE = re.compile(r"BOL Number:\s*(\S+)")
_FUEL_SUBTOTAL_RE = re.compile(r"INVOICE TOTAL:\s*\$?\s*([\d,]+\.\d{2})")
_TOTAL_DUE_RE = re.compile(r"TOTAL INVOICE AMOUNT DUE:\s*\$?\s*([\d,]+\.\d{2})")

# Rango de tolerancia para la proporción galonaje-freight / galonaje-facturación
# de cada línea -- validado contra las 5 facturas reales (siempre 1.01-1.03).
_FREIGHT_RATIO_MIN = 0.85
_FREIGHT_RATIO_MAX = 1.25


def _num(value):
    return float(str(value).replace(",", "").replace("$", "").strip())


def _find_table(tables, matcher):
    for table in tables:
        if table and table[0] and matcher(table[0]):
            return table
    return None


def extract_fuel_invoice(pdf_path):
    """
    Lee una factura de compra de combustible en PDF y devuelve un dict:
      invoice_number, invoice_date (date), due_date (date), bol_number,
      lines (lista de dicts con product_code/description/qty_billing/
      qty_freight/unit_price_wo_tax/unit_price_with_tax/line_total),
      fuel_subtotal (INVOICE TOTAL, solo combustible, informativo),
      total_amount_due (TOTAL INVOICE AMOUNT DUE -- lo que se guarda como
      "amount"), total_gallons (suma de qty_freight de todas las líneas
      -- lo que se guarda como "gallons").

    Lanza ValueError (mensaje claro, incluye el nombre del archivo) si
    falta algún campo obligatorio, si la cantidad de filas FREIGHT-FUELS
    no coincide 1 a 1 con las líneas de producto, si alguna proporción
    freight/facturación se sale de rango, o si los totales no cuadran --
    nunca adivina ni guarda con baja confianza.
    """
    ensure_pdfplumber()
    filename = os.path.basename(pdf_path)
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            raise ValueError(f"{filename}: el PDF no tiene páginas.")
        page = pdf.pages[0]
        text = page.extract_text() or ""
        tables = page.extract_tables()

    if not text.strip():
        raise ValueError(
            f"{filename}: no se encontró texto en el PDF (¿es un escaneo o una foto? "
            "este lector solo entiende facturas digitales de texto real) -- cargala a mano."
        )

    invoice_no_m = _INVOICE_NO_RE.search(text)
    date_m = _INVOICE_DATE_RE.search(text)
    due_m = _DUE_DATE_RE.search(text)
    bol_m = _BOL_RE.search(text)
    fuel_subtotal_m = _FUEL_SUBTOTAL_RE.search(text)
    due_amount_matches = _TOTAL_DUE_RE.findall(text)

    item_table = _find_table(
        tables, lambda header: header[0] == _ITEM_HEADER_FIRST_CELL and "Item" in header
    )
    freight_table = _find_table(tables, lambda header: header[0] == _FREIGHT_TABLE_TITLE)

    missing = []
    if not invoice_no_m:
        missing.append("N° de factura")
    if not date_m:
        missing.append("fecha de factura")
    if not due_m:
        missing.append("fecha de vencimiento")
    if not item_table or len(item_table) < 2:
        missing.append("tabla de productos (Qty/Item/Description/...)")
    if not freight_table:
        missing.append('tabla "FREIGHT SUMMARY"')
    if not fuel_subtotal_m:
        missing.append("INVOICE TOTAL")
    if not due_amount_matches:
        missing.append("TOTAL INVOICE AMOUNT DUE")
    if missing:
        raise ValueError(
            f"{filename}: no se pudo leer con confianza: {', '.join(missing)}. "
            "Este lector espera el formato de factura de combustible ya conocido "
            "(Qty/Item/Description/Unit Price w/o Tax/Unit Price with Tax + FREIGHT "
            "SUMMARY + TOTAL INVOICE AMOUNT DUE) -- si esta factura es de otro "
            "formato o proveedor, cargala a mano."
        )

    lines = []
    for row in item_table[1:]:
        if not row or len(row) < 7 or not (row[0] or "").strip():
            continue
        qty_billing, code, description, _contract, price_wo, price_with, total = row[:7]
        try:
            lines.append({
                "product_code": (code or "").strip(),
                "description": " ".join((description or "").split()),
                "qty_billing": _num(qty_billing),
                "unit_price_wo_tax": _num(price_wo),
                "unit_price_with_tax": _num(price_with),
                "line_total": _num(total),
            })
        except (ValueError, AttributeError) as exc:
            raise ValueError(
                f"{filename}: no se pudo leer una línea de producto ({row}): {exc}"
            ) from exc

    if not lines:
        raise ValueError(f"{filename}: la tabla de productos no tiene ninguna línea de combustible legible.")

    freight_rows = [
        row for row in freight_table if row and (row[0] or "").strip() == _FREIGHT_FUELS_CODE
    ]
    if len(freight_rows) != len(lines):
        raise ValueError(
            f"{filename}: {len(freight_rows)} fila(s) FREIGHT-FUELS contra {len(lines)} línea(s) de "
            "producto -- no se puede mapear con confianza qué galonaje-freight corresponde a cada grado."
        )

    for line, freight_row in zip(lines, freight_rows):
        try:
            freight_gallons = _num(freight_row[2])
        except (ValueError, IndexError) as exc:
            raise ValueError(
                f"{filename}: no se pudo leer el galonaje de FREIGHT-FUELS ({freight_row}): {exc}"
            ) from exc
        ratio = (freight_gallons / line["qty_billing"]) if line["qty_billing"] else None
        if ratio is None or not (_FREIGHT_RATIO_MIN <= ratio <= _FREIGHT_RATIO_MAX):
            raise ValueError(
                f"{filename}: el galonaje FREIGHT-FUELS ({freight_gallons}) y el galonaje facturado "
                f"({line['qty_billing']}) de {line['product_code']} no guardan una proporción razonable "
                "-- no se guarda nada por las dudas (revisar/cargar esta factura a mano)."
            )
        line["qty_freight"] = freight_gallons

    fuel_subtotal = _num(fuel_subtotal_m.group(1))
    calc_subtotal = round(sum(line["line_total"] for line in lines), 2)
    if abs(fuel_subtotal - calc_subtotal) > 0.05:
        raise ValueError(
            f"{filename}: INVOICE TOTAL (${fuel_subtotal:,.2f}) no coincide con la suma de las líneas "
            f"de producto (${calc_subtotal:,.2f}) -- no se guarda nada por las dudas."
        )

    due_amounts = {round(_num(value), 2) for value in due_amount_matches}
    if len(due_amounts) != 1:
        raise ValueError(
            f"{filename}: TOTAL INVOICE AMOUNT DUE aparece con valores distintos en el PDF "
            f"({sorted(due_amounts)}) -- no se guarda nada por las dudas."
        )
    total_amount_due = due_amounts.pop()

    try:
        invoice_date = datetime.strptime(date_m.group(1), "%m/%d/%Y").date()
        due_date = datetime.strptime(due_m.group(1), "%m/%d/%Y").date()
    except ValueError as exc:
        raise ValueError(f"{filename}: no se pudo interpretar una fecha ({exc}).") from exc

    total_gallons = round(sum(line["qty_freight"] for line in lines), 2)

    return {
        "invoice_number": invoice_no_m.group(1).strip(),
        "invoice_date": invoice_date,
        "due_date": due_date,
        "bol_number": (bol_m.group(1).strip() if bol_m else None),
        "lines": lines,
        "fuel_subtotal": fuel_subtotal,
        "total_amount_due": total_amount_due,
        "total_gallons": total_gallons,
    }
