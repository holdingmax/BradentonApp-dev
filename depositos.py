"""
Depósitos (Documentos -> Depósitos) -- lectura de los recibos de cajero de
Chase ("My Transaction Summary"), sin conocimiento de la base (eso vive en
depositos_db.py).

Pedido explícito del usuario (2026-09-23): un PDF de depósitos puede traer
varios recibos (uno por página, ej. "Transaccion #58-59-60 21-09-2026.pdf"),
y cada uno tiene que quedar como una fila propia con su fecha, monto,
descripción y su propio PDF. Los recibos son fotos (sin texto digital) pero
limpias: el OCR lee bien "Transaction #134", "Checking Deposit $2,973.00" y
"Business Date 09/15/2026" (validado contra los 26 recibos de septiembre-2026,
todos coinciden con su depósito en Chase).

Food Truck / Ice Machine: a veces lo escriben a mano sobre la foto (no se
puede leer por OCR), pero el nombre de archivo ya lo aclara por transacción
-- "Transaccion #133 (Ice Machine)-134" -- y se usa eso. Un depósito sin
aclaración de exactamente $1,000/$1,030 es Food Truck (misma regla que
chase_rules._split_small_deposit).
"""

import os
import re
from datetime import date, datetime

import pdfplumber
import pytesseract
from pypdf import PdfReader, PdfWriter

from ocr_utils import ensure_pytesseract, extract_largest_page_image

FOOD_TRUCK = "Food Truck"
ICE_MACHINE = "Ice Machine"
_FOOD_TRUCK_AMOUNTS = (1000.0, 1030.0)

# "$2,973.00" -- tolera los espacios que a veces mete el OCR ("$626 .00", "$2, 140.00").
_AMOUNT = r"\$\s?(\d[\d,\s]*\.\s?\d{2})"
_TX_RE = re.compile(r"Transaction\s*#\s*(\d+)", re.I)
_DEPOSIT_RE = re.compile(r"Checking\s+Deposit\s*" + _AMOUNT, re.I)
_CASH_IN_RE = re.compile(r"Cash\s+In\s*" + _AMOUNT, re.I)
_BUSINESS_DATE_RE = re.compile(r"Business\s+Date\s*(\d{2}/\d{2}/\d{4})", re.I)
# Fecha al final del nombre, DD-MM-YYYY (ej. "... 15-09-2026.pdf" o "(03-08-2026)").
_FILENAME_DATE_RE = re.compile(r"(\d{1,2})[-./](\d{1,2})[-./](\d{4})(?!.*\d{1,2}[-./]\d{1,2}[-./]\d{4})")
# "#118 (Food Truck)-119-120 (Vaccumms)" -> {118: "Food Truck", 120: "Vaccumms"}.
# El paréntesis de la fecha "(03-08-2026)" queda afuera porque empieza con dígito.
_ANNOTATION_RE = re.compile(r"(\d+)\s*\(\s*([^)\d][^)]*)\)")


def filename_date(filename):
    match = _FILENAME_DATE_RE.search(os.path.splitext(filename or "")[0])
    if not match:
        return None
    try:
        return date(int(match.group(3)), int(match.group(2)), int(match.group(1)))
    except ValueError:
        return None


def normalize_kind(text):
    lowered = (text or "").lower()
    if "food" in lowered or "truck" in lowered:
        return FOOD_TRUCK
    if "ice" in lowered or "hielo" in lowered:
        return ICE_MACHINE
    return (text or "").strip() or None


def filename_annotations(filename):
    base = os.path.splitext(filename or "")[0]
    return {int(num): normalize_kind(label) for num, label in _ANNOTATION_RE.findall(base)}


def _parse_amount(match):
    if not match:
        return None
    try:
        return float(re.sub(r"[\s,]", "", match.group(1)))
    except ValueError:
        return None


def read_receipt_text(text):
    """Campos de un recibo a partir del texto OCR de su página."""
    tx = _TX_RE.search(text)
    amount = _parse_amount(_DEPOSIT_RE.search(text))
    if amount is None:
        amount = _parse_amount(_CASH_IN_RE.search(text))
    business = _BUSINESS_DATE_RE.search(text)
    deposit_date = None
    if business:
        try:
            deposit_date = datetime.strptime(business.group(1), "%m/%d/%Y").date()
        except ValueError:
            deposit_date = None
    return {
        "tx_number": int(tx.group(1)) if tx else None,
        "amount": amount,
        "date": deposit_date,
    }


def default_description(tx_number, kind):
    text = f"Transacción #{tx_number}" if tx_number is not None else "Depósito"
    return f"{text} ({kind})" if kind else text


def extract_deposits_from_pdf(pdf_path, filename):
    """
    Un dict por página (= un recibo): page, tx_number, date, amount, kind.
    Cualquier dato que no se pudo leer queda en None (se completa a mano) --
    nunca se inventa. La fecha cae al nombre de archivo si el OCR no la leyó.
    """
    ensure_pytesseract()
    annotations = filename_annotations(filename)
    fallback_date = filename_date(filename)
    results = []
    with pdfplumber.open(pdf_path) as pdf:
        for index, page in enumerate(pdf.pages):
            fields = {"tx_number": None, "amount": None, "date": None}
            try:
                image = extract_largest_page_image(page)
                if image is not None:
                    fields = read_receipt_text(pytesseract.image_to_string(image, config="--psm 6"))
            except Exception as exc:
                print(f"[depositos] {filename} página {index + 1}: {exc}")
            kind = annotations.get(fields["tx_number"]) if fields["tx_number"] is not None else None
            if kind is None and fields["amount"] in _FOOD_TRUCK_AMOUNTS:
                kind = FOOD_TRUCK
            results.append({
                "page": index,
                "tx_number": fields["tx_number"],
                "date": fields["date"] or fallback_date,
                "amount": fields["amount"],
                "kind": kind,
            })
    return results


def write_single_page_pdf(pdf_path, page_index, dest_path):
    """Copia una sola página del PDF original (sin re-comprimir la imagen)."""
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    writer.add_page(reader.pages[page_index])
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    with open(dest_path, "wb") as handle:
        writer.write(handle)
