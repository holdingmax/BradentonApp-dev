"""
Lector de PDFs de "Pagos" de Gettel/Toyota/Kia -- pedido explícito de
Alfonso (2026-09-21): reemplaza la carga a mano de cada cupón/transacción
por la lectura automática del PDF de recibos (mismo criterio "no volver a
escribir todo a mano" ya aplicado en Combustible/Físico el mismo día -- ver
fisico_invoice_parser.py).

Cada PDF ("PagosN Empresa.pdf") es un escaneo/foto (CamScanner, sin capa de
texto -- 0 caracteres con pdfplumber.extract_text) de una tanda de recibos
de "LOCAL ACCT" de la estación 14th St. Chevron: una cuenta corriente que
factura combustible a Toyota o Kia en vez de cobrarlo en el momento, con el
nombre de la empresa escrito a mano arriba del recibo. Cada página = un
recibo = una transacción/cupón. Como no hay texto embebido, hace falta OCR
sobre la imagen de cada página -- se reusan los helpers compartidos de
ocr_utils.py (extract_largest_page_image ya corrige la rotación).

El N° de Pago y la Empresa NO aparecen en ningún recibo -- un recibo nunca
dice a qué "pago"/envío de Gettel se lo va a agrupar, eso es una referencia
externa que Gettel asigna al reembolsar. Se toman del nombre del archivo,
mismo criterio que ya usaba la herramienta vieja /gettel/pagos
(gettel_toyota_parser.extract_pago_batch_from_pdf/_parse_pagos_filename).

Bug real encontrado en esa herramienta vieja -- explica el "no esta leyendo
bien los pagos que subi en pdf" que llevó a Alfonso a pedir carga manual el
2026-09-19 (ver claude/Estado_2026-09-17_commits_y_push_pendiente.md): su
regex de nombre de archivo exigía el formato exacto "PagosN (Empresa).pdf"
CON paréntesis (`pagos?\s*(\d+)\s*\(([^)]+)\)`) -- los archivos reales de
Alfonso se llaman "Pagos1 Toyota.pdf" (sin paréntesis), así que CADA
archivo fallaba en el primer paso, antes de leer una sola página. Acá el
regex acepta espacio, guion o guion bajo como separador, con o sin
paréntesis.

Validado (2026-09-21) leyendo los 4 PDFs reales que subió Alfonso
(Pagos1/2/4 Toyota, Pagos3 Kia -- 27 recibos en total) y comparando a mano
(visualmente, página por página) contra las fotos: los 27 recibos se
leyeron exactos (fecha + N° de Transacción + Total) en la primera pasada
después de ajustar el regex del Total (ver _PAGOS_TOTAL_RE) para tolerar
ruido de OCR entre "=" y "$" (ej. "Total = | $58.09", "Total = — $58.09").

Golden rule del proyecto (igual que fisico_invoice_parser.py): nunca se
inserta un campo con baja confianza. Una página cuyo N° de Transacción,
Total o Fecha no se puede leer con confianza queda AFUERA (no se guarda) y
se reporta en `page_warnings` con el número de página -- el resto de las
páginas del mismo PDF se procesan igual (aislamiento por página, no por
archivo: un recibo ilegible no tira abajo los otros 26). Si el NOMBRE del
archivo no se puede leer (N° de Pago/Empresa), en cambio, se aborta el
archivo entero -- sin esos dos datos no hay dónde guardar ninguno de sus
recibos.
"""

import os
import re
from datetime import datetime

import pdfplumber

import ocr_utils

try:
    import pytesseract
except ImportError:  # pragma: no cover - environment guard
    pytesseract = None  # type: ignore[assignment]

try:
    from pdfplumber.utils.exceptions import MalformedPDFException, PdfminerException
except ImportError:  # pragma: no cover - older pdfplumber
    class MalformedPDFException(Exception):
        pass

    class PdfminerException(Exception):
        pass

# Mismo criterio que fisico_invoice_parser.PDF_READ_EXCEPTIONS -- cualquier
# excepción de acá se atrapa por archivo, aisla el resto del lote (ver
# webapp.py _run_carga_datos_gettel_pagos_job).
PDF_READ_EXCEPTIONS = (
    ValueError, TypeError, AttributeError, RuntimeError, OSError,
    MalformedPDFException, PdfminerException,
)

# "Pagos1 Toyota.pdf", "Pagos2_Toyota.pdf", "Pagos3 (Kia).pdf", "Pago 4 - Kia.pdf" -- todas válidas.
_PAGOS_FILENAME_RE = re.compile(r"pagos?\s*[-_ ]?\s*(\d+)\s*[-_ ]*\(?\s*([a-zA-Z][a-zA-Z ]*?)\s*\)?\s*$", re.IGNORECASE)
_PAGOS_TRANS_RE = re.compile(r"trans\s*#\s*[:;,.]*\s*(\d+)", re.IGNORECASE)
# Tolera basura de OCR entre "=" y "$" (ej. "Total = | $58.09", "Total = — $58.09") -- ver docstring del módulo.
_PAGOS_TOTAL_RE = re.compile(r"\btotal\s*=?\s*[^0-9$]{0,4}\$?\s*([\d,]+\.\d{2})", re.IGNORECASE)
_PAGOS_DATE_RE = re.compile(r"(\d{1,2}/\d{1,2}/\d{2,4})")
_DATE_INPUT_FORMATS = ("%m/%d/%Y", "%m/%d/%y")

# Probados contra los 4 PDFs reales -- psm 6/4 (texto en bloque) leyeron
# limpio "Total = $X.XX"; psm 12/11 (sparse+OSD) a veces separan "Total" y
# el monto en líneas distintas, o leen espacios de más adentro del número
# (ej. "98 .09" en vez de "58.09") -- se prueban todos y se toma el texto
# con más campos (fecha/trans/total) reconocidos.
_OCR_TESSERACT_CONFIGS = ("--psm 6", "--psm 4", "--psm 12", "--psm 11", "--psm 3")


def _parse_date(token):
    text = token.strip()
    for fmt in _DATE_INPUT_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def parse_pagos_filename(pdf_path):
    """("Pagos3 Toyota.pdf") -> (3, "Toyota"). Ver docstring del módulo para el bug de la versión vieja."""
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    match = _PAGOS_FILENAME_RE.search(base)
    if not match:
        raise ValueError(
            f'"{os.path.basename(pdf_path)}": no se pudo leer el N° de Pago ni la Empresa del nombre del '
            'archivo -- se espera algo como "Pagos3 Toyota.pdf" (N° de pago + empresa en el nombre).'
        )
    return int(match.group(1)), match.group(2).strip().title()


def _ocr_receipt_text(image):
    """Mejor texto de entre varios modos de Tesseract, elegido por cuántos de los 3 campos buscados reconoce."""
    if image is None or pytesseract is None:
        return "", -1
    best_text, best_score = "", -1
    for config in _OCR_TESSERACT_CONFIGS:
        try:
            text = pytesseract.image_to_string(image, config=config) or ""
        except Exception:
            continue
        score = sum(
            1 for pattern in (_PAGOS_TRANS_RE, _PAGOS_TOTAL_RE, _PAGOS_DATE_RE) if pattern.search(text)
        )
        if score > best_score:
            best_score, best_text = score, text
    return best_text, best_score


def extract_pagos_from_pdf(pdf_path):
    """
    Lee un "PagosN Empresa.pdf" entero -- un recibo por página -- y devuelve:
        {"pago_n": int, "empresa": str,
         "receipts": [{"fecha": date, "transc_n": str, "total_cupon": float}, ...],
         "page_warnings": [str, ...]}

    Cada receipt en `receipts` se leyó con confianza (los 3 campos
    reconocidos); las páginas que no se pudieron leer quedan afuera y
    listadas en `page_warnings`, nunca se inserta un valor adivinado.
    Lanza ValueError (dentro de PDF_READ_EXCEPTIONS) si el nombre de
    archivo no trae N° de Pago/Empresa, si el PDF no tiene páginas, o si
    NINGUNA página se pudo leer con confianza.
    """
    ocr_utils.ensure_pdfplumber()
    ocr_utils.ensure_pytesseract()
    filename = os.path.basename(pdf_path)
    pago_n, empresa = parse_pagos_filename(pdf_path)

    receipts = []
    page_warnings = []
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            raise ValueError(f'"{filename}": el PDF no tiene páginas.')
        for index, page in enumerate(pdf.pages, start=1):
            image = ocr_utils.extract_largest_page_image(page)
            if image is None:
                page_warnings.append(f"página {index}: no se encontró ninguna imagen en esa página.")
                continue

            text, _score = _ocr_receipt_text(image)
            trans_match = _PAGOS_TRANS_RE.search(text)
            total_match = _PAGOS_TOTAL_RE.search(text)
            date_match = _PAGOS_DATE_RE.search(text)
            receipt_date = _parse_date(date_match.group(1)) if date_match else None

            missing = []
            if not trans_match:
                missing.append("N° de Transacción")
            if not total_match:
                missing.append("Total")
            if not date_match or receipt_date is None:
                missing.append("Fecha")
            if missing:
                page_warnings.append(
                    f"página {index}: no se pudo leer con confianza ({', '.join(missing)}) -- revisar a mano."
                )
                continue

            receipts.append({
                "fecha": receipt_date,
                "transc_n": trans_match.group(1),
                "total_cupon": float(total_match.group(1).replace(",", "")),
            })

    if not receipts:
        detail = " ".join(page_warnings) if page_warnings else ""
        raise ValueError(f'"{filename}": no se pudo leer ningún recibo con confianza de este PDF. {detail}'.strip())

    return {
        "pago_n": pago_n,
        "empresa": empresa,
        "receipts": receipts,
        "page_warnings": page_warnings,
    }
