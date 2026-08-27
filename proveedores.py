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

import io
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
    import pytesseract
    from PIL import Image
except ImportError:  # pragma: no cover - environment guard
    pytesseract = None  # type: ignore[assignment]
    Image = None  # type: ignore[assignment]

try:
    import cv2
    import numpy as np
except ImportError:  # pragma: no cover - environment guard
    cv2 = None  # type: ignore[assignment]
    np = None  # type: ignore[assignment]

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill

    OPENPYXL_AVAILABLE = True
except ImportError:  # pragma: no cover - environment guard
    load_workbook = None  # type: ignore[assignment,misc]
    Font = None  # type: ignore[assignment,misc]
    PatternFill = None  # type: ignore[assignment,misc]
    OPENPYXL_AVAILABLE = False

_OSD_ROTATE_TO_TRANSPOSE = {
    90: Image.ROTATE_270 if Image is not None else None,
    180: Image.ROTATE_180 if Image is not None else None,
    270: Image.ROTATE_90 if Image is not None else None,
}

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


_TESSERACT_CANDIDATE_PATHS = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
)
_TESSERACT_CONFIGURED = False


def _ensure_pytesseract():
    global _TESSERACT_CONFIGURED
    if pytesseract is None or Image is None:
        raise ImportError(
            "Leer facturas escaneadas requiere pytesseract y Pillow. "
            "Instale con: pip install pytesseract pillow"
        )
    if _TESSERACT_CONFIGURED:
        return
    try:
        pytesseract.get_tesseract_version()
        _TESSERACT_CONFIGURED = True
        return
    except Exception:
        pass
    for candidate in _TESSERACT_CANDIDATE_PATHS:
        if os.path.isfile(candidate):
            pytesseract.pytesseract.tesseract_cmd = candidate
            _TESSERACT_CONFIGURED = True
            return
    raise ImportError(
        "No se encontró el motor Tesseract OCR. Instálelo (ej. con 'winget install "
        "UB-Mannheim.TesseractOCR') para poder leer facturas escaneadas."
    )


def _correct_image_orientation(image):
    """Undo whole-page rotation via la propia deteccion de orientacion de Tesseract."""
    try:
        osd = pytesseract.image_to_osd(image, output_type=pytesseract.Output.DICT)
        rotate = int(osd.get("rotate", 0) or 0)
    except Exception:
        return image
    transpose_const = _OSD_ROTATE_TO_TRANSPOSE.get(rotate)
    if transpose_const is not None:
        image = image.transpose(transpose_const)
    return image


def _extract_page_image(page):
    """La imagen mas grande incrustada en una pagina de pdfplumber, o None."""
    if not page.images:
        return None
    biggest = max(page.images, key=lambda im: im["width"] * im["height"])
    raw = biggest["stream"].get_data()
    image = Image.open(io.BytesIO(raw)).convert("RGB")
    return _correct_image_orientation(image)


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


def _ensure_cv2():
    if cv2 is None or np is None:
        raise ImportError(
            "Leer algunas facturas escaneadas requiere opencv-python y numpy. "
            "Instale con: pip install opencv-python numpy"
        )


def _crop_relative(image, left, top, right, bottom):
    """Recorte por fracción del ancho/alto (0.0-1.0), no por píxeles fijos."""
    w, h = image.size
    return image.crop((int(w * left), int(h * top), int(w * right), int(h * bottom)))


def _remove_grid_lines(image, upscale=3):
    """
    Algunas facturas meten el N°/fecha en una tabla con bordes que Tesseract
    confunde con parte del texto (o directamente no lee nada adentro).
    Sube la resolución y borra las líneas horizontales/verticales con
    morfología de OpenCV antes de OCR -- sin esto, celdas como la de fecha
    de King's salen vacías o con los dígitos mezclados con el borde.
    """
    _ensure_cv2()
    gray = np.array(image.convert("L").resize((image.width * upscale, image.height * upscale), Image.LANCZOS))
    _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    horiz_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 120))
    lines = cv2.bitwise_or(
        cv2.morphologyEx(bw, cv2.MORPH_OPEN, horiz_kernel, iterations=2),
        cv2.morphologyEx(bw, cv2.MORPH_OPEN, vert_kernel, iterations=2),
    )
    cleaned = cv2.bitwise_not(cv2.bitwise_and(bw, cv2.bitwise_not(lines)))
    return Image.fromarray(cleaned)


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


def _extract_cec_invoice(pdf_path):
    """
    Las facturas de CEC (Chinook Enterprises Corp.) son un escaneo de una
    sola página, sin capa de texto -- hace falta OCR. El N° de invoice
    puede salir con un espacio de más entre dígitos (ej. "188412 7"), y el
    total final a veces se pierde por el fondo gris de su casillero, así
    que se usa el "Subtotal:" -- numéricamente igual al total en todas las
    facturas vistas -- que siempre se lee bien.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        image = _extract_page_image(pdf.pages[0])
    if image is None:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró la imagen escaneada de la factura CEC."
        )
    text = pytesseract.image_to_string(image)

    invoice_match = re.search(r"Inv\s*#\s*([\d\s]+?)\n", text)
    date_match = re.search(r"Order taken on\s*(\d{1,2}/\d{1,2}/\d{4})", text)
    subtotal_match = re.search(r"Subtotal:?\s*\$?\s*([\d,]+\.\d{2})", text)

    if not (invoice_match and date_match and subtotal_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total "
            "del PDF de CEC (Chinook)."
        )

    invoice_no = int(re.sub(r"\s+", "", invoice_match.group(1)))
    invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%Y")
    amount = float(subtotal_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_colonial_invoice(pdf_path):
    """
    Colonial Wholesale Dist. LLC -- escaneo de una sola página. El N° y la
    fecha están en la cajita de arriba a la derecha ("INVOICE NO." /
    "INVOICE DATE"); el "BALANCE DUE" en la esquina inferior derecha, sobre
    la línea de la firma.

    Limitación conocida: si el conductor corrigió el total a mano (ej. un
    crédito tachado y reescrito), la letra manuscrita no se puede leer de
    forma confiable y esta función va a fallar con un error claro -- en ese
    caso hay que cargar esa factura a mano.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        image = _extract_page_image(pdf.pages[0])
    if image is None:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró la imagen escaneada de la factura Colonial."
        )

    top_text = pytesseract.image_to_string(_crop_relative(image, 0.68, 0.0, 1.0, 0.22))
    bottom_text = pytesseract.image_to_string(_crop_relative(image, 0.50, 0.90, 1.0, 1.0))

    invoice_match = re.search(r"\b(\d{7})\b", top_text)
    date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})", top_text)
    balance_match = re.search(r"([\d,]+\.\d{2})", bottom_text)

    if not (invoice_match and date_match and balance_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de "
            "Colonial (si tiene una corrección a mano, cargue esta factura manualmente)."
        )

    invoice_no = int(invoice_match.group(1))
    invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%y")
    amount = float(balance_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_gce_invoice(pdf_path):
    """
    Gold Coast Eagle -- escaneo; con muchos ítems, la factura se corre a
    una segunda página y la línea de confirmación ("Inv# 657065
    $1,381.00", con invoice y total juntos) queda ahí, no en la primera.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = []
        for page in pdf.pages:
            image = _extract_page_image(page)
            if image is not None:
                pages_text.append(pytesseract.image_to_string(image))
    text = "\n".join(pages_text)

    confirm_match = re.search(r"Inv.{0,4}?(\d{5,7})\D{0,3}\$?\s*([\d,]+\.\d{2})", text)
    date_match = re.search(r"(\w{3}\s+\w{3}\s+\d{1,2},\s*\d{4})", text)

    if not (confirm_match and date_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de "
            "Gold Coast Eagle."
        )

    invoice_no = int(confirm_match.group(1))
    invoice_date = datetime.strptime(date_match.group(1), "%a %b %d, %Y")
    amount = float(confirm_match.group(2).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_frito_lay_invoice(pdf_path):
    """
    Frito-Lay -- escaneo; puede traer una foto del cheque en una página
    aparte (se ignora, ninguno de los campos buscados aparece ahí). El
    total real es "TOTAL DUE", no el "GROSS SALES AMOUNT" de la tabla.

    Hay al menos dos plantillas distintas de recibo: "CASH SALE" (con
    "DATE: 27 Dec 2025" y "TOTAL DUE:") y "CHARGE SALES" (con la fecha
    suelta como "11/06/24" sin etiqueta, y "TOTAL DUE =" con igual en vez
    de dos puntos) -- se intentan ambos formatos.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = []
        for page in pdf.pages:
            image = _extract_page_image(page)
            if image is not None:
                pages_text.append(pytesseract.image_to_string(image))
    text = "\n".join(pages_text)

    invoice_match = re.search(r"INVOICE\s*#\s*[:\s]*(\d+)", text, re.IGNORECASE)
    total_match = re.search(r"TOTAL DUE\s*[:=]\s*\$?\s*([\d,]+\.\d{2})", text, re.IGNORECASE)

    date_match = re.search(r"DATE:\s*(\d{1,2}\s+\w{3}\s+\d{4})", text, re.IGNORECASE)
    if date_match:
        invoice_date = datetime.strptime(date_match.group(1), "%d %b %Y")
    else:
        date_match = re.search(r"\b(\d{1,2}/\d{1,2}/\d{2})\b", text)
        invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%y") if date_match else None

    if not (invoice_match and date_match and total_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de Frito-Lay."
        )

    invoice_no = int(invoice_match.group(1))
    amount = float(total_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_kings_invoice(pdf_path):
    """
    King's Wholesale Florists -- tanto la fecha/N° de invoice (tabla de
    arriba) como el total (cuadro "Total:" abajo a la derecha) viven en
    tablas con bordes que Tesseract confunde con el texto -- ambas franjas
    se recortan y se les quitan las líneas de grilla antes de OCR.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        image = _extract_page_image(pdf.pages[0])
    if image is None:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró la imagen escaneada de la factura King's."
        )

    info_crop = _crop_relative(image, 0.0, 0.29, 1.0, 0.35)
    info_text = pytesseract.image_to_string(_remove_grid_lines(info_crop))
    totals_crop = _crop_relative(image, 0.55, 0.72, 1.0, 0.97)
    totals_text = pytesseract.image_to_string(_remove_grid_lines(totals_crop))

    invoice_match = re.search(r"(\d{6})", info_text)
    date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", info_text)
    # "Total:" también matchea dentro de "Sub Total:" -- el importe final
    # siempre es la ÚLTIMA coincidencia (Sub Total, ..., Total, en ese orden).
    total_matches = re.findall(r"Total:\s*([\d,]+\.\d{2})", totals_text)

    if not (invoice_match and date_match and total_matches):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de King's."
        )

    invoice_no = int(invoice_match.group(1))
    invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%Y")
    amount = float(total_matches[-1].replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_red_bull_invoice(pdf_path):
    """
    Red Bull Distribution -- escaneo de una sola página. El renglón
    "TOTAL DUE" final a veces no lo lee Tesseract, pero el mismo número
    aparece en el cuadro TOTALS como "INVOICE" (Deposit/Tax siempre en
    $0.00 en las facturas vistas, así que es el mismo importe).
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        image = _extract_page_image(pdf.pages[0])
    if image is None:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró la imagen escaneada de la factura Red Bull."
        )
    text = pytesseract.image_to_string(image)

    invoice_match = re.search(r"\bInv\w{0,6}:\s*(\d{6,})", text)
    date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})\s+\d{1,2}:\d{2}\s*[AP]M", text)
    total_match = re.search(r"INVOICE\D*([\d,]+\.\d{2})", text)

    if not (invoice_match and date_match and total_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de Red Bull."
        )

    invoice_no = int(invoice_match.group(1))
    invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%Y")
    amount = float(total_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_sweetheart_invoice(pdf_path):
    """
    Sweetheart Ice Cream -- escaneo de una sola página con la info que hace
    falta; puede traer una foto del cheque en una segunda página (se
    ignora). "BALANCE DUE" a veces sale con un dígito mal leído -- se usa
    "TOTAL SALES", que es el mismo importe y lee mejor.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        image = _extract_page_image(pdf.pages[0])
    if image is None:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró la imagen escaneada de la factura Sweetheart."
        )
    text = pytesseract.image_to_string(image)

    invoice_match = re.search(r"INVOICE.{0,4}?(\d{8,})", text, re.IGNORECASE)
    date_match = re.search(r"Date:\s*(\d{1,2}/\d{1,2}/\d{4})", text)
    total_match = re.search(r"TOTAL SALES:\s*\$?\s*([\d,]+\.\d{2})", text, re.IGNORECASE)

    if not (invoice_match and date_match and total_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de Sweetheart."
        )

    invoice_no = int(invoice_match.group(1))
    invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%Y")
    amount = float(total_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_bimbo_invoice(pdf_path):
    """
    Bimbo Bakeries -- escaneo, puede traer una foto de cheque en otra
    página (se ignora). El total real ("TICKET TOTALS", ya neto de
    devoluciones) vive en una tabla con bordes:

    1) A veces hay una línea de confirmación al pie ("{invoice} {fecha}
       {total}") que trae fecha y total juntos y lee mejor que la tabla.
    2) Si esa línea quedó tapada por el recibo de "Paid Out" grapado
       encima (pasa seguido), se recorta la fila de "TICKET" -- ubicada
       dinámicamente, no a una altura fija, porque varía según la
       cantidad de ítems -- y se le quitan las líneas de grilla.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = []
        pages_images = []
        for page in pdf.pages:
            image = _extract_page_image(page)
            if image is not None:
                pages_images.append(image)
                pages_text.append(pytesseract.image_to_string(image))
    text = "\n".join(pages_text)

    invoice_match = re.search(r"INVOICE#\s*([\d\s]+?)\s*\n", text, re.IGNORECASE)
    if not invoice_match:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer el N° de invoice del PDF de Bimbo."
        )
    invoice_no = int(re.sub(r"\s+", "", invoice_match.group(1)))

    footer_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})[:.]*\s*([\d,]+)[.\s]+(\d{2})\b", text)
    if footer_match:
        invoice_date = datetime.strptime(footer_match.group(1), "%m/%d/%Y")
        amount = float(f"{footer_match.group(2).replace(',', '')}.{footer_match.group(3)}")
        return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}

    date_match = re.search(r"SDD[:;]?\s*(\d{1,2}/\d{1,2}/\d{4})", text, re.IGNORECASE)

    amount = None
    for image in pages_images:
        data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
        y = next(
            (data["top"][i] for i, word in enumerate(data["text"]) if word.strip().upper() == "TICKET"),
            None,
        )
        if y is None:
            continue
        crop = image.crop((0, max(0, y - 10), image.width, min(image.height, y + 220)))
        clean_text = pytesseract.image_to_string(_remove_grid_lines(crop, upscale=2))
        amounts = re.findall(r"([\d,]+\.\d{2})", clean_text)
        if amounts:
            amount = float(amounts[-1].replace(",", ""))
            break

    if not (date_match and amount is not None):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer fecha/total del PDF de Bimbo "
            "(si el recibo de Paid Out tapa el total, cargue esta factura manualmente)."
        )

    invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%Y")
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_midtown_invoice(pdf_path):
    """
    Midtown Wholesale -- "Sales Order" de 2 páginas + una foto de cheque,
    en cualquier orden entre las 3 (a veces el cheque va primero, a veces
    último), así que se juntan las 3 antes de buscar los campos.

    El "Notes:" / "Subtotal:" van en dos columnas separadas, y eso
    descoloca el orden de lectura del OCR (separa la etiqueta "Total" de
    su importe si se busca en el texto completo de la página) -- por eso
    se ubica "Subtotal" con coordenadas y se recorta esa columna sola.

    Limitación conocida: si el chofer tachó un ítem a mano y corrigió el
    Total a mano (le pasó al menos una vez), el OCR de esa columna sale
    ilegible (sin el punto decimal) y esta función falla con un error
    claro -- en ese caso hay que cargar la factura manualmente, nunca va
    a devolver el importe viejo por error.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = []
        totals_text = ""
        for page in pdf.pages:
            image = _extract_page_image(page)
            if image is None:
                continue
            pages_text.append(pytesseract.image_to_string(image))
            if totals_text:
                continue
            data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
            y = next((data["top"][i] for i, w in enumerate(data["text"]) if "Subtotal" in w), None)
            x = next((data["left"][i] for i, w in enumerate(data["text"]) if "Subtotal" in w), None)
            if y is not None:
                crop = image.crop((max(0, x - 50), max(0, y - 10), image.width, min(image.height, y + 250)))
                totals_text = pytesseract.image_to_string(crop)
    text = "\n".join(pages_text)

    invoice_match = re.search(r"Receipt #:?\s*(\d+)", text, re.IGNORECASE)
    date_match = re.search(r"Receipt Date\s*:?\s*(\d{1,2})-(\d{1,2})-(\d{4})", text, re.IGNORECASE)
    total_match = re.search(r"\bTotal\D{0,5}\$?\s*([\d,]+\.\d{2})", totals_text)

    if not (invoice_match and date_match and total_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de Midtown."
        )

    invoice_no = int(invoice_match.group(1))
    month, day, year = date_match.groups()
    invoice_date = datetime(int(year), int(month), int(day))
    amount = float(total_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_johnson_brothers_invoice(pdf_path):
    """
    Johnson Brothers of Florida -- puede venir en 1 o 2 páginas (el total
    real solo aparece en la última). El N° de invoice sale de "DOC {N}" al
    pie, salvo que el OCR lo lea como "boc" -- ahí se recurre a la fila de
    encabezado, anclada por la cuenta fija "146571", que trae fecha e
    invoice juntos (aunque a veces en orden invertido según la página).
    El total real está siempre pegado al texto fijo "APR 18%" del cargo
    por mora, nunca al "Gross Amount" de la tabla de arriba.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = []
        for page in pdf.pages:
            image = _extract_page_image(page)
            if image is not None:
                pages_text.append(pytesseract.image_to_string(image))
    text = "\n".join(pages_text)

    date_match = re.search(r"146571.{0,30}?(\d{1,2}/\d{1,2}/\d{2,4})", text)
    invoice_match = re.search(r"DOC\s*(\d{6,7})", text, re.IGNORECASE)
    if not invoice_match:
        invoice_match = re.search(
            r"146571.{0,30}?\d{1,2}/\d{1,2}/\d{2,4}.{0,20}?(\d{6,7})", text
        )
    amount_match = re.search(r"APR\s*18%\s+([\d,]+)\s*\.?\s*(\d{2})\b", text)

    if not (date_match and invoice_match and amount_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de "
            "Johnson Brothers."
        )

    invoice_no = int(invoice_match.group(1))
    invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%y")
    amount = float(f"{amount_match.group(1).replace(',', '')}.{amount_match.group(2)}")
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_flori_gas_invoice(pdf_path):
    """
    Flori-Gas -- formulario de matriz de punto (carbón), muy desgastado:
    el N° de invoice y la fecha impresos no se pueden leer de forma
    confiable ni recortando ni con más resolución. Por decisión del
    usuario, para ESTOS DOS CAMPOS se usa el nombre del archivo (formato
    "Invoice {N} {DD.MM.YYYY}.pdf"), que ya se viene usando como
    referencia confiable en todo este módulo.

    El monto sí se verifica contra el documento: sale del recibo de
    "Paid Out" grapado (Cash: $-X, igual al total de la factura en todos
    los casos vistos). Si no hay recibo de "Paid Out" en el PDF, falla
    con un error en vez de arriesgar leer mal el total escrito a mano.
    """
    filename_match = re.search(
        r"Invoice\s+(\d+)\s+(\d{1,2})\.(\d{1,2})\.(\d{4})",
        os.path.basename(pdf_path),
        re.IGNORECASE,
    )
    if not filename_match:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: el nombre del archivo no tiene el formato "
            '"Invoice {N} {DD.MM.YYYY}.pdf" esperado para Flori-Gas.'
        )
    invoice_no = int(filename_match.group(1))
    day, month, year = filename_match.group(2), filename_match.group(3), filename_match.group(4)
    invoice_date = datetime(int(year), int(month), int(day))

    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = []
        for page in pdf.pages:
            image = _extract_page_image(page)
            if image is not None:
                pages_text.append(pytesseract.image_to_string(image))
    text = "\n".join(pages_text)

    # El diseño en 2 columnas (recibo de Paid Out + factura lado a lado)
    # a veces separa la etiqueta "Cash:" de su valor en la lectura del
    # OCR, así que se busca directamente el patrón "$-X.XX" del recibo.
    amount_match = re.search(r"\$-\s*([\d,]+\.\d{2})", text)
    if not amount_match:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró el recibo de \"Paid Out\" para leer "
            "el monto de esta factura de Flori-Gas -- cárguela manualmente."
        )
    amount = float(amount_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


SUPPLIER_REGISTRY = {
    "ht_hackney": {
        "label": "H.T. Hackney",
        "sheet_name": "HT Hackney",
        "detect": lambda text: "h.t. hackney" in text.lower(),
        "extract": _extract_ht_hackney_invoice,
    },
    "cec": {
        "label": "CEC (Chinook Enterprises Corp.)",
        "sheet_name": "Chinook CEC",
        "detect": lambda text: "cec distributing" in text.lower(),
        "extract": _extract_cec_invoice,
    },
    "colonial": {
        "label": "Colonial Wholesale Dist. LLC",
        "sheet_name": "Colonial",
        "detect": lambda text: "colonial wholesale" in text.lower(),
        "extract": _extract_colonial_invoice,
    },
    "gce": {
        "label": "Gold Coast Eagle",
        "sheet_name": "GOLDCE",
        "detect": lambda text: "gold coast eagle" in text.lower(),
        "extract": _extract_gce_invoice,
    },
    "frito_lay": {
        "label": "Frito-Lay",
        "sheet_name": "FRITO-LAY",
        "detect": lambda text: "frito" in text.lower() and "lay" in text.lower(),
        "extract": _extract_frito_lay_invoice,
    },
    "kings": {
        "label": "King's Wholesale Florists",
        "sheet_name": "KING'S",
        "detect": lambda text: "wholesale florists" in text.lower(),
        "extract": _extract_kings_invoice,
    },
    "red_bull": {
        "label": "Red Bull Distribution Company",
        "sheet_name": "RED BULL",
        "detect": lambda text: "red bull distribution" in text.lower(),
        "extract": _extract_red_bull_invoice,
    },
    "sweetheart": {
        "label": "Sweetheart Ice Cream",
        "sheet_name": "SWEETHEART-ICE CREAM",
        "detect": lambda text: "sweetheart" in text.lower(),
        "extract": _extract_sweetheart_invoice,
    },
    "bimbo": {
        "label": "Bimbo Bakeries USA, Inc.",
        "sheet_name": "BIMBO",
        "detect": lambda text: "bimbo bakeries" in text.lower(),
        "extract": _extract_bimbo_invoice,
    },
    "midtown": {
        "label": "Midtown Wholesale LLC",
        "sheet_name": "MIDTOWN",
        "detect": lambda text: "midtown wholesale" in text.lower(),
        "extract": _extract_midtown_invoice,
    },
    "johnson": {
        "label": "Johnson Brothers of Florida",
        "sheet_name": "JOHNSON",
        "detect": lambda text: "johnson brothers" in text.lower(),
        "extract": _extract_johnson_brothers_invoice,
    },
    "flori_gas": {
        "label": "Flori-Gas",
        "sheet_name": "FLORI-GAS",
        "detect": lambda text: "305-637-9262" in text,
        "extract": _extract_flori_gas_invoice,
    },
}


def _detect_supplier(pdf_path):
    """
    Detecta el proveedor por el contenido del PDF: primero intenta con el
    texto digital (rápido, sin OCR); si el PDF es un escaneo sin capa de
    texto, recién ahí hace OCR -- de cada página hasta encontrar una
    coincidencia, porque algunos proveedores meten la foto del cheque
    ANTES que la factura (la primera página sola no alcanza).
    """
    _ensure_pdfplumber()
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not text.strip():
                image = _extract_page_image(page)
                if image is not None:
                    _ensure_pytesseract()
                    text = pytesseract.image_to_string(image)
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


def _find_last_invoice_row(sheet, from_row):
    """
    Última fila de tipo "invoice" (COMPROB) hasta from_row inclusive -- una
    fila "OP" (pago) no sirve de referencia de estilo porque su columna N
    normalmente queda en blanco, sin la negrita que sí llevan las facturas.
    """
    for row in range(from_row, 0, -1):
        comprob = sheet.cell(row=row, column=COL_COMPROB).value
        if isinstance(comprob, str) and comprob.strip().lower() == "invoice":
            return row
    return from_row


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

    # El N° de factura siempre va en negrita, sin importar el estilo que
    # haya traído style_ref_row.
    number_cell = sheet.cell(row=target_row, column=COL_NUMERO)
    number_font = number_cell.font
    number_cell.font = Font(
        name=number_font.name,
        size=number_font.sz,
        bold=True,
        italic=number_font.italic,
        color=number_font.color,
        underline=number_font.underline,
    )

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
        style_ref_row = _find_last_invoice_row(sheet, last_row)

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
