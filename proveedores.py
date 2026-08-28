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

# Pestaña de la hoja: celeste mientras el proveedor tenga saldo pendiente
# (le debemos), sin color cuando el saldo vuelve a 0. Mismo celeste que el
# usuario ya venía usando a mano en algunas hojas del libro real (ej.
# AIRGAS, KING'S) -- no es un color nuevo, solo se automatiza.
DEBT_TAB_COLOR = "FF00B0F0"


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


def _parse_mmdd_year_flexible(date_text):
    """
    Parsea "M/D/AA" o "M/D/AAAA" con el mismo patrón -- varios extractores
    permiten año de 2 o 4 dígitos en su regex pero llamaban a strptime con
    "%y" fijo, así que un año de 4 dígitos rompía con un ValueError críptico
    de la librería estándar en vez del mensaje de error propio de la función.
    """
    for fmt in ("%m/%d/%y", "%m/%d/%Y"):
        try:
            return datetime.strptime(date_text, fmt)
        except ValueError:
            continue
    raise ValueError(f'No se pudo interpretar la fecha "{date_text}".')


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
    invoice_date = _parse_mmdd_year_flexible(date_match.group(1))
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
    invoice_date = _parse_mmdd_year_flexible(date_match.group(1))
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

    # Anclar la línea de confirmación al pie ("{invoice} {fecha} {total}",
    # ver docstring) al propio N° de invoice ya confirmado -- sin esto, el
    # patrón "fecha seguida de números" no tiene ninguna palabra ancla y
    # puede matchear una fecha/monto de la tabla de ítems antes de llegar
    # a la línea real de confirmación. Se toleran espacios sueltos de OCR
    # entre los dígitos del invoice, como en el resto del texto.
    invoice_digits_pattern = r"\s*".join(re.escape(digit) for digit in str(invoice_no))
    footer_match = re.search(
        invoice_digits_pattern + r"\s*(\d{1,2}/\d{1,2}/\d{4})[:.]*\s*([\d,]+)[.\s]+(\d{2})\b",
        text,
    )
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
    invoice_date = _parse_mmdd_year_flexible(date_match.group(1))
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


def _extract_airgas_invoice(pdf_path):
    """
    Airgas National Carbonation -- texto digital (no escaneado), dos tipos
    de factura ("STANDARD INVOICE" de gas y "CYLINDER RENTAL INVOICE" de
    alquiler de tanque), ambos con la misma tabla resumen "INVOICE DATE
    PAYER INVOICE NO. DUE DATE PAY THIS AMOUNT".

    Trampa real: en facturas de alquiler con débito automático, "PAY THIS
    AMOUNT" a veces muestra $0.00 (no queda nada por pagar aparte del
    débito), pero el importe real de la factura es el que aparece más
    abajo en el campo "AMOUNT" (junto a "Sales Tax"). Por eso se toma
    siempre la ÚLTIMA coincidencia de "AMOUNT" en el texto, nunca la
    primera.
    """
    _ensure_pdfplumber()
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    header_match = re.search(
        r"(\d{1,2}/\d{1,2}/\d{4})\s+\d{6,8}\s+(\d{8,10})\s+\d{1,2}/\d{1,2}/\d{4}\s+\$?\s*[\d,]+\.\d{2}",
        text,
    )
    amount_matches = re.findall(r"AMOUNT\D{0,6}?\$?\s*([\d,]+\.\d{2})", text)

    if not (header_match and amount_matches):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de Airgas."
        )

    invoice_no = int(header_match.group(2))
    invoice_date = datetime.strptime(header_match.group(1), "%m/%d/%Y")
    amount = float(amount_matches[-1].replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_az_invoice(pdf_path):
    """
    AZ Southeast Distributors -- vale de entrega escaneado (no factura
    formal), con bastante ruido visual (grapas, dobleces, cheque o recibo
    de "Paid Out" superpuesto). El N° de factura es el "DELIVERY NO." de
    la cajita de arriba a la izquierda -- el CUST# de esa misma cajita
    (siempre el mismo número en todas las facturas del cliente) NO es el
    N° de factura, y el nombre de archivo a veces usa por error ese
    número en vez del DELIVERY NO. real.

    La fecha normalmente se lee bien del texto completo de la página, pero
    en algún escaneo esa cajita sale en blanco del todo -- ahí se recurre
    al recorte de la cajita con la grilla removida, donde sí aparece. El
    monto ("SUB TOTAL") casi nunca se lee del texto completo (queda
    perdido entre el cheque o el recibo de Paid Out superpuestos) -- se
    recorta la cajita de totales de abajo a la derecha con la grilla
    removida, ahí sí es consistente.

    LÍMITE CONOCIDO IMPORTANTE: a diferencia de los demás proveedores de
    este módulo, este extractor falla en la MAYORÍA de las facturas reales
    probadas (~5 de 32 en una validación amplia contra el archivo del
    Drive, sin importar el año) -- el recorte fijo de la cajita de totales
    no encuentra "SUB TOTAL" en muchos escaneos porque esa franja de la
    imagen varía de tamaño/posición según cómo quedó pegado el cheque o el
    recibo de "Paid Out" sobre el vale, o directamente el texto no es
    legible ahí. Cuando falla, falla limpio (no carga un valor incorrecto),
    pero la tasa de éxito real es baja -- conviene tratarlo como un
    complemento que a veces ahorra la carga manual, no como confiable en
    general.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        full_text = page.extract_text() or ""
        image = _extract_page_image(page)
    if image is None:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró la imagen escaneada del vale de "
            "AZ Southeast."
        )
    if not full_text.strip():
        full_text = pytesseract.image_to_string(image)

    top_crop = _remove_grid_lines(_crop_relative(image, 0.0, 0.0, 0.25, 0.10), upscale=4)
    top_text = pytesseract.image_to_string(top_crop)
    # Preferir el número anclado a la etiqueta "DELIVERY NO." -- el CUST#
    # (siempre el mismo número, no es la factura) convive en la misma
    # cajita recortada, así que un (\d{9}) suelto puede tomar cualquiera
    # de los dos según el orden de lectura del OCR. Si la etiqueta no se
    # lee (ruido/grapas), se cae al primer número de 9 dígitos como antes.
    invoice_match = re.search(r"DELIVERY\s*NO\.?\s*[:.]?\s*(\d{9})", top_text, re.IGNORECASE)
    if not invoice_match:
        invoice_match = re.search(r"(\d{9})", top_text)

    date_match = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", full_text)
    if not date_match:
        date_match = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", top_text)

    bottom_crop = _remove_grid_lines(_crop_relative(image, 0.55, 0.68, 1.0, 0.92), upscale=3)
    bottom_text = pytesseract.image_to_string(bottom_crop)
    total_match = re.search(r"TOTAL\D{0,10}?([\d,]+\.\d{2})", bottom_text)

    if not (invoice_match and date_match and total_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del vale de "
            "AZ Southeast."
        )

    invoice_no = int(invoice_match.group(1))
    invoice_date = datetime(int(date_match.group(3)), int(date_match.group(1)), int(date_match.group(2)))
    amount = float(total_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_express_beverage_invoice(pdf_path):
    """
    Express Beverage of Tampa -- escaneo de una sola página, impreso
    prolijo (no manuscrito). N°/fecha viven en la tabla "Date | Invoice #"
    de arriba a la derecha; el total en la celda "Total" de abajo a la
    derecha, justo arriba del cheque grapado. El N° de factura a veces
    sale con espacios sueltos entre dígitos por el OCR (ej. "01 371 1"),
    por eso se toma todo el resto de esa línea después de la fecha y se
    le quitan los espacios, en vez de exigir un ancho fijo de dígitos.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        image = _extract_page_image(pdf.pages[0])
    if image is None:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró la imagen escaneada de la factura "
            "Express Beverage."
        )

    top_crop = _crop_relative(image, 0.68, 0.08, 1.0, 0.17)
    top_text = pytesseract.image_to_string(top_crop)
    date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", top_text)
    invoice_match = None
    if date_match:
        line_after_date = top_text[date_match.end():].split("\n", 1)[0]
        invoice_match = re.search(r"(\d[\d\s]{3,7}\d)", line_after_date)

    total_crop = _crop_relative(image, 0.55, 0.69, 1.0, 0.80)
    total_text = pytesseract.image_to_string(total_crop)
    total_match = re.search(r"([\d,]+\.\d{2})", total_text)
    if total_match is None:
        total_text = pytesseract.image_to_string(_remove_grid_lines(total_crop, upscale=3))
        total_match = re.search(r"([\d,]+\.\d{2})", total_text)

    if not (date_match and invoice_match and total_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de "
            "Express Beverage."
        )

    invoice_no = int(re.sub(r"\s+", "", invoice_match.group(1)))
    invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%Y")
    amount = float(total_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_kooler_ice_invoice(pdf_path):
    """
    Kooler Ice, Inc. -- mezcla facturas con texto digital y escaneos según
    el pedido, y al menos dos plantillas: los pedidos normales/reparación/
    suscripción usan N° "INV{n}" y terminan la tabla resumen en "Applied
    Deposit" (que en esos casos coincide con el importe real -- se usa en
    vez de "Total", que en las facturas de suscripción queda separado y en
    $0.00, el saldo pendiente, no el importe de la factura); los pedidos
    de equipos usan N° "SO{n}" y no tienen "Applied Deposit" -- ahí el
    importe real sí está directamente en "Total".

    Validado contra el archivo histórico del Drive: ~75% de éxito (9/12).
    Las fallas restantes vistas fueron escaneos con el papel arrugado/
    doblado justo sobre el N° de factura ("#INV...") -- el resto del texto
    (fecha, importe) se lee bien, pero sin el N° no hay forma segura de
    cargar la fila; falla limpio en esos casos.
    """
    _ensure_pdfplumber()
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""
        if not text.strip():
            _ensure_pytesseract()
            image = _extract_page_image(page)
            if image is None:
                raise ValueError(
                    f"{os.path.basename(pdf_path)}: no se encontró la imagen escaneada de la "
                    "factura de Kooler Ice."
                )
            text = pytesseract.image_to_string(image)

    invoice_match = re.search(r"(?:INV|SO)\s*(\d{3,7})", text, re.IGNORECASE)
    date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text)
    amount_match = re.search(r"Applied Deposit\D{0,6}\$?\s*([\d,]+\.\d{2})", text, re.IGNORECASE)
    if not amount_match:
        amount_match = re.search(r"\bTotal\D{0,6}\$?\s*([\d,]+\.\d{2})", text, re.IGNORECASE)

    if not (invoice_match and date_match and amount_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de "
            "Kooler Ice."
        )

    invoice_no = int(invoice_match.group(1))
    invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%Y")
    amount = float(amount_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_sams_club_invoice(pdf_path):
    """
    Sam's Club -- recibo de caja escaneado (no factura formal). El N° de
    "factura" es el número de 4 dígitos que sigue a la hora en el
    encabezado del recibo (ej. "01/06/26 09:40 5495 08201 002 3909" -> N°
    5495), el mismo número que ya usa el usuario en el nombre de archivo.

    El monto se prefiere del recibo de "Paid Out" grapado ("Cash: $-X"),
    porque la columna de TOTAL del recibo original suele leerse mal (los
    dígitos de precio de la tabla de ítems quedan mezclados verticalmente
    por el OCR) -- pero no todas las compras tienen ese recibo (ej. una
    compra sin impuesto, solo alimentos, puede no llevarlo), así que si no
    aparece se cae al campo "TOTAL" leído directamente. El recibo y el
    "Paid Out" a veces quedan en páginas separadas del PDF (no una sola
    imagen combinada), por eso se juntan todas las páginas antes de buscar.

    Riesgo conocido (mismo tipo que Red Bull): en un caso Tesseract
    confundió un solo dígito del N° de recibo (5 leído como 6) de forma
    consistente sin importar el recorte o preprocesamiento probado --
    conviene revisar visualmente antes de guardar.

    Validado contra el archivo histórico del Drive: ~68% de éxito (13/19).
    Las fallas restantes fueron recibos viejos con ruido de OCR puntual
    (ej. la hora del encabezado ilegible, lo que rompe el patrón fecha+N°)
    -- fallan limpio, no insertan un valor incorrecto.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = []
        for page in pdf.pages:
            image = _extract_page_image(page)
            if image is not None:
                pages_text.append(pytesseract.image_to_string(image))
    if not pages_text:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró la imagen escaneada del recibo de "
            "Sam's Club."
        )
    text = "\n".join(pages_text)

    header_match = re.search(r"(\d{2}/\d{2}/\d{2})\s+\d{1,2}:\d{2}\s+(\d{4})\s+\d{5}", text)
    amount_match = re.search(r"\$-\s*([\d,]+\.\d{2})", text)
    if not amount_match:
        amount_match = re.search(r"\bTOTAL\D{0,10}?([\d,]+\.\d{2})", text, re.IGNORECASE)

    if not (header_match and amount_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del recibo de "
            "Sam's Club."
        )

    invoice_no = int(header_match.group(2))
    invoice_date = datetime.strptime(header_match.group(1), "%m/%d/%y")
    amount = float(amount_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_fs_wholesale_invoice(pdf_path):
    """
    FS Wholesale -- en realidad factura Florida Smokes Wholesale, LLC.
    Escaneo de 2-3 páginas (puede traer foto de cheque en una página
    aparte, se ignora): el N°/fecha viven en el encabezado de la primera
    página ("05 Nov 2025 98836 10 BRADENTON GAS STATION USA" -- fecha,
    invoice, cantidad de ítems, cliente, todo en la misma fila), pero el
    total real ("Tax Invoice Total (USD)") está en la página de Payments,
    no en la primera -- se buscan ambos patrones en el texto combinado de
    todas las páginas.
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

    header_match = re.search(r"(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})\s+(\d{5,7})\s+\d+", text)
    total_match = re.search(
        r"(?:Tax\s+)?Invoice Total\s*\(USD\):\s*\$?\s*([\d,]+\.\d{2})", text, re.IGNORECASE
    )

    if not (header_match and total_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de "
            "FS Wholesale."
        )

    invoice_no = int(header_match.group(2))
    invoice_date = datetime.strptime(header_match.group(1), "%d %b %Y")
    amount = float(total_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_lmt_invoice(pdf_path):
    """
    LMT Trading Company LLC -- escaneo, y el PDF casi siempre agrupa la
    factura real de LMT con una de J.J. Taylor Dist. FL, Inc. (comparten
    domicilio, teléfono y plantilla de reparto) en cualquier orden entre
    1-2 páginas. J.J. Taylor está pausado por pedido del usuario, así que
    se usa SOLO la página que trae "Paylink - LMT" (la de J.J. Taylor
    trae "Paylink - JT") -- el ancla que más limpio lee de todo el
    encabezado, incluso cuando "LMT Trading Company" sale muy garabateado.

    El N°/fecha del nombre de archivo ("[LMT. ]Invoice {N} {DD.MM.YYYY}.pdf")
    y los impresos en el encabezado del documento ("{fecha} {N} {load
    sheet}...") normalmente coinciden, pero NINGUNO de los dos es
    confiable al 100% por separado -- se encontró un caso real donde el
    nombre de archivo tenía un N°/fecha completamente distintos a los del
    documento y el libro real (archivo mal nombrado), y otro caso real
    donde el documento leyó un dígito de más en el N° por OCR. Por eso se
    cruzan ambas fuentes cuando el encabezado del documento se puede leer:
    si coinciden, hay confianza; si no, se falla con un error claro en vez
    de adivinar cuál de las dos está bien (si el documento no se puede
    leer, se sigue confiando solo en el nombre de archivo, como Flori-Gas).

    El total ("Sub Total"/"Total", mismo importe repetido) se lee del
    documento: se ubica la fila por posición vía pytesseract.image_to_data
    (ancla "Sub", con "Total" de respaldo si "Sub" no se detectó como
    palabra aparte) y se recorta esa franja para un OCR más limpio.
    """
    filename_match = re.search(
        r"Invoice\s+(\d+)\s+(\d{1,2})\.(\d{1,2})\.(\d{4})",
        os.path.basename(pdf_path),
        re.IGNORECASE,
    )
    if not filename_match:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: el nombre del archivo no tiene el formato "
            '"Invoice {N} {DD.MM.YYYY}.pdf" esperado para LMT.'
        )
    invoice_no = int(filename_match.group(1))
    day, month, year = filename_match.group(2), filename_match.group(3), filename_match.group(4)
    invoice_date = datetime(int(year), int(month), int(day))

    _ensure_pdfplumber()
    _ensure_pytesseract()
    lmt_image = None
    lmt_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            image = _extract_page_image(page)
            if image is None:
                continue
            text = pytesseract.image_to_string(image)
            if re.search(r"paylink\s*-\s*lmt", text, re.IGNORECASE):
                lmt_image = image
                lmt_text = text
                break
    if lmt_image is None:
        raise ValueError(
            f'{os.path.basename(pdf_path)}: no se encontró la página de LMT (con "Paylink - LMT") '
            "en este PDF."
        )

    header_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})\s+(\d{5,8})\b", lmt_text)
    if header_match:
        doc_invoice_no = int(header_match.group(2))
        if doc_invoice_no != invoice_no:
            raise ValueError(
                f"{os.path.basename(pdf_path)}: el N° de factura del nombre de archivo "
                f"({invoice_no}) no coincide con el N° leído del documento ({doc_invoice_no}) "
                "-- revise cuál es el correcto y cárguela manualmente."
            )
        try:
            doc_date = datetime.strptime(header_match.group(1), "%m/%d/%Y")
        except ValueError:
            doc_date = None
        if doc_date is not None and doc_date.date() != invoice_date.date():
            raise ValueError(
                f"{os.path.basename(pdf_path)}: la fecha del nombre de archivo "
                f"({invoice_date:%d/%m/%Y}) no coincide con la fecha leída del documento "
                f"({doc_date:%d/%m/%Y}) -- revise cuál es la correcta y cárguela manualmente."
            )

    upscaled = lmt_image.resize((lmt_image.width * 2, lmt_image.height * 2), Image.LANCZOS)
    data = pytesseract.image_to_data(upscaled, output_type=pytesseract.Output.DICT)

    sub_ys = [
        data["top"][i]
        for i, w in enumerate(data["text"])
        if w.strip().lower() in ("sub", "suo", "sud", "sob")
    ]
    total_ys = [
        data["top"][i]
        for i, w in enumerate(data["text"])
        if "otal" in w.lower() or w.strip().lower() in ("tol", "oul", "onl", "toul")
    ]

    amount = None
    if sub_ys:
        y = max(sub_ys)
        crop = upscaled.crop((0, max(0, y - 15), upscaled.width, min(upscaled.height, y + 150)))
        amount_match = re.search(r"([\d,]+\.\d{2})", pytesseract.image_to_string(crop))
        if amount_match:
            amount = float(amount_match.group(1).replace(",", ""))
    if amount is None and total_ys:
        y = max(total_ys)
        crop = upscaled.crop((0, max(0, y - 30), upscaled.width, min(upscaled.height, y + 250)))
        amount_match = re.search(r"([\d,]+\.\d{2})", pytesseract.image_to_string(crop))
        if amount_match:
            amount = float(amount_match.group(1).replace(",", ""))

    if amount is None:
        raise ValueError(f"{os.path.basename(pdf_path)}: no se pudo leer el total del PDF de LMT.")

    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_overflow_invoice(pdf_path):
    """
    Overflow Group Distribution -- escaneo de una sola página, texto
    limpio. N°/fecha en el encabezado ("Invoice #1500034491" /
    "11/13/2025 2:45:03 PM"), monto en "Total Sales:" (igual a "Balance:"
    en las facturas vistas, sin pagos parciales).
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        image = _extract_page_image(pdf.pages[0])
    if image is None:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró la imagen escaneada de la factura Overflow."
        )
    text = pytesseract.image_to_string(image)

    invoice_match = re.search(r"Invoice\s*#\s*(\d+)", text, re.IGNORECASE)
    date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})\s+\d{1,2}:\d{2}:\d{2}\s*[AP]M", text)
    total_match = re.search(r"Total Sales:\s*\$?\s*([\d,]+\.\d{2})", text, re.IGNORECASE)

    if not (invoice_match and date_match and total_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de Overflow."
        )

    invoice_no = int(invoice_match.group(1))
    invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%Y")
    amount = float(total_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_swisher_invoice(pdf_path):
    """
    Swisher -- recibo de caja escaneado (no factura formal, como Sam's
    Club/AZ). El N° es el "Receipt #" del encabezado (a veces con un
    espacio de más en medio, se le saca). La fecha viene en dos formatos
    según la antigüedad de la factura ("8/28/24 10:14 AM" en las viejas,
    "2026-02-23" suelto en las nuevas) -- se prueban ambos. El monto se
    prefiere del recibo de "Paid Out" grapado ("Cash: $-X", con alguna
    coma suelta de OCR que se descarta) -- el "Total"/"Cash Due Now" de la
    tabla principal casi siempre sale con la parte decimal incompleta.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        image = _extract_page_image(pdf.pages[0])
    if image is None:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró la imagen escaneada del recibo de Swisher."
        )
    text = pytesseract.image_to_string(image)

    invoice_match = re.search(r"Receipt\s*#?\s*(\d[\d\s]{5,20}\d)", text, re.IGNORECASE)
    date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{2})\s+\d{1,2}:\d{2}\s*[AP]M", text)
    if date_match:
        invoice_date = datetime.strptime(date_match.group(1), "%m/%d/%y")
    else:
        date_match = re.search(r"(\d{4}-\d{2}-\d{2})", text)
        invoice_date = datetime.strptime(date_match.group(1), "%Y-%m-%d") if date_match else None
    amount_match = re.search(r"Cash:\s*/?\s*\$-\s*([\d,]+\.\d{2})", text, re.IGNORECASE)

    if not (invoice_match and date_match and amount_match):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del recibo de Swisher."
        )

    invoice_no = int(re.sub(r"\s+", "", invoice_match.group(1)))
    amount = float(amount_match.group(1).replace(",", ""))
    return {"invoice_no": invoice_no, "date": invoice_date, "amount": amount}


def _extract_signarama_invoice(pdf_path):
    """
    Signarama (Bradenton Signs) -- escaneo, puede traer foto de cheque en
    página aparte (se ignora). N°/fecha salen limpios del texto completo
    ("IN INV-8148", "Created Date: 7/17/2026" -- coincide con el nombre
    de archivo en casi todos los casos vistos; se prefiere sobre
    "Generated On", que puede ser muchos días posterior por una reimpresión).

    El "Grand Total" vive en una plantilla a 2 columnas (etiquetas a la
    izquierda, importes a la derecha) que la mayoría de las veces
    descoloca el valor del OCR de página completa -- se ubica la palabra
    "Grand" (o "Total" de respaldo) por posición con
    pytesseract.image_to_data y se recorta esa fila hacia la derecha, con
    más resolución, para leer el importe.
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    signarama_image = None
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            image = _extract_page_image(page)
            if image is None:
                continue
            page_text = pytesseract.image_to_string(image)
            if "bradentonsigns" in page_text.lower() or "inv-" in page_text.lower():
                signarama_image = image
                text = page_text
                break
    if signarama_image is None:
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se encontró la página de factura de Signarama."
        )

    invoice_match = re.search(r"INV[-\s]?(\d{3,6})", text, re.IGNORECASE)
    date_match = re.search(r"Created Date:\s*(\d{1,2})/(\d{1,2})/(\d{4})", text, re.IGNORECASE)

    data = pytesseract.image_to_data(signarama_image, output_type=pytesseract.Output.DICT)
    label_idx = next((i for i, w in enumerate(data["text"]) if w.strip().lower() == "grand"), None)
    if label_idx is None:
        label_idx = next((i for i, w in enumerate(data["text"]) if "total" in w.lower()), None)

    amount = None
    if label_idx is not None:
        top = data["top"][label_idx]
        height = data["height"][label_idx]
        left = data["left"][label_idx]
        crop = signarama_image.crop(
            (left, max(0, top - 10), signarama_image.width, top + height + 25)
        )
        upscaled = crop.resize((crop.width * 3, crop.height * 3), Image.LANCZOS)
        amount_match = re.search(r"([\d,]+\.\d{2})", pytesseract.image_to_string(upscaled))
        if amount_match:
            amount = float(amount_match.group(1).replace(",", ""))

    if not (invoice_match and date_match and amount is not None):
        raise ValueError(
            f"{os.path.basename(pdf_path)}: no se pudo leer invoice/fecha/total del PDF de Signarama."
        )

    invoice_no = int(invoice_match.group(1))
    invoice_date = datetime(int(date_match.group(3)), int(date_match.group(1)), int(date_match.group(2)))
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
    "airgas": {
        "label": "Airgas National Carbonation",
        "sheet_name": "AIRGAS",
        "detect": lambda text: "airgas" in text.lower(),
        "extract": _extract_airgas_invoice,
    },
    "az": {
        "label": "AZ Southeast Distributors LLC",
        "sheet_name": "AZ Sout",
        "detect": lambda text: "az southeast distributors" in text.lower(),
        "extract": _extract_az_invoice,
    },
    "express": {
        "label": "Express Beverage of Tampa",
        "sheet_name": "EXPRESS ",
        "detect": lambda text: "express beverage" in text.lower(),
        "extract": _extract_express_beverage_invoice,
    },
    "kooler_ice": {
        "label": "Kooler Ice, Inc.",
        "sheet_name": "KOOLER ICE",
        "detect": lambda text: "kooler ice" in text.lower(),
        "extract": _extract_kooler_ice_invoice,
    },
    "sams_club": {
        "label": "Sam's Club",
        "sheet_name": "SAM'S",
        "detect": lambda text: re.search(r"sam.?s\s*club", text.lower()) is not None,
        "extract": _extract_sams_club_invoice,
    },
    "fs_wholesale": {
        "label": "FS Wholesale (Florida Smokes Wholesale, LLC)",
        "sheet_name": "FS WHOLESALE",
        "detect": lambda text: "florida smokes" in text.lower(),
        "extract": _extract_fs_wholesale_invoice,
    },
    "lmt": {
        "label": "LMT Trading Company LLC",
        "sheet_name": "LMT",
        "detect": lambda text: re.search(r"paylink\s*-\s*lmt", text.lower()) is not None,
        "extract": _extract_lmt_invoice,
    },
    "overflow": {
        "label": "Overflow Group Distribution",
        "sheet_name": "OVERFLOW",
        "detect": lambda text: "overflowgroupdistribution" in text.lower().replace(" ", ""),
        "extract": _extract_overflow_invoice,
    },
    "swisher": {
        "label": "Swisher",
        "sheet_name": "SWISHER",
        "detect": lambda text: "swisher" in text.lower(),
        "extract": _extract_swisher_invoice,
    },
    "signarama": {
        "label": "Signarama (Bradenton Signs)",
        "sheet_name": "SIGNARAMA",
        "detect": lambda text: "bradentonsigns" in text.lower(),
        "extract": _extract_signarama_invoice,
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


def _update_sheet_tab_color(sheet):
    """
    Pinta la pestaña de la hoja de celeste si el proveedor tiene saldo
    pendiente (le debemos), y la despinta si el saldo llega a 0 (ej. al
    cargar un pago desde Chase en la Fase 2 del módulo). El saldo se
    calcula sumando DEBE y HABER directamente en vez de leer la fórmula de
    BALANCE, porque openpyxl no evalúa fórmulas y la fila recién agregada
    todavía no tiene un valor calculado en caché.
    """
    last_row = _find_last_real_row(sheet)
    total_debe = 0.0
    total_haber = 0.0
    for row in range(1, (last_row or 0) + 1):
        debe = sheet.cell(row=row, column=COL_DEBE).value
        haber = sheet.cell(row=row, column=COL_HABER).value
        if isinstance(debe, (int, float)):
            total_debe += debe
        if isinstance(haber, (int, float)):
            total_haber += haber
    balance = total_debe - total_haber
    sheet.sheet_properties.tabColor = DEBT_TAB_COLOR if balance > 0.005 else None


def _existing_invoice_numbers(sheet):
    numbers = set()
    for row in range(1, sheet.max_row + 1):
        comprob = sheet.cell(row=row, column=COL_COMPROB).value
        if not (isinstance(comprob, str) and comprob.strip().lower() == "invoice"):
            continue
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
    # La fila anterior a target_row puede ser el separador en blanco entre
    # meses (sin fórmula de balance) -- siempre referenciar last_row, la
    # última fila real con saldo, para no perder el arrastre al cambiar de mes.
    sheet.cell(
        row=target_row,
        column=COL_BALANCE,
        value=f"=+{sheet.cell(row=last_row, column=COL_BALANCE).coordinate}"
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

        _update_sheet_tab_color(sheet)

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
