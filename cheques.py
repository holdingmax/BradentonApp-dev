"""
Lectura de cheques propios (Bradenton Gas Station USA LLC, cuenta Chase) en
PDFs escaneados/fotografiados -- sueltos o grapados junto a la factura de un
proveedor.

Lo único que importa leer es el N° de cheque impreso en la esquina superior
derecha (ej. "1775"). Todo lo demás del cheque es manuscrito y no se intenta
leer.

Cómo se ubica: se busca el encabezado impreso "BRADENTON GAS STATION" (el
ancla más confiable de todo el cheque) y el N° se lee en una zona relativa a
ese encabezado (a la derecha y un poco más arriba). Así funciona aunque la
foto tenga la factura y el cheque en la misma página, y aunque el cheque venga
girado (se prueban las 4 orientaciones si la corrección automática de
Tesseract no alcanza).

Regla de oro del proyecto: si el N° no se puede leer con confianza, el cheque
se devuelve igual pero con number=None, para completarlo a mano -- nunca se
inventa un número.
"""

import io
import re

from ocr_utils import (
    Image,
    correct_image_orientation,
    ensure_pdfplumber,
    ensure_pytesseract,
    pdfplumber,
    pytesseract,
    remove_grid_lines,
)

# Rango razonable de un N° de cheque de esta chequera (hoy ~1775). Evita
# confundir con fragmentos del código postal, la fecha o el routing number.
_MIN_CHECK_NUMBER = 100
_MAX_CHECK_NUMBER = 99999

# Confusiones típicas de Tesseract con la fuente del N° de cheque.
_DIGIT_FIXES = str.maketrans({"L": "1", "l": "1", "I": "1", "|": "1", "O": "0", "o": "0"})


def _page_images(page):
    """Imágenes candidatas de una página: la más grande incrustada, o la página renderizada."""
    if page.images:
        biggest = max(page.images, key=lambda im: im["width"] * im["height"])
        try:
            return Image.open(io.BytesIO(biggest["stream"].get_data())).convert("RGB")
        except Exception:
            pass
    try:
        return page.to_image(resolution=200).original.convert("RGB")
    except Exception:
        return None


def _ocr_words(image):
    data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
    return [
        (data["text"][i].strip().upper(), data["left"][i], data["top"][i],
         data["width"][i], data["height"][i])
        for i in range(len(data["text"]))
        if data["text"][i].strip()
    ]


def _find_anchors(image):
    """
    Devuelve [(left, top, span)] por cada cheque propio encontrado en la
    imagen. El encabezado "BRADENTON GAS STATION" solo no alcanza -- las
    facturas de los proveedores lo traen igual en "Bill To" -- así que además
    se exige la dirección impresa del cheque justo debajo ("17375 COLLINS
    AVE ... NORTH MIAMI BEACH", la del dueño -- las facturas van a la
    dirección de la estación en Bradenton).
    """
    words = _ocr_words(image)
    anchors = []
    for idx, (text, left, top, width, height) in enumerate(words):
        if "BRADENTON" not in text:
            continue
        following = " ".join(w[0] for w in words[idx + 1: idx + 4])
        if "GAS" not in following and "STATION" not in following:
            continue
        # Escala del cheque = ancho del título "BRADENTON GAS STATION USA LLC".
        # Se mide hasta el final de "STATION" (las cajas de Tesseract para
        # "USA LLC" a veces se pegan con basura de al lado) y se extrapola:
        # "BRADENTON GAS STATION" es ~75% del título en todos los cheques.
        station = next((w for w in words[idx + 1: idx + 4] if "STATION" in w[0]), None)
        if station is not None and station[1] > left:
            span = (station[1] + station[3] - left) / 0.75
        else:
            span = width / 0.42
        below = [w for w in words if top < w[2] <= top + 0.25 * span and left - span * 0.3 <= w[1] <= left + span * 1.3]
        if not any("COLLINS" in w[0] or "MIAMI" in w[0] for w in below):
            continue
        if any(abs(a[0] - left) < span * 0.2 and abs(a[1] - top) < span * 0.2 for a in anchors):
            continue
        anchors.append((left, top, span))
    return anchors


def _clean_number_tokens(text):
    candidates = []
    for token in re.split(r"\s+", text):
        token = token.strip(".,:;'\"()[]{}-_")
        if not token:
            continue
        fixed = token.translate(_DIGIT_FIXES)
        if re.fullmatch(r"\d{3,5}", fixed):
            value = int(fixed)
            if _MIN_CHECK_NUMBER <= value <= _MAX_CHECK_NUMBER:
                candidates.append(value)
    return candidates


# Zonas donde buscar el N° (fracciones del ancho del encabezado, relativas a
# su esquina izquierda). Medidas sobre cheques reales: el N° queda a ~2.6
# anchos a la derecha y apenas por encima del encabezado. Se prueban varias
# (más ajustadas y más amplias) porque el borde decorativo de arriba del
# cheque confunde a Tesseract si entra en el recorte.
_NUMBER_ZONES = (
    (2.1, -0.13, 3.05, 0.02),
    (2.2, -0.10, 3.05, 0.04),
    (2.0, -0.18, 3.10, 0.05),
    (1.9, -0.25, 3.20, 0.10),
)


def _read_number_near_anchor(image, anchor):
    left, top, span = anchor
    w, h = image.size
    votes = {}
    for x0, y0, x1, y1 in _NUMBER_ZONES:
        box = (
            max(0, int(left + x0 * span)), max(0, int(top + y0 * span)),
            min(w, int(left + x1 * span)), min(h, int(top + y1 * span)),
        )
        if box[2] - box[0] < 20 or box[3] - box[1] < 10:
            continue
        crop = image.crop(box)
        crop = crop.resize((crop.width * 2, crop.height * 2), Image.LANCZOS)
        variants = [crop]
        try:
            variants.append(remove_grid_lines(crop, upscale=1))
        except Exception:
            pass
        for variant in variants:
            text = pytesseract.image_to_string(variant, config="--psm 11")
            for value in set(_clean_number_tokens(text)):
                votes[value] = votes.get(value, 0) + 1
    if not votes:
        return None
    best = sorted(votes.items(), key=lambda kv: -kv[1])
    # Dos candidatos distintos con los mismos votos: no se adivina.
    if len(best) > 1 and best[0][1] == best[1][1]:
        return None
    return best[0][0]


def _crop_check(image, anchor):
    """Recorte aproximado del cheque completo, para guardarlo derecho aparte."""
    left, top, span = anchor
    w, h = image.size
    box = (
        max(0, int(left - 0.3 * span)),
        max(0, int(top - 0.25 * span)),
        min(w, int(left + 3.15 * span)),
        min(h, int(top + 1.0 * span)),
    )
    return image.crop(box)


def _orientations(image):
    """La orientación que sugiere Tesseract primero, después el resto."""
    first = correct_image_orientation(image)
    yield first
    for angle in (90, 180, 270):
        yield first.rotate(angle, expand=True)


def extract_checks_from_pdf(pdf_path):
    """
    Busca cheques propios en cada página del PDF. Devuelve una lista de
    {"page": índice, "number": int o None, "image": PIL.Image del cheque ya derecho}.
    Una página sin cheque no aporta nada.
    """
    ensure_pdfplumber()
    ensure_pytesseract()
    found = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages):
            raw = _page_images(page)
            if raw is None:
                continue
            for oriented in _orientations(raw):
                anchors = _find_anchors(oriented)
                if not anchors:
                    continue
                for anchor in anchors:
                    found.append({
                        "page": page_index,
                        "number": _read_number_near_anchor(oriented, anchor),
                        "image": _crop_check(oriented, anchor),
                    })
                break
    return found


def check_number_from_chase_description(description):
    """'CHECK 1775 SIGNARAMA' -> 1775; None si no es un cheque."""
    match = re.match(r"\s*CHE(?:CK|QUE)\s*#?\s*0*(\d+)", description or "", re.IGNORECASE)
    return int(match.group(1)) if match else None
