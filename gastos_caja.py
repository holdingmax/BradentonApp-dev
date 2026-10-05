"""
Gastos de Caja leídos de los comprobantes escaneados (2026-10-04).

Pedido del usuario: "implementemos un sistema de OCR para subir los gastos
del mes hechos con caja, así aunque no sean proveedores ya escaneados, que
sea un intento de cargar de forma automática gastos del mes".

Cada comprobante de "Gastos del Mes" es distinto (tickets de Gordon u
Office Depot, facturas a mano, presupuestos, notas), pero todo lo que se
pagó con la caja lleva abrochado el ticket del POS "Paid Out" -- siempre el
mismo formato, de 2023 a hoy:

    CASH PAID OUTS
    GL Account ID: 3
    08/21/26 11:32:12 AM
    Register: 1 Trans #: 8561 Op ID: 1
    Your cashier: RICK
    Cash: $-70.00

De ahí sale el gasto: la fecha (la del ticket, que es la que va en Caja --
el comprobante del proveedor puede ser de días antes), el monto y el N° de
transacción (para no cargar dos veces lo mismo). El detalle (a quién se le
pagó) sale del nombre del archivo, que el usuario ya pone así ("Gordon Food
Service 04-08.pdf"); si el nombre no dice nada ("WhatsApp Image...",
"Receipt.pdf"), de un comercio conocido impreso en la página, y si no,
queda vacío y se avisa.

Un comprobante SIN ese ticket no se carga solo: puede no ser de caja
(Petroserv, pagado con cheque) o serlo igual (una compra con débito que se
repuso de la caja, como el Home Depot del 13/07/2026 que el Excel cuenta
en efectivo). Para esos se propone la fecha (la del nombre del archivo) y
el total impreso del comprobante, y el usuario confirma o descarta.

Regla de oro del OCR: la fecha y el monto de cada ticket se leen varias
veces (recortes y umbrales distintos) y se guardan solo si todas las
lecturas coinciden, o si el valor se confirma por otro lado (ver _vote).
Un ticket dudoso no se guarda: queda "para confirmar" con lo que sí se leyó.
"""
import os
import re
from collections import Counter
from datetime import date

import ocr_utils
from ocr_utils import ensure_pdfplumber, ensure_pytesseract

try:
    from PIL import Image, ImageOps
except ImportError:  # pragma: no cover - sin Pillow falla ensure_pytesseract con un error claro
    Image = ImageOps = None  # type: ignore[assignment]

# Ancla: "Paid Out" (el rótulo en negro) y "CASH PAID OUTS", tal como los lee
# Tesseract (a veces "QUTS", "PA1D").
_ANCHOR_RE = re.compile(r"^[^\w]*(?:PAID|PA1D|PAlD|PALD|PAIO)$", re.IGNORECASE)
_OUT_RE = re.compile(r"^(?:OUTS?|QUTS?|0UTS?|OUT\w?)[^\w]*$", re.IGNORECASE)
# "08/21/26 11:32:12 AM": la fecha va seguida de la hora (así no se confunde
# con otra fecha impresa en el comprobante).
_DATE_RE = re.compile(r"(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{2})\b\s+(\d{1,2})\s*[:.]\s*(\d{2})")
# "Cash: $-70.00"; el punto a veces se pierde ("$-49 21") -- el ticket
# siempre imprime dos decimales.
_AMOUNT_NUM = r"(\d{1,3}(?:,\d{3})+|\d{1,5})\s*[.,]?\s*(\d{2})\b"
_AMOUNT_RE = re.compile(r"C[a-z]{1,3}h?\s*[:;.,]?\s*\$?\s*[-~—–]\s*\$?\s*" + _AMOUNT_NUM, re.IGNORECASE)
_AMOUNT_LOOSE_RE = re.compile(r"\$\s*[-~—–]\s*" + _AMOUNT_NUM)
_TRANS_RE = re.compile(r"Tran[a-z]?\s*[#H]?\s*[:;]?\s*(\d{1,6})", re.IGNORECASE)

# Comprobantes sin ticket: el renglón del total ("TOTAL $20.84", "Job Total",
# "Amount Due"), sin SUBTOTAL ni "TOTAL NUMBER OF ITEMS".
_TOTAL_LINE_RE = re.compile(
    r"(?<!SUB )(?<!SUB-)\b(?:GRAND\s+TOTAL|JOB\s+TOTAL|TOTAL\s+DUE|AMOUNT\s+DUE|BALANCE\s+DUE|TOTAL\s+PAID|"
    r"AMOUNT\s+PAID|TOTAL)\b(?!\s+(?:NUMBER|ITEMS|TAX|SAVINGS|DISCOUNT))([^\n]{0,40})", re.IGNORECASE)
_MONEY_RE = re.compile(r"\$?\s*(\d{1,3}(?:,\d{3})+|\d{1,6})\s*[.,]\s*(\d{2})\b")
_FILE_DATE_RE = re.compile(r"(?<!\d)(\d{1,2})[-.](\d{1,2})(?:[-.](\d{2,4}))?(?!\d)")

# Nombres de archivo que no dicen a quién se le pagó.
_GENERIC_NAME_RE = re.compile(
    r"^(?:whatsapp image.*|img\s*\d*|image\s*\d*|scan\w*|escaneo\w*|receipt\w*|recibo\w*|invoice|factura|"
    r"documento?|doc\w*|cam\s*scanner.*|pdf|estimate|\d+)$", re.IGNORECASE)
# Comercios frecuentes de Gastos del Mes (2023-2026), por si el nombre del
# archivo no sirve: (patrón sobre el texto de la página, detalle).
_KNOWN_VENDORS = (
    (r"\bGordon\b", "Gordon Food Service"),
    (r"Office\s*DEPOT|OfficeMax", "Office Depot OfficeMax"),
    (r"COMPLIANCE\s+MONITORING", "Compliance Monitoring Services"),
    (r"Charlotte\s+Monitoring", "Charlotte Monitoring Systems"),
    (r"Batteries\s*\+|Batteries\s*Plus", "BatteriesPlus"),
    (r"HARBOR\s+FREIGHT", "Harbor Freight"),
    (r"Home\s+Depot", "The Home Depot"),
    (r"Lowe'?s\b", "Lowe's"),
    (r"Walmart", "Walmart"),
    (r"Family\s+Dollar", "Family Dollar"),
    (r"Sam'?s\s+Club", "Sam's Club"),
    (r"Slush\s*Pupp", "Slush Puppie"),
    (r"Flori[\s-]*Gas", "Flori Gas"),
    (r"PETRO\s*SERV", "Petroserv"),
)


def detail_from_filename(filename):
    """
    "Gordon Food Service 04-08.pdf" -> "Gordon Food Service": el nombre sin
    fechas, sin "Inv N°" y sin el "(1)" de las copias. None si no dice nada.
    """
    stem = os.path.splitext(os.path.basename(filename or ""))[0].replace("_", " ")
    stem = re.sub(r"\(\d+\)", " ", stem)
    stem = re.sub(r"\b\d{4}[-.]\d{1,2}[-.]\d{1,2}\b", " ", stem)            # 2026-06-16
    stem = re.sub(r"\b\d{1,2}[-.]\d{1,2}(?:[-.]\d{2,4})?\b", " ", stem)      # 21-08, 13.07.2026
    stem = re.sub(r"\bat\s+\d{1,2}\s+\d{2}\s+\d{2}\b", " ", stem)            # WhatsApp "at 12.48.25"
    stem = re.sub(r"\b(?:inv(?:oice)?|factura)\b[\s#.:-]*(?:n[°º]?\s*)?[\w-]*\d[\w-]*", " ", stem, flags=re.IGNORECASE)
    stem = re.sub(r"\s+", " ", stem).strip(" -.,#")
    if not stem or _GENERIC_NAME_RE.match(stem):
        return None
    return stem


def date_from_filename(filename, default_year):
    """La fecha que el usuario pone en el nombre: "21-08", "15-08-2026", "13.07.2026" (día-mes). None si no hay."""
    stem = os.path.splitext(os.path.basename(filename or ""))[0]
    for m in _FILE_DATE_RE.finditer(stem):
        day, month, year = int(m.group(1)), int(m.group(2)), m.group(3)
        if year:
            year = int(year) + (2000 if len(year) == 2 else 0)
        try:
            return date(year or default_year, month, day)
        except (TypeError, ValueError):
            continue
    return None


def _proposed_total(images):
    """El total impreso del comprobante si dos lecturas coinciden (y ninguna otra empata); None si no hay uno claro."""
    pt = ocr_utils.pytesseract
    votes = Counter()
    for image in images[:3]:
        for config in ("--psm 4", "--psm 6"):
            found = []
            for m in _TOTAL_LINE_RE.finditer(pt.image_to_string(image, config=config)):
                money = _MONEY_RE.search(m.group(1))
                if money:
                    value = float(money.group(1).replace(",", "") + "." + money.group(2))
                    if value > 0:
                        found.append(value)
            if found:
                votes[max(found)] += 1
    if not votes:
        return None
    (best, n), *rest = votes.most_common()
    return best if n >= 2 and (not rest or n > rest[0][1]) else None


def _page_images(path):
    """Cada página (o la foto) como imagen ya orientada, a no más de 2200 px de lado."""
    ensure_pytesseract()
    ext = os.path.splitext(path)[1].lower()
    if ext in (".jpg", ".jpeg", ".png"):
        images = [ocr_utils.correct_image_orientation(Image.open(path).convert("RGB"))]
    else:
        ensure_pdfplumber()
        images = []
        with ocr_utils.pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                try:
                    image = ocr_utils.extract_largest_page_image(page)
                except Exception:  # una imagen que Pillow no decodifica (JBIG2/CCITT): se renderiza la página
                    image = None
                if image is None or image.width * image.height < 300 * 300:
                    image = ocr_utils.correct_image_orientation(page.to_image(resolution=200).original.convert("RGB"))
                images.append(image)
    out = []
    for image in images:
        scale = 2200 / max(image.size)
        if scale < 1:
            image = image.resize((int(image.width * scale), int(image.height * scale)))
        out.append(image)
    return out


def _words(image):
    """Palabras sueltas de toda la página con su caja (psm 11: texto disperso, sin asumir columnas)."""
    pt = ocr_utils.pytesseract
    data = pt.image_to_data(image, config="--psm 11", output_type=pt.Output.DICT)
    return [(data["text"][i].strip(), data["left"][i], data["top"][i], data["width"][i], data["height"][i])
            for i in range(len(data["text"])) if data["text"][i].strip()]


def _find_slips(words):
    """Caja de cada ticket "Paid Out" de la página (uno por ticket, aunque haya dos lado a lado)."""
    anchors = []
    for i, (text, left, top, width, height) in enumerate(words):
        if not _ANCHOR_RE.match(text):
            continue
        for other, oleft, otop, owidth, oheight in words[i + 1:i + 4]:
            if _OUT_RE.match(other) and abs(otop - top) < height and 0 <= oleft - (left + width) < 3 * height:
                anchors.append((left, top, oleft + owidth, max(top + height, otop + oheight)))
                break
    # "Paid Out" y "CASH PAID OUTS" son del mismo ticket: se quedan con el de más arriba.
    merged = []
    for box in sorted(anchors, key=lambda b: (b[1], b[0])):
        for m in merged:
            line_h = m[3] - m[1]
            if abs((box[0] + box[2]) / 2 - (m[0] + m[2]) / 2) < 3 * (m[2] - m[0]) and 0 <= box[1] - m[1] < 6 * line_h:
                break
        else:
            merged.append(box)
    return merged


def _slip_box(image, anchor):
    """Recorte del ticket: centrado en el ancla, unas 14 líneas hacia abajo (hasta "Signature")."""
    left, top, right, bottom = anchor
    line_h = max(bottom - top, 10)
    cx = (left + right) / 2
    half = max((right - left) * 1.6, line_h * 14)
    return (max(0, int(cx - half)), max(0, int(top - line_h)),
            min(image.width, int(cx + half)), min(image.height, int(top + line_h * 14)))


def _parse_slip_text(text):
    out = {}
    m = _DATE_RE.search(text)
    if m:
        try:
            out["date"] = date(2000 + int(m.group(3)), int(m.group(1)), int(m.group(2)))
        except ValueError:
            pass
    m = _AMOUNT_RE.search(text) or _AMOUNT_LOOSE_RE.search(text)
    # "$-04.19" es un "6" mal leído (el ticket nunca imprime un cero adelante): esa lectura no cuenta.
    if m and not (len(m.group(1)) > 1 and m.group(1)[0] == "0"):
        value = float(m.group(1).replace(",", "") + "." + m.group(2))
        if 0 < value < 100000:
            out["amount"] = value
    m = _TRANS_RE.search(text)
    if m:
        out["trans"] = m.group(1)
    return out


def _variants(crop, extra):
    scale = 2 if crop.width < 1400 else 1
    big = crop.resize((crop.width * scale, crop.height * scale), Image.LANCZOS) if scale > 1 else crop
    gray = ImageOps.grayscale(big)
    if not extra:
        return [(gray, "--psm 6"), (gray, "--psm 4"), (gray.point(lambda v: 255 if v > 150 else 0), "--psm 6")]
    gray3 = ImageOps.grayscale(crop.resize((crop.width * 3, crop.height * 3), Image.LANCZOS))
    return [(gray3, "--psm 6"), (gray.point(lambda v: 255 if v > 120 else 0), "--psm 6"),
            (gray.point(lambda v: 255 if v > 180 else 0), "--psm 4")]


def _words_text(words, box):
    """Las palabras de la página (psm 11) que caen dentro del ticket, armadas en renglones: una lectura más."""
    left, top, right, bottom = box
    inside = sorted((w for w in words if left <= w[1] <= right and top <= w[2] <= bottom), key=lambda w: (w[2], w[1]))
    lines, current, current_top = [], [], None
    for text, wl, wt, _ww, wh in inside:
        if current and abs(wt - current_top) > wh * 0.6:
            lines.append(" ".join(current))
            current = []
        if not current:
            current_top = wt
        current.append(text)
    if current:
        lines.append(" ".join(current))
    return "\n".join(lines)


def _amount_on_page(words, box, amount):
    """¿El monto está impreso en otro lado de la página (el total del comprobante del comercio)?"""
    left, top, right, bottom = box
    target = f"{amount:,.2f}"
    for text, wl, wt, _ww, _wh in words:
        if left <= wl <= right and top <= wt <= bottom:
            continue
        if text.strip("$*:;TS ").replace("$", "") in (target, target.replace(",", "")):
            return True
    return False


def _vote(readings, field, confirmed=()):
    """
    El valor de `field` si las lecturas lo sostienen, o None (dudoso).
    - Todas las lecturas que leyeron algo dicen lo mismo (y son 2 o más): vale.
    - Si no coinciden, la mayoría sola no alcanza: en las facturas reales de
      2024-2026 la mayoría se equivocó dos veces (un 3 leído como 9 y un 6
      como 8, en cuatro de seis lecturas). Vale el único valor leído 2+ veces
      que además se confirma por otro lado (`confirmed`: el monto impreso en
      el comprobante del comercio, la fecha cerca de la del nombre del
      archivo), o una mayoría aplastante (4+ lecturas y el triple que la
      segunda).
    """
    counts = Counter(r[field] for r in readings if field in r)
    if not counts:
        return None
    (best, n), *rest = counts.most_common()
    if n >= 2 and not rest:
        return best
    backed = [value for value in counts if value in confirmed and counts[value] >= 2]
    if len(backed) == 1 and not any(value in confirmed for value in counts if value != backed[0]):
        return backed[0]
    if n >= 4 and n >= 3 * rest[0][1]:
        return best
    return None


def _read_slip(image, box, words, hint_date=None):
    pt = ocr_utils.pytesseract
    crop = image.crop(box)
    readings = [_parse_slip_text(_words_text(words, box))]
    for extra in (False, True):
        for img, config in _variants(crop, extra):
            readings.append(_parse_slip_text(pt.image_to_string(img, config=config)))
        # Si todas coinciden en fecha y monto, no hace falta leer más.
        if all(len({r[f] for r in readings if f in r}) == 1 and sum(f in r for r in readings) >= 3
               for f in ("date", "amount")):
            break
    amounts = {r["amount"] for r in readings if "amount" in r}
    dates = {r["date"] for r in readings if "date" in r}
    return {
        "amount": _vote(readings, "amount", {a for a in amounts if _amount_on_page(words, box, a)}),
        "date": _vote(readings, "date", {d for d in dates if hint_date and abs((d - hint_date).days) <= 10}),
        "trans": _vote(readings, "trans"),
    }


def _image_slips(image, hint_date=None):
    """Los tickets de una imagen; si no aparece ninguno, se prueba girada (el ticket abrochado de costado)."""
    upright_words = None
    for rotation in (None, Image.ROTATE_180, Image.ROTATE_90, Image.ROTATE_270):
        img = image if rotation is None else image.transpose(rotation)
        words = _words(img)
        if upright_words is None:
            upright_words = words
        anchors = _find_slips(words)
        if anchors:
            return [_read_slip(img, _slip_box(img, a), words, hint_date) for a in anchors], words
    return [], upright_words


def _vendor_from_words(words):
    text = " ".join(w[0] for w in words)
    for pattern, name in _KNOWN_VENDORS:
        if re.search(pattern, text, re.IGNORECASE):
            return name
    return None


def extract_cash_expenses(path, filename=None, default_year=None):
    """
    Los gastos en efectivo de un comprobante (PDF o foto):
    {"detail": str | None,
     "slips": [{"date": date, "amount": float, "trans": str | None}],   # leídos con seguridad
     "doubtful": [{"date": date | None, "amount": float | None}],      # ticket encontrado pero dudoso
     "proposal": {"date": date | None, "amount": float | None} | None} # sin ticket: para confirmar a mano
    `default_year`: el año para una fecha del nombre sin año ("21-08").
    """
    filename = filename or os.path.basename(path)
    detail = detail_from_filename(filename)
    hint_date = date_from_filename(filename, default_year or date.today().year)
    slips, doubtful, seen = [], [], set()
    images = _page_images(path)
    for image in images:
        found, words = _image_slips(image, hint_date)
        if detail is None and words:
            detail = _vendor_from_words(words)
        for slip in found:
            if slip["date"] is None or slip["amount"] is None:
                doubtful.append({"date": slip["date"], "amount": slip["amount"]})
                continue
            key = (slip["date"], slip["trans"] or slip["amount"])
            if key in seen:  # el mismo ticket escaneado dos veces
                continue
            seen.add(key)
            slips.append(slip)
    proposal = None
    if not slips and not doubtful:
        proposal = {"date": hint_date, "amount": _proposed_total(images)}
    return {"detail": detail, "slips": slips, "doubtful": doubtful, "proposal": proposal}
