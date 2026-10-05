"""
Productos de las facturas de proveedores -- pedido explícito del usuario
(2026-09-28): "crear un sistema organizado en el que extraigamos los
productos que compramos con su precio, para así utilizar esa info para ver
cómo van cambiando los costos de los mismos productos en cada compra y
también poder compararlo con el CMV... a qué proveedor se le compra ese
producto por su UPC". Lectura/extracción + cálculo, sin UI -- el guardado
vive en proveedores_db (tabla supplier_invoice_lines) y las pantallas en la
sección Proveedores de webapp.py.

H.T. Hackney trae texto real. CEC y Colonial (2026-10-04, pedido del
usuario: "mejorar la precisión de los OCR de las facturas de los
proveedores... así se puedan extraer los productos al igual que H.T.")
son escaneos y se leen por OCR, con más controles (ver más abajo). El
resto de los proveedores todavía no -- ver los relevamientos del
2026-09-28 y 2026-10-04 en HISTORIAL.md. Cada proveedor nuevo se suma en
LINE_EXTRACTORS.

Regla de oro: el detalle de una factura se guarda solo si cierra al
centavo -- cantidad x neto = total de cada renglón, la suma de renglones =
INVOICE SUBTOTAL impreso, y subtotal + cargos = Total impreso. Si algo no
cierra, ValueError y no se guarda ningún renglón (nunca un detalle a medias).
"""

import re

import pdfplumber

import ocr_utils
from ocr_utils import (
    crop_relative,
    ensure_cv2,
    ensure_pdfplumber,
    ensure_pytesseract,
    extract_largest_page_image,
    remove_grid_lines,
)

try:
    import cv2
    import numpy as np
    from PIL import Image
except ImportError:  # sin OpenCV/Pillow: solo fallan CEC y Colonial, con un error claro (ensure_cv2)
    cv2 = None  # type: ignore[assignment]
    np = None  # type: ignore[assignment]
    Image = None  # type: ignore[assignment]

# Columnas de la grilla de H.T., en el orden de los 16 bloques de guiones
# que la factura imprime debajo del encabezado ("- ------- -------------- ...").
_HT_COLUMNS = [
    "flag", "qty", "upc", "pack", "size", "description", "item_no", "srp_pct",
    "srp", "srp_ext", "price", "allowance", "tax", "net", "ext", "tx",
]
# Un cartón de cigarrillos (size "CTN") son 10 paquetes, y el POS tiene el
# costo por paquete: Marlboro Gold $102.95 el cartón = $10.30 en el POS.
_UNITS_PER_CARTON = 10


def normalize_upc(value):
    """Solo dígitos, sin ceros adelante -- así cruza con el UPC del POS (CMV)."""
    digits = re.sub(r"\D", "", str(value or ""))
    return digits.lstrip("0")


def product_key(supplier_key, upc, item_no):
    """Clave de producto: el UPC; sin UPC, el Item # propio del proveedor."""
    return upc if upc else f"item-{supplier_key}-{item_no}"


def _num(text):
    text = (text or "").replace(",", "").replace("$", "").strip()
    if not text:
        return None
    negative = text.endswith("-") or text.startswith("-")
    text = text.strip("-").strip()
    try:
        value = float(text)
    except ValueError:
        return None
    return -value if negative else value


def _page_rows(page):
    """Renglones de la página: palabras agrupadas por altura, de izquierda a derecha."""
    rows = []
    last_top = None
    for word in sorted(page.extract_words(), key=lambda w: (round(w["top"]), w["x0"])):
        top = round(word["top"])
        if last_top is not None and top - last_top <= 2:
            rows[-1].append(word)
        else:
            rows.append([word])
            last_top = top
    for row in rows:
        row.sort(key=lambda w: w["x0"])
    return rows


def _column_index(dash_ranges, word):
    center = (word["x0"] + word["x1"]) / 2

    def distance(index):
        start, end = dash_ranges[index]
        if start - 1 <= center <= end + 1:
            return 0
        return min(abs(center - start), abs(center - end))

    return min(range(len(dash_ranges)), key=distance)


def extract_ht_hackney_lines(pdf_path, invoices=None):
    """
    Renglones de producto de una factura de H.T. Hackney. Las columnas se
    leen por posición (los bloques de guiones del encabezado marcan dónde
    empieza y termina cada una), porque varias vienen vacías según el
    producto (SRP%, Alw, Tax, y a veces el UPC) y leer el texto corrido las
    correría de lugar.
    """
    invoice_no = None
    lines = []
    charges = {}
    subtotal = None
    total = None
    category = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            dash_ranges = None
            for row in _page_rows(page):
                text = " ".join(w["text"] for w in row)
                if invoice_no is None:
                    match = re.search(r"Invoice #:\s*(\d+)", text)
                    if match:
                        invoice_no = match.group(1)
                # El encabezado repite el total en cada página ("**For
                # Reference Purposes Only** Total: ..."): no es el renglón final.
                if "For Reference Purposes Only" in text:
                    continue
                # Cargos, subtotal y total se buscan en toda la página, no solo
                # debajo de la grilla: a veces el Total cae solo en una última
                # página que ya no repite la grilla de columnas.
                match = re.match(r"^(FUEL SURCHARGE|DELIVERY CHARGE)\s+([\d,.]+-?)$", text)
                if match:
                    charges[match.group(1)] = _num(match.group(2))
                    continue
                # Al final del renglón y no al principio: a veces viene pegado
                # a un aviso de retiro de producto de la FDA ("...RECALLS Total: 4995.51").
                # El monto puede venir sin cero adelante ("Tax: .66").
                match = re.search(r"(?:Tax:\s*([\d,]*\.\d{2}-?)\s+)?Total:\s*([\d,]*\.\d{2}-?)$", text)
                if match:
                    if match.group(1):
                        charges["TAX"] = _num(match.group(1))
                    total = _num(match.group(2))
                    continue
                match = re.search(r"\*\*\* INVOICE SUBTOTAL \*\*\*.*?([\d,]+\.\d{2}-?)$", text)
                if match:
                    subtotal = _num(match.group(1))
                    continue
                if dash_ranges is None:
                    if len(row) >= 14 and all(set(w["text"]) == {"-"} for w in row):
                        if len(row) != len(_HT_COLUMNS):
                            raise ValueError(
                                "la grilla de productos de H.T. cambió de formato "
                                f"({len(row)} columnas en vez de {len(_HT_COLUMNS)})."
                            )
                        dash_ranges = [(w["x0"], w["x1"]) for w in row]
                    continue

                cells = {name: [] for name in _HT_COLUMNS}
                for word in row:
                    cells[_HT_COLUMNS[_column_index(dash_ranges, word)]].append(word["text"])
                cells = {name: " ".join(parts) for name, parts in cells.items()}
                if (re.fullmatch(r"\d+", cells["qty"]) and re.fullmatch(r"\d{5,8}", cells["item_no"])
                        and _num(cells["ext"]) is not None and _num(cells["net"]) is not None):
                    lines.append(_ht_line(cells, category, len(lines) + 1))
                    continue
                # Encabezado de sección ("CANDY", "CIGS, FULL PRICE"): en
                # mayúsculas y sin números. Las notas ("ABOVE IS ...", "THE
                # ABOVE ITEM WAS ...") no son sección.
                if text.isupper() and not re.search(r"\d", text) and "ABOVE" not in text and "***" not in text:
                    category = text

    if not lines:
        raise ValueError("no se encontró ningún renglón de producto en la factura de H.T.")
    for line in lines:
        if abs(line["qty"] * line["net"] - line["ext"]) > 0.011:
            raise ValueError(
                f"el renglón {line['line_no']} ({line['description']}) no cierra: "
                f"{line['qty']:g} x {line['net']:.2f} ≠ {line['ext']:.2f}."
            )
    lines_total = round(sum(line["ext"] for line in lines), 2)
    if subtotal is None or abs(lines_total - subtotal) >= 0.01:
        raise ValueError(
            f"los renglones suman ${lines_total:,.2f} y el INVOICE SUBTOTAL de la factura es "
            f"{'$' + format(subtotal, ',.2f') if subtotal is not None else 'ilegible'}."
        )
    if total is None or abs(subtotal + sum(charges.values()) - total) >= 0.01:
        raise ValueError("el subtotal más los cargos (fuel, delivery, tax) no da el Total de la factura.")
    if invoice_no is None:
        raise ValueError("no se encontró el N° de invoice en la factura de H.T.")
    return {"invoice_no": invoice_no, "lines": lines, "subtotal": subtotal, "total": total}


def _ht_line(cells, category, line_no):
    pack = int(_num(cells["pack"].lstrip("*")) or 1)
    size = cells["size"]
    units = pack * (_UNITS_PER_CARTON if size.upper() == "CTN" else 1)
    net = _num(cells["net"])
    return {
        "line_no": line_no,
        "upc": normalize_upc(cells["upc"]),
        "item_no": cells["item_no"],
        "description": cells["description"],
        "category": category,
        "qty": _num(cells["qty"]),
        "pack": pack,
        "size": size,
        "units": units,
        "price": _num(cells["price"]),
        "allowance": _num(cells["allowance"]),
        "tax": _num(cells["tax"]),
        "net": net,
        "ext": _num(cells["ext"]),
        "unit_cost": round(net / units, 4),
        "srp": _num(cells["srp"]),
    }


# ---------------------------------------------------------------------------
# Facturas ESCANEADAS: CEC y Colonial (2026-10-04)
# ---------------------------------------------------------------------------
# Mismo contrato y misma regla de oro que H.T. Hackney: el detalle se guarda
# solo si cierra al centavo. Como estas facturas son escaneos (OCR), además:
# - cada renglón tiene que cerrar (cantidad x neto = total del renglón);
# - la cantidad se toma del OCR, y solo si ese número no cierra (las líneas
#   de la grilla lo ensucian: "|42" por 12) se deduce de total / precio,
#   siempre que dé un número entero Y lo confirme otra fuente (el conteo
#   impreso al pie o una relectura de esa celda sola);
# - los renglones tienen que sumar el subtotal/total impreso (o, en
#   Colonial, el importe que la app ya leyó del encabezado);
# - el UPC, cuando lo hay, tiene que pasar el dígito verificador.
# Validado sobre los PDFs reales del Drive (2025-2026): CEC 83 de 83, todas
# cerrando contra el Subtotal; Colonial 28 de 80 con el BALANCE DUE nuevo
# de proveedores.py (el resto da ValueError, y ninguna de las que pasan
# tiene un valor mal leído en la revisión a ojo).

_TOLERANCE = 0.005  # medio centavo: precio de 4 decimales x cantidad, redondeado


# ---------------------------------------------------------------------------
# Helpers comunes de OCR
# ---------------------------------------------------------------------------

def _pytesseract():
    ensure_pytesseract()
    return ocr_utils.pytesseract


def _page_images(pdf_path):
    """Imagen escaneada (ya orientada) de cada página, salteando las que no tienen."""
    ensure_pdfplumber()
    with pdfplumber.open(pdf_path) as pdf:
        images = [extract_largest_page_image(page) for page in pdf.pages]
    return [image for image in images if image is not None]


def _ocr_rows(image, config="--psm 6", box=None, scale=1, clean=False):
    """
    Renglones del OCR con la posición de cada palabra. psm 6 lee la página
    como un único bloque, así cada renglón de la tabla sale entero (con psm 3
    Tesseract separa las columnas en bloques y desarma los renglones).
    Con box/scale se lee solo ese recorte agrandado, pero las posiciones
    vuelven en coordenadas de la página original. clean borra el sombreado
    de puntitos (renglones alternados de Colonial) antes de leer, de dos
    formas distintas (dos lecturas independientes): "blur" (desenfoque leve
    y corte fijo: los puntitos quedan más claros que la letra) o "median"
    (filtro de mediana y blanco/negro automático).
    """
    pytesseract = _pytesseract()
    x_off = y_off = 0
    source = image
    if box is not None:
        x_off, y_off = max(0, int(box[0])), max(0, int(box[1]))
        source = image.crop((x_off, y_off, min(image.width, int(box[2])), min(image.height, int(box[3]))))
        source = source.convert("L")
    if scale != 1:
        source = source.convert("L").resize((source.width * scale, source.height * scale))
    if clean:
        ensure_cv2()
        gray = np.array(source.convert("L"))
        if clean == "blur":
            gray = np.where(cv2.GaussianBlur(gray, (3, 3), 0) < 120, 0, 255).astype(np.uint8)
        else:
            gray = cv2.medianBlur(gray, 7 if scale >= 3 else 5)
            _, gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        source = Image.fromarray(gray)
    data = pytesseract.image_to_data(source, config=config, output_type=pytesseract.Output.DICT)
    rows = {}
    for i, text in enumerate(data["text"]):
        text = (text or "").strip()
        if not text:
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        rows.setdefault(key, []).append({
            "text": text,
            "x0": x_off + data["left"][i] / scale,
            "x1": x_off + (data["left"][i] + data["width"][i]) / scale,
            "top": y_off + data["top"][i] / scale,
            "bottom": y_off + (data["top"][i] + data["height"][i]) / scale,
        })
    result = []
    for words in rows.values():
        words.sort(key=lambda w: w["x0"])
        result.append(words)
    result.sort(key=lambda ws: min(w["top"] for w in ws))
    return result


def _row_text(words):
    return " ".join(w["text"] for w in words)


def _ocr_cell(image, box, whitelist="0123456789", psm=7):
    """Relee una celda sola, sin las líneas de la grilla (que el OCR confunde con 1, 4 o |)."""
    pytesseract = _pytesseract()
    x0, top, x1, bottom = box
    crop = image.crop((max(0, int(x0)), max(0, int(top)), min(image.width, int(x1)), min(image.height, int(bottom))))
    if crop.width < 4 or crop.height < 4:
        return ""
    cleaned = remove_grid_lines(crop, upscale=3)
    config = f"--psm {psm}" + (f" -c tessedit_char_whitelist={whitelist}" if whitelist else "")
    return pytesseract.image_to_string(cleaned, config=config).strip()


_DIGIT_LOOKALIKES = str.maketrans({"O": "0", "o": "0", "g": "9", "S": "5", "s": "5", "l": "1", "I": "1", "B": "8"})


def _fix_digits(text):
    """Letras que el OCR confunde con dígitos (g.0000 -> 9.0000); solo para celdas que son números."""
    return text.translate(_DIGIT_LOOKALIKES)


def _money(text):
    """'$1,234.56' / '1.234,56' mal leído / '34.9000' / '9,.0000' -> float, o None."""
    text = (text or "").replace("$", "").replace(" ", "").strip()
    if not text:
        return None
    # El OCR a veces lee la coma de miles como punto o el punto decimal como
    # coma: el separador decimal es siempre el ÚLTIMO, seguido de 2 o 4 dígitos.
    match = re.fullmatch(r"(-?)([\d.,]*?)[.,](\d{2}|\d{4})", text)
    if not match:
        return None
    sign, whole, decimals = match.groups()
    whole = re.sub(r"[.,]", "", whole) or "0"
    value = float(f"{whole}.{decimals}")
    return -value if sign else value


def _gtin_ok(digits):
    """Dígito verificador de un UPC-A/EAN (GTIN 8, 12, 13 o 14): detecta cualquier dígito mal leído."""
    if not re.fullmatch(r"\d{8}|\d{12,14}", digits or ""):
        return False
    body, check = digits[:-1], int(digits[-1])
    total = sum(int(d) * (3 if i % 2 == 0 else 1) for i, d in enumerate(reversed(body)))
    return (10 - total % 10) % 10 == check


def _closes(qty, net, ext):
    return abs(qty * net - ext) <= _TOLERANCE + 1e-9


def _whole_qty(net, ext):
    """La única cantidad entera que cierra total / precio, o None."""
    if not net or net <= 0:
        return None
    qty = round(ext / net)
    return qty if qty > 0 and _closes(qty, net, ext) else None


def _clean_token(text):
    return text.strip("|[](){}_—–~=‘’'\"`«»,;:!*")


def _line(line_no, *, upc="", item_no="", description="", category=None, qty, pack, size, units,
          price, allowance=None, tax=None, net, ext, srp=None):
    """Renglón con exactamente las claves de proveedores_productos._ht_line."""
    return {
        "line_no": line_no,
        "upc": normalize_upc(upc),
        "item_no": item_no,
        "description": description,
        "category": category,
        "qty": qty,
        "pack": pack,
        "size": size,
        "units": units,
        "price": price,
        "allowance": allowance,
        "tax": tax,
        "net": net,
        "ext": ext,
        "unit_cost": round(net / units, 4),
        "srp": srp,
    }


def _check_sum(lines, printed, label):
    lines_total = round(sum(line["ext"] for line in lines), 2)
    if printed is None or abs(lines_total - printed) >= 0.01:
        raise ValueError(
            f"los renglones suman ${lines_total:,.2f} y el {label} de la factura es "
            f"{'$' + format(printed, ',.2f') if printed is not None else 'ilegible'}."
        )


def _median(values):
    values = sorted(values)
    return values[len(values) // 2] if values else None


def _upc_candidate(digits, lengths=(12,)):
    """
    UPC válido dentro de una tira de dígitos leída por OCR. La línea de la
    grilla a veces se lee como un dígito más pegado ("1709215205322"): se
    prueban las ventanas del largo esperado y vale solo si UNA pasa el
    dígito verificador.
    """
    found = set()
    for length in lengths:
        for start in range(0, len(digits) - length + 1):
            window = digits[start:start + length]
            if _gtin_ok(window):
                found.add(window)
    return found.pop() if len(found) == 1 else None


def _reread_digits(image, box, min_len, lengths=(12,)):
    """Relee una celda con solo dígitos y devuelve un UPC válido, o None."""
    digits = re.sub(r"\D", "", _ocr_cell(image, box))
    return _upc_candidate(digits, lengths) if len(digits) >= min_len else None


def _invoice_no_from_crop(image, box, pattern):
    """Relee el N° de factura en un recorte agrandado al doble (el OCR de página entera a veces pierde un dígito)."""
    pytesseract = _pytesseract()
    x0, top, x1, bottom = (int(v) for v in box)
    crop = image.crop((max(0, x0), max(0, top), min(image.width, x1), min(image.height, bottom)))
    crop = crop.convert("L").resize((crop.width * 2, crop.height * 2))
    for psm in (7, 6):
        text = pytesseract.image_to_string(crop, config=f"--psm {psm}")
        text = re.sub(r"(?<=\d) (?=\d)", "", text)
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    return None


def _resolve_quantities(rows, printed_sum, supplier):
    """
    Cantidad y total de cada renglón (claves price, ext, qty_token). En orden:
    1) la cantidad leída cierra (cantidad x precio = total): listo;
    2) si no, total / precio da un entero: cantidad deducida (qty_derived),
       que después tiene que confirmar el conteo impreso o una relectura;
    3) si UN solo renglón tiene el total ilegible o mal leído ("$27200.",
       "204.90" por 204.00) pero su cantidad se leyó limpia, su total es el
       subtotal impreso menos los demás renglones, y vale solo si eso da
       exactamente cantidad x precio.
    Si algo no cierra, ValueError.
    """
    bad = []
    for row in rows:
        row["qty_derived"] = False
        price, ext, qty = row["price"], row["ext"], row["qty_token"]
        if price is None:
            raise ValueError(f"no se pudo leer el precio del renglón {row['line_no']} de {supplier}.")
        if qty is not None and ext is not None and _closes(qty, price, ext):
            row["qty"] = qty
            continue
        whole = _whole_qty(price, ext) if ext is not None else None
        if whole is not None:
            row["qty"] = whole
            row["qty_derived"] = True
            continue
        bad.append(row)
    if not bad:
        return
    row = bad[0]
    if len(bad) == 1 and printed_sum is not None and row["qty_token"]:
        rest = sum(other["ext"] for other in rows if other is not row)
        ext = round(printed_sum - rest, 2)
        if _closes(row["qty_token"], row["price"], ext):
            row["qty"] = row["qty_token"]
            row["ext"] = ext
            return
    raise ValueError(
        f"el renglón {row['line_no']} de {supplier} ({row.get('description', '')}) no cierra: "
        "cantidad x precio no da el total del renglón."
    )


def _qty_reread_ok(image, row):
    """
    Relee la celda de cantidad sola (sin grilla, a la altura del precio, que
    siempre se lee limpio) y confirma la cantidad deducida. Se prueba como
    renglón (psm 7) y como palabra suelta (psm 8).
    """
    box = (max(0, row["row_left"] - 40), row["cell_top"], row["qty_right"] - 4, row["cell_bottom"])
    for psm in (7, 8):
        text = _ocr_cell(image, box, psm=psm)
        if text.isdigit() and int(text) == row["qty"]:
            return True
    return False


# ---------------------------------------------------------------------------
# CEC (Chinook Enterprises Corp.)
# ---------------------------------------------------------------------------

# CEC vende solo por cartón: cigarrillos 305's y filter cigars 305's, ambos en
# cartones de 10 paquetes (el cuadro "Cartons" del pie los cuenta igual). El
# costo por unidad queda por paquete, como el POS (igual que el CTN de H.T.).
_CEC_UNITS_PER_CARTON = 10
# Precio por cartón con 4 decimales; el OCR a veces mete ",." o lee el 9 como "g".
_CEC_PRICE = re.compile(r"[\dgOoSlIB]{1,4}[.,]{1,2}\d{4}(?!\d)")
_CEC_MONEY = re.compile(r"\d[\d.,]*[.,]\d{2}(?!\d)")
_CEC_ITEM = re.compile(r"(?:FC\d{3}[A-Z]|\d{4})(?![\dA-Za-z])")


def _cec_kind(description, item_no):
    """
    'cig' (cigarrillos), 'other' (filter cigars, el "Others" del cuadro
    Cartons) o None si no se reconoce -- en ese caso no se guarda (no se
    sabe cuántas unidades trae).
    """
    text = description.upper()
    if "CIGAR" in text or item_no.startswith("FC"):
        return "other"
    # Los dos tipos son cartones de 10: confundirlos no cambia las unidades
    # (y el cuadro Cartons lo detectaría); lo que no es 305's no se guarda.
    if re.search(r"BOX|SOFT|5['’]S", text) or re.fullmatch(r"\d{4}", item_no):
        return "cig"
    return None


def extract_cec_lines(pdf_path, invoices=None):
    """
    Renglones de una factura de CEC: QTY, Product, Description, UPC, Price
    (4 decimales, por cartón) y Subtotal. Cada renglón se ancla en el precio
    de 4 decimales y el UPC (12 dígitos con verificador). La cantidad y el
    código de producto son las celdas más sucias (la línea de la grilla o
    el tilde del chofer pegados: "|42" por 12, "[4", "v//s08"): el código es
    solo informativo (la clave del producto es el UPC) y la cantidad, si no
    cierra contra precio x total, se deduce de total / precio y tiene que
    confirmarla el cuadro "Cartons" del pie (cigarrillos / otros / total) o
    una relectura de esa celda sola.

    Si la lectura de la página entera no cierra (un renglón que el OCR pegó
    al encabezado de la tabla, por ejemplo), se relee solo la tabla al doble
    de tamaño y se vuelve a validar todo; si tampoco cierra, ValueError.
    """
    images = _page_images(pdf_path)
    if not images:
        raise ValueError("no se encontró la imagen escaneada de la factura de CEC.")
    image = images[0]
    info = _cec_parse(_ocr_rows(image))
    try:
        return _cec_finish(image, info)
    except ValueError as first_error:
        if not info["rows"]:
            raise
        tops = [row["top"] for row in info["rows"]]
        height = _median([row["bottom"] - row["top"] for row in info["rows"]])
        bottom = info["subtotal_bottom"] or max(row["bottom"] for row in info["rows"]) + 3 * height
        retry = _cec_parse(_ocr_rows(image, box=(0, min(tops) - 3 * height, image.width, bottom + 4), scale=2))
        for key in ("invoice_no", "inv_word", "subtotal", "cartons", "total"):
            if retry[key] is None:
                retry[key] = info[key]
        try:
            return _cec_finish(image, retry)
        except ValueError:
            raise first_error from None


def _cec_parse(rows):
    """Encabezado, pie y renglones de producto (sin validar) de una lectura OCR."""
    info = {"invoice_no": None, "inv_word": None, "subtotal": None, "subtotal_bottom": None,
            "cartons": None, "total": None, "rows": []}
    for words in rows:
        text = _row_text(words)
        if info["invoice_no"] is None and info["inv_word"] is None:
            info["inv_word"] = next((w for w in words if w["text"].startswith("Inv")), None)
            if info["inv_word"] is not None:
                match = re.search(r"Inv\s*#?\s*(\d{7})\b", re.sub(r"(?<=\d) (?=\d)", "", text))
                info["invoice_no"] = match.group(1) if match else None
        match = re.search(r"Subtotal\W*(\d[\d.,]*[.,]\d{2})(?!\d)", text)
        if match:
            info["subtotal"] = _money(match.group(1))
            info["subtotal_bottom"] = max(w["bottom"] for w in words)
            continue
        match = re.search(r"Cartons?\W+(\d+)\W+(\S+?)\W+(\d+)\b", text)
        if match:
            cig, other_text, all_ = int(match.group(1)), match.group(2), int(match.group(3))
            other = int(re.sub(r"\D", "", other_text)) if re.search(r"\d", other_text) else None
            if other is None and cig <= all_:
                other = all_ - cig  # "(¢]": el 0 de "Others" ilegible
            info["cartons"] = (cig, other, all_) if other is not None and cig + other == all_ else None
            continue
        match = re.search(r"\bTotal\W*(\d[\d.,]*[.,]\d{2})(?!\d)", text)
        if match:
            info["total"] = _money(match.group(1))
            continue
        price_index = max((i for i, w in enumerate(words) if _CEC_PRICE.search(w["text"])), default=None)
        if price_index is not None:
            info["rows"].append(_cec_row(words, price_index, len(info["rows"]) + 1))
    return info


def _cec_finish(image, info):
    """Valida una lectura: N° de invoice, UPCs, cantidades, Subtotal y cuadro Cartons."""
    parsed = info["rows"]
    invoice_no = info["invoice_no"]
    if invoice_no is None:
        # Misma lectura que el encabezado de la app (página entera, psm 3):
        # el psm 6 a veces pierde un dígito del N° ("Inv #1/64904").
        text = _pytesseract().image_to_string(image)
        match = re.search(r"Inv\s*#\s*([\d\s]+?)\n", text)
        digits = re.sub(r"\s+", "", match.group(1)) if match else ""
        invoice_no = info["invoice_no"] = digits if re.fullmatch(r"\d{7}", digits) else None
    if invoice_no is None:
        inv_word = info["inv_word"]
        # Recorte angosto: a la derecha suele haber un número escrito a mano.
        box = ((inv_word["x0"] - 10, inv_word["top"] - 15, inv_word["x0"] + image.width * 0.24, inv_word["bottom"] + 15)
               if inv_word is not None else (image.width * 0.55, 0, image.width, image.height * 0.08))
        invoice_no = info["invoice_no"] = _invoice_no_from_crop(image, box, r"(?:Inv\s*#?\s*)?(\d{7})\b")
    if invoice_no is None:
        raise ValueError("no se encontró el N° de invoice en la factura de CEC.")
    if not parsed:
        raise ValueError("no se encontró ningún renglón de producto en la factura de CEC.")
    subtotal = info["subtotal"]

    # UPC ilegible (un dígito mal o la celda vacía): se relee la celda sola,
    # entre el borde de la columna UPC (tomado de los otros renglones) y el precio.
    upc_x = _median([row["upc_x"] for row in parsed if row["upc_x"] is not None])
    for row in parsed:
        if row["upc"] is None:
            left = row["upc_x"] if row["upc_x"] is not None else upc_x
            if left is not None:
                row["upc"] = _reread_digits(
                    image, (left - 12, row["cell_top"], row["price_x"] - 4, row["cell_bottom"]), 11
                )
        if row["upc"] is None:
            raise ValueError(f"no se pudo leer el UPC del renglón {row['line_no']} de CEC.")

    _resolve_quantities(parsed, subtotal, "CEC")
    lines = []
    for row in parsed:
        kind = _cec_kind(row["description"], row["item_no"])
        # El código de producto no lo valida ninguna cuenta ("(3403" por
        # 3103): se guarda solo si una relectura de la celda sola coincide.
        if row["item_no"] and _ocr_cell(
            image, (row["item_x"] - 6, row["cell_top"], row["desc_x"] - 4, row["cell_bottom"]),
            whitelist="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
        ) != row["item_no"]:
            row["item_no"] = ""
        if kind is None:
            raise ValueError(
                f"el renglón {row['line_no']} de CEC no es un producto conocido ({row['description']})."
            )
        row["kind"] = kind
        lines.append(_line(
            row["line_no"], upc=row["upc"], item_no=row["item_no"], description=row["description"],
            category="CIGARETTES" if kind == "cig" else "FILTER CIGARS",
            qty=row["qty"], pack=_CEC_UNITS_PER_CARTON, size="CTN", units=_CEC_UNITS_PER_CARTON,
            price=row["price"], net=row["price"], ext=row["ext"],
        ))
    _check_sum(lines, subtotal, "Subtotal")

    # Cuadro "Cartons" (cigarrillos, otros, total): confirma las cantidades.
    cartons = info["cartons"]
    if cartons is not None:
        cig = sum(row["qty"] for row in parsed if row["kind"] == "cig")
        other = sum(row["qty"] for row in parsed if row["kind"] == "other")
        if (cig, other) != cartons[:2]:
            raise ValueError(
                f"los renglones suman {cig} cartones de cigarrillos y {other} de otros, y el cuadro "
                f"Cartons de la factura dice {cartons[0]} y {cartons[1]}."
            )
    for row in parsed:
        if row["qty_derived"] and cartons is None and not _qty_reread_ok(image, row):
            raise ValueError(
                f"la cantidad del renglón {row['line_no']} ({row['description']}) de CEC no se pudo "
                "leer con seguridad."
            )
    # El "Total" del pie es igual al Subtotal en todas las facturas vistas.
    total = info["total"]
    if total is not None and abs(total - subtotal) >= 0.01:
        total = None
    return {"invoice_no": invoice_no, "lines": lines, "subtotal": subtotal,
            "total": total if total is not None else subtotal}


def _cec_row(words, price_index, line_no):
    price_word = words[price_index]
    price = _money(_fix_digits(_CEC_PRICE.search(price_word["text"]).group(0)))
    exts = _CEC_MONEY.findall(" ".join(w["text"] for w in words[price_index + 1:]))
    ext = _money(exts[-1]) if exts else None
    if ext is None:
        # "$27200": el punto decimal no salió; los centavos son los 2 últimos
        # dígitos (las cuentas del renglón y del Subtotal lo confirman o no).
        match = re.search(r"\$\s*(\d{3,7})(?![\d.,])", " ".join(w["text"] for w in words[price_index + 1:]))
        ext = int(match.group(1)) / 100 if match else None

    # UPC: la palabra más a la derecha (antes del precio) con 11 o más dígitos.
    upc = None
    upc_index = None
    for i in range(price_index - 1, -1, -1):
        digits = re.sub(r"\D", "", words[i]["text"])
        if len(digits) >= 11:
            upc_index = i
            upc = _upc_candidate(digits)
            break
    left_words = words[:upc_index] if upc_index is not None else words[:price_index]

    # Cantidad, código y descripción. "40-3103": cantidad y código pegados.
    tokens = []
    for word in left_words:
        # "40-3103": cantidad y código pegados; "3306.~=S305's-": código y descripción.
        text = re.sub(r"(?<=\d)[-—](?=\d{4})", " ", word["text"])
        text = re.sub(r"(?<=[^\s\d])(?=(?:[38]05|05|5)['’]s)", " ", text)
        for part in text.split():
            part = _clean_token(part.strip("."))
            if part:
                tokens.append((part, word))
    desc_index = next((i for i, (t, _w) in enumerate(tokens) if re.search(r"5['’]s", t)), None)
    item_index = next(
        (i for i, (t, _w) in enumerate(tokens[:desc_index]) if _CEC_ITEM.match(t)), None
    )
    if desc_index is None:
        desc_index = item_index + 1 if item_index is not None else len(tokens)
    item_no = _CEC_ITEM.match(tokens[item_index][0]).group(0) if item_index is not None else ""
    description = " ".join(t for t, _w in tokens[desc_index:] if re.search(r"\w", t))
    qty_end = item_index if item_index is not None else desc_index
    qty_token = next((t for t, _w in tokens[:qty_end] if re.fullmatch(r"\d{1,3}", t)), None)
    first_kept = tokens[qty_end][1] if qty_end < len(tokens) else price_word
    upc_x = words[upc_index]["x0"] if upc_index is not None else None
    desc_word = tokens[desc_index][1] if desc_index < len(tokens) else None
    return {
        "line_no": line_no,
        "price": price,
        "ext": ext,
        "qty_token": int(qty_token) if qty_token else None,
        "upc": upc,
        "upc_x": upc_x,
        "item_x": tokens[item_index][1]["x0"] if item_index is not None else None,
        "desc_x": desc_word["x0"] if desc_word is not None else (upc_x or price_word["x0"]),
        "item_no": item_no,
        "description": description,
        "price_x": price_word["x0"],
        "qty_right": first_kept["x0"],
        "row_left": min(w["x0"] for w in words),
        "top": min(w["top"] for w in words),
        "bottom": max(w["bottom"] for w in words),
        # Alto de las celdas para releerlas: el del precio, que siempre sale
        # limpio (el renglón entero se estira con los tildes a mano).
        "cell_top": price_word["top"] - 10,
        "cell_bottom": price_word["bottom"] + 10,
    }


# ---------------------------------------------------------------------------
# Colonial Wholesale Dist. LLC
# ---------------------------------------------------------------------------

# Colonial factura por UNIT (BX caja, CT cartón, CS bulto, PK paquete, RL
# rollo, EA unidad) y la columna SIZE dice cuántas unidades de venta trae:
# "10CT" = 10, "10PK" = cartón de 10 paquetes, "500CT" = 500 sorbetes. Esas
# son las unidades del renglón (BLACK & MILD 5PK en BX 10CT = 10 paquetes de
# 5, costo por paquete de 5). No trae UPC: la clave del producto es el ITEM#
# de Colonial, así que el ITEM# y el SIZE se guardan solo si dos lecturas
# distintas coinciden (la de la tabla tal cual y otra al doble de tamaño,
# sin sombreado; si no coinciden, desempata la celda sola).
_COLONIAL_SIZE = re.compile(r"^([\dSOILB]{1,4})(CT|PK)$")
_COLONIAL_MONEY = re.compile(r"^\d[\d,]*[.,]\d{2}$")
_COLONIAL_PROMO = re.compile(r"^[-.,]?\d{2}$|^\d+[.,]\d{2}$")
# Delivery fee que Colonial suma al pie (BALANCE DUE = renglones + fee).
_COLONIAL_FEES = (7.99, 9.99, 11.99)
_ONE_LIKE = str.maketrans({"l": "1", "I": "1", "i": "1", "]": "1", "[": "1", "!": "1", "|": "1",
                           "O": "0", "o": "0", "S": "5", "s": "5", "B": "8", "Z": "2", "z": "2"})


def _colonial_int(token):
    """Cantidad o ITEM# con las letras que el OCR confunde con dígitos; None si no es un número."""
    text = _clean_token(token or "").translate(_ONE_LIKE)
    return text if re.fullmatch(r"\d+", text) else None


def _colonial_size(token):
    """'10CT' / 'SOCT' (50CT) / '1OPK' -> ('10CT', 10), o None."""
    text = re.sub(r"[^\w]", "", token or "").upper().replace("CCT", "CT")  # "6&CT", "6CcT"
    match = _COLONIAL_SIZE.match(text)
    if not match:
        return None
    count = match.group(1).replace("L", "1").translate(_ONE_LIKE)
    if not count.isdigit() or int(count) == 0:
        return None
    return f"{int(count)}{match.group(2).upper()}", int(count)


def _without_vertical_lines(image):
    """
    Copia en grises sin las líneas verticales de la grilla. El borde
    izquierdo de la tabla queda pegado al ITEM# y el OCR lo lee como un "1"
    de más (7098 -> 17098): las lecturas de confirmación se hacen sin esas
    líneas, así ese error no se repite en las dos y no pasa como confirmado.
    """
    ensure_cv2()
    gray = np.array(image.convert("L"))
    ink = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 25, 15)
    # Más largo que cualquier letra (~1,2% del alto) pero corto para seguir
    # las líneas que el escaneo dobla.
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(30, gray.shape[0] // 60)))
    lines = cv2.morphologyEx(ink, cv2.MORPH_OPEN, kernel)
    lines = cv2.dilate(lines, cv2.getStructuringElement(cv2.MORPH_RECT, (5, 3)))
    # Se rellena con el fondo de alrededor, no con blanco: una franja blanca
    # en un renglón sombreado deja dos bordes que el OCR lee como un "1"
    # (15CT -> 115CT) en todas las lecturas sin grilla a la vez.
    gray = cv2.inpaint(gray, lines, 3, cv2.INPAINT_TELEA)
    return Image.fromarray(gray)


def _colonial_row(words):
    """
    Un renglón de producto de Colonial, sin validar: ITEM#, ORDERED,
    SHIPPED, UNIT, SIZE, DESCRIPTION, PRICE, PROMO y AMOUNT. Se lee de
    derecha a izquierda (los 3 importes del final) y la izquierda se ancla en
    el SIZE ("10CT"), que siempre viene después de UNIT. None si no es un
    renglón de producto ("Box 1 of 4", notas).
    """
    text = re.sub(r"(?<![\d.,])(\d+)\s+([.,]\d{2})(?!\d)", r"\1\2", _row_text(words))  # "158 .88"
    tokens = [t for t in re.split(r"[\s|]+", text) if t]
    nums = [re.sub(r"^[^\d.,-]+|[^\d]+$", "", t) for t in tokens]
    digit_idx = [i for i, t in enumerate(tokens) if re.search(r"\d", t)]
    if len(digit_idx) < 3:
        return None
    amt_i, promo_i, price_i = digit_idx[-1], digit_idx[-2], digit_idx[-3]
    if not _COLONIAL_MONEY.match(nums[amt_i]):
        return None
    amount = _money(nums[amt_i])
    promo = None
    if _COLONIAL_PROMO.match(nums[promo_i]):
        promo_text = nums[promo_i].lstrip("-")
        promo = _money(promo_text if re.search(r"[.,]", promo_text) else "." + promo_text)
    price = None
    if _COLONIAL_MONEY.match(nums[price_i]):
        price = _money(nums[price_i])
    elif re.fullmatch(r"\W?\d{2}", tokens[price_i]):
        price = int(tokens[price_i][-2:]) / 100  # ".99" leído "‘99" (la cuenta lo confirma o no)
    if price is None and _COLONIAL_MONEY.match(nums[promo_i]):
        # Promo ilegible y pegada: "14.99 =| 14.99" -> el penúltimo es el precio.
        price, price_i, promo = _money(nums[promo_i]), promo_i, None
    if price is None:
        return None

    left = tokens[:price_i]
    # El SIZE va antes de la descripción: no se busca después de una palabra
    # ("BACKWOOD 5PK SWEET" no es un SIZE 5PK).
    first_word = next((i for i, t in enumerate(left) if re.search(r"[A-Za-z]{3}", t)), len(left))
    # Candidatos a SIZE antes de la primera palabra de la descripción. Si
    # hay dos ("2CT 10PK"), el primero es SHIPPED y UNIT pegados. Si hay uno
    # solo con el mismo número que la cantidad de al lado y sin UNIT antes
    # ("9059 2 2CT EDGEFIELD"), puede ser eso mismo: no se toma.
    candidates = [i for i in range(1, min(len(left), 8, first_word + 1)) if _colonial_size(left[i])]
    size_i = candidates[-1] if candidates else None
    if size_i is not None and len(candidates) == 1:
        count = _colonial_size(left[size_i])[1]
        previous = _colonial_int(left[size_i - 1])
        has_unit = any(re.fullmatch(r"[A-Za-z]{2}", _clean_token(t)) for t in left[:size_i])
        if previous is not None and int(previous) == count and not has_unit:
            size_i = None
    if size_i is None:
        size_text, before, desc_tokens = None, left[:4], left[4:]
    else:
        size_text = _colonial_size(left[size_i])[0]
        before = left[:size_i]
        desc_tokens = left[size_i + 1:]
    numbers = [_colonial_int(t) for t in before]
    item_i = next((i for i, n in enumerate(numbers) if n and 2 <= len(n) <= 6), None)
    item = numbers[item_i] if item_i is not None else None
    # La cantidad que cuenta es SHIPPED (la última antes de UNIT); si solo se
    # leyó una de las dos, igual se prueba: la cuenta del renglón decide.
    qtys = [n for n in numbers[item_i + 1:] if n is not None and len(n) <= 3] if item_i is not None else []
    shipped = int(qtys[-1]) if qtys else None
    description = " ".join(t for t in desc_tokens if re.search(r"\w", t))
    description = re.sub(r"^[^\w]+", "", description)
    if size_text is None and not (item and re.search(r"[A-Za-z]{3}", description)):
        return None  # el pie (ORD/SHIP/FEE/BALANCE) o una nota con números, no un producto
    item_word = next((w for w in words if _colonial_int(w["text"]) == item), None) if item else None
    size_word = next((w for w in words if size_text and _colonial_size(w["text"])), None)
    return {
        "item": item,
        "shipped": shipped,
        "size": size_text,
        "description": description,
        "price": price,
        "promo": promo,
        "amount": amount,
        "y": (min(w["top"] for w in words) + max(w["bottom"] for w in words)) / 2,
        "row_left": min(w["x0"] for w in words),
        "item_x1": item_word["x1"] if item_word else None,
        "size_x": (size_word["x0"], size_word["x1"]) if size_word else None,
        "cell_top": min(w["top"] for w in words) - 6,
        "cell_bottom": max(w["bottom"] for w in words) + 6,
    }


def _colonial_qty(row):
    """
    Cantidad y neto de un renglón: SHIPPED x (PRICE - PROMO) = AMOUNT. Si la
    promo no se leyó, se prueba sin promo (la cuenta lo confirma o no). Si
    SHIPPED no cierra, la cantidad se deduce del total (qty_derived) y hay
    que confirmarla después.
    """
    nets = []
    if row["promo"] is not None:
        nets.append(round(row["price"] - row["promo"], 2))
    if row["promo"] != 0:
        nets.append(row["price"])
    for net in nets:
        if row["shipped"] and net > 0 and _closes(row["shipped"], net, row["amount"]):
            return row["shipped"], net, False
    for net in nets:
        qty = _whole_qty(net, row["amount"]) if net > 0 else None
        if qty is not None:
            return qty, net, True
    return None, None, None


def _colonial_footer(image, rows):
    """
    Pie de la última página: ORD, SHIP, DELIVERY FEE y BALANCE DUE, en un
    recuadro que el OCR de página no lee (la grilla pegada a los números):
    se recorta la fila de valores, se borra la grilla y se prueban dos
    lecturas. Devuelve [(ship, fee, balance), ...] -- lecturas posibles, que
    valen solo si cierran con los renglones -- o [].
    """
    label = next((ws for ws in rows if re.search(r"DELIVER|BALANCE", _row_text(ws), re.I)), None)
    customer = next((ws for ws in rows if re.search(r"Cust\w*\W*Sig", _row_text(ws), re.I)), None)
    if label is None and customer is None:
        return []
    height = _median([w["bottom"] - w["top"] for ws in rows for w in ws if re.search(r"\d", w["text"])]) or 30
    if label is not None:
        top = max(w["bottom"] for w in label) + 2
        bottom = min(w["top"] for w in customer) - 2 if customer is not None else top + 3 * height
    else:
        bottom = min(w["top"] for w in customer) - 2
        top = bottom - 3 * height
    crop = image.crop((0, int(max(0, top)), image.width, int(min(image.height, bottom))))
    if crop.height < 8:
        return []
    cleaned = remove_grid_lines(crop, upscale=3)
    readings = []
    for psm in (6, 11):
        text = _pytesseract().image_to_string(cleaned, config=f"--psm {psm}")
        numbers = re.findall(r"\d[\d,]*[.,]\d{2}(?!\d)|\b\d{1,3}\b", text)
        ints = [int(n) for n in numbers if re.fullmatch(r"\d{1,3}", n)]
        money = [_money(n) for n in numbers if not re.fullmatch(r"\d{1,3}", n)]
        if len(ints) >= 2 and len(money) == 2:
            readings.append((ints[1], money[0], money[1]))
        elif len(ints) >= 2 and len(money) == 1:
            readings.append((ints[1], 0.0, money[0]))
    return readings


def extract_colonial_lines(pdf_path, invoices=None):
    """
    Renglones de una factura de Colonial (1 o 2 páginas). El OCR de página
    entera lee mal la tabla (renglones alternados con sombreado de puntitos),
    así que la tabla se lee cuatro veces (ver _colonial_page) y cada renglón
    junta sus lecturas:
    - ITEM# y SIZE valen solo con las reglas de _colonial_vote (no alcanza
      con "dos de tres"); si no, ValueError;
    - la cantidad es SHIPPED si cierra la cuenta SHIPPED x (PRICE - PROMO) =
      AMOUNT; si hay que deducirla del total, otra lectura o el pie tienen
      que confirmarla.

    Totales: los renglones más el delivery fee tienen que dar al centavo el
    importe que la app ya leyó del encabezado (`invoices`). El fee subió con
    el tiempo ($7.99 en 2025, $9.99 en mar-2026, $11.99 desde jun-2026): si
    el pie se puede leer (SHIP, fee y BALANCE DUE, que tienen que cerrar con
    los renglones), vale el fee leído; si no, uno de esos tres. Sin
    `invoices`, solo vale si el pie se leyó y cierra.
    """
    images = _page_images(pdf_path)
    if not images:
        raise ValueError("no se encontró la imagen escaneada de la factura de Colonial.")
    # N° de invoice: misma lectura que el encabezado de la app (cajita de
    # arriba a la derecha de la primera página).
    top_text = _pytesseract().image_to_string(crop_relative(images[0], 0.68, 0.0, 1.0, 0.22))
    match = re.search(r"\b(\d{7})\b", top_text)
    if not match:
        raise ValueError("no se encontró el N° de invoice en la factura de Colonial.")
    invoice_no = match.group(1)

    groups = []
    footer_readings = []
    for page in images:
        page_groups, page_footer = _colonial_page(page)
        groups.extend(page_groups)
        footer_readings = page_footer or footer_readings
    if not groups:
        raise ValueError("no se encontró ningún renglón de producto en la factura de Colonial.")

    lines = []
    unconfirmed_qty = []
    for group in groups:
        line_no = len(lines) + 1
        qty, net, amount, confirmed = _colonial_group_qty(group)
        if qty is None:
            description = next((r["description"] for r in group if r["description"]), "")
            raise ValueError(
                f"el renglón {line_no} de Colonial ({description}) no cierra: "
                "cantidad x precio no da el total del renglón."
            )
        if not confirmed:
            unconfirmed_qty.append(line_no)
        item = _colonial_vote(group, "item", line_no)
        size = _colonial_vote(group, "size", line_no)
        units = int(re.match(r"\d+", size).group(0))
        price = round(net + _colonial_promo_of(group, net), 2)
        lines.append(_line(
            line_no, item_no=item, description=_colonial_description(group),
            qty=qty, pack=units, size=size, units=units,
            price=price, allowance=round(price - net, 2) or None, net=net, ext=amount,
        ))

    lines_total = round(sum(line["ext"] for line in lines), 2)
    shipped_total = sum(line["qty"] for line in lines)
    # Pie leído (SHIP, fee y BALANCE DUE): vale la lectura que cierra con los renglones.
    footer = next(
        (f for f in footer_readings if f[0] == shipped_total and abs(lines_total + f[1] - f[2]) < 0.01), None
    )
    header_amounts = [inv["amount"] for inv in (invoices or []) if str(inv.get("invoice_no")) == invoice_no]
    if header_amounts:
        fees = [footer[1]] if footer else list(_COLONIAL_FEES)
        if not any(abs(lines_total + fee - amount) < 0.01 for fee in fees for amount in header_amounts):
            raise ValueError(
                f"los renglones suman ${lines_total:,.2f} y la factura es de ${header_amounts[0]:,.2f} "
                "(ni sumando el delivery fee da)."
            )
    elif footer is None:
        raise ValueError(
            f"los renglones suman ${lines_total:,.2f} y no hay un total de la factura para controlarlos."
        )
    if unconfirmed_qty and footer is None:
        raise ValueError(
            f"la cantidad del renglón {unconfirmed_qty[0]} de Colonial no se pudo leer con seguridad."
        )
    total = footer[2] if footer else next(
        a for a in header_amounts if any(abs(lines_total + fee - a) < 0.01 for fee in _COLONIAL_FEES)
    )
    return {"invoice_no": invoice_no, "lines": lines, "subtotal": lines_total, "total": total}


def _colonial_page(page):
    """
    Las cuatro lecturas de la tabla de una página, agrupadas por renglón, y
    las lecturas posibles del pie. Lecturas (medido contra el consenso en 20
    facturas): 0 = hoja tal cual (la mejor para SIZE, pero pega el borde de
    la tabla al ITEM#), 1 = sin líneas verticales (la mejor para ITEM#),
    2 = sin líneas al doble y sin sombreado, 3 = hoja tal cual al doble.
    """
    clean_page = _without_vertical_lines(page)
    lower = (0, page.height * 0.25, page.width, page.height)
    first_rows = _ocr_rows(page, box=lower)
    reads = [first_rows, _ocr_rows(clean_page, box=lower)]
    found = [r for rows in reads for r in (_colonial_row(ws) for ws in rows) if r]
    footer = _colonial_footer(page, first_rows)
    if not found:
        return [], footer
    height = _median([r["cell_bottom"] - r["cell_top"] for r in found])
    box = (0, min(r["cell_top"] for r in found) - height, page.width, max(r["cell_bottom"] for r in found) + height)
    reads.append(_ocr_rows(clean_page, box=box, scale=2, clean="blur"))
    reads.append(_ocr_rows(page, box=box, scale=2))

    # Renglones de las cuatro lecturas agrupados por altura (uno por lectura).
    groups = []
    for read_index, rows in enumerate(reads):
        for words in rows:
            row = _colonial_row(words)
            if row is None:
                continue
            row["read"] = read_index
            row["image"] = clean_page
            row["page"] = page
            group = min(groups, key=lambda g: abs(g[0]["y"] - row["y"]), default=None)
            if (group is not None and abs(group[0]["y"] - row["y"]) < height * 0.5
                    and all(r["read"] != read_index for r in group)):
                group.append(row)
            else:
                groups.append([row])
    groups.sort(key=lambda g: g[0]["y"])
    # El pie a veces sale en la lectura sin grilla como un renglón suelto de
    # números: "51 51 7.99 1436.71" (ORD, SHIP, DELIVERY FEE, BALANCE DUE).
    last_y = groups[-1][0]["y"] if groups else 0
    for rows in reads:
        for words in rows:
            if min(w["top"] for w in words) <= last_y:
                continue
            match = re.fullmatch(
                r"\W*(\d{1,3})\W+(\d{1,3})\W+(?:(\d+[.,]\d{2})\W+)?(\d[\d,]*[.,]\d{2})\W*", _row_text(words)
            )
            if match:
                fee = _money(match.group(3)) if match.group(3) else 0.0
                footer.append((int(match.group(2)), fee, _money(match.group(4))))
    return groups, footer


def _majority(values):
    """El valor que se repite al menos dos veces (el más repetido), o None."""
    counts = {}
    for value in values:
        if value is not None:
            counts[value] = counts.get(value, 0) + 1
    best = max(counts.items(), key=lambda kv: kv[1], default=(None, 0))
    return best[0] if best[1] >= 2 else None


def _colonial_group_qty(group):
    """
    (cantidad, neto, total, confirmada) de un renglón con varias lecturas.
    Vale la cantidad leída que cierra la cuenta (la que más lecturas
    repiten); si ninguna cierra, la deducida del total, y queda sin
    confirmar salvo que otra lectura haya leído esa misma cantidad.
    """
    closing = []
    derived = []
    amount = _majority([r["amount"] for r in group])
    for row in group:
        if amount is not None and row["amount"] != amount:
            continue  # la lectura que leyó otro total no cuenta
        qty, net, is_derived = _colonial_qty(row)
        if qty is None:
            continue
        (derived if is_derived else closing).append((qty, net, row["amount"]))
    for options, from_token in ((closing, True), (derived, False)):
        if not options:
            continue
        best = max(set(options), key=options.count)
        if options.count(best) == 1 and len(set(options)) > 1:
            return None, None, None, None  # lecturas que cierran con cuentas distintas
        qty = best[0]
        confirmed = from_token or sum(1 for r in group if r["shipped"] == qty) >= 1
        return best[0], best[1], best[2], confirmed
    return None, None, None, None


def _colonial_promo_of(group, net):
    """La promo (descuento por unidad) que explica el neto: la leída, o 0."""
    for row in group:
        if row["promo"] is not None and abs(row["price"] - row["promo"] - net) < 0.005:
            return row["promo"]
    return 0.0


def _colonial_description(group):
    """La descripción que más lecturas repiten; si no, la de la hoja tal cual (solo informativa)."""
    texts = [r["description"] for r in sorted(group, key=lambda r: r["read"]) if r["description"]]
    return _majority(texts) or (texts[0] if texts else "")


# Lectura principal de cada campo (la que nunca leyó mal en la medición) y
# de qué lecturas puede venir la confirmación.
_COLONIAL_PRIMARY = {"size": 0, "item": 1}


def _colonial_cell(row, field):
    """
    Relectura de la celda sola (sin grilla, agrandada): solo sirve para
    CONFIRMAR el valor de la lectura principal cuando las otras no lo
    leyeron. SIZE se relee de la hoja tal cual; ITEM#, de la hoja sin líneas.
    """
    if field == "size":
        if not row["size_x"]:
            return None
        box = (row["size_x"][0] - 10, row["cell_top"], row["size_x"][1] + 10, row["cell_bottom"])
        size = _colonial_size(_ocr_cell(row["page"], box, whitelist="0123456789CTPK"))
        return size[0] if size else None
    if not row["item_x1"]:
        return None
    box = (row["row_left"] - 20, row["cell_top"], row["item_x1"] + 10, row["cell_bottom"])
    text = _ocr_cell(row["image"], box)
    return text if re.fullmatch(r"\d{2,6}", text) else None


def _colonial_vote(group, field, line_no):
    """
    ITEM# / SIZE. No alcanza con "dos de tres" cualquiera: las lecturas sin
    grilla repiten a veces el mismo error (15CT -> "115CT" en dos) y un SIZE
    o un ITEM# mal leído es un costo por unidad o un producto equivocado.
    Vale el valor si:
    - la lectura principal del campo (la que nunca leyó mal en la medición:
      SIZE la hoja tal cual, ITEM# la hoja sin líneas) lo confirma otra
      lectura, o la celda sola si las demás no lo leyeron; o
    - la principal no lo leyó y lo dicen igual todas las demás que lo
      leyeron (al menos dos); o
    - lo dicen al menos tres lecturas y como mucho una dice otra cosa.
    Si no, ValueError.
    """
    primary_read = _COLONIAL_PRIMARY[field]
    by_read = {r["read"]: r[field] for r in group if r[field]}
    if field == "item":
        # El borde izquierdo pegado al ITEM# en las lecturas con grilla (0 y
        # 3): "17098" por 7098. No cuenta como lectura distinta.
        clean_values = {v for read, v in by_read.items() if read in (1, 2)}
        by_read = {read: v for read, v in by_read.items()
                   if not (read in (0, 3) and any(v[1:] == c for c in clean_values))}
    primary = by_read.get(primary_read)
    others = [v for read, v in by_read.items() if read != primary_read]
    values = list(by_read.values())
    value = None
    if primary is not None and primary in others:
        value = primary
    elif primary is not None and not others:
        row = next(r for r in group if r["read"] == primary_read)
        value = primary if _colonial_cell(row, field) == primary else None
    elif primary is None and len(others) >= 2 and len(set(others)) == 1:
        value = others[0]
    if value is None:
        best = max(set(values), key=values.count, default=None)
        if best is not None and values.count(best) >= 3 and len(values) - values.count(best) <= 1:
            value = best
    if value is None:
        label = "ITEM#" if field == "item" else "SIZE (unidades)"
        raise ValueError(f"el {label} del renglón {line_no} de Colonial no se pudo leer con seguridad.")
    return value


# supplier_key (el de SUPPLIER_REGISTRY en proveedores.py) -> extractor de renglones.
LINE_EXTRACTORS = {
    "ht_hackney": extract_ht_hackney_lines,
    "cec": extract_cec_lines,
    "colonial": extract_colonial_lines,
}


def extract_lines(supplier_key, pdf_path, invoices=None):
    """
    Detalle de productos de la factura, o None si el proveedor todavía no
    tiene extractor. `invoices`: los encabezados que la app ya leyó de ese
    PDF ({"invoice_no", "date", "amount"}) -- Colonial controla sus renglones
    contra ese importe, porque el total impreso de su factura no se lee por OCR.
    """
    extractor = LINE_EXTRACTORS.get(supplier_key)
    return extractor(pdf_path, invoices=invoices) if extractor else None


# ---------------------------------------------------------------------------
# Resumen por producto y cruce con el POS (CMV)
# ---------------------------------------------------------------------------

def _pct(part, whole):
    return round(part / whole * 100, 1) if whole else None


def _pos_diff(pos_cost, unit_cost):
    if pos_cost is None:
        return None
    diff = pos_cost - unit_cost
    return 0.0 if abs(diff) <= 0.0051 else round(diff, 2)


def _group_lines(lines):
    """{clave de producto: [renglones del más viejo al más nuevo]}."""
    by_product = {}
    for line in lines:
        key = product_key(line["supplier_key"], line["upc"], line["item_no"])
        by_product.setdefault(key, []).append(line)
    return by_product


def _pos_index(pos_costs):
    return {normalize_upc(row.get("upc")): row for row in pos_costs if normalize_upc(row.get("upc"))}


def _summary(key, product_lines, supplier_labels, pos_by_upc):
    last = product_lines[-1]
    # Compra anterior = el último renglón de OTRA factura (el mismo producto
    # puede venir dos veces en una misma factura).
    previous = next(
        (line for line in reversed(product_lines[:-1])
         if (line["supplier_key"], line["invoice_no"]) != (last["supplier_key"], last["invoice_no"])),
        None,
    )
    change = round(last["unit_cost"] - previous["unit_cost"], 4) if previous else None
    pos = pos_by_upc.get(last["upc"]) if last["upc"] else None
    pos_cost = pos.get("cost") if pos else None
    pos_price = pos.get("price") if pos else None
    costs = [line["unit_cost"] for line in product_lines]
    return {
        "key": key,
        "upc": last["upc"],
        "item_no": last["item_no"],
        "description": last["description"],
        "category": last["category"],
        "size": last["size"],
        "units_per_case": last["units"],
        "suppliers": sorted({supplier_labels.get(line["supplier_key"], line["supplier_key"]) for line in product_lines}),
        "purchases": len({(line["supplier_key"], line["invoice_no"]) for line in product_lines}),
        "units_bought": sum((line["qty"] or 0) * (line["units"] or 0) for line in product_lines),
        "first_date": product_lines[0]["invoice_date"],
        "last_date": last["invoice_date"],
        "last_cost": last["unit_cost"],
        "previous_cost": previous["unit_cost"] if previous else None,
        "previous_date": previous["invoice_date"] if previous else None,
        "change": change,
        "change_pct": _pct(change, previous["unit_cost"]) if previous else None,
        "min_cost": min(costs),
        "max_cost": max(costs),
        "srp": last["srp"],
        "pos_name": pos.get("name") if pos else None,
        "pos_dept": pos.get("dept_name") if pos else None,
        "pos_cost": pos_cost,
        "pos_price": pos_price,
        # Diferencia entre el costo que tiene cargado el POS y el costo real
        # de la última compra -- si no es 0, el CMV está calculando el margen
        # con un costo viejo. Medio centavo o menos es redondeo del POS
        # (ej. $12.35 la caja de 10 = $1.235 la unidad, en el POS $1.24).
        "pos_cost_diff": _pos_diff(pos_cost, last["unit_cost"]),
        "margin_pct": _pct(pos_price - last["unit_cost"], pos_price) if pos_price else None,
    }


def build_product_list(lines, supplier_labels, pos_costs):
    """
    Un renglón por producto comprado (clave = UPC, o Item # si la factura no
    trae UPC), del comprado más recientemente al más viejo.
    """
    pos_by_upc = _pos_index(pos_costs)
    products = [
        _summary(key, product_lines, supplier_labels, pos_by_upc)
        for key, product_lines in _group_lines(lines).items()
    ]
    products.sort(key=lambda p: (p["last_date"], p["description"] or ""), reverse=True)
    return products


# ---------------------------------------------------------------------------
# Carpetas: proveedor -> facturas por fecha -> productos (2026-10-02)
# ---------------------------------------------------------------------------
# Pedido del usuario: "distintos módulos con fechas dentro de ellos como si
# fueran carpetas", solo de los proveedores que se leen bien (los de
# LINE_EXTRACTORS -- los únicos que guardan renglones), y cada factura
# compara sus precios contra la factura ANTERIOR POR FECHA, no por orden de
# carga: si se sube una del medio después, la más nueva pasa a compararse
# contra esa. Por eso todo se recalcula al mostrar, ordenando por fecha.

def _compare_supplier_invoices(supplier_lines):
    """
    Facturas de un proveedor del más vieja a la más nueva (fecha y N°), cada
    renglón con el costo de la última compra ANTERIOR de ese producto (de
    una factura con fecha previa; dos renglones de una misma factura nunca
    se comparan entre sí). Devuelve [(fecha, N°, [renglones])].
    """
    by_invoice = {}
    for line in sorted(supplier_lines, key=lambda l: (l["invoice_date"], str(l["invoice_no"]), l["line_no"] or 0)):
        by_invoice.setdefault((line["invoice_date"], line["invoice_no"]), []).append(line)
    last_by_product = {}
    result = []
    for (invoice_date, invoice_no), invoice_lines in sorted(by_invoice.items(), key=lambda item: (item[0][0], str(item[0][1]))):
        rows = []
        seen = {}
        for line in invoice_lines:
            key = product_key(line["supplier_key"], line["upc"], line["item_no"])
            previous = last_by_product.get(key)
            change = round(line["unit_cost"] - previous["unit_cost"], 4) if previous else None
            if change is None:
                state = "Nuevo"
            elif change > 0.0001:
                state = "Subió"
            elif change < -0.0001:
                state = "Bajó"
            else:
                state = "Igual"
            rows.append({
                **line,
                "product_key": key,
                "previous_cost": previous["unit_cost"] if previous else None,
                "previous_date": previous["invoice_date"] if previous else None,
                "previous_invoice_no": previous["invoice_no"] if previous else None,
                "change": change,
                "change_pct": _pct(change, previous["unit_cost"]) if previous else None,
                "state": state,
            })
            seen[key] = line
        last_by_product.update(seen)
        result.append((invoice_date, invoice_no, rows))
    return result


def _invoice_summary(invoice_date, invoice_no, rows, previous):
    count = lambda state: sum(1 for row in rows if row["state"] == state)
    return {
        "invoice_date": invoice_date,
        "invoice_no": invoice_no,
        "lines": len(rows),
        "total": round(sum(row["ext"] or 0.0 for row in rows), 2),
        "up": count("Subió"),
        "down": count("Bajó"),
        "same": count("Igual"),
        "new": count("Nuevo"),
        "previous_date": previous[0] if previous else None,
        "previous_invoice_no": previous[1] if previous else None,
    }


def build_supplier_invoices(supplier_key, lines):
    """Facturas del proveedor con su resumen contra la anterior, de la más nueva a la más vieja."""
    compared = _compare_supplier_invoices([line for line in lines if line["supplier_key"] == supplier_key])
    summaries = []
    for index, (invoice_date, invoice_no, rows) in enumerate(compared):
        previous = compared[index - 1][:2] if index > 0 else None
        summaries.append(_invoice_summary(invoice_date, invoice_no, rows, previous))
    summaries.reverse()
    return summaries


def build_invoice_products(supplier_key, invoice_date, invoice_no, lines):
    """(resumen, renglones) de una factura con cada producto contra su compra anterior; None si no existe."""
    compared = _compare_supplier_invoices([line for line in lines if line["supplier_key"] == supplier_key])
    for index, (date_value, number, rows) in enumerate(compared):
        if date_value == invoice_date and str(number) == str(invoice_no):
            previous = compared[index - 1][:2] if index > 0 else None
            return _invoice_summary(date_value, number, rows, previous), rows
    return None


def build_supplier_folders(lines, supplier_labels):
    """Una carpeta por proveedor con renglones guardados: cuántas facturas, de qué fechas y cómo vino la última."""
    keys = sorted({line["supplier_key"] for line in lines}, key=lambda k: supplier_labels.get(k, k).lower())
    folders = []
    for key in keys:
        invoices = build_supplier_invoices(key, lines)
        folders.append({
            "key": key,
            "label": supplier_labels.get(key, key),
            "invoices": len(invoices),
            "first_date": invoices[-1]["invoice_date"] if invoices else None,
            "last": invoices[0] if invoices else None,
            "years": sorted({inv["invoice_date"][:4] for inv in invoices}, reverse=True),
        })
    return folders


def build_product_detail(key, lines, supplier_labels, pos_costs, monthly_sales):
    """
    Resumen + historial de compras de un producto, y sus ventas por mes del
    POS (CMV) con el costo real de ese momento -- para ver a cuánto se vende
    contra lo que costó. None si ese producto nunca se compró.
    """
    product_lines = _group_lines(lines).get(key)
    if not product_lines:
        return None
    summary = _summary(key, product_lines, supplier_labels, _pos_index(pos_costs))

    purchases = []
    previous_cost = None
    for line in product_lines:
        change = round(line["unit_cost"] - previous_cost, 4) if previous_cost is not None else None
        purchases.append({
            **line,
            "supplier_label": supplier_labels.get(line["supplier_key"], line["supplier_key"]),
            "change": change,
            "change_pct": _pct(change, previous_cost) if change is not None else None,
        })
        previous_cost = line["unit_cost"]
    purchases.reverse()

    sales = []
    if summary["upc"]:
        for row in monthly_sales:
            if normalize_upc(row.get("upc")) != summary["upc"] or not row.get("count"):
                continue
            month_end = f"{row['year']:04d}-{row['month']:02d}-31"
            # Costo vigente en ese mes = el de la última compra hasta fin de mes.
            cost_lines = [line for line in product_lines if line["invoice_date"] <= month_end]
            cost = cost_lines[-1]["unit_cost"] if cost_lines else None
            avg_price = row["amount"] / row["count"]
            sales.append({
                "year": row["year"],
                "month": row["month"],
                "count": row["count"],
                "amount": row["amount"],
                "avg_price": round(avg_price, 2),
                "cost": cost,
                "margin_pct": _pct(avg_price - cost, avg_price) if cost is not None and avg_price else None,
            })
        sales.reverse()

    return {"product": summary, "purchases": purchases, "sales": sales}


# ---------------------------------------------------------------------------
# Reporte para el manager (pedido del usuario, 2026-10-02): "enviar un
# reporte en Excel o PDF directamente al manager de la estación de servicio
# con los productos que cambiaron de precio". Solo los que subieron o
# bajaron contra su compra anterior; con el precio del POS y el margen que
# deja el costo nuevo, para que sepa qué precio revisar.
# ---------------------------------------------------------------------------

def price_change_rows(rows, pos_costs):
    """Renglones de una factura (build_invoice_products) que cambiaron de costo: primero los que subieron."""
    pos_by_upc = _pos_index(pos_costs)
    changed = []
    for row in rows:
        if row["state"] not in ("Subió", "Bajó"):
            continue
        pos = pos_by_upc.get(row["upc"]) if row["upc"] else None
        pos_price = pos.get("price") if pos else None
        margin = _pct(pos_price - row["unit_cost"], pos_price) if pos_price else None
        changed.append({**row, "pos_price": pos_price, "margin_pct": margin})
    changed.sort(key=lambda r: (0 if r["state"] == "Subió" else 1, r.get("category") or "", r.get("description") or ""))
    return changed


def _ddmmyyyy(iso):
    return f"{iso[8:10]}/{iso[5:7]}/{iso[0:4]}" if iso else ""


def _pack_label(row):
    return " ".join(str(part) for part in (row.get("pack"), row.get("size")) if part not in (None, ""))


def build_price_change_pdf(supplier_label, invoice, changed, dest_path):
    from pdf_export import build_simple_table_pdf

    def money(value):
        # Hasta 4 decimales (costos por unidad de centavos), nunca menos de 2.
        if value is None:
            return ""
        text = f"{abs(value):,.4f}".rstrip("0")
        if len(text.split(".")[1]) < 2:
            text = f"{abs(value):,.2f}"
        return f"{'-' if value < 0 else ''}${text}"

    table_rows = []
    for row in changed:
        sign = "+" if row["change"] > 0 else ""
        table_rows.append([
            row["upc"] or "",
            row.get("description") or "",
            _pack_label(row),
            f"{money(row['previous_cost'])} ({_ddmmyyyy(row['previous_date'])})",
            money(row["unit_cost"]),
            f"{sign}{money(row['change'])} ({sign}{row['change_pct']}%)",
            money(row.get("srp")),
            money(row.get("pos_price")),
            "" if row.get("margin_pct") is None else f"{row['margin_pct']}%",
        ])
    up = sum(1 for row in changed if row["state"] == "Subió")
    build_simple_table_pdf(
        dest_path,
        f"{supplier_label} — Cambios de precio — Factura N° {invoice['invoice_no']}",
        ["UPC", "Producto", "Pack", "Costo anterior (fecha)", "Costo nuevo", "Cambio", "SRP factura", "Precio POS", "Margen POS"],
        table_rows,
        col_widths_mm=[30, 82, 22, 40, 24, 36, 22, 22, 22],
        company_header=True,
        period_label=(
            f"Factura del {_ddmmyyyy(invoice['invoice_date'])} — {up} subieron, {len(changed) - up} bajaron"
            + (" — comparado con la compra anterior de cada producto" if changed else " — ningún producto cambió de precio")
        ),
        footer_note="Costo por unidad. Margen POS = (precio del POS − costo nuevo) / precio del POS.",
    )
    return dest_path


def build_price_change_workbook(supplier_label, invoice, changed, dest_path):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Cambios de precio"
    sheet.append([f"{supplier_label} — Factura N° {invoice['invoice_no']} del {_ddmmyyyy(invoice['invoice_date'])}"])
    sheet["A1"].font = Font(bold=True, size=12)
    sheet.append([])
    headers = ["Cambio", "UPC", "Producto", "Categoría", "Pack", "Costo anterior", "Fecha anterior",
               "Costo nuevo", "Diferencia", "Diferencia %", "SRP factura", "Precio POS", "Margen POS %"]
    sheet.append(headers)
    for cell in sheet[3]:
        cell.font = Font(bold=True)
    up_fill = PatternFill("solid", fgColor="FEE2E2")
    down_fill = PatternFill("solid", fgColor="DCFCE7")
    money_format = '"$"#,##0.00##'
    for row in changed:
        sheet.append([
            row["state"], row["upc"], row.get("description"), row.get("category"), _pack_label(row),
            row["previous_cost"], _ddmmyyyy(row["previous_date"]), row["unit_cost"], row["change"],
            row["change_pct"], row.get("srp"), row.get("pos_price"), row.get("margin_pct"),
        ])
        new_row = sheet[sheet.max_row]
        new_row[0].fill = up_fill if row["state"] == "Subió" else down_fill
        for index in (5, 7, 8, 10, 11):
            new_row[index].number_format = money_format
        for cell in new_row[5:]:
            cell.alignment = Alignment(horizontal="center")
    for col_letter, width in zip("ABCDEFGHIJKLM", (9, 15, 38, 16, 10, 14, 14, 12, 12, 12, 12, 12, 13)):
        sheet.column_dimensions[col_letter].width = width
    if changed:
        sheet.auto_filter.ref = f"A3:M{sheet.max_row}"
    sheet.freeze_panes = "A4"
    workbook.save(dest_path)
    return dest_path
