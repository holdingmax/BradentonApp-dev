"""
Helpers de OCR/PDF compartidos entre los módulos que necesitan leer PDFs
escaneados o fotografiados (Reporte Diario, Gettel/Toyota Pagos, Proveedores).

Antes esta lógica vivía casi al carácter triplicada en esos tres módulos —
cualquier fix (ej. el auto-detect de la ruta de Tesseract en Windows, o el
manejo de rotación via Tesseract OSD) había que acordarse de aplicarlo en los
tres por separado. Ahora vive acá una sola vez; cada módulo solo importa lo
que necesita.
"""

import copy
import functools
import hashlib
import io
import os
import threading
from collections import OrderedDict

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

_TESSERACT_CANDIDATE_PATHS = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
)
_TESSERACT_CONFIGURED = False

_OSD_ROTATE_TO_TRANSPOSE = {
    90: Image.ROTATE_270 if Image is not None else None,
    180: Image.ROTATE_180 if Image is not None else None,
    270: Image.ROTATE_90 if Image is not None else None,
}


# Lecturas de Tesseract guardadas por el contenido exacto de la imagen
# (2026-10-14). Varios lectores le pasaban el OCR a la misma imagen con la
# misma configuración más de una vez (el reconocimiento del proveedor y
# después su lector, el encabezado y después los renglones, la orientación de
# la misma hoja en cada pasada): Frito-Lay repetía 5 de sus 7 lecturas,
# Midtown 5 de 13. Tesseract da siempre lo mismo para la misma imagen y la
# misma configuración, así que la segunda vez se devuelve lo ya leído. Se
# instala una sola vez sobre pytesseract.image_to_string/_data/_osd (todos los
# módulos llaman por ahí); la clave es la función, sus argumentos y el MD5 de
# los píxeles (una imagen recortada, girada o retocada es otra imagen). Una
# ruta de archivo en vez de imagen no se guarda.
_OCR_MEMO = OrderedDict()
_OCR_MEMO_LOCK = threading.Lock()
_OCR_MEMO_MAX_BYTES = 64 * 1024 * 1024
_ocr_memo_bytes = 0


def _image_digest(image):
    if Image is not None and isinstance(image, Image.Image):
        palette = image.getpalette() if image.mode in ("P", "PA") else None
        return ("pil", image.mode, image.size, hashlib.md5(image.tobytes()).hexdigest(),
                hashlib.md5(bytes(palette)).hexdigest() if palette else None, str(image.info.get("transparency")))
    if np is not None and isinstance(image, np.ndarray):
        return ("np", str(image.dtype), image.shape, hashlib.md5(np.ascontiguousarray(image).tobytes()).hexdigest())
    return None


def _memo_size(value):
    if isinstance(value, dict):
        return 64 + sum(48 * len(v) if isinstance(v, list) else 64 for v in value.values())
    return 64 + len(value) if isinstance(value, (str, bytes)) else 256


def _memoized_ocr(func):
    # Los errores no se guardan: retendrían en su traceback las imágenes de
    # quien llamó, y uno pasajero (Tesseract sin memoria, el antivirus con el
    # PNG temporal) quedaría fijo hasta reiniciar (revisión 2026-10-14).
    @functools.wraps(func)
    def wrapper(image, *args, **kwargs):
        global _ocr_memo_bytes
        digest = _image_digest(image)
        try:
            key = (func.__name__, digest, args, tuple(sorted(kwargs.items())))
            hash(key)
        except TypeError:
            digest = None
        if digest is None:
            return func(image, *args, **kwargs)
        with _OCR_MEMO_LOCK:
            cached = _OCR_MEMO.get(key)
            if cached is not None:
                _OCR_MEMO.move_to_end(key)
        if cached is not None:
            return copy.deepcopy(cached[0])
        value = func(image, *args, **kwargs)
        size = _memo_size(value)
        if size <= _OCR_MEMO_MAX_BYTES // 8:
            with _OCR_MEMO_LOCK:
                if key not in _OCR_MEMO:
                    _OCR_MEMO[key] = (copy.deepcopy(value), size)
                    _ocr_memo_bytes += size
                while _ocr_memo_bytes > _OCR_MEMO_MAX_BYTES and _OCR_MEMO:
                    _, (_, old_size) = _OCR_MEMO.popitem(last=False)
                    _ocr_memo_bytes -= old_size
        return value

    wrapper.bradenton_memo = True
    return wrapper


def _install_ocr_memo():
    if pytesseract is None:
        return
    for name in ("image_to_string", "image_to_data", "image_to_osd"):
        func = getattr(pytesseract, name)
        if not getattr(func, "bradenton_memo", False):
            setattr(pytesseract, name, _memoized_ocr(func))


_install_ocr_memo()


def ensure_pdfplumber():
    if pdfplumber is None:
        raise ImportError(
            "Esta operación requiere pdfplumber. Instale con: pip install pdfplumber"
        )


def ensure_pytesseract():
    """Auto-detecta el path de Tesseract en Windows si no está en el PATH."""
    global _TESSERACT_CONFIGURED
    if pytesseract is None or Image is None:
        raise ImportError(
            "Leer PDFs escaneados/fotografiados requiere pytesseract y Pillow. "
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
        "UB-Mannheim.TesseractOCR') para poder leer PDFs escaneados/fotografiados."
    )


def ensure_cv2():
    if cv2 is None or np is None:
        raise ImportError(
            "Esta operación requiere opencv-python y numpy. "
            "Instale con: pip install opencv-python numpy"
        )


def correct_image_orientation(image):
    """
    Undo whole-page rotation (scans o fotos de celular tomadas de costado o
    al revés) via la propia detección de orientación de Tesseract, antes de
    correr el OCR real. Usa Image.transpose (remapeo exacto en múltiplos de
    90°) en vez de Image.rotate -- .rotate() resamplea cada pixel incluso en
    giros de ángulo recto, lo que difumina texto (sobre todo el "/" de
    fechas) lo suficiente como para romper la lectura. Nunca lanza por OCR -- una
    página que Tesseract no puede leer con confianza queda como está.

    ensure_pytesseract() va primero y fuera del try: sin él, en una PC donde
    Tesseract no está en el PATH, el primer OCR del proceso fallaba en
    silencio y la imagen volvía sin rotar (auditoría 2026-09). Si Tesseract
    no está instalado, ese error sí se propaga.
    """
    ensure_pytesseract()
    return apply_osd_rotation(image, detect_osd_rotation(image) or 0)


def ocr_process_limit():
    """
    Cuántos procesos de Tesseract (o archivos) leer a la vez (2026-10-14): uno
    por núcleo, hasta 4. Cada uno ocupa ~100 MB y en un contenedor
    os.cpu_count() es el del servidor entero, no el del plan: en Render el
    valor por defecto es 2. BRADENTON_OCR_PROCESOS lo fija a mano.
    """
    try:
        configured = int(os.environ.get("BRADENTON_OCR_PROCESOS", "0"))
    except ValueError:
        configured = 0
    if configured > 0:
        return configured
    return 2 if os.environ.get("RENDER") else max(1, min(os.cpu_count() or 1, 4))


# PDFium (pypdfium2, lo que usa pdfplumber para dibujar una página) no admite
# llamadas desde varios hilos a la vez: con las cargas que leen varios
# archivos en paralelo daba "access violation" o tiraba el proceso
# (reproducido en la revisión del 2026-10-14). Toda página que se dibuja pasa
# por acá, de a una por vez en todo el proceso.
_RENDER_LOCK = threading.Lock()


def render_page(page, resolution):
    """La página de pdfplumber dibujada como imagen RGB (con el candado de PDFium)."""
    with _RENDER_LOCK:
        return page.to_image(resolution=resolution).original.convert("RGB")


def detect_osd_rotation(image):
    """El giro (0/90/180/270) que pide la OSD de Tesseract, o None si no la pudo leer."""
    ensure_pytesseract()
    try:
        osd = pytesseract.image_to_osd(image, output_type=pytesseract.Output.DICT)
        return int(osd.get("rotate", 0) or 0)
    except Exception:
        return None


def apply_osd_rotation(image, rotate):
    """Endereza la imagen con el giro de detect_osd_rotation (transpose: remapeo exacto, ver arriba)."""
    transpose_const = _OSD_ROTATE_TO_TRANSPOSE.get(rotate)
    if transpose_const is not None:
        image = image.transpose(transpose_const)
    return image


def largest_page_image(page):
    """La imagen más grande incrustada en una página de pdfplumber, tal cual (sin orientar), o None."""
    if not page.images:
        return None
    biggest = max(page.images, key=lambda im: im["width"] * im["height"])
    raw = biggest["stream"].get_data()
    return Image.open(io.BytesIO(raw)).convert("RGB")


def extract_largest_page_image(page):
    """La imagen más grande incrustada en una página de pdfplumber, ya orientada, o None."""
    image = largest_page_image(page)
    return None if image is None else correct_image_orientation(image)


def crop_relative(image, left, top, right, bottom):
    """Recorte por fracción del ancho/alto (0.0-1.0), no por píxeles fijos."""
    w, h = image.size
    return image.crop((int(w * left), int(h * top), int(w * right), int(h * bottom)))


def remove_grid_lines(image, upscale=3):
    """
    Algunas facturas/reportes meten el texto buscado en una tabla con bordes
    que confunden al OCR (o directamente no lee nada adentro). Sube la
    resolución y borra las líneas horizontales/verticales con morfología de
    OpenCV antes de correr Tesseract.
    """
    ensure_cv2()
    gray = np.array(
        image.convert("L").resize((image.width * upscale, image.height * upscale), Image.LANCZOS)
    )
    _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    horiz_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 120))
    lines = cv2.bitwise_or(
        cv2.morphologyEx(bw, cv2.MORPH_OPEN, horiz_kernel, iterations=2),
        cv2.morphologyEx(bw, cv2.MORPH_OPEN, vert_kernel, iterations=2),
    )
    cleaned = cv2.bitwise_not(cv2.bitwise_and(bw, cv2.bitwise_not(lines)))
    return Image.fromarray(cleaned)
