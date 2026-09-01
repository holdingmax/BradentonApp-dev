"""
Helpers de OCR/PDF compartidos entre los módulos que necesitan leer PDFs
escaneados o fotografiados (Reporte Diario, Gettel/Toyota Pagos, Proveedores).

Antes esta lógica vivía casi al carácter triplicada en esos tres módulos —
cualquier fix (ej. el auto-detect de la ruta de Tesseract en Windows, o el
manejo de rotación via Tesseract OSD) había que acordarse de aplicarlo en los
tres por separado. Ahora vive acá una sola vez; cada módulo solo importa lo
que necesita.
"""

import io
import os

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
    fechas) lo suficiente como para romper la lectura. Nunca lanza -- una
    página que Tesseract no puede leer con confianza queda como está.
    """
    try:
        osd = pytesseract.image_to_osd(image, output_type=pytesseract.Output.DICT)
        rotate = int(osd.get("rotate", 0) or 0)
    except Exception:
        return image
    transpose_const = _OSD_ROTATE_TO_TRANSPOSE.get(rotate)
    if transpose_const is not None:
        image = image.transpose(transpose_const)
    return image


def extract_largest_page_image(page):
    """La imagen más grande incrustada en una página de pdfplumber, ya orientada, o None."""
    if not page.images:
        return None
    biggest = max(page.images, key=lambda im: im["width"] * im["height"])
    raw = biggest["stream"].get_data()
    image = Image.open(io.BytesIO(raw)).convert("RGB")
    return correct_image_orientation(image)


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
