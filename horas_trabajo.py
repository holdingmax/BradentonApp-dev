"""
Horas de Trabajo -- lee el "Clock In/Out Detail Report" semanal (Chevron
POS) que llega como PDF cada lunes -- pedido explícito del usuario
(2026-09-18): "antes solia cargarlos en el excel [BDT. HOURS...]... ahora
deberian empezar a aparecer en un cuadrito asi como salen los eft uno abajo
del otro que solo va a haber 1 a la semana".

Formato real del reporte (confirmado contra varios PDFs reales de agosto-26,
todos escaneados con CamScanner -- sin texto digital, hace falta OCR): un
encabezado con "REPORT PRINTED: MM/DD/YY h:mm:ss AM/PM" (la fecha real de
esa semana) y "PERIOD FROM: Mon DD, YYYY ... TO: Mon DD, YYYY", seguido de
una tabla "Employee ID | Employee Name | ClockIn Time | ClockOut Time |
Hours Worked | Total" -- cada empleado trae una o más filas de detalle (un
turno cada una) y, al pie de su propio bloque, un renglón APARTE (alineado
bajo la columna "Total", visualmente más abajo y a la derecha de su última
fila) con la suma de horas de la semana para ese empleado -- ese valor
(HH:MM) es el que este módulo extrae como "horas trabajadas" -- el pago se
calcula sobre esa suma semanal, nunca sobre cada turno suelto.

Regla de oro del proyecto: si no se puede leer con confianza el total de un
empleado (el bloque se detectó por ID+Nombre pero nunca apareció su renglón
de Total), no se inventa ningún valor -- se devuelve en
`unresolved_employees` para que quede a la vista y el usuario lo agregue a
mano ("+ Agregar empleado", ya editable de todos modos en la pantalla).
"""

import re
from datetime import datetime

try:
    import pdfplumber
except ImportError:  # pragma: no cover - environment guard
    pdfplumber = None  # type: ignore[assignment]

try:
    import pytesseract
except ImportError:  # pragma: no cover - environment guard
    pytesseract = None  # type: ignore[assignment]

from ocr_utils import (
    ensure_pdfplumber as _ensure_pdfplumber,
    ensure_pytesseract as _ensure_pytesseract,
    extract_largest_page_image as _extract_page_image,
)

DEFAULT_HOURLY_RATE = 15.0

_REPORT_PRINTED_RE = re.compile(r"REPORT\s*PRINTED:?\s*(\d{1,2}/\d{1,2}/\d{2,4})", re.IGNORECASE)
_PERIOD_RE = re.compile(
    r"PERIOD\s*FROM:?\s*([A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4}).*?TO:?\s*([A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4})",
    re.IGNORECASE | re.DOTALL,
)
# Cada fila de detalle trae una fecha de ClockIn bien legible (dos dígitos
# grandes de mes/día/año) -- mucho más confiable que la columna Employee ID,
# que confirmado contra un PDF real (semana del 10/08) se lee mal seguido
# (un solo dígito chico -- "2" salió "PD" en una fila y "NNMNN" en otra del
# mismo empleado). En vez de depender del ID, se ubica esa fecha y se toman
# las DOS palabras en mayúscula que la preceden inmediatamente -- los
# nombres reales de este negocio son siempre Nombre + Apellido -- cualquier
# basura de OCR que haya quedado pegada más a la izquierda (el ID mal
# leído) queda afuera sola, sin hacer falta reconocerla como tal.
_DATE_TOKEN_RE = re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}")
_CAPS_WORD_RE = re.compile(r"^[A-Z][A-Z.'\-]*$")
# El renglón de Total de un empleado queda solo en su propia línea -- nada
# más que un valor H:MM (1-3 dígitos de horas, sin límite real de 24hs
# porque es una suma semanal).
_TOTAL_LINE_RE = re.compile(r"^\s*(\d{1,3}):(\d{2})\s*$")


def _extract_employee_name(line):
    m_date = _DATE_TOKEN_RE.search(line)
    if not m_date:
        return None
    prefix_words = line[: m_date.start()].split()
    if len(prefix_words) < 2:
        return None
    name_words = prefix_words[-2:]
    if not all(_CAPS_WORD_RE.match(w) for w in name_words):
        return None
    return " ".join(name_words)


def _parse_short_date(text):
    text = text.strip()
    for fmt in ("%m/%d/%y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _parse_long_date(text):
    text = re.sub(r"\s+", " ", text.strip().rstrip(","))
    text = text.replace(",", "")
    for fmt in ("%b %d %Y", "%B %d %Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _ocr_text_by_rows(image):
    """
    OCR reconstruyendo el orden de lectura por posición real (Y/X de cada
    palabra), en vez de confiar en `pytesseract.image_to_string()` -- se
    encontró contra un reporte real (semana del 10/08) que, con columnas muy
    separadas (ID/Nombre/ClockIn/ClockOut/Horas/Total), el análisis de
    layout automático de Tesseract a veces agrupa el texto POR COLUMNA
    entera (todos los ID, después todos los Nombre, después...) en vez de
    por fila -- mismo tipo de problema ya documentado en este proyecto para
    otras tablas anchas (ver CLAUDE.md, Bimbo/Midtown), resuelto siempre
    reconstruyendo la posición real con `image_to_data()` en vez de confiar
    en el orden de lectura que Tesseract elige solo.
    """
    _ensure_pytesseract()
    data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
    words = []
    for i, raw_text in enumerate(data["text"]):
        text = (raw_text or "").strip()
        if not text:
            continue
        height = data["height"][i]
        words.append(
            {
                "text": text,
                "left": data["left"][i],
                "cy": data["top"][i] + height / 2.0,
                "height": height,
            }
        )
    if not words:
        return ""

    words.sort(key=lambda w: w["cy"])
    rows = []
    current_row = [words[0]]
    row_cy_values = [words[0]["cy"]]
    for w in words[1:]:
        row_center = sum(row_cy_values) / len(row_cy_values)
        tolerance = max(12, current_row[-1]["height"] * 0.6)
        if abs(w["cy"] - row_center) <= tolerance:
            current_row.append(w)
            row_cy_values.append(w["cy"])
        else:
            rows.append(current_row)
            current_row = [w]
            row_cy_values = [w["cy"]]
    rows.append(current_row)

    lines = []
    for row in rows:
        row.sort(key=lambda w: w["left"])
        lines.append(" ".join(w["text"] for w in row))
    return "\n".join(lines)


def _page_text(page):
    """Texto digital si el PDF lo trae; si no (el caso real de este
    reporte -- viene de un escaneo CamScanner, sin capa de texto), OCR sobre
    la imagen más grande de la página, reconstruyendo el orden por fila."""
    text = (page.extract_text() or "").strip()
    if text:
        return text
    image = _extract_page_image(page)
    if image is None:
        return ""
    return _ocr_text_by_rows(image)


def extract_hours_report(pdf_path):
    """
    Devuelve:
    {
        "report_date": date|None,       # de "REPORT PRINTED"
        "period_from": date|None,
        "period_to": date|None,
        "employees": [
            {"employee_name": str, "hours_label": "20:48", "hours": 20.8},
            ...
        ],
        "unresolved_employees": [str, ...],  # detectados pero sin Total legible
    }
    """
    _ensure_pdfplumber()
    report_date = None
    period_from = None
    period_to = None
    employees = []
    unresolved = []
    current = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = _page_text(page)
            if not text:
                continue
            if report_date is None:
                m = _REPORT_PRINTED_RE.search(text)
                if m:
                    report_date = _parse_short_date(m.group(1))
            if period_from is None:
                m = _PERIOD_RE.search(text)
                if m:
                    period_from = _parse_long_date(m.group(1))
                    period_to = _parse_long_date(m.group(2))

            for raw_line in text.splitlines():
                line = raw_line.strip()
                if not line:
                    continue

                employee_name = _extract_employee_name(line)
                if employee_name:
                    if current is None or current["employee_name"] != employee_name:
                        if current is not None and "hours_label" not in current:
                            unresolved.append(current["employee_name"])
                        current = {"employee_name": employee_name}
                    continue

                m_total = _TOTAL_LINE_RE.match(line)
                if m_total and current is not None and "hours_label" not in current:
                    hours_label = f"{int(m_total.group(1))}:{m_total.group(2)}"
                    hours = int(m_total.group(1)) + int(m_total.group(2)) / 60.0
                    current["hours_label"] = hours_label
                    current["hours"] = round(hours, 2)
                    employees.append(
                        {
                            "employee_name": current["employee_name"],
                            "hours_label": hours_label,
                            "hours": round(hours, 2),
                        }
                    )
                    current = None

    if current is not None and "hours_label" not in current:
        unresolved.append(current["employee_name"])

    return {
        "report_date": report_date,
        "period_from": period_from,
        "period_to": period_to,
        "employees": employees,
        "unresolved_employees": unresolved,
    }
