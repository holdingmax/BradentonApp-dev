"""
Horas de Trabajo -- lee el "Clock In/Out Detail Report" semanal (Chevron
POS) que llega como PDF cada lunes -- pedido explícito del usuario
(2026-09-18): "antes solia cargarlos en el excel [BDT. HOURS...]... ahora
deberian empezar a aparecer en un cuadrito asi como salen los eft uno abajo
del otro que solo va a haber 1 a la semana".

Formato real del reporte (escaneado con CamScanner, sin texto digital: hace
falta OCR): un encabezado con "REPORT PRINTED: MM/DD/YY h:mm:ss AM/PM" (la
fecha real de esa semana) y "PERIOD FROM: Mon DD, YYYY ... TO: Mon DD,
YYYY", seguido de una tabla "Employee ID | Employee Name | ClockIn Time |
ClockOut Time | Hours Worked | Total" -- cada empleado trae una fila por
turno y, al pie de su bloque, un renglón APARTE (bajo la columna "Total")
con la suma de horas de la semana: ese valor (HH:MM) es el que se paga.

2026-10-05 (pedido del usuario: "que nunca se ponga mal la cantidad de horas
y el nombre del empleado"): antes se leía una sola vez y se tomaba el Total
tal cual, sin control. Ahora el reporte se controla consigo mismo:
- cada turno trae su entrada, su salida y las horas (Hours Worked): la
  cuenta salida - entrada tiene que dar esas horas;
- el Total de cada empleado tiene que ser la suma de sus turnos;
- el nombre se repite en cada turno: vale el más votado y, si se parece casi
  letra por letra a un empleado de semanas anteriores, el conocido;
- la página se lee hasta cinco veces con distintos arreglos de la imagen
  (de a una, solo mientras algo no cierre) y cada dato se vota; la última
  lectura es de texto disperso, que encuentra los Totales que las otras
  pierden;
- si falta el Total de un empleado, se avisa por su nombre (no junto con el
  siguiente).

Regla de oro del proyecto: si las horas de un empleado no cierran (la suma
de sus turnos no da el Total, o hay un dato ilegible que no se puede
deducir), no se inventa ningún valor -- va a `unresolved_employees` para que
el usuario lo cargue a mano.
"""

import difflib
import itertools
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

try:
    from PIL import Image, ImageFilter
except ImportError:  # pragma: no cover - environment guard
    Image = ImageFilter = None  # type: ignore[assignment]

from ocr_utils import (
    ensure_pdfplumber as _ensure_pdfplumber,
    ensure_pytesseract as _ensure_pytesseract,
    extract_largest_page_image as _extract_page_image,
)

DEFAULT_HOURLY_RATE = 15.0

_HOURS_LABEL_RE = re.compile(r"^\s*(\d{1,3}):([0-5]\d)\s*$")
_REPORT_PRINTED_RE = re.compile(r"REPORT\s*PRINTED:?\s*(\d{1,2}/\d{1,2}/\d{2,4})", re.IGNORECASE)
_PERIOD_RE = re.compile(
    r"PERIOD\s*FROM:?\s*([A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4}).*?TO:?\s*([A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4})",
    re.IGNORECASE,
)
# Un renglón de turno: "10 CRYSTAL RHODEN 08/04/26 9:49:00 AM 08/04/26 4:04:00 PM 06:15" (o
# "CLOCKED IN" en lugar de la salida, si el turno quedó abierto).
_STAMP = r"(\d{1,2})/(\d{1,2})/(\d{2,4})\s+(\d{1,2}):(\d{2})(?::\d{2})?\s*([AP])\.?\s*[MN]"
_DETAIL_RE = re.compile(
    rf"^(?P<pre>.*?){_STAMP}\s+(?:{_STAMP}|(?P<open>CLOCKED\s*IN))\s*(?P<rest>.*)$", re.IGNORECASE
)
# El Total de cada empleado: solo un H:MM en su propio renglón, debajo de la columna Total.
_TOTAL_RE = re.compile(r"^\W*(\d{1,3}):([0-5]\d)\W*$")
_SHIFT_RE = re.compile(r"^\W*(\d{1,2}):([0-5]\d)(?!\d)")


def hours_to_label(hours):
    """26.8 -> "26:48" (minutos redondeados)."""
    total_minutes = int(round((hours or 0.0) * 60))
    return f"{total_minutes // 60}:{total_minutes % 60:02d}"


def parse_hours_input(text):
    """
    Horas tipeadas a mano: "26:48" (HH:MM, como el reporte) o "26.8"
    (horas decimales). Regla de pago confirmada por el usuario (2026-10-01):
    se paga el tiempo real, 26:48 = 26.8 h -- nunca 26.48. Devuelve
    (horas, etiqueta HH:MM). ValueError si no se entiende.
    """
    cleaned = (text or "").strip()
    match = _HOURS_LABEL_RE.match(cleaned)
    if match:
        hours = int(match.group(1)) + int(match.group(2)) / 60.0
        return hours, f"{int(match.group(1))}:{match.group(2)}"
    hours = float(cleaned.replace(",", "."))
    return hours, hours_to_label(hours)


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


def _sharpen(image):
    return image.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))


def _otsu(image):
    gray = image.convert("L")
    histogram = gray.histogram()
    total = sum(histogram)
    sum_all = sum(i * h for i, h in enumerate(histogram))
    best, threshold, weight, sum_back = 0.0, 128, 0, 0
    for i, h in enumerate(histogram):
        weight += h
        if weight == 0 or weight == total:
            continue
        sum_back += i * h
        mean_back = sum_back / weight
        mean_fore = (sum_all - sum_back) / (total - weight)
        between = weight * (total - weight) * (mean_back - mean_fore) ** 2
        if between > best:
            best, threshold = between, i
    return gray.point(lambda v: 255 if v > threshold else 0)


# Lecturas de cada página, de la que más renglones lee bien a la que menos;
# se hacen de a una y solo hasta que todos los empleados cierran.
_PASSES = (
    ("--psm 6", 1.0, None),
    ("--psm 4", 1.0, None),
    ("--psm 6", 1.4, _sharpen),
    ("--psm 6", 1.0, _otsu),
    # Texto disperso: lee el Total (en otra letra, solo en la columna de la
    # derecha) cuando las otras lecturas lo pierden (semana del 31/08/2026).
    ("--psm 11", 1.0, None),
)


def _ocr_rows(image, config, scale, prep):
    """
    Renglones (altura, texto) de una lectura de la página, reconstruyendo el
    orden por la posición real de cada palabra: con las columnas tan
    separadas, el análisis de Tesseract a veces agrupa el texto por columna
    entera en vez de por fila (visto en la semana del 10/08).
    """
    if prep is not None:
        image = prep(image)
    if scale != 1.0:
        image = image.resize((int(image.width * scale), int(image.height * scale)), Image.LANCZOS)
    data = pytesseract.image_to_data(image, config=config, output_type=pytesseract.Output.DICT)
    words = []
    for i, raw in enumerate(data["text"]):
        text = (raw or "").strip()
        if text:
            height = data["height"][i] / scale
            words.append({"text": text, "left": data["left"][i] / scale,
                          "cy": data["top"][i] / scale + height / 2, "height": height})
    words.sort(key=lambda w: w["cy"])
    rows = []
    for word in words:
        if rows and abs(word["cy"] - rows[-1]["cy"]) <= max(8, word["height"] * 0.6):
            row = rows[-1]
            row["words"].append(word)
            row["cy"] = sum(w["cy"] for w in row["words"]) / len(row["words"])
        else:
            rows.append({"cy": word["cy"], "words": [word]})
    result = []
    for row in rows:
        row["words"].sort(key=lambda w: w["left"])
        heights = sorted(w["height"] for w in row["words"])
        result.append({"y": row["cy"], "height": heights[len(heights) // 2],
                       "text": " ".join(w["text"] for w in row["words"])})
    return result


def _stamp(groups):
    """(mes, día, año, hora, minuto, A/P) -> datetime, o None si no es una fecha y hora válida."""
    month, day, year, hour, minute, half = groups
    year = int(year) if len(year) == 4 else 2000 + int(year)
    hour, minute = int(hour), int(minute)
    if not 1 <= hour <= 12:
        return None
    hour = hour % 12 + (12 if half.upper() == "P" else 0)
    try:
        return datetime(year, int(month), int(day), hour, minute)
    except ValueError:
        return None


def _parse_row(page, row):
    """Un renglón leído -> turno, Total de empleado o None."""
    text = row["text"].strip(" |")
    match = _TOTAL_RE.match(text)
    if match:
        return {"page": page, "y": row["y"], "kind": "total",
                "minutes": int(match.group(1)) * 60 + int(match.group(2))}
    match = _DETAIL_RE.match(text)
    if match is None:
        return None
    groups = match.groups()
    clock_in = _stamp(groups[1:7])
    clock_out = None if match.group("open") else _stamp(groups[7:13])
    words = [w for w in re.split(r"\s+", match.group("pre").upper()) if re.fullmatch(r"[A-Z][A-Z.'\-]*", w)]
    ids = re.findall(r"(?<![\d/])\d{1,3}(?![\d/:])", match.group("pre"))
    shift = _SHIFT_RE.match(match.group("rest") or "")
    return {
        "page": page, "y": row["y"], "kind": "shift",
        "name": " ".join(words) or None, "id": ids[0] if ids else None,
        "in": clock_in, "out": clock_out, "open": bool(match.group("open")),
        "minutes": int(shift.group(1)) * 60 + int(shift.group(2)) if shift else None,
    }


def _votes(items, field):
    counts = {}
    for item in items:
        value = item.get(field)
        if value is not None:
            counts[value] = counts.get(value, 0) + 1
    return counts


def _winner(counts):
    """El valor más votado; None si no hay o si dos empatan arriba."""
    ranked = sorted(counts.items(), key=lambda kv: -kv[1])
    if not ranked or (len(ranked) > 1 and ranked[0][1] == ranked[1][1]):
        return None
    return ranked[0][0]


def _clusters(items, tolerance):
    """Junta las lecturas del mismo renglón (misma página, misma altura) de todas las pasadas."""
    clusters = []
    for item in sorted(items, key=lambda i: (i["page"], i["y"])):
        if clusters and clusters[-1][0]["page"] == item["page"] and abs(clusters[-1][-1]["y"] - item["y"]) <= tolerance:
            clusters[-1].append(item)
        else:
            clusters.append([item])
    return clusters


def _label(minutes):
    return f"{minutes // 60}:{minutes % 60:02d}"


def _snap_name(name, known):
    """El nombre leído, o el ya conocido (de semanas anteriores) que se le parece casi letra por letra."""
    if not name or not known:
        return name
    best = max(known, key=lambda k: difflib.SequenceMatcher(None, name, k).ratio())
    return best if difflib.SequenceMatcher(None, name, best).ratio() >= 0.85 else name


def _other_employee(previous, current):
    """True si dos renglones seguidos son de empleados distintos (cambian el nombre y el ID leídos)."""
    names = [_winner(_votes(c, "name")) for c in (previous, current)]
    ids = [_winner(_votes(c, "id")) for c in (previous, current)]
    return None not in names and None not in ids and names[0] != names[1] and ids[0] != ids[1]


def _employees(parsed, tolerance, known):
    """
    Arma los empleados con todas las lecturas: los turnos entre un Total y el
    siguiente son del mismo empleado. Devuelve (empleados, sin resolver, todo cerró).
    - Cada turno vale lo que dicen la columna Hours Worked y la cuenta
      ClockOut - ClockIn; si dicen lo mismo, está confirmado.
    - El Total del empleado tiene que ser la suma de sus turnos. Si un turno
      tiene dos valores posibles (la columna y la cuenta no coinciden), vale
      el único que hace cerrar la suma contra el Total leído.
    - Nombre: el más votado entre todos sus renglones y lecturas.
    """
    employees, unresolved, settled = [], [], True
    block = []
    for cluster in _clusters(parsed, tolerance):
        kinds = _votes(cluster, "kind")
        if kinds.get("total", 0) > kinds.get("shift", 0):
            total_votes = _votes(cluster, "minutes")
            if block:
                result = _employee(block, total_votes, known)
                if result is None:
                    name = _winner(_votes([c for cluster_ in block for c in cluster_], "name"))
                    unresolved.append(f"{_snap_name(name, known) or 'Empleado sin nombre legible'} "
                                      "(las horas de los turnos no cierran contra el Total)")
                    settled = False
                elif "unresolved" in result:
                    unresolved.append(result["unresolved"])
                    settled = settled and result.get("final", False)
                else:
                    employees.append(result)
                    settled = settled and result["confirmed"]
            block = []
        else:
            shifts = [c for c in cluster if c["kind"] == "shift"]
            if block and _other_employee(block[-1], shifts):
                # Otro empleado sin que se haya leído el Total del anterior: se
                # avisa cada uno por su nombre (no los dos juntos como uno solo).
                name = _winner(_votes([c for cluster_ in block for c in cluster_], "name"))
                unresolved.append(f"{_snap_name(name, known) or 'Empleado sin nombre legible'} (no se encontró su Total)")
                settled = False
                block = []
            block.append(shifts)
    if block:
        name = _winner(_votes([c for cluster in block for c in cluster], "name"))
        unresolved.append(f"{_snap_name(name, known) or 'Empleado sin nombre legible'} (no se encontró su Total)")
        settled = False
    return employees, unresolved, settled


def _employee(block, total_votes, known):
    name = _winner(_votes([c for cluster in block for c in cluster], "name"))
    name = _snap_name(name, known)
    if any(_winner(_votes(cluster, "open")) for cluster in block):
        # Turno abierto (no marcó la salida): el Total es el tiempo hasta que se
        # imprimió el reporte, no horas trabajadas. No se paga solo (auditoría
        # 2026-09: Rick Leal 20:48 se pagaba $312).
        return {"unresolved": f"{name or 'Empleado sin nombre legible'} (turno abierto, sin salida marcada)",
                "final": True}
    if name is None:
        return None
    options, confirmed = [], True
    for cluster in block:
        clock_in = _winner(_votes(cluster, "in"))
        clock_out = _winner(_votes(cluster, "out"))
        computed = None
        if clock_in and clock_out and clock_out > clock_in:
            computed = int((clock_out - clock_in).total_seconds() // 60)
        read = _winner(_votes(cluster, "minutes"))
        values = {v for v in (computed, read) if v is not None}
        confirmed = confirmed and computed is not None and computed == read
        options.append(sorted(values))
    total = _winner(total_votes)
    sums = {}
    if all(options):
        for combo in itertools.islice(itertools.product(*options), 4096):
            sums.setdefault(sum(combo), []).append(combo)
    if total is not None and len(sums.get(total, [])) == 1:
        minutes = total
    elif total is None and confirmed and len(sums) == 1:
        minutes = next(iter(sums))  # el Total ilegible, pero cada turno confirmado por las dos cuentas
    else:
        return None
    confirmed = confirmed and total is not None
    return {"employee_name": name, "hours_label": _label(minutes), "hours": minutes / 60.0, "confirmed": confirmed}


def extract_hours_report(pdf_path, known_names=()):
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
        "unresolved_employees": [str, ...],  # detectados pero sin horas seguras
    }
    known_names: los empleados de semanas anteriores (un nombre leído con
    una letra cambiada se toma como el conocido).
    """
    _ensure_pdfplumber()
    _ensure_pytesseract()
    with pdfplumber.open(pdf_path) as pdf:
        pages = []
        for page in pdf.pages:
            text = (page.extract_text() or "").strip()
            image = None if text else _extract_page_image(page)
            if text or image is not None:
                pages.append((text, image))
    known = [re.sub(r"\s+", " ", n.strip().upper()) for n in known_names if n and n.strip()]

    parsed, headers, heights = [], [], []
    result = None
    for config, scale, prep in _PASSES:
        for page_no, (text, image) in enumerate(pages):
            if text:
                rows = [{"y": float(i) * 10, "height": 4.0, "text": line} for i, line in enumerate(text.splitlines())]
            else:
                rows = _ocr_rows(image, config, scale, prep)
            headers.append("\n".join(r["text"] for r in rows[:25]))
            heights.extend(r["height"] for r in rows)
            for row in rows:
                item = _parse_row(page_no, row)
                if item is not None:
                    parsed.append(item)
        heights.sort()
        tolerance = max(4.0, heights[len(heights) // 2] * 0.6) if heights else 8.0
        employees, unresolved, settled = _employees(parsed, tolerance, known)
        result = (employees, unresolved)
        if settled or all(text for text, _ in pages):
            break

    report_dates, periods = {}, {}
    for header in headers:
        match = _REPORT_PRINTED_RE.search(header)
        found = _parse_short_date(match.group(1)) if match else None
        if found:
            report_dates[found] = report_dates.get(found, 0) + 1
        match = _PERIOD_RE.search(header.replace("\n", " "))
        if match:
            period = (_parse_long_date(match.group(1)), _parse_long_date(match.group(2)))
            if all(period):
                periods[period] = periods.get(period, 0) + 1
    period = _winner(periods) or (None, None)
    employees, unresolved = result
    return {
        "report_date": _winner(report_dates),
        "period_from": period[0],
        "period_to": period[1],
        "employees": [{k: e[k] for k in ("employee_name", "hours_label", "hours")} for e in employees],
        "unresolved_employees": unresolved,
    }
