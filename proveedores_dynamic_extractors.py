"""
Motor de extracción genérico para proveedores agregados desde la web (sin
tocar código), ver /proveedores/nuevo en webapp.py.

Los 24 proveedores de proveedores.py::SUPPLIER_REGISTRY tienen un extractor
Python a medida, escrito por un programador. Este módulo persiste reglas
más simples (un "ancla" de texto por campo + qué aparición usar) en
proveedores_dynamic_suppliers.json y las aplica de forma genérica -- pensado
para facturas de texto limpio (digital o escaneo prolijo), no manuscritas
ni con layouts en columnas que necesiten anclaje posicional.

Algoritmo central: quien arma la regla tipea el valor "tal cual lo ve" en
la factura de muestra (ej. "188412"); en vez de buscar ese texto literal
(que puede estar escrito distinto en el documento -- "1,234.56" vs
"1234.56"), se escanea TODO el texto buscando substrings con la FORMA del
tipo de dato, se parsea cada candidato encontrado, y se compara el VALOR ya
parseado contra lo que tipeó el usuario -- mucho más robusto que comparar
texto crudo.
"""

import json
import os
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
    ensure_pdfplumber,
    ensure_pytesseract,
    extract_largest_page_image,
)

DYNAMIC_SUPPLIERS_FILENAME = "proveedores_dynamic_suppliers.json"

FIELDS = ("invoice_no", "date", "amount")
FIELD_LABELS = {
    "invoice_no": "N° de Factura",
    "date": "Fecha",
    "amount": "Monto Total",
}

# Excepciones que la lectura de un PDF de muestra puede tirar por un
# archivo genuinamente corrupto/ilegible -- mismo criterio de aislamiento
# que _PDF_EXTRACTION_EXCEPTIONS en proveedores.py.
PDF_READ_EXCEPTIONS = (ValueError, TypeError, AttributeError, RuntimeError, OSError)


# ---- Persistencia (mismo patrón atómico que chase_rules.py) ----

def _dynamic_suppliers_file_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), DYNAMIC_SUPPLIERS_FILENAME)


def _load_dynamic_suppliers_file():
    path = _dynamic_suppliers_file_path()
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(payload, dict):
        return {}
    suppliers = payload.get("suppliers", {})
    return suppliers if isinstance(suppliers, dict) else {}


def _save_dynamic_suppliers_file(suppliers):
    path = _dynamic_suppliers_file_path()
    directory = os.path.dirname(path)
    os.makedirs(directory, exist_ok=True)
    payload = {"suppliers": suppliers}
    temp_path = f"{path}.tmp"
    with open(temp_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, ensure_ascii=False)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp_path, path)


def load_dynamic_suppliers():
    return _load_dynamic_suppliers_file()


_SLUG_INVALID_RE = re.compile(r"[^a-z0-9]+")


def _slugify(label):
    base = _SLUG_INVALID_RE.sub("_", (label or "").strip().lower()).strip("_")
    return base or "proveedor"


def _unique_slug(label, existing_keys):
    base = _slugify(label)
    if base not in existing_keys:
        return base
    n = 2
    while f"{base}_{n}" in existing_keys:
        n += 1
    return f"{base}_{n}"


def add_dynamic_supplier(entry, created_by=None):
    """
    entry: {"label", "sheet_name", "resumen_label", "detect_keyword",
    "fields": {campo: {"anchor", "occurrence_index", ["date_format"]}}}.
    La clave (id interno) se genera sola a partir del label, evitando
    colisión con proveedores dinámicos ya guardados.
    """
    for field in FIELDS:
        if field not in entry.get("fields", {}):
            raise ValueError(f"Falta resolver el campo {FIELD_LABELS[field]} antes de guardar.")
    label = (entry.get("label") or "").strip()
    sheet_name = (entry.get("sheet_name") or "").strip()
    if not label or not sheet_name:
        raise ValueError("Nombre del proveedor y nombre de hoja son obligatorios.")
    detect_keyword = (entry.get("detect_keyword") or "").strip().lower()
    if not detect_keyword:
        raise ValueError("Falta la palabra de detección del proveedor.")

    suppliers = _load_dynamic_suppliers_file()
    clave = _unique_slug(label, suppliers.keys())
    suppliers[clave] = {
        "label": label,
        "sheet_name": sheet_name,
        "resumen_label": (entry.get("resumen_label") or sheet_name).strip(),
        "detect_keyword": detect_keyword,
        "fields": entry["fields"],
        "created_by": created_by,
        "created_at": datetime.utcnow().isoformat() + "Z",
    }
    _save_dynamic_suppliers_file(suppliers)
    return clave


def delete_dynamic_supplier(clave):
    suppliers = _load_dynamic_suppliers_file()
    if clave not in suppliers:
        raise ValueError("Ese proveedor ya no existe (probablemente ya se borró desde otra pestaña).")
    del suppliers[clave]
    _save_dynamic_suppliers_file(suppliers)


def list_dynamic_suppliers_display():
    suppliers = _load_dynamic_suppliers_file()
    return [
        {"clave": clave, **entry}
        for clave, entry in sorted(suppliers.items(), key=lambda kv: kv[1].get("label", "").lower())
    ]


# ---- Parseo genérico (monto/fecha) ----

_AMOUNT_CLEAN_RE = re.compile(r"[^0-9,.\-]")


def parse_amount_text(text):
    """
    Convierte un monto (tipeado por el usuario o leído del documento, con o
    sin "$", con separador de miles o decimal) a un float. Misma heurística
    ya usada en cmv_costo.py para la ambigüedad de la coma: 3 dígitos
    después de la ÚLTIMA coma es agrupamiento de miles, 1-2 sigue siendo
    separador decimal.
    """
    cleaned = _AMOUNT_CLEAN_RE.sub("", (text or "").strip())
    if not cleaned or cleaned in ("-", "."):
        raise ValueError(f'No se pudo interpretar "{text}" como un monto.')

    negative = cleaned.startswith("-")
    cleaned = cleaned.lstrip("-")

    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        digits_after = len(cleaned.rsplit(",", 1)[1])
        cleaned = cleaned.replace(",", "") if digits_after == 3 else cleaned.replace(",", ".")

    try:
        value = float(cleaned)
    except ValueError:
        raise ValueError(f'No se pudo interpretar "{text}" como un monto.')
    return -value if negative else value


_DATE_FORMATS = (
    "%m/%d/%Y", "%m/%d/%y",
    "%m-%d-%Y", "%m-%d-%y",
    "%m.%d.%Y", "%m.%d.%y",
    "%d/%m/%Y", "%d/%m/%y",
    "%d-%m-%Y", "%d-%m-%y",
    "%d.%m.%Y", "%d.%m.%y",
    "%Y-%m-%d",
)


def parse_date_text(text):
    """Prueba formatos MM/DD primero (todas las facturas de este proyecto son de EE.UU.)."""
    cleaned = (text or "").strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(cleaned, fmt), fmt
        except ValueError:
            continue
    raise ValueError(f'No se pudo interpretar "{text}" como una fecha.')


def _parse_target_value(field, raw_input):
    raw_input = (raw_input or "").strip()
    if not raw_input:
        raise ValueError("No se escribió ningún valor.")
    if field == "amount":
        return parse_amount_text(raw_input)
    if field == "date":
        parsed, _fmt = parse_date_text(raw_input)
        return parsed.date()
    if field == "invoice_no":
        digits = re.sub(r"\D", "", raw_input)
        if not digits:
            raise ValueError(f'"{raw_input}" no parece un número de factura.')
        return int(digits)
    raise ValueError(f"Campo desconocido: {field}")


_VALUE_SHAPE_PATTERNS = {
    "invoice_no": re.compile(r"\d{3,12}"),
    "amount": re.compile(r"-?\$?\s*\d[\d,]*\.\d{2}"),
    "date": re.compile(r"\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}"),
}


def _parse_shape_match(raw, field):
    if field == "amount":
        return parse_amount_text(raw)
    if field == "date":
        parsed, _fmt = parse_date_text(raw)
        return parsed.date()
    if field == "invoice_no":
        return int(raw)
    raise ValueError(field)


# ---- Extracción de texto (mismo patrón digital-primero-OCR-de-respaldo
# que _detect_supplier en proveedores.py, generalizado a todas las páginas) ----

def _extract_full_text(pdf_path):
    """
    Texto de todas las páginas concatenado. Resuelve gratis el caso de
    facturas multi-página donde el total está en la última (ej. H.T.
    Hackney): "occurrence_index: -1" sobre el texto completo ya encuentra
    la última aparición, sin lógica especial de "última página".
    """
    ensure_pdfplumber()
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not text.strip():
                image = extract_largest_page_image(page)
                if image is not None:
                    ensure_pytesseract()
                    text = pytesseract.image_to_string(image)
            pages_text.append(text)
    return "\n".join(pages_text)


def _derive_anchor(full_text, start_index, window=40):
    prefix = full_text[max(0, start_index - window):start_index]
    if "\n" in prefix:
        prefix = prefix.rsplit("\n", 1)[1]
    # Solo se recorta ruido de PRINCIPIO (espacios, separadores de columna
    # sueltos) y espacio en blanco al final -- nunca puntuación final como
    # ":" o "|", que suele ser justo lo que separa la etiqueta del valor y
    # forma parte del ancla útil (ej. "Invoice #:", "Total:").
    prefix = prefix.lstrip(" \t\r-|,")
    return prefix.rstrip(" \t\r")


def find_value_occurrences(full_text, target_value, field):
    """
    Todas las apariciones de `target_value` (ya parseado) en full_text,
    comparando VALORES -- no texto literal -- para tolerar formato distinto
    entre lo tipeado y lo impreso en el documento.
    """
    pattern = _VALUE_SHAPE_PATTERNS[field]
    occurrences = []
    for match in pattern.finditer(full_text):
        try:
            candidate_value = _parse_shape_match(match.group(0), field)
        except ValueError:
            continue
        if candidate_value != target_value:
            continue
        anchor = _derive_anchor(full_text, match.start())
        if len(anchor) < 2:
            continue
        context_start = max(0, match.start() - 60)
        context_end = min(len(full_text), match.end() + 20)
        occurrences.append({
            "anchor": anchor,
            "position": match.start(),
            "context": full_text[context_start:context_end].replace("\n", " ").strip(),
        })
    return occurrences


def _dedupe_occurrences_by_anchor(occurrences):
    """
    Una factura multi-página suele reimprimir el mismo encabezado (mismo
    ancla, mismo valor) en cada página -- sin esto, el usuario vería 5
    candidatos idénticos para desambiguar algo que en realidad no es
    ambiguo. Se muestra una sola vez por texto de ancla distinto, quedándose
    con la ÚLTIMA aparición de cada grupo como representante -- da lo mismo
    cuál se use como representante porque, por construcción, todas las
    apariciones de un mismo grupo ya comparten el mismo valor (se llega acá
    después de filtrar por igualdad de valor en find_value_occurrences).
    """
    representative_by_anchor = {}
    order = []
    for occ in occurrences:
        key = occ["anchor"].lower()
        if key not in representative_by_anchor:
            order.append(key)
        representative_by_anchor[key] = occ
    return [representative_by_anchor[key] for key in order]


def _anchor_occurrence_index(full_text, anchor, target_position):
    """
    En qué aparición del propio ANCLA (no del valor) cae la ocurrencia
    elegida -- lo que se guarda en la regla es "la N-ésima vez que aparece
    este ancla", que generaliza a facturas futuras mucho mejor que "la
    N-ésima vez que aparece este valor" (coincidencia de esta muestra
    puntual). 0 = primera aparición del ancla, -1 = última.
    """
    anchor_matches = [m.start() for m in re.finditer(re.escape(anchor), full_text, re.IGNORECASE)]
    if len(anchor_matches) <= 1:
        # Única aparición del ancla -- "primera" y "última" son lo mismo,
        # preferir 0 porque se lee más natural en la UI ("primera aparición").
        return 0
    best_index = 0
    for i, pos in enumerate(anchor_matches):
        if pos <= target_position:
            best_index = i
    return -1 if best_index == len(anchor_matches) - 1 else best_index


def analyze_sample(pdf_path, sample_values, chosen_occurrence_index=None):
    """
    sample_values: {"invoice_no": "188412", "date": "08/15/2026", "amount": "1,234.56"}
    (texto tal cual lo tipeó el usuario).
    chosen_occurrence_index: {"amount": 1, ...} -- índice, dentro de la
    lista de apariciones candidatas de ese VALOR, que el usuario ya eligió
    en una vuelta anterior de desambiguación (ver /proveedores/nuevo/analizar).

    Devuelve {"preview_text": ..., "fields": {campo: {...}}} -- por campo,
    "status" es "resolved" (con anchor/occurrence_index listos para
    guardar), "ambiguous" (con "candidates" para mostrarle al usuario), o
    "not_found"/"error" (con "message").
    """
    chosen_occurrence_index = chosen_occurrence_index or {}
    full_text = _extract_full_text(pdf_path)

    fields = {}
    for field in FIELDS:
        raw_input = sample_values.get(field)
        try:
            target_value = _parse_target_value(field, raw_input)
        except ValueError as exc:
            fields[field] = {"status": "error", "message": str(exc)}
            continue

        occurrences = _dedupe_occurrences_by_anchor(find_value_occurrences(full_text, target_value, field))
        if not occurrences:
            fields[field] = {
                "status": "not_found",
                "message": "No encontramos ese valor en el texto leído de la factura.",
            }
            continue

        if len(occurrences) == 1:
            chosen = occurrences[0]
        else:
            idx = chosen_occurrence_index.get(field)
            if idx is None or not (0 <= idx < len(occurrences)):
                fields[field] = {
                    "status": "ambiguous",
                    "candidates": [
                        {"index": i, "context": occ["context"]} for i, occ in enumerate(occurrences)
                    ],
                }
                continue
            chosen = occurrences[idx]

        occurrence_index = _anchor_occurrence_index(full_text, chosen["anchor"], chosen["position"])
        result = {
            "status": "resolved",
            "anchor": chosen["anchor"],
            "occurrence_index": occurrence_index,
            "context": chosen["context"],
        }
        if field == "date":
            _, fmt = parse_date_text(raw_input)
            result["date_format"] = fmt
        fields[field] = result

    return {"preview_text": full_text[:1500], "fields": fields}


def build_rule_fields(resolved_fields):
    """
    resolved_fields: {"invoice_no": {"anchor":..., "occurrence_index":...}, ...}
    (la parte de analyze_sample ya resuelta, sin "status"/"context"/"candidates").
    Devuelve la forma final para guardar en la regla, validando que los 3
    campos estén completos.
    """
    rule_fields = {}
    for field in FIELDS:
        info = resolved_fields.get(field) or {}
        if not info.get("anchor"):
            raise ValueError(f"Falta resolver el campo {FIELD_LABELS[field]} antes de guardar.")
        entry = {"anchor": info["anchor"], "occurrence_index": info.get("occurrence_index", 0)}
        if field == "date" and info.get("date_format"):
            entry["date_format"] = info["date_format"]
        rule_fields[field] = entry
    return rule_fields


# ---- Detección / extracción para una factura real (equivalente al
# detect()/extract() de una entrada de SUPPLIER_REGISTRY) ----

def detect_with_dynamic_rule(text, rule):
    keyword = (rule.get("detect_keyword") or "").strip().lower()
    if not keyword:
        return False
    return keyword in text.lower()


def _parse_date_with_preference(text, preferred_fmt):
    if preferred_fmt:
        try:
            return datetime.strptime(text.strip(), preferred_fmt)
        except ValueError:
            pass
    parsed, _fmt = parse_date_text(text)
    return parsed


def extract_with_dynamic_rule(pdf_path, rule):
    full_text = _extract_full_text(pdf_path)
    filename = os.path.basename(pdf_path)
    values = {}
    for field in FIELDS:
        field_rule = (rule.get("fields") or {}).get(field)
        if not field_rule:
            raise ValueError(f"La regla de este proveedor no tiene configurado el campo {field}.")
        anchor = field_rule["anchor"]
        occurrence_index = field_rule.get("occurrence_index", 0)

        anchor_matches = list(re.finditer(re.escape(anchor), full_text, re.IGNORECASE))
        if not anchor_matches:
            raise ValueError(f'{filename}: no se encontró el texto "{anchor}" para leer {field}.')
        try:
            chosen = anchor_matches[occurrence_index]
        except IndexError:
            raise ValueError(f'{filename}: "{anchor}" aparece menos veces de lo esperado.')

        window = full_text[chosen.end():chosen.end() + 80]
        value_match = _VALUE_SHAPE_PATTERNS[field].search(window)
        if value_match is None:
            raise ValueError(f'{filename}: no se pudo leer {field} cerca de "{anchor}".')
        raw_value = value_match.group(0)
        try:
            if field == "amount":
                values[field] = parse_amount_text(raw_value)
            elif field == "date":
                values[field] = _parse_date_with_preference(raw_value, field_rule.get("date_format"))
            else:
                values[field] = int(re.sub(r"\D", "", raw_value))
        except ValueError as exc:
            raise ValueError(f"{filename}: {exc}")

    return {"invoice_no": values["invoice_no"], "date": values["date"], "amount": values["amount"]}
