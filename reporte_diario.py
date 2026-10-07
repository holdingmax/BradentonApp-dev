"""
Reporte Diario — daily closure PDF extraction into Bradenton C-Store master.

Reads department totals from a user-selected PDF page of a daily closure
report, maps headers dynamically on sheet \"CARGA AQUI\" (row 3, column C onward),
and injects count/amount pairs on the first eligible operational row.
"""

import copy
import functools
import io
import itertools
import os
import re
import statistics
import tempfile
import unicodedata
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from difflib import SequenceMatcher, get_close_matches
from datetime import date, datetime, time, timedelta

try:
    import pdfplumber
except ImportError:
    pdfplumber = None  # type: ignore[assignment]

try:
    import pytesseract
    from PIL import Image
except ImportError:
    pytesseract = None  # type: ignore[assignment]
    Image = None  # type: ignore[assignment]

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    OPENPYXL_AVAILABLE = True
    COUNT_CELL_ALIGNMENT = Alignment(horizontal="center")
except ImportError:
    load_workbook = None  # type: ignore[assignment,misc]
    Alignment = None  # type: ignore[assignment,misc]
    Font = None  # type: ignore[assignment,misc]
    get_column_letter = None  # type: ignore[assignment,misc]
    COUNT_CELL_ALIGNMENT = None  # type: ignore[assignment,misc]
    OPENPYXL_AVAILABLE = False

SHEET_NAME = "CARGA AQUI"
HEADER_ROW = 3
HEADER_START_COLUMN = 3  # Column C
DATA_START_ROW = 5
DATE_SCAN_COLUMN = 1  # Column A only — calendar day matching
MANUAL_REPORT_COUNT_COLUMN = 71  # Column BS — total de unidades del reporte, cargado a mano

DEFAULT_PDF_PAGE_INDEX = 3  # Fourth page (0-based) for full un-cropped daily PDF

# Fixed 1-based PDF column layout (daily closure report)
PDF_DEPT_NAME_COL = 0  # 1st column — Dept.Name
PDF_NET_COUNT_COL = 4  # 5th column — Net Count
PDF_NET_SALES_COL = 7  # 8th column — Net Sales $
PDF_MIN_COLUMNS = 8

PDF_HEADER_DEPT_TOKENS = ("dept", "name")
PDF_HEADER_NET_COUNT_TOKENS = ("net", "count")
PDF_HEADER_NET_SALES_TOKENS = ("net", "sales")

DEPARTMENT_SALES_REPORT_ANCHOR = "Department Sales Report"
SAFE_DROP_REPORT_ANCHOR = "Safe Drop Report"

PROTECTED_DEPARTMENT_LABELS = frozenset(
    {
        "gift card",
        "varios/bolsa",
        "varios / bolsa",
        "varios",
        "bolsa",
    }
)

SUMMARY_STOP_MARKERS = (
    "total sales",
    "total cost",
    "gross profit",
    "grand total",
    "qty sold",
    "total qty",
    "department total",
    "net sales total",
)

SUB_HEADER_LABELS = frozenset(
    {
        "count",
        "qty",
        "units",
        "amount",
        "sales",
        "net",
        "net sales",
        "units sold",
        "qty sold",
        "cant",
        "cantidad",
        "importe",
    }
)

DEPARTMENT_ALIASES = {
    "hot dog": "hot dogs",
    "hotdog": "hot dogs",
    "hot dogs": "hot dogs",
    "beer/wine": "beer/wine",
    "beer wine": "beer/wine",
    "beer-wine": "beer/wine",
    "beer": "beer/wine",
    "e-cig": "e-gigarette",
    "e cig": "e-gigarette",
    "ecig": "e-gigarette",
    "e-cigarette": "e-gigarette",
    "e cigarette": "e-gigarette",
    "e-gigarette": "e-gigarette",
    "coffee": "coffe",
    "coffe": "coffe",
    "fountain": "foutain",
    "foutain": "foutain",
    "boiled peanuts": "boiled peanuts",
    "boiled-peanuts": "boiled peanuts",
    "automotive": "auto",
    "auto": "auto",
    "ice cream": "ice cream",
    "ice-cream": "ice cream",
    "prop hd": "propane",
    "prop-hd": "propane",
    "propane": "propane",
}

# Parsed PDF labels -> exact CARGA AQUI row 3 worksheet titles
DEPARTMENT_NAME_NORMALIZATION = {
    "e-gigarette": "E-GIGARETTE",
    "e-cig": "E-GIGARETTE",
    "e cig": "E-GIGARETTE",
    "ecig": "E-GIGARETTE",
    "e-cigarette": "E-GIGARETTE",
    "e cigarette": "E-GIGARETTE",
    "foutain": "FOUTAIN",
    "fountain": "FOUTAIN",
    "coffe": "COFFE",
    "coffee": "COFFE",
    "beer/wine": "BEER/WINE",
    "beer wine": "BEER/WINE",
    "beer-wine": "BEER/WINE",
    "beer": "BEER/WINE",
    "hot dog": "HOT DOGS",
    "hotdog": "HOT DOGS",
    "hot dogs": "HOT DOGS",
    "boiled peanuts": "BOILED PEANUTS",
    "boiled-peanuts": "BOILED PEANUTS",
    "automotive": "AUTO",
    "auto": "AUTO",
    "ice cream": "ICE CREAM",
    "ice-cream": "ICE CREAM",
    "propane": "PROPANE",
    "prop hd": "PROPANE",
    "prop-hd": "PROPANE",
    "local acct": "GETTEL/TOYOTA",
}

# Department row layout (right to left): ... | Net Count | ... | Net Sales $ | % of sales
NET_COUNT_REVERSE_INDEX = -5
NET_SALES_REVERSE_INDEX = -2
MIN_ROW_SPLIT_PARTS = 5

SALES_ALERT_FONT_COLOR = "FF0000"
GROUP_500_DEPARTMENTS = frozenset(
    {
        "AUTO",
        "BOILED PEANUTS",
        "HBA",
        "ICE CREAM",
        "MILK",
        "PROPANE",
        "GROCERIES",
        "NONTAX",
        "CANDY",
        "COFFE",
        "JUICE",
        "SNACK",
        "FOUTAIN",
        "WATER",
    }
)
GROUP_1400_DEPARTMENTS = frozenset(
    {
        "BEER/WINE",
        "CIGARS",
        "E-GIGARETTE",
        "GEN-CTN",
        "GEN-PAK",
        "MAJ CR",
        "MAJ PAK",
        "SNUFF",
        "SODA",
        "ONLINE",
        "SKOFF",
    }
)
GROUP_500_THRESHOLD = 500.00
GROUP_1400_THRESHOLD = 1400.00
GETTEL_TOYOTA_THRESHOLD = 17000.00

FUSED_DEPT_TRAILING_COUNT_RE = re.compile(
    r"^(?P<name>[A-Za-z][A-Za-z0-9\s/\-.'&]*?)\s*(?P<count>\d+)$"
)


from ocr_utils import (
    ensure_pytesseract as _ensure_pytesseract,
    extract_largest_page_image as _extract_page_image,
)


def _ensure_pdfplumber():
    if pdfplumber is None:
        raise ImportError(
            "Reporte Diario requiere pdfplumber. Instale con: pip install pdfplumber"
        )


def _ensure_openpyxl():
    if not OPENPYXL_AVAILABLE:
        raise ImportError(
            "Reporte Diario requiere openpyxl. Instale con: pip install openpyxl"
        )


class _LazyPdfPageImages:
    """
    Decode and orient a PDF page's embedded photo only the first time it's
    actually requested, caching the result for any later reuse.

    A daily closure PDF can run 50+ pages, but every report only ever needs
    2-3 of them (the anchor page found on the first or second try, plus one
    continuation page) — decoding and running Tesseract's orientation check
    on every other page was pure wasted work driving up processing time,
    especially across a multi-PDF batch. This makes cost proportional to
    pages actually read instead of pages in the file.
    """

    def __init__(self, pdf_path):
        _ensure_pdfplumber()
        _ensure_pytesseract()
        self._pdf = pdfplumber.open(pdf_path)
        self._cache = {}

    def __len__(self):
        return len(self._pdf.pages)

    def __getitem__(self, index):
        if index not in self._cache:
            self._cache[index] = self._load(index)
        return self._cache[index]

    def _load(self, index):
        return _extract_page_image(self._pdf.pages[index])

    def close(self):
        self._pdf.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()


def _strip_cell(value):
    if value is None:
        return ""
    return str(value).strip()


def _clean_dept_name(value):
    """Col 1 — strip trailing spaces, newlines, and collapse internal whitespace."""
    text = _strip_cell(value)
    text = text.replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _sanitize_parsed_dept_name(value):
    """
    Strip phantom leading/trailing punctuation from Dept.Name: dots, hyphens,
    spaces, and also commas/semicolons or any other OCR junk glued to the
    edges ('SODA ,' quedaba como un departamento aparte y caía en RESTO,
    auditoría 2026-09). Inside the name everything stays; '/' and '&' are
    kept at the end too.
    """
    text = _clean_dept_name(value)
    text = re.sub(r"^[^A-Za-z0-9]+|[^A-Za-z0-9/&]+$", "", text)
    return text


def _normalize_department_spacing(value):
    """
    Collapse mixed whitespace inside department labels to one space.

    Handles repeated spaces, tabs, and embedded newlines from un-cropped PDF text.
    """
    text = _strip_cell(value)
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _fallback_department_alias(dept_name):
    """
    Force-map partially corrupted labels to known worksheet departments.

    Used by the never-skip policy to avoid dropping rows due to OCR spacing shifts.
    """
    raw = _normalize_department_spacing(dept_name).upper()
    key = _normalize_department_label(raw)
    if not raw:
        return ""
    if "LOCAL" in raw:
        return "GETTEL/TOYOTA"
    if "BEER" in raw:
        return "BEER/WINE"
    if "GIGARETTE" in raw or "CIGARETTE" in raw:
        return "E-GIGARETTE"
    if "MAJ" in raw and "PAK" in raw:
        return "MAJ PAK"
    if "MAJ" in raw and "CR" in raw:
        return "MAJ CR"
    if "GEN" in raw and "PAK" in raw:
        return "GEN-PAK"
    if "GEN" in raw and "CTN" in raw:
        return "GEN-CTN"
    if "FLOWER" in raw:
        return "FLOWERS"
    if "FOUNTAIN" in raw:
        return "FOUTAIN"
    if key in DEPARTMENT_NAME_NORMALIZATION:
        return DEPARTMENT_NAME_NORMALIZATION[key]
    return raw


def normalize_parsed_department_name(value):
    """
    Map parsed PDF department text to exact CARGA AQUI row 3 titles.
    """
    text = _normalize_department_spacing(_sanitize_parsed_dept_name(value)).upper()
    if not text:
        return ""
    forced = _fallback_department_alias(text)
    if forced:
        text = forced
    key = _normalize_department_label(text)
    if key in DEPARTMENT_NAME_NORMALIZATION:
        return DEPARTMENT_NAME_NORMALIZATION[key]
    if key in DEPARTMENT_ALIASES:
        alias_target = _normalize_department_label(DEPARTMENT_ALIASES[key])
        if alias_target in DEPARTMENT_NAME_NORMALIZATION:
            return DEPARTMENT_NAME_NORMALIZATION[alias_target]
    return text


def _row_cells_to_line(cells):
    """Join table cells into one line for reverse-index whitespace parsing."""
    return " ".join(_strip_cell(cell) for cell in (cells or []) if _strip_cell(cell))


def _strip_discount_tokens(text):
    """
    Remove negative Discount $ tokens (e.g. -$5.26) from a raw line.

    Ensures discount amounts never interfere with Net Sales $ / Net Count parsing.
    """
    if not text:
        return text
    tokens = str(text).split()
    kept = []
    for token in tokens:
        plain = token.replace(",", "")
        if "-$" in plain or re.search(r"-\$\s*\d", plain):
            continue
        kept.append(token)
    return " ".join(kept)


def _is_pdf_header_parts(parts):
    joined = " ".join(parts).lower()
    return "dept" in joined and "name" in joined


def _is_alphabetic_dept_token(token):
    if not token:
        return False
    cleaned = token.replace(",", "")
    if re.fullmatch(r"[\d.$%-]+", cleaned):
        return False
    return bool(re.search(r"[A-Za-z]", token))


def _split_dept_name_from_fused_tail(text):
    """Isolate GEN-PAK from GEN-PAK31 when count was not a separate token."""
    dept_text = _sanitize_parsed_dept_name(text)
    if not dept_text:
        return "", None
    match = FUSED_DEPT_TRAILING_COUNT_RE.match(dept_text)
    if match:
        return (
            _sanitize_parsed_dept_name(match.group("name")),
            match.group("count"),
        )
    compact = dept_text.replace(" ", "")
    match = re.match(r"^(.+?)(\d+)$", compact)
    if match and re.search(r"[A-Za-z/]", match.group(1)):
        return (
            _sanitize_parsed_dept_name(match.group(1)),
            match.group(2),
        )
    return dept_text, None


def _resolve_reverse_indices(part_count):
    """Return (net_count_index, net_sales_index) for the row width."""
    if part_count >= 8:
        return NET_COUNT_REVERSE_INDEX, NET_SALES_REVERSE_INDEX
    if part_count >= 5:
        return -3, NET_SALES_REVERSE_INDEX
    return None, None


def _parse_row_by_reverse_index(line):
    """
    Stable reverse-index row parser after Department Sales Report.

    Whitespace split, then:
      [-2] = Net Sales $ (skip [-1] % of sales)
      [-5] = Net Count on full-width rows (or [-3] on compact rows)
      left tokens = Dept.Name (alphabetic parts only)
    """
    text = _normalize_department_spacing(_strip_cell(line))
    text = _strip_discount_tokens(text)
    # Keep department labels detached from trailing numeric tokens on dense PDF rows.
    text = re.sub(r"(?i)(BEER/WINE|E-?CIGARETTE)(?=\d)", r"\1 ", text)
    if not text:
        return None

    parts = text.split()
    filtered_parts = [
        token
        for token in parts
        if token and not token.startswith("-$") and "-$" not in token
    ]
    if len(filtered_parts) < MIN_ROW_SPLIT_PARTS or _is_pdf_header_parts(filtered_parts):
        return None

    tail_tokens = (
        filtered_parts[:-1]
        if filtered_parts and "%" in filtered_parts[-1]
        else list(filtered_parts)
    )

    amount_index = None
    amount = 0.0
    count = 0
    count_index = None
    # Stable reverse-index parse after dynamic discount-token removal.
    # Net Sales: [-2] (skip trailing percent), Net Count: [-5] on full rows.
    try:
        if len(tail_tokens) >= 5:
            amount_index = len(tail_tokens) - 2
            amount = float(_sanitize_sales_float(tail_tokens[amount_index]))
            if len(tail_tokens) >= 8:
                count_index = len(tail_tokens) - 5
            else:
                count_index = len(tail_tokens) - 3
            count = int(_safe_parse_count(tail_tokens[count_index]))
        else:
            count_idx, sales_index = _resolve_reverse_indices(len(tail_tokens))
            if sales_index is not None:
                amount_raw = tail_tokens[sales_index]
                amount = float(_sanitize_sales_float(amount_raw))
                amount_index = (
                    sales_index
                    if sales_index >= 0
                    else len(tail_tokens) + sales_index
                )
            if count_idx is not None and -len(tail_tokens) <= count_idx < len(tail_tokens):
                count = int(_safe_parse_count(tail_tokens[count_idx]))
                count_index = (
                    count_idx if count_idx >= 0 else len(tail_tokens) + count_idx
                )
    except Exception:
        amount = float(_safe_parse_amount(tail_tokens[-1] if tail_tokens else 0.0))
        count = int(_safe_parse_count(tail_tokens[-2] if len(tail_tokens) >= 2 else 0))
        amount_index = max(len(tail_tokens) - 1, 1)
        count_index = max(amount_index - 1, 0)

    if amount_index is None:
        # _resolve_reverse_indices() solo se llama con part_count < 5 (ver
        # el if/else de arriba), pero sus dos condiciones (>=8, >=5) exigen
        # part_count >= 5 -- siempre cae a su "return None, None", así que
        # amount_index/count_index quedan sin asignar acá. Sin este chequeo,
        # la línea de abajo (amount_index - 1 con amount_index=None) tira un
        # TypeError sin atrapar que, en el camino de texto plano
        # (_parse_tables_with_line_anchor, sin su propio try/except),
        # tumbaba el archivo entero en vez de descartar esta fila como no
        # parseable -- mismo trato que ya reciben el resto de filas que no
        # se pueden leer con confianza (bug real, auditoría 2026-09-06).
        return None
    if count_index is None:
        count_index = max(amount_index - 1, 0)

    left_end = count_index
    if left_end <= 0:
        return None

    split_tokens = [chunk.strip() for chunk in re.split(r"\s{2,}", text) if chunk.strip()]
    cleaned_from_split = (
        _normalize_department_spacing(split_tokens[0]).upper() if split_tokens else ""
    )
    dept_tokens = [
        token for token in tail_tokens[:left_end] if _is_alphabetic_dept_token(token)
    ]
    dept_raw = cleaned_from_split or _normalize_department_spacing(" ".join(dept_tokens))
    dept_raw = _normalize_department_spacing(dept_raw).upper()
    dept_text, fused_count = _split_dept_name_from_fused_tail(dept_raw)
    if fused_count is not None and count == 0:
        count = int(_safe_parse_count(fused_count))

    department = normalize_parsed_department_name(_fallback_department_alias(dept_text))
    if not department or _is_summary_row(department):
        return None
    # _is_protected_department (GIFT CARD/VARIOS/BOLSA) ya NO se filtra acá
    # -- bug real encontrado 2026-09-17: ese chequeo existía para que
    # inject_daily_sales no pise columnas del Excel real con fórmula propia
    # (ver esa función, más abajo, que sigue teniendo su propio chequeo
    # independiente), pero al vivir también en el parser de filas se perdía
    # la fila ENTERA antes de llegar a Carga de Datos -- "GIFT CARD" nunca
    # aparecía en Ventas por Departamento ni se sumaba a RESTO, pese a ser
    # una venta real. Pedido explícito del usuario: "CHEVRON GIFT CARD...
    # tambien es un departamento que se suma a VARIOS, asi como milk o
    # icecream" -- ahora fluye igual que cualquier otro departamento sin
    # categoría conocida (group_department_sales ya lo suma a RESTO solo).

    return {
        "department": department,
        "count": count,
        "amount": amount,
    }


def _safe_parse_count(value):
    try:
        return _parse_count(value)
    except (TypeError, ValueError, OverflowError):
        return 0


def _safe_parse_amount(value):
    try:
        return _parse_amount(value)
    except (TypeError, ValueError, OverflowError):
        return 0.0


def _normalize_department_label(value):
    text = _clean_dept_name(value)
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def _canonical_department_key(value):
    key = _normalize_department_label(value)
    if not key:
        return ""
    if key in DEPARTMENT_ALIASES:
        return _normalize_department_label(DEPARTMENT_ALIASES[key])
    for alias, canonical in DEPARTMENT_ALIASES.items():
        if key == alias or key.startswith(alias + " ") or key.endswith(" " + alias):
            return _normalize_department_label(canonical)
    return key


def _department_keys_match(pdf_key, header_key):
    if not pdf_key or not header_key:
        return False
    if pdf_key == header_key:
        return True
    pdf_canon = _canonical_department_key(pdf_key)
    header_canon = _canonical_department_key(header_key)
    if pdf_canon and pdf_canon == header_canon:
        return True
    if pdf_canon in header_canon or header_canon in pdf_canon:
        return True
    pdf_compact = pdf_canon.replace(" ", "").replace("-", "").replace("/", "")
    header_compact = header_canon.replace(" ", "").replace("-", "").replace("/", "")
    return pdf_compact == header_compact


def _sanitize_numeric_text(value):
    """Strip whitespace, currency symbols, and grouping separators."""
    text = _strip_cell(value)
    if not text:
        return ""
    text = text.replace("$", "").replace("€", "").replace("£", "")
    text = text.replace(" ", "").replace("\u00a0", "").replace("\n", "").replace("\r", "")
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    text = text.replace(",", "")
    return text


def _parse_count(value):
    """Col 5 — Net Count: strip symbols, default empty to 0, return int."""
    if value is None:
        return 0
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return 0
    text = _sanitize_numeric_text(value)
    if not text:
        return 0
    if re.fullmatch(r"-?\d+", text):
        return int(text)
    try:
        number = float(text)
        if number.is_integer():
            return int(number)
    except ValueError:
        return 0
    return 0


def _parse_amount(value):
    """Col 8 — Net Sales $: strip currency, handle negatives, return float."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = _sanitize_numeric_text(value)
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def _is_sub_header(label):
    return _normalize_department_label(label) in SUB_HEADER_LABELS


def _is_summary_row(label):
    key = _normalize_department_label(label)
    if not key:
        return True
    # "dept " o "dept." (el encabezado de la tabla, con o sin el punto).
    if re.match(r"dept\W", key):
        return True
    if key in PROTECTED_DEPARTMENT_LABELS:
        return False
    return any(marker in key for marker in SUMMARY_STOP_MARKERS)


def _is_protected_department(label):
    key = _normalize_department_label(label)
    return key in PROTECTED_DEPARTMENT_LABELS


def _cell_has_formula(cell):
    if getattr(cell, "data_type", None) == "f":
        return True
    value = cell.value
    return isinstance(value, str) and value.startswith("=")


def _create_temp_workbook_path():
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="reporte_diario_")
    os.close(fd)
    return temp_path



def _get_carga_aqui_sheet(workbook):
    target = SHEET_NAME.strip().lower()
    for name in workbook.sheetnames:
        if name.strip().lower() == target:
            return workbook[name]
    raise ValueError(
        f'Hoja "{SHEET_NAME}" no encontrada. Disponibles: {", ".join(workbook.sheetnames)}'
    )


def _pad_row(row, min_columns=PDF_MIN_COLUMNS):
    cells = [_strip_cell(cell) for cell in (row or [])]
    if len(cells) < min_columns:
        cells.extend([""] * (min_columns - len(cells)))
    return cells


def _header_cell_matches(label, required_tokens):
    key = _normalize_department_label(label).replace(".", " ")
    return all(token in key for token in required_tokens)


def _is_pdf_department_header_row(cells):
    padded = _pad_row(cells)
    return (
        _header_cell_matches(padded[PDF_DEPT_NAME_COL], PDF_HEADER_DEPT_TOKENS)
        and _header_cell_matches(padded[PDF_NET_COUNT_COL], PDF_HEADER_NET_COUNT_TOKENS)
        and _header_cell_matches(padded[PDF_NET_SALES_COL], PDF_HEADER_NET_SALES_TOKENS)
    )


def _extract_tables_from_page(page):
    """Extract tables from a PDF page (silent — no console output)."""
    tables = page.extract_tables() or []
    if tables:
        return tables

    line_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 4,
        "join_tolerance": 4,
    }
    tables = page.extract_tables(line_settings) or []
    if tables:
        return tables

    text_settings = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "snap_tolerance": 4,
    }
    return page.extract_tables(text_settings) or []


FILENAME_DAY_MONTH_PATTERN = re.compile(
    r"(?<!\d)(\d{1,2})[-/.](\d{1,2})(?!\d)"
)


def extract_day_from_filename(file_path):
    """
    Extract the calendar day-of-month from a daily PDF filename.

    Examples: \"Close Store 01-05.pdf\" -> 1, \"Close Store 07-05.pdf\" -> 7.
    """
    base = os.path.splitext(os.path.basename(file_path))[0]
    for match in FILENAME_DAY_MONTH_PATTERN.finditer(base):
        target_day = int(match.group(1))
        month = int(match.group(2))
        if 1 <= target_day <= 31 and 1 <= month <= 12:
            return target_day

    trailing = re.search(r"[^\d](\d{1,2})\s*$", base)
    if trailing:
        target_day = int(trailing.group(1))
        if 1 <= target_day <= 31:
            return target_day

    raise ValueError(
        f"No se pudo extraer el día del nombre de archivo: {os.path.basename(file_path)}. "
        "Se esperaba un patrón como 'Close Store 01-05.pdf'."
    )


def _cell_calendar_day(value):
    """
    Return the day-of-month (1-31) from a Column A cell value, or None.

    datetime -> .day; plain integers 1-31 -> day-of-month; larger numbers -> Excel serial.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, datetime):
        return int(value.day)
    if isinstance(value, int):
        day = int(value)
        if 1 <= day <= 31:
            return day
        try:
            from openpyxl.utils.datetime import from_excel

            converted = from_excel(day)
            if isinstance(converted, datetime):
                return int(converted.day)
        except (ValueError, TypeError, OverflowError):
            return None
        return None
    if isinstance(value, float):
        if value.is_integer():
            whole = int(value)
            if 1 <= whole <= 31:
                return whole
        try:
            from openpyxl.utils.datetime import from_excel

            converted = from_excel(value)
            if isinstance(converted, datetime):
                return int(converted.day)
        except (ValueError, TypeError, OverflowError):
            return None
        return None
    text = _strip_cell(value)
    if not text:
        return None
    if re.fullmatch(r"\d{1,2}", text):
        day = int(text)
        if 1 <= day <= 31:
            return day
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return int(datetime.strptime(text, fmt).day)
        except ValueError:
            continue
    if re.match(r"^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$", text):
        for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y"):
            try:
                return int(datetime.strptime(text, fmt).day)
            except ValueError:
                continue
    month_match = re.match(r"^(\d{1,2})[-/]([a-zA-Z]{3,9})(?:[-/]\d{2,4})?$", text)
    if month_match:
        return int(month_match.group(1))
    if re.match(r"^\d{1,2}-[a-zA-Z]{3}(-\d{2,4})?$", text):
        return int(text.split("-", 1)[0])
    return None


def find_row_for_calendar_day(sheet, target_day, start_row=DATA_START_ROW):
    """
    Locate the row whose Column A date matches target_day (strict integer compare).

    Scans column A only from start_row downward so each PDF locks one unique row.
    """
    target_day = int(target_day)
    max_row = max(sheet.max_row, start_row)
    for row in range(start_row, max_row + 1):
        cell_value = sheet.cell(row=row, column=DATE_SCAN_COLUMN).value
        cell_day = _cell_calendar_day(cell_value)
        if cell_day is not None and int(cell_day) == target_day:
            return row
    raise ValueError(
        f"No se encontró una fila con el día {target_day} en la columna A de la hoja {SHEET_NAME} "
        f"(desde la fila {start_row})."
    )


def _parse_pdf_data_row(cells):
    """
    Extract one department row using reverse-index whitespace splitting.

    Col layout: Dept.Name | ... | Net Count | ... | Net Sales $ | % of sales
    """
    try:
        line = _row_cells_to_line(cells)
        if not line:
            return None
        parsed = _parse_row_by_reverse_index(line)
        if parsed is None:
            return None
        return {
            "department": parsed["department"],
            "count": int(parsed["count"]),
            "amount": float(parsed["amount"]),
        }
    except (TypeError, ValueError, OverflowError, IndexError, AttributeError):
        return None


def _line_contains_anchor(line):
    return DEPARTMENT_SALES_REPORT_ANCHOR.lower() in _strip_cell(line).lower()


def _line_contains_stop_marker(line):
    """
    True at 'Safe Drop Report' (daily "Close Store" PDF) or 'Method of
    Payment Totals Report' (the standalone report of that name a monthly
    bundle PDF -- Store Sales Summary + Department Sales + Method of
    Payment Totals + Inventory, printed together for month-end -- prints
    right after Department Sales Report, with no Safe Drop Report at all)
    -- everything from there on is out of scope. The full "... totals
    report" phrase is required (not just "method of payment") so this never
    matches the unrelated "Method of Payment Totals Count $ Sales" line
    that Store Sales Summary Report itself already prints, on an earlier
    page, as part of its own Store Tender Reading section.
    """
    text = _strip_cell(line).lower()
    return "safe drop" in text or "method of payment totals report" in text


_OCR_HEADER_COMPANION_TOKENS = ("sales", "count", "refund", "item")


def _is_ocr_header_line(text):
    """
    True for the "Dept. Name ... Net Count ... Net Sales $ ..." header row,
    which OCR sometimes repeats mid-table as its own smashed-together line —
    including right before a same-day Refund sub-table, which reprints its
    own header. "dept" is checked as normal, but OCR also regularly drops
    its leading "D" ("Dept." -> "Ept."), so "name" alongside any other
    column-header word is treated as the same header line too — no real
    department is ever literally named "Name".
    """
    key = _normalize_department_label(text)
    if "dept" in key and "name" in key:
        return True
    return "name" in key and any(token in key for token in _OCR_HEADER_COMPANION_TOKENS)


def _anchor_passed_in_page_text(page):
    """True once 'Department Sales Report' appears in line-by-line page text."""
    for line in (page.extract_text() or "").splitlines():
        if _line_contains_anchor(line):
            return True
    return False


def _crop_page_below_anchor(page):
    """
    Return a pdfplumber page cropped to content below the anchor phrase.

    Falls back to the full page when search geometry is unavailable.
    """
    hits = page.search(DEPARTMENT_SALES_REPORT_ANCHOR, case=False) or []
    if hits:
        anchor_bottom = max(hit["bottom"] for hit in hits)
        return page.crop((0, anchor_bottom, page.width, page.height))
    return page


def _parse_department_table_rows(rows, require_header=True):
    if not rows:
        return []

    header_index = None
    if require_header:
        for idx, row in enumerate(rows):
            if _is_pdf_department_header_row(row):
                header_index = idx
                break
        start = (header_index + 1) if header_index is not None else 0
    else:
        start = 0

    records = []
    for row in rows[start:]:
        if not row:
            continue
        record = _parse_pdf_data_row(row)
        if record is None:
            if records and _clean_dept_name(_pad_row(row)[PDF_DEPT_NAME_COL]):
                break
            continue
        records.append(record)
    return records


def _parse_tables_with_line_anchor(page):
    """
    Parse department rows only after the Department Sales Report anchor.

    Scans raw text line-by-line to locate the anchor, then reads table rows
    from the cropped region below it (or skips pre-anchor table rows).
    """
    if not _anchor_passed_in_page_text(page):
        raise ValueError(
            f'No se encontró el ancla "{DEPARTMENT_SALES_REPORT_ANCHOR}" en la página del PDF.'
        )

    cropped = _crop_page_below_anchor(page)
    tables = _extract_tables_from_page(cropped)
    records = []
    for table in tables:
        parsed = _parse_department_table_rows(table)
        if parsed:
            records = parsed
            break

    if records:
        return records

    past_anchor = False
    for table in _extract_tables_from_page(page):
        for row in table or []:
            row_text = _row_cells_to_line(row)
            if not past_anchor:
                if _line_contains_anchor(row_text):
                    past_anchor = True
                continue
            if _is_pdf_department_header_row(row):
                continue
            record = _parse_pdf_data_row(row)
            if record is None:
                if records:
                    probe = _row_cells_to_line(row)
                    if probe and (
                        _is_summary_row(probe)
                        or "total sales" in probe.lower()
                    ):
                        break
                continue
            records.append(record)
        if records:
            break

    if not records:
        past_anchor = False
        for line in (page.extract_text() or "").splitlines():
            if not past_anchor:
                if _line_contains_anchor(line):
                    past_anchor = True
                continue
            if not _strip_cell(line):
                continue
            if _is_pdf_department_header_row([line]):
                continue
            parsed = _parse_row_by_reverse_index(line)
            if parsed is None:
                if records:
                    break
                continue
            records.append(parsed)

    return records


# No single Tesseract page-segmentation mode is reliably best across every
# photo, so each page is tried under each of these and scored (see
# _score_ocr_page_text) rather than trusting one fixed mode.
_OCR_TEXT_CONFIGS = ("--psm 6", "--psm 4", "--psm 3", "--psm 11")

# WATER is always the last department printed on a real report; when it was
# a $0.00 day it's omitted and TAXABLE (the second-to-last) becomes the last
# one. Used only as an informational cross-check, never to stop parsing.
LAST_DEPARTMENT_CANDIDATES = ("WATER", "TAXABLE")

# Fields after Dept.Name, right to left: % of Sales (dropped separately),
# Net Sales $, Discount $, Refund $, Net Count, Refund Count, Item Count,
# Gross Sales $ — 7 fields once the trailing % token is stripped.
_OCR_ROW_TAIL_FIELDS = 7
_OCR_NET_SALES_REVERSE_INDEX = -1
_OCR_NET_COUNT_REVERSE_INDEX = -4

# Tesseract occasionally hallucinates a lone "_" as its own whitespace-
# separated token between two real columns (seen in practice between two
# "$0.00" cells) -- never a real column value in this report, but it does
# add one extra token, which is enough to throw off every reverse-index
# field lookup below it (e.g. "ONLINE $59.00 10 0 10 $0.00 _ $0.00 $59.00
# 1.64%" reads Net Count off the wrong column and folds "$59" into the
# department name). Dropped before indexing so a lone noise token can't
# shift the count.
_OCR_NOISE_TOKEN_RE = re.compile(r"^_+$")
# Token con forma de monto o conteo ('$1,234.56', '-12.00', '(5.00)', '511').
_OCR_NUMERIC_TOKEN_RE = re.compile(r"^[(\-~]*\$?\d[\d,]*(\.\d+)?\)?$")


def _parse_ocr_department_row(line):
    """
    Parse one OCR'd "Department Sales Report" line.

    Tesseract's image_to_string collapses the original column gaps to single
    spaces, so — unlike the pdfplumber text parser, which relies on
    multi-space gaps to isolate the department name — this indexes strictly
    from the right against the report's fixed layout: Dept.Name | Gross
    Sales $ | Item Count | Refund Count | Net Count | Refund $ | Discount $ |
    Net Sales $ | % of Sales. Returns None for anything that isn't a
    parseable row; the printed grand-total line (no department name) comes
    back with department="" and is_total=True for the OCR subtotal cross-check.
    """
    text = _normalize_department_spacing(_strip_cell(line))
    if not text:
        return None
    if _is_ocr_header_line(text):
        return None

    parts = [token for token in text.split() if token and not _OCR_NOISE_TOKEN_RE.match(token)]
    if len(parts) < _OCR_ROW_TAIL_FIELDS + 1:
        return None

    # The trailing "% of Sales" column is always present in the source table
    # even when Tesseract fails to read its literal "%" character (e.g. a
    # garbled "OBI" instead of "0.02%") — so it is always dropped, rather
    # than only when the "%" glyph itself came through.
    tail_tokens = parts[:-1]

    net_sales_raw = tail_tokens[_OCR_NET_SALES_REVERSE_INDEX]
    net_count_raw = tail_tokens[_OCR_NET_COUNT_REVERSE_INDEX]
    dept_tokens = tail_tokens[:-_OCR_ROW_TAIL_FIELDS]

    # Una línea de ruido o un encabezado mal leído no trae ni Net Sales ni
    # Net Count con forma de número: antes quedaba como un departamento de
    # $0 ('PE EE LN', 'DEPT. N GROSS ...'). Auditoría 2026-09.
    if not (_OCR_NUMERIC_TOKEN_RE.match(net_sales_raw) or _OCR_NUMERIC_TOKEN_RE.match(net_count_raw)):
        return None

    amount = _sanitize_sales_float(net_sales_raw)
    count = _safe_parse_count(net_count_raw)

    # La fila del gran total no tiene nombre de departamento. Si el OCR le
    # mete un carácter suelto adelante ('|', '—') o lee una etiqueta 'Total',
    # sigue siendo el total: nunca se guarda como departamento.
    dept_has_letters = any(re.search(r"[A-Za-z]", token) for token in dept_tokens)
    dept_key = _normalize_department_label(" ".join(dept_tokens))
    if not dept_tokens or not dept_has_letters or dept_key in ("total", "grand total", "totals"):
        return {"department": "", "count": int(count), "amount": float(amount), "is_total": True}

    dept_raw = _normalize_department_spacing(" ".join(dept_tokens)).upper()
    dept_text, _fused_count = _split_dept_name_from_fused_tail(dept_raw)
    department = normalize_parsed_department_name(_fallback_department_alias(dept_text))
    # _is_protected_department ya no se filtra acá -- ver el mismo comentario
    # en _parse_row_by_reverse_index, más arriba en este archivo.
    if not department or _is_summary_row(department):
        return None

    return {"department": department, "count": int(count), "amount": float(amount), "is_total": False}


def _score_ocr_page_text(text):
    """
    Score one OCR attempt by how many department rows it yields.

    If the anchor line is present, only rows between it and the
    "Safe Drop Report" stop marker count (skips preamble tables read on the
    same page). If not, the whole page is treated as table continuation —
    real reports sometimes spill the department table onto a second page.
    """
    lines = text.splitlines() if text else []
    start = 0
    has_anchor = False
    for index, line in enumerate(lines):
        if _line_contains_anchor(line):
            start = index + 1
            has_anchor = True
            break

    count = 0
    for line in lines[start:]:
        if _line_contains_stop_marker(line):
            break
        parsed = _parse_ocr_department_row(line)
        if parsed is not None and not parsed["is_total"]:
            count += 1
    return (int(has_anchor), count)


# ---------------------------------------------------------------------------
# Department Sales Report por OCR: lectura por votación (2026-10-04)
# ---------------------------------------------------------------------------
# Pedido explícito del usuario (2026-10-04): "muchas veces me falló errando a
# la cantidad de productos que tenía algún que otro departamento". Antes se
# guardaba la fila de UNA sola pasada de OCR (la de más filas), leída por
# posición desde la derecha, y la suma contra el total impreso toleraba hasta
# 2 unidades de diferencia. Relevamiento contra el Excel de Ventas (hoja CARGA
# AQUI, cargada a mano) en 142 días reales de 2025-2026: los errores eran casi
# todos de unidades y de tres tipos:
#   1. Tesseract funde dos dígitos iguales ("77" -> "7", "11" -> "1") en las
#      TRES pasadas a la vez, así que votar no alcanza. Pero la caja de esa
#      palabra en la imagen mide lo que miden dos dígitos: se recorta esa
#      celda, se agranda 3x y se relee solo con dígitos.
#   2. Basura pegada a un número ("4.", "40°", "1—", "$0.00.") o suelta entre
#      columnas ("-", "~", "*") que corría todas las columnas de lugar.
#   3. Un dígito mal leído en una sola columna ("215 0 216"): la fila trae su
#      propia ecuación (Item Count - Refund Count = Net Count; Gross - Refund
#      - Discount = Net Sales) y el % de ventas confirma el monto.
# Ahora: se leen las filas de 3 pasadas (--psm 6 con cajas por palabra, 4 y
# 3; la 11 parte cada celda en su propia línea y no sirve para filas), cada
# valor se vota entre ellas dando más peso a las filas que cierran su
# ecuación, y el día se cierra EXACTO contra la fila del total impreso
# (unidades y monto): si no cierra, se prueba cambiar 1 o 2 valores dudosos
# por su otra lectura y se aplica solo si hay una única salida clara. Lo que
# sigue dudoso queda VACÍO y se avisa (regla de oro del OCR).

_OCR_ROW_CONFIGS = ("--psm 6", "--psm 4", "--psm 3")

# Confusiones típicas de Tesseract dentro de un número.
_OCR_DIGIT_FIX = str.maketrans({
    "O": "0", "o": "0", "D": "0", "Q": "0", ")": "0", "(": "0",
    "l": "1", "I": "1", "|": "1", "i": "1", "!": "1", "]": "1",
    "S": "5", "s": "5", "§": "5", "B": "8", "Z": "2", "z": "2", "g": "9",
})
_OCR_MONEY_TOKEN_RE = re.compile(r"^[-~—–=]?\$?-?\(?\$?\d{1,3}(,?\d{3})*\.\d{2}\)?$")
# El reporte siempre imprime centavos: "$3480" es $34.80 con el punto perdido.
_OCR_MONEY_NO_POINT_RE = re.compile(r"^(-?)\$(\d{1,3}(?:,\d{3})+|\d{1,5})[^\d,]?(\d{2})$")
_OCR_INT_TOKEN_RE = re.compile(r"^\d{1,5}$")
_OCR_EDGE_JUNK = "\"'`‘’“”~—–\\-_.,:;°*"
_OCR_EDGE_JUNK_RE = re.compile(f"^[{_OCR_EDGE_JUNK}]+|[{_OCR_EDGE_JUNK}]+$")

# Departamentos que imprime el POS, EN EL ORDEN en que los imprime (alfabético
# por su nombre en el POS, con HOT DOGS primero y LOCAL ACCT -- acá
# "GETTEL/TOYOTA" -- en la L), tal como quedan después de normalizar el
# nombre leído (ver _ocr_row_department). Solo para corregir un nombre
# garabateado ("J TAXABLE", "TAY AELE", "GAN" por HBA). Un departamento
# nuevo de verdad que las tres pasadas leen igual se guarda con su nombre.
_OCR_KNOWN_DEPARTMENTS = (
    "HOT DOGS SANDWICH", "AUTO", "BEER/WINE", "BOILED PEANUTS", "CANDY", "CHEVRON GIFT CARD",
    "CIGARS", "COFFE", "E-GIGARETTE", "FEES", "FLOWERS", "FOUTAIN", "GEN-CTN", "GEN-PAK",
    "GIFT CARD", "GROCERIES", "HBA", "ICECREAM", "JUICE", "GETTEL/TOYOTA", "MAJ CR", "MAJ PAK",
    "MILK", "NONTAX", "ONLINE", "PROPANE", "SKOFF", "SNACK", "SNUFF", "SODA", "STORE COUPON",
    "TAXABLE", "WATER",
)


def _ocr_money_value(token):
    text = token.replace("—", "-").replace("–", "-").replace("~", "-").replace("=", "-")
    negative = text.startswith("-") or "(" in text or "-$" in text
    try:
        value = float(re.sub(r"[^\d.]", "", text))
    except ValueError:
        return None
    return -value if negative else value


def _classify_ocr_token(token):
    """
    Un token de una fila de departamento -> ("M", monto), ("I", entero),
    ("P", porcentaje) o None (nombre, ruido o ilegible).
    """
    raw = token.strip(",;:'\"`")
    if not raw:
        return None
    if raw.endswith("%"):
        try:
            return ("P", float(re.sub(r"[^\d.]", "", raw) or "x"))
        except ValueError:
            return ("P", None)
    if "$" in raw or re.search(r"\.\d{2}\)?$", raw):
        # Un signo menos garabateado ("—$3.00", o un carácter ilegible antes del "$") queda como "-";
        # una comilla suelta adelante ("‘$10.96") es ruido, no un signo.
        lead = re.match(r"^[^\d$]+", raw)
        if lead:
            raw = ("-" if re.search(r"[-—–~=\ufffd]", lead.group(0)) else "") + raw[lead.end():]
        fixed = raw.translate(_OCR_DIGIT_FIX) if re.search(r"\d", raw) else raw
        if _OCR_MONEY_TOKEN_RE.match(fixed) or _OCR_MONEY_TOKEN_RE.match(fixed.replace("$", "")):
            return ("M", _ocr_money_value(fixed))
    elif re.search(r"\d", raw) and len(raw) <= 5:
        fixed = raw.translate(_OCR_DIGIT_FIX)
        if _OCR_INT_TOKEN_RE.match(fixed):
            return ("I", int(fixed))
    if raw in ("O", "o", ")", "(", "D", "Q"):
        return ("I", 0)
    # Basura pegada adelante/atrás: "4.", "40°", "1—", "$0.00.", "$3480".
    if token.startswith("-$"):
        trimmed = "-" + _OCR_EDGE_JUNK_RE.sub("", token[1:])
    else:
        trimmed = _OCR_EDGE_JUNK_RE.sub("", token)
    if "$" in trimmed:
        fixed = trimmed.translate(_OCR_DIGIT_FIX)
        match = _OCR_MONEY_NO_POINT_RE.match(fixed)
        if match:
            value = float(f"{match.group(2).replace(',', '')}.{match.group(3)}")
            return ("M", -value if match.group(1) else value)
        if trimmed != token and _OCR_MONEY_TOKEN_RE.match(fixed):
            return ("M", _ocr_money_value(fixed))
        return None
    if trimmed and trimmed != token and re.fullmatch(r"[\dOoDQlI|SBZ]{1,5}", trimmed) and re.search(r"\d", trimmed):
        fixed = trimmed.translate(_OCR_DIGIT_FIX)
        if _OCR_INT_TOKEN_RE.match(fixed):
            return ("I", int(fixed))
    return None


def _parse_ocr_sales_row(line, allow_unknown=False):
    """
    Una línea de OCR -> dict con TODOS los campos de la fila (no solo Net
    Count/Net Sales), asignados por la forma de cada valor y no por posición:
    monto, 3 enteros, 3 montos y el % al final (M I I I M M M [P]). Así un
    token de ruido suelto ya no corre las columnas. Si falta exactamente uno
    de los tres conteos, se deduce de la ecuación de la fila. Con
    allow_unknown, una fila con dos o más conteos ilegibles vuelve igual
    ("partial") para releer esas celdas en la imagen.
    """
    text = _normalize_department_spacing(_strip_cell(line))
    if not text or _is_ocr_header_line(text):
        return None
    tokens = text.split()
    name_tokens, entries = [], []
    merged = set()
    for index, token in enumerate(tokens):
        if index in merged or _OCR_NOISE_TOKEN_RE.match(token):
            continue
        if (entries and index + 1 < len(tokens) and re.fullmatch(r"-?\$\d{1,3}(,\d{3})*", token)
                and re.fullmatch(r"\d{2}", tokens[index + 1])):
            token = f"{token}.{tokens[index + 1]}"
            merged.add(index + 1)
        kind = _classify_ocr_token(token)
        if kind is None and entries and not re.search(r"[0-9A-Za-z§]", token):
            # Ruido suelto ("-", "~-", "*", "<~")... o un conteo garabateado ("€" por "7").
            kind = ("J", None)
        if kind is None:
            if not entries:
                if re.search(r"[A-Za-z]", token):
                    name_tokens.append(token)
                continue
            kind = ("X", None)
        elif not entries and kind[0] == "I" and name_tokens and re.fullmatch(r"\d+[A-Za-z]+|[A-Za-z]+\d+", token):
            name_tokens.append(token)
            continue
        entries.append((kind, token, index))
    # Primero el ruido suelto no cuenta; si así la fila no cierra su forma,
    # se prueba de nuevo tomándolo como un valor ilegible.
    row = _assign_ocr_row_values([e for e in entries if e[0][0] != "J"], allow_unknown)
    if row is None and any(e[0][0] == "J" for e in entries):
        row = _assign_ocr_row_values([(("X", None) if e[0][0] == "J" else e[0], e[1], e[2]) for e in entries], allow_unknown)
    if row is None:
        return None
    row.update({"name": " ".join(name_tokens), "tokens": tokens})
    return row


def _assign_ocr_row_values(entries, allow_unknown):
    """Los valores de una fila (kind, token, posición) -> campos de la fila, o None si no tiene la forma."""
    values = [kind for kind, _token, _index in entries]
    raw_values = [token for _kind, token, _index in entries]
    positions = [index for _kind, _token, index in entries]
    pct = None
    if values and values[-1][0] == "P":
        pct = values.pop()[1]
        positions.pop()
    elif len(values) == 8 and values[-1][0] in ("M", "X") and "$" not in raw_values[-1]:
        # El % sin su "%" ("0.20") o ilegible en su lugar.
        values.pop()
        positions.pop()
    if len(values) == 8 and "".join(kind for kind, _ in values).endswith("MMMI"):
        values.pop()
        positions.pop()
    kinds = "".join(kind for kind, _ in values)
    vals = [value for _, value in values]
    if len(kinds) > 7 and kinds.endswith("MIIIMMM"):
        kinds, vals, positions = kinds[-7:], vals[-7:], positions[-7:]
    inferred = partial = False
    count_positions = positions[1:4]
    if kinds == "MIIIMMM":
        gross, item, refund_count, net_count, refund, discount, net = vals
    elif (len(kinds) == 7 and kinds[0] == "M" and kinds[4:] == "MMM"
            and kinds[1:4].count("X") == 1 and kinds[1:4].count("I") == 2):
        gross, a, b, c, refund, discount, net = vals
        if a is None:
            item, refund_count, net_count = c + b, b, c
        elif b is None:
            item, refund_count, net_count = a, a - c, c
        else:
            item, refund_count, net_count = a, b, a - b
        inferred = True
        if min(item, refund_count, net_count) < 0:
            return None
    elif allow_unknown and len(kinds) == 7 and kinds[0] == "M" and kinds[4:] == "MMM" and set(kinds[1:4]) <= {"I", "X"}:
        gross, item, refund_count, net_count, refund, discount, net = vals
        partial = True
    elif kinds == "MIIMMM":
        # Falta un entero: si los dos que hay son iguales, el que falta es Refund Count = 0.
        gross, a, b, refund, discount, net = vals
        if a != b:
            return None
        item, refund_count, net_count = a, 0, b
        count_positions = [positions[1], None, positions[2]]
    else:
        return None
    return {
        "count_positions": count_positions,
        "gross": gross, "item": item, "refund_count": refund_count, "net_count": net_count,
        "refund": refund, "discount": discount, "net": net, "pct": pct,
        "inferred": inferred, "partial": partial,
        "count_ok": not partial and item - refund_count == net_count,
        "amount_ok": abs(gross - abs(refund) - abs(discount) - net) < 0.011,
    }


def _ocr_row_department(name):
    """Nombre leído -> departamento normalizado; "" para la fila del total; None si no es una fila."""
    if not name or not re.search(r"[A-Za-z]", name):
        return ""
    if _normalize_department_label(name) in ("total", "grand total", "totals"):
        return ""
    text, _fused = _split_dept_name_from_fused_tail(_normalize_department_spacing(name).upper())
    department = normalize_parsed_department_name(_fallback_department_alias(text))
    if not department or _is_summary_row(department):
        return None
    if department in _OCR_KNOWN_DEPARTMENTS:
        return department
    # Basura corta pegada adelante o atrás del nombre ("J TAXABLE", "FOUTAIN I").
    tokens = department.split()
    for start in range(0, min(2, len(tokens))):
        for end in range(len(tokens), max(len(tokens) - 2, start), -1):
            trimmed = " ".join(tokens[start:end])
            extra = tokens[:start] + tokens[end:]
            if extra and trimmed in _OCR_KNOWN_DEPARTMENTS and all(len(token) <= 2 for token in extra):
                return trimmed
    close = get_close_matches(department, _OCR_KNOWN_DEPARTMENTS, n=2, cutoff=0.75)
    if len(close) == 1 or (close and SequenceMatcher(None, department, close[0]).ratio()
                           - SequenceMatcher(None, department, close[1]).ratio() > 0.1):
        return close[0]
    return department


def _ocr_page_readings(image):
    """
    Las lecturas de una página: el texto de cada configuración de
    _OCR_TEXT_CONFIGS, las palabras de --psm 6 con su caja en la imagen
    (image_to_data; de ahí sale también su texto) y el mejor texto según
    _score_ocr_page_text (el que se usa para el ancla y el período).
    """
    readings = {"texts": {}, "words": [], "best": "", "image": image}
    if image is None:
        return readings
    _ensure_pytesseract()
    for config in _OCR_TEXT_CONFIGS:
        try:
            if config == "--psm 6":
                data = pytesseract.image_to_data(image, config=config, output_type=pytesseract.Output.DICT)
                words = [
                    (str(data["text"][i]), data["left"][i], data["top"][i], data["width"][i], data["height"][i],
                     (data["block_num"][i], data["par_num"][i], data["line_num"][i]))
                    for i in range(len(data["text"])) if str(data["text"][i]).strip()
                ]
                readings["words"] = words
                text = "\n".join(" ".join(word[0] for word in line) for line in _ocr_word_lines(words))
            else:
                text = pytesseract.image_to_string(image, config=config) or ""
        except Exception:
            continue
        readings["texts"][config] = text
    best_score = (-1, -1)
    for config in _OCR_TEXT_CONFIGS:
        text = readings["texts"].get(config)
        if text is None:
            continue
        score = _score_ocr_page_text(text)
        if score > best_score:
            best_score = score
            readings["best"] = text
    return readings


def _ocr_word_lines(words):
    """Palabras de image_to_data agrupadas en líneas, en el orden de lectura de Tesseract."""
    lines = {}
    for word in words:
        lines.setdefault(word[5], []).append(word)
    return [sorted(line, key=lambda word: word[1]) for line in lines.values()]


def _ocr_reread_digits(image, box):
    """Relee una celda de conteo recortada y agrandada 3x, solo con dígitos ("" si no lee nada)."""
    _ensure_pytesseract()
    left, top, width, height = box
    pad = 6
    crop = image.convert("L").crop((left - pad, top - pad, left + width + pad, top + height + pad))
    big = crop.resize((crop.width * 3, crop.height * 3), Image.LANCZOS)
    digits = ""
    for psm in ("7", "10"):
        try:
            text = pytesseract.image_to_string(big, config=f"--psm {psm} -c tessedit_char_whitelist=0123456789")
        except Exception:
            continue
        digits = re.sub(r"\D", "", text or "")
        if digits:
            break
    return digits


def _ocr_digit_width(rows):
    """Ancho de un dígito en la página: la mediana de las cajas de los conteos de un solo dígito."""
    widths = []
    for row, words in rows:
        for position in row["count_positions"]:
            if position is not None and re.fullmatch(r"\d", row["tokens"][position]):
                widths.append(words[position][3])
    return statistics.median(widths) if len(widths) >= 5 else None


def _ocr_fix_counts_from_image(row, words, image, digit_width):
    """
    Relee en la imagen los conteos dudosos de una fila de --psm 6: la caja
    más ancha que los dígitos leídos ("7" con el ancho de "77"), un conteo
    ilegible, o las tres celdas si la fila no cierra su ecuación (siempre en
    la fila del total). Aplica la relectura entera si la fila cierra; si no,
    solo las celdas corregidas por ancho.
    """
    if image is None or len(words) != len(row["tokens"]):
        return
    names = ("item", "refund_count", "net_count")
    reread_all = not row["count_ok"] or row["inferred"] or row["department"] == ""
    reread = {}
    for name, position in zip(names, row["count_positions"]):
        if position is None:
            continue
        word = words[position]
        expected = round(word[3] / digit_width) if digit_width else None
        too_wide = row[name] is not None and expected is not None and expected > len(str(row[name]))
        if not (too_wide or reread_all or row[name] is None):
            continue
        digits = _ocr_reread_digits(image, word[1:5])
        if digits and (not too_wide or len(digits) == expected):
            reread[name] = (int(digits), too_wide)
    if not reread:
        return
    if "net_count" in reread and reread["net_count"][1]:
        row["net_confirmed"] = reread["net_count"][0]
    candidate = {name: reread[name][0] if name in reread else row[name] for name in names}
    if row["count_positions"][1] is None:
        candidate["refund_count"] = 0
    missing = [name for name in names if candidate[name] is None]
    inferred = False
    if len(missing) == 1:
        item, refund_count, net_count = (candidate[name] for name in names)
        value = {"item": lambda: refund_count + net_count, "refund_count": lambda: item - net_count,
                 "net_count": lambda: item - refund_count}[missing[0]]()
        if value >= 0:
            candidate[missing[0]] = value
            inferred = True
    if None not in candidate.values() and candidate["item"] - candidate["refund_count"] == candidate["net_count"]:
        new = candidate
    elif row["department"] == "":
        return
    else:
        new = {name: reread[name][0] if name in reread and reread[name][1] else row[name] for name in names}
    if None in new.values() or new == {name: row[name] for name in names}:
        return
    row.update(new)
    row["count_ok"] = new["item"] - new["refund_count"] == new["net_count"]
    row["image_checked"] = True
    row["partial"] = False
    row["inferred"] = inferred and new is candidate


def _ocr_table_rows(pages):
    """
    Filas de la tabla de las 3 pasadas, desde el ancla hasta "Safe Drop
    Report". `pages` = lecturas de _ocr_page_readings, la del ancla primero.
    """
    rows = []
    for config in _OCR_ROW_CONFIGS:
        for number, page in enumerate(pages):
            past_anchor = number != 0
            stop = False
            if config == "--psm 6":
                lines = [(" ".join(word[0] for word in line), line) for line in _ocr_word_lines(page["words"])]
            else:
                lines = [(line, None) for line in page["texts"].get(config, "").splitlines()]
            parsed = []
            for text, words in lines:
                if not past_anchor:
                    past_anchor = _line_contains_anchor(text)
                    continue
                if _line_contains_stop_marker(text):
                    stop = True
                    break
                row = _parse_ocr_sales_row(text, allow_unknown=words is not None)
                if row is None:
                    continue
                department = _ocr_row_department(row["name"])
                if department is None:
                    continue
                row["department"] = department
                row["config"] = config
                parsed.append((row, words))
            if config == "--psm 6":
                digit_width = _ocr_digit_width(parsed)
                for row, words in parsed:
                    _ocr_fix_counts_from_image(row, words, page["image"], digit_width)
            rows.extend(row for row, _words in parsed if not row["partial"])
            if stop:
                break
    return rows


def _ocr_candidates(rows, field):
    """
    {valor: puntaje} de un campo entre las lecturas de una fila: la lectura
    que cierra su ecuación vale 3; la deducida, 2; una que no cierra vale 1,
    y también suma 1 el valor que daría su ecuación.
    """
    score = Counter()
    for row in rows:
        if field == "count":
            if row["inferred"]:
                score[row["net_count"]] += 2
            elif row["count_ok"]:
                score[row["net_count"]] += 3
            else:
                score[row["net_count"]] += 1
                score[row["item"] - row["refund_count"]] += 1
        else:
            value = round(row["net"], 2)
            if row["amount_ok"]:
                score[value] += 3
            else:
                score[value] += 1
                score[round(row["gross"] - abs(row["refund"]) - abs(row["discount"]), 2)] += 1
    return score


def _ocr_resolve_unknown_names(order, by_department):
    """
    Nombres ilegibles que no se parecen a ninguno ("TAY AELE"; "GAN" y "GEA"
    para el mismo renglón de HBA, cada pasada lo garabatea distinto). El POS
    imprime los departamentos siempre en el mismo orden
    (_OCR_KNOWN_DEPARTMENTS), así que el renglón se ubica por sus vecinos:
    si entre ellos entra un solo departamento conocido que falta ese día (o
    uno solo con la misma inicial), es ese. Si no se puede saber: un nombre
    que las tres pasadas leen igual es un departamento nuevo de verdad y se
    guarda así; si no, el renglón NO se guarda con un nombre inventado y se
    devuelve aparte (unnamed) para avisar sus valores.
    Devuelve (order, unnamed) y deja by_department con los nombres resueltos.
    """
    known = _OCR_KNOWN_DEPARTMENTS
    groups = []
    for index, department in enumerate(order):
        if department in known:
            continue
        nets = {round(row["net"], 2) for row in by_department[department]}
        # Ilegibles seguidos con el mismo monto son el mismo renglón leído distinto.
        if groups and groups[-1]["end"] == index - 1 and groups[-1]["nets"] & nets:
            groups[-1]["names"].append(department)
            groups[-1]["nets"] |= nets
            groups[-1]["end"] = index
        else:
            groups.append({"start": index, "end": index, "names": [department], "nets": nets})
    resolved = {}
    unnamed = []
    for group in groups:
        before = next((d for d in reversed(order[:group["start"]]) if d in known), None)
        after = next((d for d in order[group["end"] + 1:] if d in known), None)
        low = known.index(before) if before else -1
        high = known.index(after) if after else len(known)
        free = [d for d in known[low + 1:high] if d not in by_department]
        if len(free) > 1:
            same_letter = [d for d in free if any(d[0] == name[0] for name in group["names"])]
            if len(same_letter) == 1:
                free = same_letter
        rows = [row for name in group["names"] for row in by_department.pop(name)]
        if len(free) == 1:
            target = free[0]
        else:
            stable = [name for name in group["names"] if len(name) >= 4
                      and len({row["config"] for row in rows if row["department"] == name}) == len(_OCR_ROW_CONFIGS)]
            target = stable[0] if len(stable) == 1 else None
        for name in group["names"]:
            resolved[name] = target
        if target:
            for row in rows:
                row["department"] = target
            by_department[target] = rows
        else:
            unnamed.append(rows)
    new_order = []
    for department in order:
        department = resolved.get(department, department)
        if department and department not in new_order:
            new_order.append(department)
    return new_order, unnamed


def _vote_ocr_departments(pages):
    """
    (records, printed_totals, status): un registro por departamento con el
    valor votado de Net Count y Net Sales, el total impreso ya confirmado, y
    el resultado del cierre contra ese total ({"count"/"amount": "ok" |
    "corregido" | "no_cierra" | "sin_total"}).
    """
    rows = _ocr_table_rows(pages)
    by_department, order, totals = {}, [], []
    previous = {}
    for row in rows:
        if row["department"] == "":
            totals.append(row)
            continue
        if row["department"] not in by_department:
            # Un departamento que solo leyó otra pasada va donde lo leyó esa
            # pasada (después del mismo vecino), no al final de la lista.
            if row["config"] not in previous:
                position = 0
            else:
                after = previous[row["config"]]
                position = order.index(after) + 1 if after in order else len(order)
            order.insert(position, row["department"])
        previous[row["config"]] = row["department"]
        by_department.setdefault(row["department"], []).append(row)
    order, unnamed = _ocr_resolve_unknown_names(order, by_department)

    records = []
    for department in order:
        record = {"department": department, "is_total": False}
        for field in ("count", "amount"):
            department_rows = by_department[department]
            if field == "count" and any(r.get("image_checked") and r["count_ok"] for r in department_rows):
                # La relectura en la imagen manda: las otras pasadas fundieron los mismos dígitos.
                department_rows = [r for r in department_rows if r.get("image_checked") and r["count_ok"]]
            candidates = _ocr_candidates(department_rows, field)
            if field == "count":
                for row in by_department[department]:
                    if row.get("net_confirmed") is not None:
                        candidates[row["net_confirmed"]] += 6
            ranked = candidates.most_common()
            record[field] = ranked[0][0]
            record[field + "_options"] = candidates
            record[field + "_sure"] = len(ranked) == 1 or (ranked[0][1] >= 2 * ranked[1][1] and ranked[0][1] >= 3)
        records.append(record)

    total_count = _ocr_candidates(totals, "count")
    for row in totals:
        if row.get("image_checked") and row["count_ok"]:
            total_count[row["net_count"]] += 3
    total_amount = _ocr_candidates(totals, "amount")
    printed_amount = total_amount.most_common(1)[0][0] if total_amount else None
    # % de ventas: confirma cuál de las lecturas del monto es la buena.
    if printed_amount:
        for record in records:
            pcts = {r["pct"] for r in by_department[record["department"]] if r["pct"] is not None}
            options = record["amount_options"]
            confirmed = [v for v in options if any(abs(v / printed_amount * 100 - p) < 0.006 for p in pcts)]
            if len(confirmed) == 1 and len(options) > 1:
                options[confirmed[0]] += 10
                record["amount"] = confirmed[0]
                record["amount_sure"] = True

    # Un renglón sin nombre legible no se guarda, pero sus valores sí se
    # conocen: se descuentan del total impreso para cerrar los demás.
    unnamed_values = [
        {"count": _ocr_candidates(rows, "count").most_common(1)[0][0],
         "amount": _ocr_candidates(rows, "amount").most_common(1)[0][0]}
        for rows in unnamed
    ]
    unnamed_count = sum(item["count"] for item in unnamed_values)
    unnamed_amount = sum(item["amount"] for item in unnamed_values)
    status = _reconcile_ocr_departments(records, {
        "count": Counter({value - unnamed_count: score for value, score in total_count.items()}),
        "amount": Counter({round(value - unnamed_amount, 2): score for value, score in total_amount.items()}),
    })
    for field, offset in (("count", unnamed_count), ("amount", unnamed_amount)):
        if field + "_target" in status:
            status[field + "_target"] = round(status[field + "_target"] + offset, 2)
    status["unnamed"] = unnamed_values
    printed_totals = None
    if totals:
        printed_totals = {
            "department": "", "is_total": True,
            "count": int(status.get("count_target", total_count.most_common(1)[0][0])),
            "amount": float(status.get("amount_target", printed_amount)),
        }
    return records, printed_totals, status


def _ocr_total_fixes(records, field, target, tolerance):
    """Cambios de 1 o 2 valores dudosos que cierran contra target, del más barato al más caro."""
    current = sum(record[field] for record in records)
    if abs(current - target) <= tolerance:
        return [(0, [])]
    doubtful = [record for record in records if len(record[field + "_options"]) > 1]
    fixes = []
    for size in (1, 2):
        for combo in itertools.combinations(doubtful, size):
            alternatives = [[v for v in record[field + "_options"] if v != record[field]] for record in combo]
            for choice in itertools.product(*alternatives):
                delta = sum(value - record[field] for value, record in zip(choice, combo))
                if abs(current + delta - target) <= tolerance:
                    cost = sum(record[field + "_options"][record[field]] - record[field + "_options"][value]
                               for value, record in zip(choice, combo))
                    fixes.append((cost, list(zip(combo, choice))))
        if fixes:
            break
    return sorted(fixes, key=lambda fix: fix[0])


def _reconcile_ocr_departments(records, total_options):
    """
    Cierra unidades y monto EXACTO contra el total impreso. Prueba las
    lecturas posibles del total (también puede estar mal leído: "593" como
    "693"), de la más votada a la menos, y aplica el cambio más barato solo
    si es claramente mejor que el siguiente. Si no cierra con ninguna, los
    valores que no estaban seguros quedan marcados como dudosos.
    """
    status = {}
    for record in records:
        record["doubt"] = []
    for field, tolerance in (("count", 0), ("amount", 0.005)):
        targets = [value for value, _score in total_options[field].most_common()]
        closed = False
        for target in targets:
            fixes = _ocr_total_fixes(records, field, target, tolerance)
            if not fixes or (len(fixes) > 1 and fixes[1][0] - fixes[0][0] < 2):
                continue
            for record, value in fixes[0][1]:
                record[field] = value
            status[field] = "ok" if not fixes[0][1] else "corregido"
            status[field + "_target"] = target
            closed = True
            break
        if not closed:
            status[field] = "no_cierra" if targets else "sin_total"
            for record in records:
                if not record[field + "_sure"]:
                    record["doubt"].append(field)
    return status


def parse_elistar_daily_pdf_ocr(pdf_path, start_page_index=DEFAULT_PDF_PAGE_INDEX):
    """
    OCR-based extraction for photographed/scanned daily PDFs (no
    extractable text). The "Department Sales Report" table doesn't always
    land on the same page and can spill onto the next one, so this scans
    forward from start_page_index (wrapping around) for the anchor,
    corrects page rotation via Tesseract OSD, and keeps reading department
    rows across pages until the "Safe Drop Report" anchor is reached.
    Each value is voted across three OCR passes and the day is closed
    against the printed grand-total line (see _vote_ocr_departments).

    Returns:
        tuple[list[dict], dict]: (records, diagnostics) — each record is
        {"department", "count", "amount", "is_total": False}, with count
        and/or amount None when it couldn't be read with confidence (never
        a guessed value). diagnostics has keys "pages_used",
        "last_department", "subtotal_mismatch" (the sum of the rows doesn't
        close against the printed grand total even after trying the other
        readings — e.g. a whole row OCR skipped), "doubtful_departments"
        ([{"department", "fields"}] left empty), "printed_totals" (the
        grand-total line as {"count", "amount"}, when the report printed
        one) and "period" (the {"from_date", "to_date", ...} dict parsed
        from this same page's own "PERIOD FROM: ... TO: ..." line, or None).
    """
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF no encontrado: {pdf_path}")

    images = _LazyPdfPageImages(pdf_path)
    total_pages = len(images)
    if not (0 <= start_page_index < total_pages):
        start_page_index = 0

    search_order = list(range(start_page_index, total_pages)) + list(
        range(0, start_page_index)
    )
    readings_by_page = {}

    def page_readings(idx):
        if idx not in readings_by_page:
            readings_by_page[idx] = _ocr_page_readings(images[idx])
        return readings_by_page[idx]

    anchor_page = None
    for idx in search_order:
        if _line_contains_anchor(page_readings(idx)["best"]):
            anchor_page = idx
            break

    if anchor_page is None:
        images.close()
        raise ValueError(
            f'No se encontró el ancla "{DEPARTMENT_SALES_REPORT_ANCHOR}" en ninguna '
            "página del PDF (vía OCR). Verifique que el reporte no esté demasiado "
            "borroso o girado."
        )

    table_pages = []
    pages_used = []
    idx = anchor_page
    while idx < total_pages:
        readings = page_readings(idx)
        table_pages.append(readings)
        pages_used.append(idx + 1)
        past_anchor = idx != anchor_page
        reached_stop = False
        for line in readings["best"].splitlines():
            if not past_anchor:
                past_anchor = _line_contains_anchor(line)
                continue
            if _line_contains_stop_marker(line):
                reached_stop = True
                break
        if reached_stop:
            break
        idx += 1

    try:
        voted, printed_totals, status = _vote_ocr_departments(table_pages)
    finally:
        images.close()

    if not voted:
        raise ValueError(
            f"No se encontraron registros de departamento vía OCR después de "
            f'"{DEPARTMENT_SALES_REPORT_ANCHOR}". Verifique que el reporte no esté '
            "demasiado borroso o girado."
        )

    # Lo que sigue dudoso después de votar y cerrar contra el total queda
    # VACÍO (None) y se avisa -- regla de oro del OCR: nunca un valor dudoso.
    records = []
    doubtful = []
    for record in voted:
        fields = record["doubt"]
        records.append({
            "department": record["department"],
            "count": None if "count" in fields else int(record["count"]),
            "amount": None if "amount" in fields else round(float(record["amount"]), 2),
            "is_total": False,
        })
        if fields:
            doubtful.append({"department": record["department"], "fields": list(fields)})
    for item in status.get("unnamed", []):
        doubtful.append({"department": None, "fields": ["count", "amount"], "values": item})

    subtotal_mismatch = None
    if printed_totals is not None and "no_cierra" in (status.get("count"), status.get("amount")):
        subtotal_mismatch = {
            "computed_amount": round(sum(r["amount"] for r in records if r["amount"] is not None), 2),
            "printed_amount": printed_totals["amount"],
            "computed_count": sum(r["count"] for r in records if r["count"] is not None),
            "printed_count": printed_totals["count"],
        }

    period = None
    anchor_readings = readings_by_page[anchor_page]
    for text in [anchor_readings["best"]] + list(anchor_readings["texts"].values()):
        for line in text.splitlines():
            if PERIOD_FROM_ANCHOR in line.lower():
                period = _parse_period_from_to_line(line)
                if period:
                    break
        if period:
            break

    diagnostics = {
        "pages_used": pages_used,
        "last_department": records[-1]["department"],
        "subtotal_mismatch": subtotal_mismatch,
        "doubtful_departments": doubtful,
        # Sin la fila del total legible no hay contra qué verificar la suma.
        "total_unverified": status.get("count") == "sin_total",
        "printed_totals": printed_totals,
        "period": period,
    }
    return records, diagnostics


def _page_is_scanned(page):
    """La página es una foto/escaneo: una imagen incrustada que cubre al menos la mitad de la hoja."""
    page_area = float(page.width * page.height) or 1.0
    return any(
        (image["x1"] - image["x0"]) * (image["bottom"] - image["top"]) >= 0.5 * page_area
        for image in page.images
    )


def _parse_elistar_daily_pdf_page_uncached(pdf_path, page_index=DEFAULT_PDF_PAGE_INDEX):
    """
    Extract department records from the selected PDF page (0-based index).

    Only rows after the \"Department Sales Report\" anchor are parsed.
    Uses fixed columns: Dept.Name (1), Net Count (5), Net Sales $ (8).
    Falls back to OCR automatically when the page has no extractable text
    at all (a photographed/scanned report, as opposed to a digital export).

    Returns:
        tuple[list[dict], dict]: (records, diagnostics) — diagnostics has
        keys "used_ocr", "pages_used" (OCR only) and "last_department"
        (OCR only, may be absent for text-based extraction).
    """
    _ensure_pdfplumber()
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF no encontrado: {pdf_path}")

    if page_index < 0:
        raise ValueError("El índice de página del PDF debe ser 0 o mayor.")

    with pdfplumber.open(pdf_path) as pdf:
        if page_index >= len(pdf.pages):
            raise ValueError(
                f"Página {page_index + 1} del PDF no encontrada; "
                f"el documento tiene {len(pdf.pages)} página(s)."
            )
        page = pdf.pages[page_index]
        # Una foto/escaneo puede traer igual una capa de texto que le agrega
        # la app del escáner, inservible ("R E T A W E L B A X A T ...", al
        # revés y letra por letra): con ella no se encontraba el ancla y se
        # perdía el día entero, o salían 5 departamentos de 23 sin aviso
        # (relevamiento 2026-10-04: 11 de 189 PDFs reales). Si la página es
        # una imagen, se lee siempre por OCR.
        has_text = bool((page.extract_text() or "").strip()) and not _page_is_scanned(page)
        records = _parse_tables_with_line_anchor(page) if has_text else []

    if records:
        return [dict(record) for record in records], {"used_ocr": False}

    if not has_text:
        ocr_records, ocr_diagnostics = parse_elistar_daily_pdf_ocr(
            pdf_path, start_page_index=page_index
        )
        diagnostics = {"used_ocr": True}
        diagnostics.update(ocr_diagnostics)
        return [dict(record) for record in ocr_records], diagnostics

    raise ValueError(
        f"No se encontraron registros de departamento en la página {page_index + 1} del PDF después de "
        f'"{DEPARTMENT_SALES_REPORT_ANCHOR}". '
        "Se esperaba Dept.Name (col 1), Net Count (col 5), Net Sales $ (col 8)."
    )


@functools.lru_cache(maxsize=64)
def _parse_elistar_daily_pdf_page_cached(pdf_path, page_index):
    return _parse_elistar_daily_pdf_page_uncached(pdf_path, page_index)


def parse_elistar_daily_pdf_page(pdf_path, page_index=DEFAULT_PDF_PAGE_INDEX):
    """
    Wrapper con caché sobre _parse_elistar_daily_pdf_page_uncached -- pedido
    explícito del usuario (2026-09-14): "los pdf siguen cargando... se
    deberia ignorar las paginas que no sean relevantes de leer". La página
    de Department Sales Report YA se leía de forma perezosa (solo hasta
    encontrar el ancla, nunca las 50+ páginas del PDF -- ver
    _LazyPdfPageImages) -- el problema real, encontrado investigando, era
    otro: para el MISMO PDF, esta función se llamaba 3 veces independientes
    dentro de un mismo job (una desde `process_reporte_diario`, otra desde
    `extract_department_sales_for_day`, otra desde
    `extract_lottery_department_fields_from_pdf`), cada una re-decodificando
    y re-OCR'ando la página del ancla desde cero. `lru_cache` (acotado a 64
    PDFs distintos, cada request usa una ruta de archivo temporal única, así
    que nunca queda un resultado viejo pisando uno nuevo del mismo path)
    evita repetir ese trabajo -- se devuelve una copia nueva en cada llamada
    (`copy.deepcopy`) para que ningún caller pueda mutar el resultado
    cacheado de otro por accidente.
    """
    key = (os.path.abspath(pdf_path), page_index)
    if key in _PREFETCH_ERRORS:
        raise _PREFETCH_ERRORS[key]
    records, diagnostics = _parse_elistar_daily_pdf_page_cached(*key)
    return copy.deepcopy(records), copy.deepcopy(diagnostics)


# Lectura por adelantado (pedido del usuario, 2026-10-06: "que los reportes
# diarios se carguen más velozmente sin perder calidad"). Casi todo el tiempo
# de un Reporte Diario es el OCR de Ventas por Departamento (unas 26 llamadas
# a Tesseract, una detrás de otra), y la carga leía los PDF de a uno. Ahora,
# mientras se guarda uno, los siguientes ya se están leyendo en otros hilos
# (Tesseract corre en su propio proceso, así que usa los otros núcleos). La
# lectura es exactamente la misma -- mismas pasadas, misma votación, mismo
# cierre contra el total --; solo cambia cuándo se hace: el resultado queda
# en la caché de parse_elistar_daily_pdf_page y la carga lo toma de ahí, en
# el orden de siempre. Un error también se guarda, para no leer dos veces un
# PDF que no se puede leer.
_PREFETCH_ERRORS = {}


def _prefetch_workers():
    # Medido con 6 PDFs de septiembre en la PC de la oficina (4 núcleos): en
    # serie 219 s, 2 hilos 140 s, 3 hilos 148 s (Tesseract ya usa más de un
    # núcleo por lectura), con resultados idénticos. La mitad de los núcleos.
    return max(1, min((os.cpu_count() or 2) // 2, _MAX_CONCURRENT_PDF_WORKERS))


class DepartmentPagePrefetch:
    def __init__(self, pdf_paths, workers=None):
        self._keys = [(os.path.abspath(p), DEFAULT_PDF_PAGE_INDEX) for p in pdf_paths]
        workers = min(workers or _prefetch_workers(), len(self._keys))
        self._ahead = workers * 2  # nunca más adelante que lo que entra en la caché (64)
        self._executor = ThreadPoolExecutor(max_workers=workers) if len(self._keys) > 1 else None
        self._futures = {}
        self._submitted = 0
        self._fill(0)

    @staticmethod
    def _warm(key):
        try:
            _parse_elistar_daily_pdf_page_cached(*key)
        except Exception as exc:
            _PREFETCH_ERRORS[key] = exc

    def _fill(self, current):
        while self._executor and self._submitted < len(self._keys) and self._submitted <= current + self._ahead:
            self._futures[self._submitted] = self._executor.submit(self._warm, self._keys[self._submitted])
            self._submitted += 1

    def wait(self, index):
        """Espera a que el PDF `index` (desde 0) esté leído y encola los siguientes."""
        self._fill(index)
        future = self._futures.pop(index, None)
        if future is not None:
            future.result()

    def close(self):
        if self._executor:
            self._executor.shutdown(wait=False, cancel_futures=True)
        for key in self._keys:
            _PREFETCH_ERRORS.pop(key, None)


def extract_department_sales_for_day(pdf_path):
    """
    Extracción "pura" (sin tocar ningún Excel) de los departamentos de un
    PDF de cierre diario + la fecha real de negocio a la que pertenecen --
    para guardarlos en reportes_db.py. Reusa parse_elistar_daily_pdf_page
    (el mismo motor ya usado por Reporte Diario y por el control
    Controles→Cierre Mensual→Ventas por Departamento) en vez de duplicar
    lógica de parseo.

    El PDF imprime la fecha un día antes del día de negocio real (mismo bug
    ya conocido y corregido en Reporte Diario/Cierre Mensual) -- por eso
    esta función devuelve la fecha YA con el +1 día aplicado, para que
    quien la use nunca tenga que acordarse de hacerlo por su cuenta.

    La fecha sale gratis de los diagnósticos cuando el PDF se leyó vía OCR
    (parse_elistar_daily_pdf_ocr ya la calcula al pasar); si el PDF tiene
    texto digital (no hizo falta OCR), se llama una vez más a
    extract_store_info_from_pdf solo para sacar la fecha de esa misma
    página "PERIOD FROM: ...".
    """
    records, diagnostics = parse_elistar_daily_pdf_page(pdf_path)
    period = diagnostics.get("period")
    if period is None:
        # Solo hace falta el período: lo demás de Store Info puede faltar.
        period = extract_store_info_from_pdf(pdf_path, strict=False)
    business_date = period["from_date"] + timedelta(days=1)

    # Cross-chequeo contra el total impreso al pie del Department Sales
    # Report -- pedido explícito del usuario (2026-09-17), tras confirmar
    # con un PDF real que Tesseract puede saltearse una fila entera sin
    # dejar ningún rastro de texto (ni garabateado) -- "GIFT CARD" ese día
    # tenía 5 unidades y $0.00, y ninguna de las 4 pasadas de OCR la leyó.
    # `subtotal_mismatch` (solo se calcula del lado OCR, ver
    # parse_elistar_daily_pdf_ocr) ya detectaba justo este caso (diferencia
    # de conteo/monto contra el total impreso) pero se descartaba acá sin
    # usarlo -- ahora se devuelve, para que el caller (webapp.py) pueda
    # avisarle al usuario "revisá este día a mano" en vez de guardar un
    # total silenciosamente incompleto.
    subtotal_mismatch = diagnostics.get("subtotal_mismatch")

    # El departamento real del reporte es "LOCAL ACCT" -- la normalización
    # compartida (DEPARTMENT_NAME_NORMALIZATION/_fallback_department_alias)
    # lo renombra a "GETTEL/TOYOTA" porque así se llama la columna real del
    # Excel Cierre donde ese monto se escribe (ver _write_amount_cell) --
    # eso no se toca, sigue igual para la escritura del Excel real.
    #
    # Acá, del lado Carga de Datos, se renombra de vuelta a su nombre real
    # ("LOCAL ACCT") y se deja como una fila más de `records` -- pedido
    # explícito del usuario (2026-09-18): hasta acá se descartaba por
    # completo (no aparecía en el detalle crudo del día ni en "Total del mes
    # por departamento", y sus unidades se perdían) porque no es una venta
    # de mercadería como cualquier otra -- pero eso hacía que "no apareciera
    # al lado de AUTO/BEER/CANDY al editar el día" y que la categoría
    # "Gettel" quedara con el monto bien pero sin ninguna unidad. Ahora se
    # comporta como cualquier departamento real (visible, con sus propias
    # unidades) -- lo único especial que conserva es NO caer en RESTO si no
    # matcheara ninguna categoría (ver DEPARTMENT_GROUPS/group_department_
    # sales más abajo, "LOCAL ACCT" quedó agregado a la categoría "Gettel").
    for record in records:
        if record.get("department") == "GETTEL/TOYOTA":
            record["department"] = "LOCAL ACCT"

    return {
        "date": business_date,
        "records": records,
        "subtotal_mismatch": subtotal_mismatch,
        # Lectura por votación (2026-10-04): lo que quedó vacío por dudoso,
        # el total impreso confirmado y si no hubo total contra qué verificar.
        "doubtful_departments": [
            f"un renglón con el nombre ilegible ({item['values']['count']} unidades, ${item['values']['amount']:,.2f})"
            if item.get("values") else
            f"{item['department']} ({' y '.join('unidades' if f == 'count' else 'monto' for f in item['fields'])})"
            for item in diagnostics.get("doubtful_departments") or []
        ],
        "printed_totals": diagnostics.get("printed_totals"),
        "total_unverified": bool(diagnostics.get("total_unverified")),
    }


# Agrupamiento de departamentos en las 6 categorías de "Resumen Venta"/
# "Reply to Report c-Store" (hojas 2 y 3 de "... Ventas ... ANALISIS.xlsx")
# -- confirmado 2026-09-13 leyendo las fórmulas reales de las DOS hojas con
# openpyxl (se cruzaron entre sí, coinciden exactas) en vez de asumirlo:
# antes el usuario copiaba estos 6 totales a mano, día por día, desde
# "Reply to Report c-Store" hacia Store Info -- pedido explícito: "ahora se
# va a poder hacer de forma automatica con solo cargar un reporte diario".
# "Gettel" en la hoja real (llamada "CAR WASH/ICE" en Resumen Venta, columna
# "VS" en Store Info) sale de dos departamentos reales del Department Sales
# Report: "GETTEL" (rara vez aparece) y "LOCAL ACCT" (esporádico, ver
# extract_department_sales_for_day) -- los dos suman a la misma categoría,
# nunca se reemplazan entre sí. Los grupos siguen el orden real de las
# hojas.
DEPARTMENT_GROUPS = (
    # "MAJ PAK" (con espacio) -- así queda guardado por extract_department_
    # sales_for_day/parse_elistar_daily_pdf_page, aunque el PDF y el Excel
    # real lo impriman/nombren "MAJPAK" sin espacio -- confirmado leyendo el
    # dato ya extraído, no asumido.
    ("TABACCO", ("CIGARS", "GEN-CTN", "GEN-PAK", "MAJ CR", "MAJ PAK", "SNUFF")),
    ("SODA", ("SODA",)),
    ("BEER/WINE", ("BEER/WINE",)),
    ("LOTERY/LOTTO", ("ONLINE", "SKOFF")),
    ("Gettel", ("GETTEL", "LOCAL ACCT")),
    (
        "RESTO",
        ("COFFE", "TAXABLE", "NONTAX", "SNACK", "JUICE", "WATER", "E-GIGARETTE", "CANDY", "VARIOS/BOLSA", "FOUTAIN"),
    ),
)


def group_department_sales(records):
    """
    Agrupa una lista de departamentos (cada uno {"department", "count",
    "amount"} -- una fila de daily_report_departments, de un día o ya
    sumados por mes) en las 6 categorías de arriba.

    Devuelve (groups, unmatched): `groups` en el mismo orden que las hojas
    reales, cada uno {"label", "count", "amount"}. Un departamento real que
    no matchea NINGÚN grupo conocido (ej. AUTO/BOILED PEANUTS/FLOWER(S)/
    HBA/ICECREAM) se suma dentro de RESTO -- pedido explícito del usuario
    (2026-09-14): "estas categorias se suman a RESTO" -- a diferencia del
    Excel real (que tampoco los suma a ninguna de las 6 categorías), acá sí
    se incluyen para no dejar ventas reales fuera de las categorías. `
    unmatched` queda siempre vacío -- se conserva en la firma para no tener
    que tocar los callers/templates que todavía lo reciben.

    "LOCAL ACCT" -- pedido explícito del usuario (2026-09-15, tras ver un
    día real con RESTO disparado a $8.457 por esto, y de nuevo el
    2026-09-18 al notar que perdía sus unidades y no aparecía al editar el
    día): no es una venta de mercadería como BOILED PEANUTS/FLOWER/HBA/
    ICECREAM, es un monto de cuenta/pago (mismo concepto que el campo
    "Local Accounts" de Store Info, ver reportes_db.get_month_local_
    accounts, que alimenta el DIF de Gettel/Toyota) -- sumarlo a RESTO lo
    distorsiona con un cargo puntual grande que no tiene nada que ver con
    ventas reales de otros departamentos. Por eso está listado como
    miembro de la categoría "Gettel" en DEPARTMENT_GROUPS de arriba, en vez
    de dejarlo caer en RESTO como el resto de los departamentos sin
    categoría propia.
    """
    totals_by_department = {}
    for record in records:
        key = (record.get("department") or "").strip().upper()
        if not key:
            continue
        bucket = totals_by_department.setdefault(key, {"count": 0, "amount": 0.0})
        bucket["count"] += record.get("count") or 0
        bucket["amount"] += record.get("amount") or 0.0

    matched_departments = set()
    groups = []
    for label, members in DEPARTMENT_GROUPS:
        count = sum(totals_by_department.get(member, {}).get("count", 0) for member in members)
        amount = sum(totals_by_department.get(member, {}).get("amount", 0.0) for member in members)
        matched_departments.update(member for member in members if member in totals_by_department)
        groups.append({"label": label, "count": count, "amount": round(amount, 2)})

    resto_group = next(g for g in groups if g["label"] == "RESTO")
    for department, values in totals_by_department.items():
        if department in matched_departments:
            continue
        resto_group["count"] += values["count"]
        resto_group["amount"] = round(resto_group["amount"] + values["amount"], 2)

    return groups, []


def real_store_info_total_sales(info, department_detail):
    """
    Total Sales real de Store Info!R: Total Fuel (Sales Fuel + Desc. Comb) +
    Non Fuel + Desc. Otros + Tax Collect - VS, donde VS es la categoría
    "Gettel" de los departamentos del día. El "Total Sales" impreso en el PDF
    no resta VS. Compartido entre Store Info (webapp) y Caja para que las dos
    pantallas muestren lo mismo (auditoría 2026-09, caja.py:253).

    Devuelve (total_sales, gettel_amount); total_sales es None si falta
    algún componente.
    """
    groups, _unmatched = group_department_sales(department_detail or [])
    gettel_amount = next((g["amount"] for g in groups if g["label"] == "Gettel"), 0.0)
    parts = [info.get(k) for k in ("sales_fuel", "desc_comb", "non_fuel_total", "desc_otros", "tax_collect")]
    if None in parts:
        return None, gettel_amount
    return round(sum(parts) - gettel_amount, 2), gettel_amount


def extract_store_info_for_day(pdf_path):
    """
    Wrapper delgado sobre extract_store_info_from_pdf que además aplica el
    mismo +1 día que ya aplica write_store_info_row al escribir la columna
    A/C del Excel -- centralizado acá para que ningún llamador nuevo se
    olvide del ajuste (ya pasó más de una vez en este proyecto).
    """
    fields = extract_store_info_from_pdf(pdf_path, strict=False)
    business_date = fields["from_date"] + timedelta(days=1)
    return {"date": business_date, "fields": fields}


def build_store_info_export_workbook(rows, year, month, dest_path):
    """
    Excel NUEVO (no toca ningún archivo real ni ninguna plantilla) con
    Store Info de un mes ya guardado en reportes_db -- pedido explícito del
    usuario (2026-09-19): replicar el mismo diseño/colores/columnas/orden
    que la hoja "Store Info" real del Excel Cierre (títulos, agrupación de
    colores por bloque, anchos de columna), no un formato genérico propio.
    `rows` es la lista tal cual la arma _build_store_info_rows (webapp.py)
    -- mismos campos ya calculados que se muestran en pantalla, más
    `category_amounts` (Tabacco/SODA/BEER-WINE/LOTTO/VS/Resto, columnas I-N
    del Excel real -- sale de Ventas por Departamento, no de Store Info).

    Columnas I-Q en adelante (A-H el bloque de horario/combustible) llevan
    el mismo texto y color que el Excel real (incluido "Tax collet", así
    tal cual está tipeado en el archivo real -- no se corrige el typo para
    que sea el mismo texto que el usuario ya conoce). Wherever un valor de
    la pantalla es en realidad una fórmula (Total = Sales Fuel + Desc.
    Comb; Total Ventas = Total + Non Fuel + desc otros + Tax collet − VS;
    Total Revenue = Cash + TC + Other + Local Account) queda como fórmula
    real de Excel -- el resto son valores crudos leídos del PDF (o tipeados
    a mano), sin ninguna fórmula que los produzca.
    """
    import openpyxl
    from openpyxl.styles import Alignment as XlAlignment, Border as XlBorder, Font as XlFont, PatternFill, Side as XlSide

    HEADER_GRAY = PatternFill("solid", fgColor="FFD9D9D9")
    HEADER_ORANGE = PatternFill("solid", fgColor="FFFFC000")
    HEADER_YELLOW = PatternFill("solid", fgColor="FFFFFF00")
    HEADER_PEACH = PatternFill("solid", fgColor="FFF8CBAD")
    # Colores de las CELDAS DE DATOS (no del header) -- pedido explícito del
    # usuario (2026-09-19), confirmados columna por columna abriendo el
    # Excel real (Cierre 07-25, hoja Store Info) y resolviendo sus colores
    # de tema (theme+tint) a RGB real: H=gris (mismo que el header),
    # I-N=amarillo (no naranja -- el naranja es solo el header), R=verde
    # claro (Accent6 con tint +0.4), V=amarillo (igual que I-N), W/X=azul
    # grisáceo claro (Text2/dk2 con tint +0.6). La primera versión de este
    # export no coloreaba ninguna celda de datos, solo el header.
    DATA_GREEN = PatternFill("solid", fgColor="FFA9D18E")
    DATA_BLUEGRAY = PatternFill("solid", fgColor="FFADB9CA")
    DATA_FILL_BY_COL = {
        8: HEADER_GRAY,
        9: HEADER_YELLOW,
        10: HEADER_YELLOW,
        11: HEADER_YELLOW,
        12: HEADER_YELLOW,
        13: HEADER_YELLOW,
        14: HEADER_YELLOW,
        18: DATA_GREEN,
        22: HEADER_YELLOW,
        23: DATA_BLUEGRAY,
        24: DATA_BLUEGRAY,
    }
    MONEY_FMT = '"$" #,##0.00'
    VOLUME_FMT = '#,##0.00_ ;[Red]\\-#,##0.00\\ '
    # Bordes finos en TODAS las celdas (header + datos) -- pedido explícito
    # del usuario (2026-09-19), confirmado contra el Excel real (Cierre
    # 07-25, hoja Store Info): tiene bordes en toda la grilla, la primera
    # versión de este export no ponía ninguno.
    THIN_SIDE = XlSide(style="thin", color="FF000000")
    THIN_BORDER = XlBorder(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Store Info"

    # (texto, fill, bold, font_color) -- tal cual el encabezado real, columna por columna.
    header_spec = [
        ("Fecha", HEADER_GRAY, True, None),
        ("hs", HEADER_GRAY, True, None),
        ("Fecha", HEADER_GRAY, True, None),
        ("hs", HEADER_GRAY, True, None),
        ("Volume", HEADER_GRAY, True, None),
        ("SALES FUEL", HEADER_GRAY, True, None),
        ("Desc. Comb", HEADER_GRAY, False, None),
        ("Total", HEADER_GRAY, False, None),
        ("Tabacco", HEADER_ORANGE, True, None),
        ("SODA", HEADER_ORANGE, True, None),
        ("BEER / WINE", HEADER_ORANGE, True, None),
        ("LOTTO", HEADER_ORANGE, True, None),
        ("VS", HEADER_ORANGE, True, None),
        ("Resto", HEADER_ORANGE, True, None),
        ("C-Store=Total Non Fuel", HEADER_ORANGE, True, None),
        ("desc otros", HEADER_ORANGE, False, None),
        ("Tax collet", HEADER_ORANGE, True, "FFFF0000"),
        ("Total Ventas", None, True, None),
        ("Cash", HEADER_YELLOW, True, None),
        ("TC", HEADER_PEACH, True, None),
        ("Other", HEADER_PEACH, True, None),
        ("Local Account", HEADER_PEACH, True, None),
        ("Total Revenue", HEADER_PEACH, True, None),
        ("Network Revenue", HEADER_PEACH, True, None),
    ]
    for col, (text, fill, bold, font_color) in enumerate(header_spec, start=1):
        cell = sheet.cell(row=1, column=col, value=text)
        cell.font = XlFont(bold=bold, color=font_color)
        if fill is not None:
            cell.fill = fill
        cell.alignment = XlAlignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    sheet.row_dimensions[1].height = 44.1
    sheet.freeze_panes = "A2"

    CATEGORY_LABELS = ("TABACCO", "SODA", "BEER/WINE", "LOTERY/LOTTO", "Gettel", "RESTO")

    for row in rows:
        r = sheet.max_row + 1
        business_date = row.get("date")
        if isinstance(business_date, str):
            business_date = datetime.strptime(business_date, "%Y-%m-%d").date()
        col_a_date = datetime(business_date.year, business_date.month, business_date.day) if business_date else None
        col_c_date = col_a_date + timedelta(days=1) if col_a_date else None
        cat = row.get("category_amounts") or {}
        credit_amounts = row.get("credit_terms") or []

        # Un día sin ningún Store Info queda en blanco (no tuvo $0 de ventas,
        # está sin cargar). En un día cargado, un dato que falta va como 0 y
        # las fórmulas se escriben igual -- pedido del usuario (2026-10-06):
        # "es mejor que aparezca como 0 y que luego te diga si hay
        # diferencias antes que no muestre nada". Las diferencias las avisan
        # el control Cierre (datos vacíos) y el cruce con el reporte mensual.
        has_store_info = bool(row.get("store_info_source"))
        has_total_fuel = has_total_ventas = has_total_revenue = has_store_info

        def num(key):
            value = row.get(key)
            return 0.0 if value is None and has_store_info else value

        values = {
            1: col_a_date,
            2: _parse_hhmm(row.get("from_time")),
            3: col_c_date,
            4: _parse_hhmm(row.get("to_time")),
            5: num("volume"),
            6: num("sales_fuel"),
            7: num("desc_comb"),
            8: f"=SUM(F{r}:G{r})" if has_total_fuel else None,
            9: cat.get(CATEGORY_LABELS[0], 0.0),
            10: cat.get(CATEGORY_LABELS[1], 0.0),
            11: cat.get(CATEGORY_LABELS[2], 0.0),
            12: cat.get(CATEGORY_LABELS[3], 0.0),
            13: cat.get(CATEGORY_LABELS[4], 0.0),
            14: cat.get(CATEGORY_LABELS[5], 0.0),
            15: num("non_fuel_total"),
            16: num("desc_otros"),
            17: num("tax_collect"),
            18: f"=+H{r}+O{r}+P{r}+Q{r}-M{r}" if has_total_ventas else None,
            19: num("cash"),
            20: _build_credit_terms_formula(credit_amounts) if (credit_amounts or has_store_info) else None,
            21: num("other_amount"),
            22: num("local_accounts"),
            23: f"=SUM(S{r}:V{r})" if has_total_revenue else None,
            24: num("network_revenue"),
        }
        for col, value in values.items():
            cell = sheet.cell(row=r, column=col, value=value)
            cell.border = THIN_BORDER
            fill = DATA_FILL_BY_COL.get(col)
            if fill is not None:
                cell.fill = fill
            if col in (1, 3):
                cell.number_format = "mm-dd-yy"
            elif col in (2, 4):
                cell.number_format = "h:mm"
            elif col == 5:
                cell.number_format = VOLUME_FMT
            elif col not in (1, 2, 3, 4):
                cell.number_format = MONEY_FMT
            if col in (8, 18, 23):
                cell.font = XlFont(bold=True)

    for col_letter, width in zip(
        "ABCDEFGHIJKLMNOPQRSTUVWX",
        (10.5, 6, 10.3, 5.5, 10.4, 12.1, 9.3, 12.4, 10.9, 10.3, 11.7, 10.9, 11.3,
         10.9, 12.4, 10.1, 10.7, 13.3, 12.6, 13.3, 8.3, 11.4, 14.3, 13.7),
    ):
        sheet.column_dimensions[col_letter].width = width

    workbook.save(dest_path)
    return dest_path


def _fmt_money_pdf(value):
    return "—" if value is None else "{:,.2f}".format(value)


def _fmt_day_month_pdf(value):
    """'01-08' (DD-MM, sin año) -- pedido explícito del usuario (2026-09-19)."""
    if value is None:
        return "—"
    if isinstance(value, str):
        return f"{value[8:10]}-{value[5:7]}" if len(value) == 10 else value
    return value.strftime("%d-%m")


# Campos numéricos de Store Info + su label en el PDF resumen (pedido
# explícito del usuario, 2026-09-17: reportes de Reportes "mas resumido
# y con los totales bien hecho") -- mismo orden y mismos campos que ya
# usa build_store_info_export_pdf fila por fila, hoisteados acá para que
# el resumen (build_store_info_pdf_resumen, más abajo) use exactamente
# los mismos campos/agregación, sin duplicar la lista y arriesgar que se
# desincronicen.
_STORE_INFO_TOTAL_FIELDS = (
    ("volume", "Volume"),
    ("sales_fuel", "Sales Fuel"),
    ("desc_comb", "Desc. Comb"),
    ("total_fuel", "Total Fuel"),
    ("non_fuel_total", "Non Fuel"),
    ("desc_otros", "Desc. Otros"),
    ("tax_collect", "Tax Collect"),
    ("total_sales", "Total Sales"),
    ("cash", "Cash"),
    ("tc", "Tarjeta/Créd."),
    ("local_accounts", "Local Acc."),
    ("other_amount", "Other"),
    ("network_revenue", "Network Rev."),
    ("total_revenue", "Total Rev."),
)


def build_store_info_export_pdf(rows, year, month, dest_path, company_header=False):
    """
    PDF (líneas/bordes + colores, sin fórmulas) con Store Info del mes --
    pedido explícito del usuario (2026-09-19): alternativa liviana al
    Excel (build_store_info_export_workbook), con las mismas columnas que
    ya se ven en /reporte/store-info/historial (sin las 6 categorías de
    respaldo del Non Fuel, que no se muestran ahí tampoco). `rows` es la
    misma lista que arma _build_store_info_rows (webapp.py) -- ya trae
    total_fuel/total_sales calculados. Colores idénticos a los del export
    a Excel (mismos hex, ver build_store_info_export_workbook) -- pedido
    explícito del usuario tras ver la primera versión sin color.

    `company_header` -- opcional, default `False` (sin cambios para el
    botón "Exportar PDF" ya existente de /reporte/store-info/historial).
    `company_header=True` agrega el membrete (logo + nombre de la empresa)
    y una línea de "Período" con la fecha real mínima/máxima de `rows` --
    usado por el módulo nuevo "Reportes" (pedido explícito del usuario,
    2026-09-17: "Con los Reportes diarios tambien"), reutilizando este
    mismo PDF ya validado en vez de duplicar la lógica de columnas.
    """
    from pdf_export import build_simple_table_pdf

    # Campos numéricos por fila, en el mismo orden que las columnas 2-15
    # (Volume en adelante) -- "tc" se calcula aparte (suma de credit_terms,
    # no un campo guardado directo, así que no está en `row` crudo -- no
    # afecta el chequeo de "¿este día tiene algo cargado?" de más abajo,
    # alcanza con que CUALQUIERA de los demás campos esté presente). Se
    # reusa tanto para el "Período" como para cada fila y la fila de
    # totales de abajo.
    numeric_fields = tuple(field for field, _label in _STORE_INFO_TOTAL_FIELDS)

    period_label = None
    if company_header:
        # `rows` trae UN renglón por CADA día del mes calendario, con o sin
        # datos (ver reportes_db.get_month_store_info) -- así que no basta
        # con mirar `row["date"]` (siempre está, aunque el día esté vacío).
        # Se usa la fecha del día con ALGÚN valor real cargado, para que el
        # "Período" refleje hasta dónde llegó la carga de verdad ("hay que
        # aclarar hasta que dia llega el reporte"), no el mes completo.
        dates = sorted(
            row["date"]
            for row in rows
            if row.get("date") and any(row.get(field) is not None for field in numeric_fields)
        )
        if dates:
            start_d = datetime.strptime(dates[0], "%Y-%m-%d")
            end_d = datetime.strptime(dates[-1], "%Y-%m-%d")
            period_label = f"Período: {start_d.strftime('%d/%m/%Y')} al {end_d.strftime('%d/%m/%Y')}"
        else:
            period_label = f"Período: sin días cargados todavía en {month:02d}/{year}"

    GRAY, ORANGE, YELLOW, PEACH = "#D9D9D9", "#FFC000", "#FFFF00", "#F8CBAD"
    GREEN_LIGHT, BLUEGRAY = "#A9D18E", "#ADB9CA"
    header_fill_by_col = {
        0: GRAY, 1: GRAY, 2: GRAY, 3: GRAY, 4: GRAY, 5: GRAY,
        6: ORANGE, 7: ORANGE, 8: ORANGE,
        10: YELLOW,
        11: PEACH, 12: PEACH, 13: PEACH, 14: PEACH, 15: PEACH,
    }
    data_fill_by_col = {5: GRAY, 9: GREEN_LIGHT, 12: YELLOW, 14: BLUEGRAY, 15: BLUEGRAY}

    headers = [
        "Día", "Hora", "Volume", "Sales Fuel", "Desc. Comb", "Total Fuel",
        "Non Fuel", "Desc. Otros", "Tax Collect", "Total Sales", "Cash",
        "Tarjeta/Créd.", "Local Acc.", "Other", "Network Rev.", "Total Rev.",
    ]
    table_rows = []
    totals = {field: 0.0 for field in numeric_fields}
    any_value = {field: False for field in numeric_fields}
    for row in rows:
        credit_terms = row.get("credit_terms") or []
        values = dict(row)
        values["tc"] = round(sum(credit_terms), 2) if credit_terms else None
        for field in numeric_fields:
            if values.get(field) is not None:
                totals[field] += values[field]
                any_value[field] = True
        table_rows.append(
            [
                _fmt_day_month_pdf(row.get("date")),
                f"{row.get('from_time') or '—'}–{row.get('to_time') or '—'}",
                *[_fmt_money_pdf(values.get(field)) for field in numeric_fields],
            ]
        )
    # Fila de totales -- pedido explícito del usuario (2026-09-19): "no
    # tenemos una fila extra despues del ultimo dia... deberia agregarse
    # eso" (mismo criterio que ya tiene el PDF de Caja).
    table_rows.append(
        [
            "Total", "",
            *[
                _fmt_money_pdf(round(totals[field], 2)) if any_value[field] else "—"
                for field in numeric_fields
            ],
        ]
    )
    col_widths_mm = [19, 24, 19, 22, 19, 19, 19, 19, 19, 22, 19, 22, 19, 17, 22, 22]
    return build_simple_table_pdf(
        dest_path,
        f"Store Info — {month:02d}/{year}",
        headers,
        table_rows,
        col_widths_mm,
        header_fill_by_col=header_fill_by_col,
        data_fill_by_col=data_fill_by_col,
        bold_last_row=True,
        company_header=company_header,
        period_label=period_label,
    )


def build_store_info_pdf_resumen(rows, year, month, dest_path):
    """
    PDF resumido de Store Info para el módulo "Reportes" -- pedido
    explícito del usuario (2026-09-17, sesión siguiente): "quiero que
    pongas los pdf de reportes en reporte diario y lottery como los
    otros dos, mas resumido y con los totales bien hecho". A diferencia
    de build_store_info_export_pdf (una fila por CADA día del mes, 16
    columnas -- la que sigue usando el botón "Exportar PDF" ya existente
    de /reporte/store-info/historial, sin tocar), acá se arma una sola
    tabla de dos columnas (Detalle/Total), una fila por cada campo de
    _STORE_INFO_TOTAL_FIELDS, igual de "resumido" que el reporte de Chase
    (chase_rules.build_chase_pdf_report).

    Los totales son una SUMA simple de cada campo a lo largo de los días
    del mes que tengan ese dato cargado -- correcto acá porque los 14
    campos de Store Info son todos importes/cantidades del día (ventas,
    volumen, impuestos, etc.), nunca un saldo corrido -- a diferencia de
    Lottery (ver build_lottery_pdf_resumen en lottery_db.py), Store Info
    no tiene ninguna columna "snapshot" que haya que excluir de la suma.
    """
    from pdf_export import build_simple_table_pdf

    numeric_fields = tuple(field for field, _label in _STORE_INFO_TOTAL_FIELDS)

    # Período real (mismo criterio que build_store_info_export_pdf con
    # company_header=True): la fecha del día con ALGÚN valor cargado, no
    # el mes calendario completo (rows trae un renglón por cada día del
    # mes exista o no dato, ver reportes_db.get_month_store_info).
    dates = sorted(
        row["date"]
        for row in rows
        if row.get("date") and any(row.get(field) is not None for field in numeric_fields)
    )
    if dates:
        start_d = datetime.strptime(dates[0], "%Y-%m-%d")
        end_d = datetime.strptime(dates[-1], "%Y-%m-%d")
        period_label = f"Período: {start_d.strftime('%d/%m/%Y')} al {end_d.strftime('%d/%m/%Y')}"
    else:
        period_label = f"Período: sin días cargados todavía en {month:02d}/{year}"

    totals = {field: 0.0 for field in numeric_fields}
    any_value = {field: False for field in numeric_fields}
    for row in rows:
        credit_terms = row.get("credit_terms") or []
        values = dict(row)
        values["tc"] = round(sum(credit_terms), 2) if credit_terms else None
        for field in numeric_fields:
            if values.get(field) is not None:
                totals[field] += values[field]
                any_value[field] = True

    table_rows = [
        [label, _fmt_money_pdf(round(totals[field], 2)) if any_value[field] else "—"]
        for field, label in _STORE_INFO_TOTAL_FIELDS
    ]

    return build_simple_table_pdf(
        dest_path,
        f"Store Info — Resumen — {month:02d}/{year}",
        ["Detalle", "Total"],
        table_rows,
        col_widths_mm=[110, 80],
        bold_last_row=True,
        company_header=True,
        period_label=period_label,
    )


def _parse_hhmm(value):
    """'22:43' (24h, tal cual reportes_db._time_str la guarda) -> datetime.time, o None."""
    if not value:
        return None
    try:
        return datetime.strptime(value, "%H:%M").time()
    except ValueError:
        return None


def _format_date_ddmmyyyy(value):
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%d-%m-%Y")
    text = str(value)
    if len(text) == 10 and text[4] == "-" and text[7] == "-":
        return f"{text[8:10]}-{text[5:7]}-{text[0:4]}"
    return text


def build_department_column_map(sheet):
    """
    Scan row 3 from column C outward and map department labels to (count_col, amount_col).

    Each department header marks a COUNT | NET SALES column pair (first = Net Count,
    second = Net Sales $). Position-independent — keyed by normalized department label.
    """
    mapping = {}
    protected = set()
    col = HEADER_START_COLUMN
    max_col = max(sheet.max_column, HEADER_START_COLUMN)
    sub_header_row = HEADER_ROW + 1

    while col <= max_col:
        header_value = sheet.cell(row=HEADER_ROW, column=col).value
        label = _strip_cell(header_value)
        if not label or _is_sub_header(label):
            col += 1
            continue

        norm = _normalize_department_label(label)
        count_col = col
        amount_col = col + 1

        sub_count = _normalize_department_label(
            sheet.cell(row=sub_header_row, column=count_col).value
        )
        sub_amount = _normalize_department_label(
            sheet.cell(row=sub_header_row, column=amount_col).value
        )
        if sub_count and sub_amount:
            first_is_sales = "net" in sub_count and "sales" in sub_count
            second_is_count = "count" in sub_amount
            if first_is_sales and second_is_count:
                count_col, amount_col = amount_col, count_col

        if _is_protected_department(label):
            protected.add(norm)
            col += 2
            continue

        mapping[norm] = (count_col, amount_col)
        mapping[_canonical_department_key(label)] = (count_col, amount_col)

        canonical = _canonical_department_key(label)
        if canonical and canonical not in mapping:
            mapping[canonical] = (count_col, amount_col)

        col += 2

    if not mapping:
        raise ValueError(
            f"No se encontraron encabezados de departamento en la fila {HEADER_ROW} desde la columna C."
        )

    return mapping, protected


def _resolve_department_columns(dept_name, column_map):
    key = _normalize_department_label(dept_name)
    if key in column_map:
        return column_map[key]

    canon = _canonical_department_key(dept_name)
    if canon in column_map:
        return column_map[canon]

    for header_key, coords in column_map.items():
        if _department_keys_match(key, header_key):
            return coords
    forced = _fallback_department_alias(dept_name)
    forced_key = _normalize_department_label(forced)
    if forced_key in column_map:
        return column_map[forced_key]
    normalized_header_keys = [k for k in column_map.keys() if k]
    if normalized_header_keys:
        fuzzy = get_close_matches(forced_key or key, normalized_header_keys, n=1, cutoff=0.6)
        if fuzzy:
            return column_map[fuzzy[0]]
    return None


def _department_label_keys(dept_name):
    """Normalized keys used to match worksheet / PDF department labels."""
    keys = set()
    key = _normalize_department_label(dept_name)
    if key:
        keys.add(key)
    canon = _canonical_department_key(dept_name)
    if canon:
        keys.add(canon)
    return keys


def _apply_sales_font_alert(cell):
    """Bright-red font only — never alters cell background fill."""
    if Font is None:
        return
    base = cell.font
    cell.font = Font(
        color=SALES_ALERT_FONT_COLOR,
        name=base.name,
        size=base.size,
    )


def _write_count_cell(sheet, row, column, value):
    cell = sheet.cell(row=row, column=column, value=int(value))
    cell.number_format = "0"
    if COUNT_CELL_ALIGNMENT is not None:
        cell.alignment = COUNT_CELL_ALIGNMENT


def _sanitize_sales_float(raw_value):
    """
    Convert parsed Net Sales token into a native float safely.

    Removes currency symbols/grouping separators and defaults to 0.00 on failure.
    """
    token = _strip_cell(raw_value)
    token = token.replace("$", "").replace(",", "").strip()
    try:
        return float(token)
    except (TypeError, ValueError):
        return 0.00


def _write_amount_cell(sheet, row, column, value, department=None):
    amount = _sanitize_sales_float(value)
    if amount.is_integer():
        cell = sheet.cell(row=row, column=column, value=int(amount))
    else:
        cell = sheet.cell(row=row, column=column, value=amount)
    cell.number_format = "0.00"

    if department is not None:
        dept_label = re.sub(r"\s+", " ", _strip_cell(department)).strip().upper()
        if dept_label == "LOCAL ACCT":
            dept_label = "GETTEL/TOYOTA"

        if dept_label == "GETTEL/TOYOTA":
            if amount >= GETTEL_TOYOTA_THRESHOLD:
                _apply_sales_font_alert(cell)
            return

        if dept_label in GROUP_1400_DEPARTMENTS:
            if amount >= GROUP_1400_THRESHOLD:
                _apply_sales_font_alert(cell)
            return

        if dept_label in GROUP_500_DEPARTMENTS:
            if amount >= GROUP_500_THRESHOLD:
                _apply_sales_font_alert(cell)
            return


def inject_daily_sales(sheet, pdf_records, column_map, target_row):
    """
    Write parsed PDF Net Count / Net Sales $ into the mapped COUNT | NET SALES columns.

    Skips protected/formula-backed columns (e.g. GIFT CARD, VARIOS/BOLSA).
    Only writes to target_row — never clears or shifts other rows.
    """
    target_row = int(target_row)
    written = []
    skipped = []

    for record in pdf_records:
        dept_name = record["department"]
        if _is_protected_department(dept_name):
            skipped.append(dept_name)
            continue

        coords = _resolve_department_columns(dept_name, column_map)
        if coords is None:
            skipped.append(dept_name)
            continue

        count_col, amount_col = coords
        count_cell = sheet.cell(row=target_row, column=count_col)
        amount_cell = sheet.cell(row=target_row, column=amount_col)

        if _cell_has_formula(count_cell) or _cell_has_formula(amount_cell):
            skipped.append(dept_name)
            continue

        if record["count"] is None or record["amount"] is None:
            # El OCR no lo pudo leer con seguridad: la celda queda como está.
            skipped.append(dept_name)
            continue
        amount_value = _sanitize_sales_float(record["amount"])
        _write_count_cell(sheet, target_row, count_col, int(record["count"]))
        _write_amount_cell(
            sheet,
            target_row,
            amount_col,
            amount_value,
            department=dept_name,
        )
        written.append(
            {
                "department": dept_name,
                "count_col": get_column_letter(count_col),
                "amount_col": get_column_letter(amount_col),
                "count": int(record["count"]),
                "amount": float(record["amount"]),
            }
        )

    return written, skipped


def _read_manual_report_count(sheet, row):
    """
    BS — el total de unidades impreso en el reporte, cargado a mano por el
    usuario (BR, al lado, es una fórmula que suma lo ya cargado en la fila
    y no se toca acá).

    Un día real nunca imprime 0 unidades vendidas, así que un 0 literal se
    trata igual que una celda vacía (todavía no se cargó el dato a mano).
    Una fórmula sin calcular también se ignora, por las dudas.
    """
    raw = sheet.cell(row=row, column=MANUAL_REPORT_COUNT_COLUMN).value
    if raw is None or (isinstance(raw, str) and (not raw.strip() or raw.strip().startswith("="))):
        return None
    return _safe_parse_count(raw) or None


def check_loaded_count_against_manual_entry(sheet, target_row, column_map):
    """
    Compare BS (cargado a mano desde el reporte impreso) contra la suma de
    Net Count ya presente en todas las columnas de departamento de esa fila.

    Devuelve None si BS está vacío (todavía no se cargó) o si ambos totales
    coinciden; si no, un dict con los dos totales para armar el aviso.
    """
    manual_count = _read_manual_report_count(sheet, target_row)
    if manual_count is None:
        return None

    seen_columns = set()
    loaded_count = 0
    for count_col, _amount_col in column_map.values():
        if count_col in seen_columns:
            continue
        seen_columns.add(count_col)
        loaded_count += _safe_parse_count(sheet.cell(row=target_row, column=count_col).value)

    if loaded_count == manual_count:
        return None
    return {"manual_count": manual_count, "loaded_count": loaded_count}


def _normalize_pdf_paths(pdf_paths):
    if pdf_paths is None:
        return []
    if isinstance(pdf_paths, str):
        text = pdf_paths.strip()
        if not text:
            return []
        for separator in (";", "|"):
            if separator in text:
                return [
                    os.path.abspath(part.strip())
                    for part in text.split(separator)
                    if part.strip()
                ]
        return [os.path.abspath(text)]
    return [
        os.path.abspath(str(path).strip())
        for path in pdf_paths
        if path is not None and str(path).strip()
    ]


_MAX_CONCURRENT_PDF_WORKERS = 4


def _parse_pdfs_concurrently(paths, parse_fn, progress_callback=None, **kwargs):
    """
    Run parse_fn(path, **kwargs) for every path, in parallel when there's
    more than one.

    Each call ultimately shells out to Tesseract as its own OS process, so
    this isn't fighting Python's GIL — a batch of daily PDFs really can be
    read at the same time on a multi-core machine instead of strictly one
    after another, with no change to how any single PDF is read. Capped at
    a handful of workers so a big batch doesn't overwhelm the machine.

    `progress_callback(done, total)`, si viene, se llama cada vez que un PDF
    más termina de leerse (éxito o error) -- el paso lento de un lote (ver
    docstring de arriba), así que es la señal más representativa de avance
    real para un job en segundo plano (ver jobs.py/CLAUDE.md, "progreso
    real"). Opcional, sin efecto si no se pasa.

    Returns (results, errors) -- two dicts keyed by path -- instead of
    raising on the first bad PDF. A single unreadable/unrecognizable PDF
    used to abort the entire batch (losing the already-good work of every
    other PDF) and its error text embedded the failing filename, which
    violates the sitewide "never a filename in an aviso" policy. Now every
    path gets its own outcome, and it's up to the caller to isolate a
    failed path into a short, count-only notice instead of a raw message.
    """

    def _run(path):
        try:
            return ("ok", parse_fn(path, **kwargs))
        except Exception as exc:
            return ("error", exc)

    total = len(paths)
    if total <= 1:
        outcomes = {}
        for path in paths:
            outcomes[path] = _run(path)
            if progress_callback:
                progress_callback(len(outcomes), total)
    else:
        max_workers = min(len(paths), os.cpu_count() or _MAX_CONCURRENT_PDF_WORKERS, _MAX_CONCURRENT_PDF_WORKERS)
        outcomes = {}
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(_run, path): path for path in paths}
            for future in as_completed(futures):
                outcomes[futures[future]] = future.result()
                if progress_callback:
                    progress_callback(len(outcomes), total)

    results = {}
    errors = {}
    for path, (status, payload) in outcomes.items():
        if status == "ok":
            results[path] = payload
        else:
            errors[path] = payload
    return results, errors


def _partition_paths_by_extractable_day(paths):
    """
    Split `paths` into (sorted_by_day, unrecognized_count) -- a PDF whose
    filename doesn't match the expected "... DD-MM.pdf" pattern used to
    abort sorting (and therefore the whole batch) for every other PDF too,
    via the ValueError from extract_day_from_filename propagating out of
    sorted()'s key function. Now it's isolated: that one file is counted
    and excluded, the rest still get sorted and processed normally.
    """
    dated_paths = []
    unrecognized = 0
    for path in paths:
        try:
            day = extract_day_from_filename(path)
        except ValueError:
            unrecognized += 1
            continue
        dated_paths.append((day, path))
    dated_paths.sort(key=lambda item: item[0])
    return [path for _day, path in dated_paths], unrecognized


def process_reporte_diario(
    master_path, pdf_paths, page_index=DEFAULT_PDF_PAGE_INDEX, progress_callback=None
):
    """
    Parse one or more daily PDFs and inject each into its calendar day row.

    Day-of-month is read from each PDF filename (e.g. Close Store 07-05.pdf -> day 7).
    Each PDF is read (OCR included) in parallel across a small worker pool,
    since that read — not the Excel write that follows — is what makes a
    multi-PDF batch slow. Workbook is saved once after the full batch completes.

    Returns:
        tuple: (temp_path, summary dict)
    """
    _ensure_openpyxl()
    master_path = os.path.abspath(str(master_path).strip())
    paths = _normalize_pdf_paths(pdf_paths)

    if not os.path.isfile(master_path):
        raise FileNotFoundError(f"Excel maestro no encontrado: {master_path}")
    if not paths:
        raise ValueError("No se proporcionaron PDF diarios.")

    extension = os.path.splitext(master_path)[1].lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise ValueError("El Excel maestro debe ser .xlsx o .xlsm.")

    # Aislado por archivo: antes, un solo PDF con un nombre que no matchea
    # "... DD-MM.pdf" tiraba el sorted() de TODO el lote (ValueError de
    # extract_day_from_filename propagándose desde la key function),
    # perdiendo el trabajo de los demás PDFs que sí estaban bien nombrados.
    sorted_paths, files_with_bad_filename = _partition_paths_by_extractable_day(paths)
    if not sorted_paths:
        raise ValueError(
            "No se pudo extraer el día del mes del nombre de ningún PDF -- se esperaba "
            "un patrón como 'Close Store 01-05.pdf'."
        )

    # Aislado por archivo también acá: antes, un PDF ilegible (OCR sin
    # ancla, escaneo corrupto, etc.) tiraba TODO el lote via
    # future.result() sin capturar, perdiendo el trabajo ya bueno de los
    # demás PDFs -- y el mensaje de esa excepción venía con el nombre de
    # archivo incrustado (violaba la política de "nunca un nombre de
    # archivo en un aviso"). Ahora cada PDF que falla se cuenta aparte y el
    # resto sigue procesándose normal.
    parsed_by_path, parse_errors = _parse_pdfs_concurrently(
        sorted_paths, parse_elistar_daily_pdf_page, progress_callback=progress_callback, page_index=page_index
    )
    files_failed_to_parse = len(parse_errors)
    ok_paths = [path for path in sorted_paths if path in parsed_by_path]

    keep_vba = extension == ".xlsm"
    workbook = load_workbook(master_path, data_only=False, keep_vba=keep_vba)
    sheet = _get_carga_aqui_sheet(workbook)
    column_map, protected = build_department_column_map(sheet)

    batch_results = []
    days_failed_to_write = 0
    total_written = 0
    total_skipped = 0
    total_departments = 0

    for pdf_path in ok_paths:
        pdf_records, pdf_diagnostics = parsed_by_path[pdf_path]
        try:
            target_day = int(extract_day_from_filename(pdf_path))
            target_row = find_row_for_calendar_day(sheet, target_day)
        except (ValueError, TypeError) as exc:
            days_failed_to_write += 1
            continue

        try:
            written, skipped = inject_daily_sales(
                sheet, pdf_records, column_map, target_row
            )
        except (ValueError, TypeError, AttributeError) as exc:
            days_failed_to_write += 1
            continue

        total_written += len(written)
        total_skipped += len(skipped)
        total_departments += len(pdf_records)

        last_department = pdf_diagnostics.get("last_department")
        warnings = []
        if pdf_diagnostics.get("used_ocr") and last_department not in (
            None,
        ) + LAST_DEPARTMENT_CANDIDATES:
            warnings.append(
                f"El último departamento leído fue \"{last_department}\", no "
                f"{'/'.join(LAST_DEPARTMENT_CANDIDATES)} — revise si la tabla se cortó."
            )
        mismatch = pdf_diagnostics.get("subtotal_mismatch")
        if mismatch:
            warnings.append(
                "El total impreso en el PDF no coincide con lo leído: "
                f"{mismatch['computed_count']} vs {mismatch['printed_count']} unidades, "
                f"${mismatch['computed_amount']:.2f} vs ${mismatch['printed_amount']:.2f} — "
                "revise el OCR."
            )
        manual_mismatch = check_loaded_count_against_manual_entry(
            sheet, target_row, column_map
        )
        if manual_mismatch:
            warnings.append(
                f"La columna BS dice {manual_mismatch['manual_count']} unidades, pero "
                f"en la fila hay {manual_mismatch['loaded_count']} cargadas — revise "
                "los departamentos."
            )
        warning = " | ".join(warnings) if warnings else None

        batch_results.append(
            {
                "pdf_path": pdf_path,
                "filename": os.path.basename(pdf_path),
                "calendar_day": target_day,
                "target_row": target_row,
                "departments_written": len(written),
                "departments_skipped": len(skipped),
                "written": list(written),
                "skipped": list(skipped),
                "used_ocr": bool(pdf_diagnostics.get("used_ocr")),
                "pages_used": pdf_diagnostics.get("pages_used"),
                "warning": warning,
            }
        )
        pdf_records = None

    if not batch_results:
        workbook.close()
        raise ValueError(
            "No se pudo cargar ningún día: "
            f"{files_with_bad_filename} archivo(s) con nombre no reconocido, "
            f"{files_failed_to_parse} archivo(s) no se pudieron leer, "
            f"{days_failed_to_write} día(s) no matchearon ninguna fila."
        )

    temp_path = _create_temp_workbook_path()
    workbook.save(os.path.abspath(temp_path))
    workbook.close()

    summary = {
        "files_processed": len(batch_results),
        "departments_written": total_written,
        "departments_skipped": total_skipped,
        "pdf_departments": total_departments,
        "protected_headers": sorted(protected),
        "files_with_bad_filename": files_with_bad_filename,
        "files_failed_to_parse": files_failed_to_parse,
        "days_failed_to_write": days_failed_to_write,
        "batch_results": batch_results,
    }
    return temp_path, summary


# ---------------------------------------------------------------------------
# Store Info — a second extraction from the same daily PDF (pages 3
# and the start of 4, up to "Network Revenue") into a different workbook's
# "Store Info" sheet: one summary row per day, appended after the last one.
# ---------------------------------------------------------------------------

STORE_INFO_SHEET_NAME = "Store Info"
STORE_INFO_DATA_START_ROW = 2
DEFAULT_STORE_INFO_PAGE_INDEX = 2  # Third page (0-based) — "PERIOD FROM:" anchor
STORE_INFO_MAX_CONTINUATION_PAGES = 3  # Anchor page + up to 2 more, to reach Network Revenue

PERIOD_FROM_ANCHOR = "period from"
NETWORK_REVENUE_ANCHOR = "network revenue"

# 1-based column numbers for the fields this extraction is allowed to touch.
# Everything else (H, I..N, R, W) is either a formula or filled in some
# other way and must never be written here. U (Other) USED to be in that
# "never touch" group too -- ver STORE_INFO_COL_OTHER más abajo, agregado
# 2026-09-12 a pedido explícito del usuario.
STORE_INFO_COL_FROM_DATE = 1  # A — Fecha (FROM date + 1 day)
STORE_INFO_COL_FROM_TIME = 2  # B — hs (FROM time, unchanged)
STORE_INFO_COL_TO_DATE = 3  # C — Fecha (FROM date + 2 days)
STORE_INFO_COL_TO_TIME = 4  # D — hs (TO time, unchanged)
STORE_INFO_COL_VOLUME = 5  # E — Volume
STORE_INFO_COL_SALES_FUEL = 6  # F — SALES FUEL
STORE_INFO_COL_DESC_COMB = 7  # G — Desc. Comb
STORE_INFO_COL_NON_FUEL = 15  # O — C-Store=Total Non Fuel
STORE_INFO_COL_DESC_OTROS = 16  # P — desc otros
STORE_INFO_COL_TAX_COLLECT = 17  # Q — Tax collet
STORE_INFO_COL_CASH = 19  # S — Cash
STORE_INFO_COL_CREDIT = 20  # T — TC (written as a "=a+b+c" formula, like the sheet's own history)
STORE_INFO_COL_OTHER = 21  # U — Other (fila "Other" del Method of Payment Totals, después de LOCAL ACCOUNTS)
STORE_INFO_COL_LOCAL_ACCOUNTS = 22  # V — Local Account
STORE_INFO_COL_NETWORK_REVENUE = 24  # X — Network Revenue

_MONTH_NAME_TO_NUMBER = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

_PERIOD_FROM_TO_RE = re.compile(
    r"([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})\s+(\d{1,2}):(\d{2})\s*([AP])\.?M\.?"
    r".*?TO:?\s*([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})\s+(\d{1,2}):(\d{2})\s*([AP])\.?M\.?",
    re.IGNORECASE,
)

# A label line is "words... <first numeric-looking token> ...": everything
# before that split point is the label, everything from it on are the values.
_NUMERIC_TOKEN_RE = re.compile(r"^[-~(]*\$?\(?-?[\d,]+\.?\d*\)?%?$")

_STORE_INFO_SCORING_ANCHORS = (
    "period from",
    "total fuel sales",
    "fuel discounts",
    "total non fuel sales",
    "other discounts",
    "total taxes collected",
    "local accounts",
    "network revenue",
)


def _to_24h_time(hour12, minute, ampm):
    hour12 = int(hour12) % 12
    if ampm.upper().startswith("P"):
        return time(hour12 + 12, int(minute))
    return time(hour12, int(minute))


def _parse_period_from_to_line(line):
    """
    Parse "PERIOD FROM: Aug 18, 2026 10:51 PM TO: Aug 19, 2026 10:42 PM".

    Returns a dict with from_date/from_time/to_date/to_time, or None.
    """
    match = _PERIOD_FROM_TO_RE.search(line)
    if not match:
        return None
    (
        from_mon, from_day, from_year, from_hh, from_mm, from_ampm,
        to_mon, to_day, to_year, to_hh, to_mm, to_ampm,
    ) = match.groups()
    from_month = _MONTH_NAME_TO_NUMBER.get(from_mon[:3].lower())
    to_month = _MONTH_NAME_TO_NUMBER.get(to_mon[:3].lower())
    if not from_month or not to_month:
        return None
    try:
        from_date_value = date(int(from_year), from_month, int(from_day))
        to_date_value = date(int(to_year), to_month, int(to_day))
    except ValueError:
        return None
    return {
        "from_date": from_date_value,
        "from_time": _to_24h_time(from_hh, from_mm, from_ampm),
        "to_date": to_date_value,
        "to_time": _to_24h_time(to_hh, to_mm, to_ampm),
    }


def _split_label_and_values(line):
    """
    Split "Total Fuel Sales 1,236.070 $5,152.10" into label + value tokens.

    Strips stray punctuation off either edge of the label (e.g. OCR
    sometimes reads a faint column rule next to "Cash" as "Cash :" or a
    graphic behind the text as "Cash |)" -- trailing junk; a rule just
    *before* the label can just as easily read as "| Network Revenue" --
    leading junk) so it still matches the real label exactly, without
    loosening the match enough to also catch a longer label like "Cash
    Acceptor Cash". Value tokens are filtered down to ones that actually
    look numeric, so a trailing OCR artifact (e.g. "Total Taxes Collected
    $157.89 ;") doesn't get picked up as the amount.
    """
    tokens = [t for t in _normalize_department_spacing(_strip_cell(line)).split() if t]
    for index, token in enumerate(tokens):
        if _NUMERIC_TOKEN_RE.match(token) and index > 0:
            label = " ".join(tokens[:index])
            label = re.sub(r"[^A-Za-z]+$", "", label)
            label = re.sub(r"^[^A-Za-z]+", "", label).strip()
            values = [t for t in tokens[index:] if _NUMERIC_TOKEN_RE.match(t)]
            return label, values
    return None, []


def _sanitize_store_info_float(raw_value):
    """Like _sanitize_sales_float, but also treats a leading '~' as a minus sign (OCR glyph noise)."""
    text = _strip_cell(raw_value)
    # El OCR ensucia el signo: "--$92.82", "—$65.24", "~$5.00", "($92.82)".
    # Todo eso es un solo negativo. Antes cualquier token así caía en 0.00
    # sin avisar (auditoría 2026-09, caso real Desc. Comb del 12/09).
    text = text.replace("—", "-").replace("–", "-").replace("−", "-")
    negative = False
    stripped = text.lstrip("-~ ")
    if stripped != text:
        negative = True
        text = stripped
    if text.startswith("(") or text.startswith("$("):
        negative = True
    text = text.replace("(", "").replace(")", "").replace("$", "").replace(",", "").strip()
    try:
        value = float(text)
    except ValueError:
        raise ValueError(f'Monto ilegible en Store Info: "{raw_value}".')
    return -abs(value) if negative else value


def _find_label_values(lines, *target_labels):
    """Return the value tokens of the first line whose label exactly matches one of target_labels."""
    targets = {label.lower() for label in target_labels}
    for line in lines:
        label, values = _split_label_and_values(line)
        if label and label.lower() in targets:
            return values
    return None


def _score_store_info_page_text(text):
    lower = text.lower() if text else ""
    return sum(1 for anchor in _STORE_INFO_SCORING_ANCHORS if anchor in lower)


def _ocr_store_info_page_text(image):
    """Best-effort OCR of a Store Info page — scored by how many known anchors it recovers."""
    if image is None:
        return ""
    _ensure_pytesseract()
    best_text = ""
    best_score = -1
    for config in _OCR_TEXT_CONFIGS:
        try:
            text = pytesseract.image_to_string(image, config=config) or ""
        except Exception:
            continue
        score = _score_store_info_page_text(text)
        if score > best_score:
            best_score = score
            best_text = text
    return best_text


def _force_positive(value):
    return abs(value)


def _force_negative(value):
    return -abs(value)


def _extract_store_info_fields(lines, require_period=True):
    """
    Pull every Store Info value out of the OCR'd lines of pages 3(-4).

    Every field's sign is fixed by its known business meaning rather than
    trusted from the OCR'd "-" glyph, which is one of the easiest characters
    for Tesseract to drop or invent on a blurry photo: Desc. Comb (G) and
    desc otros (P) are always a discount (negative); every other dollar
    figure and the fuel Volume are always positive.
    """
    period = None
    for line in lines:
        if PERIOD_FROM_ANCHOR in line.lower():
            period = _parse_period_from_to_line(line)
            if period:
                break
    if period is None and require_period:
        raise ValueError('No se encontró la línea "PERIOD FROM: ... TO: ..." en el PDF.')
    # require_period=False: una hoja suelta del reporte mensual, pedida de
    # nuevo porque salió borrosa (reporte_mensual.extract_replacement_sheet);
    # la continuación de Store Info no imprime el período.
    period = period or {"from_date": None, "from_time": None, "to_date": None, "to_time": None}

    # Cada campo se lee por separado (2026-10-02, pedido del usuario: "lo que
    # se pueda cargar de forma automática bien, y si está medio borroso y no
    # se sabe bien lo que se está cargando, no lo carga y listo"): un campo
    # que no se encuentra o no se puede leer queda en None y se anota en
    # "missing_fields" en vez de tirar todo el día. Solo el período (la
    # fecha) es obligatorio. Antes, un "Total Sales" borroso (caso real
    # 04/08) dejaba el día sin Store Info, sin departamentos y sin Lottery,
    # porque los tres sacan la fecha de acá.
    missing = []

    def read(label, reader):
        try:
            return reader()
        except ValueError:
            missing.append(label)
            return None

    def required_last(label, message):
        def reader():
            values = _find_label_values(lines, label)
            if not values:
                raise ValueError(message)
            return _force_positive(_sanitize_store_info_float(values[-1]))
        return reader

    def optional_last(label, sign):
        # Ausente = 0.0 (el día no tuvo ese concepto); presente pero ilegible = None.
        def reader():
            values = _find_label_values(lines, label)
            return sign(_sanitize_store_info_float(values[-1])) if values else 0.0
        return reader

    def fuel_reader():
        fuel_values = _find_label_values(lines, "Total Fuel Sales")
        if not fuel_values or len(fuel_values) < 2:
            raise ValueError('No se encontró "Total Fuel Sales" con Volume y Sales.')
        return (
            _force_positive(_sanitize_store_info_float(fuel_values[0])),
            _force_positive(_sanitize_store_info_float(fuel_values[-1])),
        )

    fuel = read("Total Fuel Sales", fuel_reader)
    volume, sales_fuel = fuel if fuel is not None else (None, None)
    desc_comb = read("Fuel Discounts", optional_last("Fuel Discounts", _force_negative))
    non_fuel_total = read("Total Non Fuel Sales", required_last("Total Non Fuel Sales", 'No se encontró "Total Non Fuel Sales".'))
    desc_otros = read("Other Discounts", optional_last("Other Discounts", _force_negative))
    tax_collect = read("Total Taxes Collected", required_last("Total Taxes Collected", 'No se encontró "Total Taxes Collected".'))

    # "Total Sales" -- la misma cifra que Store Info!R ("Total Ventas")
    # recalcula con una fórmula (Total Fuel + Non Fuel + Desc Otros + Tax
    # Collect - VS) -- se lee directo del PDF en vez de reconstruir esa
    # fórmula acá (VS no es un campo que este parser capture).
    total_sales = read("Total Sales", required_last("Total Sales", 'No se encontró "Total Sales".'))
    cash = read("Cash", required_last("Cash", 'No se encontró la fila "Cash" bajo Method of Payment Totals.'))

    # Every payment-method row strictly between "Cash" and "LOCAL ACCOUNTS"
    # (Credit, Crind CREDIT/DEBIT, CRIND P97, Debit, etc.) gets summed into
    # the credit-card total — zero-valued rows are dropped, same as the
    # sheet's own historical "=a+b+c" formulas. Sin "Cash" o sin "LOCAL
    # ACCOUNTS" el rango no se sabe dónde empieza o termina: las dos cosas
    # quedan sin leer, nunca una suma a medias.
    def payments_reader():
        terms = []
        local = None
        in_range = False
        saw_cash = False
        found_local = False
        for line in lines:
            label, values = _split_label_and_values(line)
            if label is None:
                continue
            norm_label = label.lower()
            if norm_label == "cash":
                in_range = True
                saw_cash = True
                continue
            if norm_label == "local accounts":
                if values:
                    local = _force_positive(_sanitize_store_info_float(values[-1]))
                    found_local = True
                break
            if in_range and values:
                amount = _force_positive(_sanitize_store_info_float(values[-1]))
                if amount:
                    terms.append(amount)
        if not saw_cash or not found_local:
            raise ValueError('No se encontró la fila "LOCAL ACCOUNTS".')
        return terms, local

    payments = read("Tarjetas / LOCAL ACCOUNTS", payments_reader)
    credit_terms, local_accounts = payments if payments is not None else (None, None)

    # Fila "Other" del Method of Payment Totals (pagos de 1 o 2 dólares,
    # pedido del usuario 2026-09-12): ausente cuando el día no tuvo ninguno
    # -- 0.0, no error. _find_label_values matchea el label EXACTO ("Other"),
    # nunca "Other Discounts".
    other_amount = read("Other", optional_last("Other", _force_positive))
    network_revenue = read("Network Revenue", required_last("Network Revenue", 'No se encontró "Network Revenue".'))
    # "Total Revenue" tal como lo imprime el POS (lo usa el control de Cierre
    # mensual para cruzar contra el total del mes).
    total_revenue = read("Total Revenue", required_last("Total Revenue", 'No se encontró "Total Revenue".'))

    # Cruce contra el "Total Sales" impreso: en todos los días reales cargados
    # se cumple exacto. Si no cierra, algún componente se leyó mal (o una
    # etiqueta opcional como "Fuel Discounts" no se reconoció y quedó en 0)
    # -- se guarda igual pero se avisa con la diferencia. Sin alguno de los
    # números no hay cruce posible.
    total_sales_mismatch = None
    parts = (sales_fuel, desc_comb, non_fuel_total, desc_otros, tax_collect, total_sales)
    if None not in parts:
        components = sales_fuel + desc_comb + non_fuel_total + desc_otros + tax_collect
        total_sales_mismatch = round(components - total_sales, 2)
        if abs(total_sales_mismatch) <= 0.02:
            total_sales_mismatch = None

    return {
        "total_sales_mismatch": total_sales_mismatch,
        "missing_fields": missing,
        "from_date": period["from_date"],
        "from_time": period["from_time"],
        "to_date": period["to_date"],
        "to_time": period["to_time"],
        "volume": volume,
        "sales_fuel": sales_fuel,
        "desc_comb": desc_comb,
        "non_fuel_total": non_fuel_total,
        "desc_otros": desc_otros,
        "tax_collect": tax_collect,
        "total_sales": total_sales,
        "cash": cash,
        "credit_terms": credit_terms,
        "local_accounts": local_accounts,
        "other_amount": other_amount,
        "network_revenue": network_revenue,
        "total_revenue": total_revenue,
    }


def _extract_store_info_from_pdf_uncached(pdf_path, start_page_index=DEFAULT_STORE_INFO_PAGE_INDEX):
    """
    OCR the "PERIOD FROM:" page (generally page 3) plus as many following
    pages as needed to reach "Network Revenue" (generally the start of page
    4), and pull every Store Info field out of the combined text.
    """
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF no encontrado: {pdf_path}")

    images = _LazyPdfPageImages(pdf_path)
    try:
        total_pages = len(images)
        if not (0 <= start_page_index < total_pages):
            start_page_index = 0

        search_order = list(range(start_page_index, total_pages)) + list(
            range(0, start_page_index)
        )

        anchor_page = None
        anchor_text = None
        for idx in search_order:
            text = _ocr_store_info_page_text(images[idx])
            if PERIOD_FROM_ANCHOR in text.lower():
                anchor_page = idx
                anchor_text = text
                break

        if anchor_page is None:
            raise ValueError(
                'No se encontró el ancla "PERIOD FROM:" en ninguna página del PDF (vía OCR). '
                "Verifique que el reporte no esté demasiado borroso o girado."
            )

        lines = list(anchor_text.splitlines())
        pages_used = [anchor_page + 1]
        idx = anchor_page + 1
        pages_tried = 1
        while (
            NETWORK_REVENUE_ANCHOR not in "\n".join(lines).lower()
            and idx < total_pages
            and pages_tried < STORE_INFO_MAX_CONTINUATION_PAGES
        ):
            text = _ocr_store_info_page_text(images[idx])
            if _score_store_info_page_text(text) == 0 and images[idx] is not None:
                # Hoja escaneada al revés (Resumen de Ventas de septiembre
                # 2026: la página de Network/Total Revenue vino girada 180°).
                rotated = _ocr_store_info_page_text(images[idx].rotate(180))
                if _score_store_info_page_text(rotated) > 0:
                    text = rotated
            if any(_line_contains_anchor(line) for line in text.splitlines()) and (
                NETWORK_REVENUE_ANCHOR not in text.lower()
            ):
                # Ya es el Department Sales Report, no Store Info: leerla como
                # continuación hacía que los departamentos se buscaran desde
                # la página siguiente, dando la vuelta a todo el PDF (3 min).
                break
            lines.extend(text.splitlines())
            pages_used.append(idx + 1)
            pages_tried += 1
            idx += 1
    finally:
        images.close()

    fields = _extract_store_info_fields(lines)
    fields["pages_used"] = pages_used
    return fields


@functools.lru_cache(maxsize=64)
def _extract_store_info_from_pdf_cached(pdf_path, start_page_index):
    return _extract_store_info_from_pdf_uncached(pdf_path, start_page_index)


def extract_store_info_from_pdf(pdf_path, start_page_index=DEFAULT_STORE_INFO_PAGE_INDEX, strict=True):
    """
    strict=True (default, los usos viejos de Herramientas/Controles): si
    algún campo no se pudo leer, ValueError como siempre. strict=False
    (Carga de Datos): devuelve lo que se leyó, con None en lo que no y la
    lista en fields["missing_fields"] -- solo falla si no hay período.

    Wrapper con caché sobre _extract_store_info_from_pdf_uncached -- mismo
    motivo/criterio que parse_elistar_daily_pdf_page de arriba (ver ese
    docstring): esta función se llama más de una vez para el MISMO PDF
    dentro de un mismo job (ej. desde extract_department_sales_for_day Y
    desde extract_lottery_department_fields_from_pdf), cada una repitiendo
    el mismo OCR de las páginas de Store Info -- el paso más lento de todo
    el pipeline. `lru_cache` acotado + devolver una copia nueva en cada
    llamada, mismo criterio de seguridad que el otro wrapper.
    """
    fields = _extract_store_info_from_pdf_cached(os.path.abspath(pdf_path), start_page_index)
    if strict and fields.get("missing_fields"):
        raise ValueError("No se pudo leer del PDF: " + ", ".join(fields["missing_fields"]) + ".")
    return copy.deepcopy(fields)


def _find_store_info_sheet(workbook):
    target = STORE_INFO_SHEET_NAME.strip().lower()
    for name in workbook.sheetnames:
        if name.strip().lower() == target:
            return workbook[name]
    raise ValueError(
        f'Hoja "{STORE_INFO_SHEET_NAME}" no encontrada. Disponibles: {", ".join(workbook.sheetnames)}'
    )


def _store_info_row_for_day(day_of_month):
    """
    Row N holds day (N-1) of the month — row 2 is always day 1 — matching
    how CARGA AQUI itself pins one calendar day to one fixed row. This is
    what lets a later, out-of-order PDF (day 10 after day 1) land on the row
    that actually corresponds to it instead of just the next blank one,
    leaving days 2-9 correctly blank until their own reports arrive.
    """
    return STORE_INFO_DATA_START_ROW + (day_of_month - 1)


def _build_credit_terms_formula(amounts):
    """Mirror the sheet's own history: '=a+b+c', omitting zero-valued rows."""
    if not amounts:
        return 0.0
    return "=" + "+".join(f"{amount:.2f}" for amount in amounts)


def write_store_info_row(sheet, fields):
    """Write one Store Info row on the row matching this report's calendar day."""
    from_date = fields["from_date"]
    to_date = fields["to_date"]
    col_a_date = datetime(from_date.year, from_date.month, from_date.day) + timedelta(days=1)
    col_c_date = datetime(to_date.year, to_date.month, to_date.day) + timedelta(days=1)
    row = _store_info_row_for_day(col_a_date.day)

    sheet.cell(row=row, column=STORE_INFO_COL_FROM_DATE, value=col_a_date)
    sheet.cell(row=row, column=STORE_INFO_COL_FROM_TIME, value=fields["from_time"])
    sheet.cell(
        row=row,
        column=STORE_INFO_COL_TO_DATE,
        value=col_c_date,
    )
    sheet.cell(row=row, column=STORE_INFO_COL_TO_TIME, value=fields["to_time"])
    sheet.cell(row=row, column=STORE_INFO_COL_VOLUME, value=fields["volume"])
    sheet.cell(row=row, column=STORE_INFO_COL_SALES_FUEL, value=fields["sales_fuel"])
    sheet.cell(row=row, column=STORE_INFO_COL_DESC_COMB, value=fields["desc_comb"])
    sheet.cell(row=row, column=STORE_INFO_COL_NON_FUEL, value=fields["non_fuel_total"])
    sheet.cell(row=row, column=STORE_INFO_COL_DESC_OTROS, value=fields["desc_otros"])
    sheet.cell(row=row, column=STORE_INFO_COL_TAX_COLLECT, value=fields["tax_collect"])
    sheet.cell(row=row, column=STORE_INFO_COL_CASH, value=fields["cash"])
    sheet.cell(
        row=row,
        column=STORE_INFO_COL_CREDIT,
        value=_build_credit_terms_formula(fields["credit_terms"]),
    )
    sheet.cell(row=row, column=STORE_INFO_COL_OTHER, value=fields.get("other_amount", 0.0))
    sheet.cell(row=row, column=STORE_INFO_COL_LOCAL_ACCOUNTS, value=fields["local_accounts"])
    sheet.cell(
        row=row, column=STORE_INFO_COL_NETWORK_REVENUE, value=fields["network_revenue"]
    )
    return row


def process_store_info(master_path, pdf_paths, progress_callback=None):
    """
    Parse one or more daily PDFs and write one Store Info row per day to the
    "Store Info" sheet of a separate workbook — each on the row matching its
    own calendar day (row 2 = day 1), not just the next blank row, so PDFs
    for non-consecutive days land where they belong and gaps stay blank.

    Each PDF is read (OCR included) in parallel across a small worker pool,
    since that read — not the Excel write that follows — is what makes a
    multi-PDF batch slow.

    Returns:
        tuple: (temp_path, summary dict)
    """
    _ensure_openpyxl()
    master_path = os.path.abspath(str(master_path).strip())
    paths = _normalize_pdf_paths(pdf_paths)

    if not os.path.isfile(master_path):
        raise FileNotFoundError(f"Excel de Store Info no encontrado: {master_path}")
    if not paths:
        raise ValueError("No se proporcionaron PDF diarios.")

    extension = os.path.splitext(master_path)[1].lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise ValueError("El Excel de Store Info debe ser .xlsx o .xlsm.")

    sorted_paths, files_with_bad_filename = _partition_paths_by_extractable_day(paths)
    if not sorted_paths:
        raise ValueError(
            "No se pudo extraer el día del mes del nombre de ningún PDF -- se esperaba "
            "un patrón como 'Close Store 01-05.pdf'."
        )

    fields_by_path, parse_errors = _parse_pdfs_concurrently(
        sorted_paths, extract_store_info_from_pdf, progress_callback=progress_callback
    )
    files_failed_to_parse = len(parse_errors)
    ok_paths = [path for path in sorted_paths if path in fields_by_path]

    keep_vba = extension == ".xlsm"
    workbook = load_workbook(master_path, data_only=False, keep_vba=keep_vba)
    sheet = _find_store_info_sheet(workbook)

    batch_results = []
    days_failed_to_write = 0

    for pdf_path in ok_paths:
        fields = fields_by_path[pdf_path]
        # Aislado por día, igual que ya hacen process_reporte_diario/
        # process_lottery con su propio paso de escritura -- sin esto, un
        # solo PDF con un campo mal parseado (ej. from_date en None) tiraba
        # TODO el lote via una excepción sin atrapar, perdiendo también los
        # días de otros PDFs ya escritos en memoria (bug real, auditoría
        # 2026-09-06: inconsistencia con los otros dos flujos del mismo
        # archivo, que sí tienen este aislamiento desde 2026-09-03).
        try:
            row = write_store_info_row(sheet, fields)
        except (ValueError, TypeError, AttributeError) as exc:
            days_failed_to_write += 1
            continue
        # El "from_date" tal como lo imprime el PDF viene siempre un día
        # antes del día real que cubre el reporte (ver write_store_info_row)
        # -- se lee de vuelta la fecha que realmente quedó en columna A en
        # vez de reformatear el "from_date" crudo, para no mostrar una
        # fecha corrida un día si este resumen alguna vez se muestra.
        written_date = sheet.cell(row=row, column=STORE_INFO_COL_FROM_DATE).value
        batch_results.append(
            {
                "pdf_path": pdf_path,
                "filename": os.path.basename(pdf_path),
                "target_row": row,
                "from_date": written_date.strftime("%d/%m/%Y") if written_date else None,
                "pages_used": fields.get("pages_used"),
            }
        )

    if not batch_results:
        workbook.close()
        raise ValueError(
            "No se pudo cargar ningún día: "
            f"{files_with_bad_filename} archivo(s) con nombre no reconocido, "
            f"{files_failed_to_parse} archivo(s) no se pudieron leer, "
            f"{days_failed_to_write} día(s) no matchearon ninguna fila."
        )

    temp_path = _create_temp_workbook_path()
    workbook.save(os.path.abspath(temp_path))
    workbook.close()

    summary = {
        "files_processed": len(batch_results),
        "files_with_bad_filename": files_with_bad_filename,
        "files_failed_to_parse": files_failed_to_parse,
        "days_failed_to_write": days_failed_to_write,
        "batch_results": batch_results,
    }
    return temp_path, summary


# ---------------------------------------------------------------------------
# Lottery — a third extraction from the same daily PDF, this time from its
# last two pages (Florida Lottery terminal receipts: "Daily Terminal Games
# Sales" for ONLINE and "Daily Scratch-Off Games Sales" for SKOFF), combined
# with the ONLINE/SKOFF rows already read off the Department Sales Report
# for the Ventas pipeline, into one row of the monthly Lottery workbook.
# ---------------------------------------------------------------------------

LOTTERY_DATA_START_ROW = 4  # Row 3 is the header; data starts at row 4
LOTTERY_COL_DATE_A = 1  # A — Fecha (period start, informational only)
LOTTERY_COL_DATE_B = 2  # B — Fecha (the date every row is actually keyed on)
LOTTERY_COL_ONLINE_COUNT = 4  # D — ONLINE Net Count (from Department Sales Report)
LOTTERY_COL_ONLINE_NET_SALES = 5  # E — ONLINE Net Sales (from Department Sales Report)
LOTTERY_COL_SALES = 6  # F — NET SALES off the terminal-games receipt
LOTTERY_COL_PAGOS = 7  # G — PAYS off the terminal-games receipt
LOTTERY_COL_CASH_BALANCE = 8  # H — F - G, pasted as a value, never a formula
LOTTERY_COL_COMIS = 9  # I — the NET SALES nested under TOTAL SALES COMM
LOTTERY_COL_PRIZE_FREE_PLAYS = 11  # K — PRIZE FREE PLAYS (+ PROMO FREE PLAYS if any)
LOTTERY_COL_SKOFF_COUNT = 14  # N — SKOFF Net Count (from Department Sales Report)
LOTTERY_COL_SKOFF_NET_SALES = 15  # O — SKOFF Net Sales (from Department Sales Report)
LOTTERY_COL_SKOFF_PAYS_UNITS = 16  # P — PAYS unit count off the scratch-off receipt
LOTTERY_COL_SKOFF_PAYS_AMOUNT = 17  # Q — PAYS $ amount off the scratch-off receipt
LOTTERY_COL_SKOFF_SALES_AMOUNT = 18  # R — "Instant Sales Amount" off the Daily Sales Report
LOTTERY_COL_SALES_COMM = 19  # S — SALES COMM off the scratch-off receipt

LOTTERY_ONLINE_DEPARTMENT = "ONLINE"
LOTTERY_SKOFF_DEPARTMENT = "SKOFF"


def _find_department_record(records, department_name):
    for record in records:
        if record["department"] == department_name:
            return record
    return None


def extract_lottery_department_fields_from_pdf(pdf_path, page_index=DEFAULT_PDF_PAGE_INDEX):
    """
    D/E/N/O — the ONLINE/SKOFF rows off the Department Sales Report (same
    page the Ventas pipeline reads) — plus this PDF's own business date.

    The date is read the same way Store Info reads it (PERIOD FROM/TO, +1
    day) rather than off the Lottery receipt pages: those are small, dense,
    watermark-obscured text where OCR has repeatedly misread a single digit
    (e.g. "08" -> "09"), while PERIOD FROM/TO is a larger, cleaner line that
    Store Info already reads reliably from this exact same PDF.

    F/G/H/I/K and P/Q/R/S no longer come from here: see
    extract_lottery_receipt_fields_from_sales_report, which reads those
    columns from the Florida Lottery portal's own "Daily Sales Report"
    PDF — real embedded text instead of a photographed receipt.

    Lo que no se pudo leer con seguridad (o el departamento que no apareció)
    vuelve en None y su nombre en "missing" (2026-10-05): antes un solo valor
    dudoso tiraba error y el día quedaba sin ONLINE ni SKOFF, sin aviso.
    Nunca se adivina; quien guarda deja lo que ya había y avisa.
    """
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF no encontrado: {pdf_path}")

    dept_records, _dept_diagnostics = parse_elistar_daily_pdf_page(pdf_path, page_index=page_index)
    values = {}
    missing = []
    for department, prefix in ((LOTTERY_ONLINE_DEPARTMENT, "online"), (LOTTERY_SKOFF_DEPARTMENT, "skoff")):
        record = _find_department_record(dept_records, department) or {}
        count, amount = record.get("count"), record.get("amount")
        values[f"{prefix}_count"] = int(count) if count is not None else None
        values[f"{prefix}_net_sales"] = float(amount) if amount is not None else None
        if count is None or amount is None:
            missing.append(department)

    # Solo la fecha: un campo de Store Info ilegible no frena ONLINE/SKOFF.
    store_info_fields = extract_store_info_from_pdf(pdf_path, strict=False)
    from_date = store_info_fields["from_date"]
    report_date = date(from_date.year, from_date.month, from_date.day) + timedelta(days=1)

    return {"report_date": report_date, **values, "missing": missing}


_LOTTERY_SALES_REPORT_START_DATE_RE = re.compile(r"Start Date:\s*(\d{4})-(\d{2})-(\d{2})")

# The Lottery's own commission rate on ONLINE net sales — column J's own
# "=+I/F" formula is meant to always read -6.00%; confirmed exactly against
# three independent real days (comis/sales = -15.66/261, -15.0/250, -5.82/97).
_LOTTERY_ONLINE_COMMISSION_RATE = 0.06

# label -> (result key, sign rule) for every value this module needs off the
# Florida Lottery portal's "Daily Sales Report" PDF. Unlike the photographed
# terminal receipt, this document has a real embedded text layer, so each
# label is matched once, exactly, with no fuzzy/OCR tolerance needed.
_LOTTERY_SALES_REPORT_FIELDS = (
    ("sales", "Net Terminal Sales Amount", _force_positive),
    ("pagos", "Terminal Pay Amount", _force_negative),
    ("comis", "Terminal Sales Commission", _force_negative),
    ("pays_units", "Instant Tickets Paid", None),
    ("pays_amount", "Instant Pay Amount", _force_negative),
    ("skoff_sales_amount", "Instant Sales Amount", _force_positive),
    ("sales_comm", "Instant Sales Commission", _force_negative),
)


def _parse_lottery_sales_report_date(text):
    match = _LOTTERY_SALES_REPORT_START_DATE_RE.search(text)
    if not match:
        return None
    year, month, day = (int(part) for part in match.groups())
    try:
        return date(year, month, day)
    except ValueError:
        return None


def _extract_sales_report_value(text, label):
    """Numeric token immediately after an exact label — safe here since the label set has no overlapping substrings."""
    match = re.search(re.escape(label) + r"\s*\$?(-?[\d,]+\.?\d*)", text)
    if not match:
        return None
    return match.group(1).replace(",", "")


def extract_lottery_receipt_fields_from_sales_report(pdf_path):
    """
    F/G/H/I/K (ONLINE) and P/Q/R/S (SKOFF) off the Florida Lottery portal's
    own "Daily Sales Report" PDF — real embedded text, not a photographed
    terminal receipt, so there's no OCR and no ambiguity. This source only
    prints the combined "Terminal Sales Commission" total, not the NET
    SALES / PRIZE FREE PLAYS split the physical receipt shows — but I/F is
    fixed at -6.00% by the Lottery's own commission rate (column J's own
    "=+I/F" formula), so I is rebuilt from that fixed rate against F and K
    takes whatever's left of the total, reproducing the same split the
    receipt shows without needing to read it.
    """
    _ensure_pdfplumber()
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF no encontrado: {pdf_path}")

    with pdfplumber.open(pdf_path) as pdf:
        raw_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    text = _normalize_department_spacing(raw_text)

    report_date = _parse_lottery_sales_report_date(text)
    if report_date is None:
        raise ValueError('No se encontró "Start Date" en el Daily Sales Report de Lottery.')

    warnings = []
    values = {}
    for key, label, sign_rule in _LOTTERY_SALES_REPORT_FIELDS:
        raw = _extract_sales_report_value(text, label)
        if raw is None:
            values[key] = None
            warnings.append(f'No se pudo leer "{label}" en el Daily Sales Report de Lottery — revisar manualmente.')
            continue
        number = float(raw)
        values[key] = sign_rule(number) if sign_rule else _safe_parse_count(raw)

    sales = values["sales"]
    pagos = values["pagos"]
    cash_balance = round(sales + pagos, 2) if sales is not None and pagos is not None else None

    total_comm = values["comis"]
    if total_comm is not None and sales is not None:
        comis = round(-_LOTTERY_ONLINE_COMMISSION_RATE * sales, 2)
        prize_free_plays = round(total_comm - comis, 2)
    elif total_comm is not None:
        comis = total_comm
        prize_free_plays = 0.0
        warnings.append(
            'No se pudo leer "Net Terminal Sales Amount" (columna F), así que no se pudo dividir '
            "la comisión entre I y K — se cargó el total completo en I."
        )
    else:
        comis = None
        prize_free_plays = None

    return {
        "report_date": report_date,
        "sales": sales,
        "pagos": pagos,
        "cash_balance": cash_balance,
        "comis": comis,
        "prize_free_plays": prize_free_plays,
        "pays_units": values["pays_units"],
        "pays_amount": values["pays_amount"],
        "skoff_sales_amount": values["skoff_sales_amount"],
        "sales_comm": values["sales_comm"],
        "warning": " | ".join(warnings) if warnings else None,
    }


def _find_lottery_sheet(workbook):
    """
    The Lottery workbook has one sheet per month, named after the month
    (e.g. "08.2026") — so unlike CARGA AQUI or Store Info there's no fixed
    name to look for. Falls back to the sheet matching the known header
    signature (Fecha/Fecha/.../COUNT in row 3) when there's more than one.
    """
    names = workbook.sheetnames
    if len(names) == 1:
        return workbook[names[0]]

    for name in names:
        sheet = workbook[name]
        header_a = _strip_cell(sheet.cell(row=3, column=LOTTERY_COL_DATE_A).value).lower()
        header_d = _strip_cell(sheet.cell(row=3, column=LOTTERY_COL_ONLINE_COUNT).value).lower()
        if header_a == "fecha" and header_d == "count":
            return sheet

    raise ValueError(
        f'No se pudo identificar la hoja de Lottery. Disponibles: {", ".join(names)}'
    )


def _find_lottery_row_for_date(sheet, target_date, start_row=LOTTERY_DATA_START_ROW):
    """
    Exact (day, month, year) match first; if that fails, fall back to
    (day, month) alone. A single Lottery workbook only ever spans one
    month, so a day/month match is already unambiguous — and this recovers
    from a lone OCR-misread year digit (e.g. "26" read as "25"), which
    doesn't affect which real-world day the report is for.
    """
    max_row = max(sheet.max_row, start_row)
    day_month_row = None
    for row in range(start_row, max_row + 1):
        cell_value = sheet.cell(row=row, column=LOTTERY_COL_DATE_B).value
        if not isinstance(cell_value, datetime):
            continue
        if cell_value.date() == target_date:
            return row
        if day_month_row is None and (cell_value.month, cell_value.day) == (
            target_date.month,
            target_date.day,
        ):
            day_month_row = row
    if day_month_row is not None:
        return day_month_row
    raise ValueError(
        f"No se encontró una fila con la fecha {target_date.strftime('%d/%m/%Y')} en la "
        f"columna B de la hoja de Lottery."
    )


def write_lottery_row(sheet, fields):
    row = _find_lottery_row_for_date(sheet, fields["report_date"])

    if fields.get("online_count") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_ONLINE_COUNT, value=fields["online_count"])
    if fields.get("online_net_sales") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_ONLINE_NET_SALES, value=fields["online_net_sales"])
    if fields.get("sales") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_SALES, value=fields["sales"])
    if fields.get("pagos") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_PAGOS, value=fields["pagos"])
    if fields.get("cash_balance") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_CASH_BALANCE, value=fields["cash_balance"])
    if fields.get("comis") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_COMIS, value=fields["comis"])
    if fields.get("prize_free_plays") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_PRIZE_FREE_PLAYS, value=fields["prize_free_plays"])
    if fields.get("skoff_count") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_SKOFF_COUNT, value=fields["skoff_count"])
    if fields.get("skoff_net_sales") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_SKOFF_NET_SALES, value=fields["skoff_net_sales"])
    if fields.get("pays_units") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_SKOFF_PAYS_UNITS, value=fields["pays_units"])
    if fields.get("pays_amount") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_SKOFF_PAYS_AMOUNT, value=fields["pays_amount"])
    if fields.get("skoff_sales_amount") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_SKOFF_SALES_AMOUNT, value=fields["skoff_sales_amount"])
    if fields.get("sales_comm") is not None:
        sheet.cell(row=row, column=LOTTERY_COL_SALES_COMM, value=fields["sales_comm"])
    return row


def process_lottery(master_path, department_pdf_paths, sales_report_pdf_paths):
    """
    Write one Lottery row per business date, merging two independent PDF
    sources by that date: the shared daily PDF's Department Sales Report
    (D/E/N/O — ONLINE/SKOFF counts and sales, same source the Ventas
    pipeline reads) and the Florida Lottery portal's own "Daily Sales
    Report" PDF (F/G/H/I/K, P/Q/R/S). Either source can be given alone —
    a date present in only one still gets a row written with just that
    source's columns, and a warning noting the other is missing — since a
    given batch of uploads won't always include both for every day.

    Each PDF set is parsed in parallel (within its own set) before any
    writing happens, since that read is what makes a multi-PDF batch slow.

    Returns:
        tuple: (temp_path, summary dict)
    """
    _ensure_openpyxl()
    master_path = os.path.abspath(str(master_path).strip())
    department_paths = _normalize_pdf_paths(department_pdf_paths)
    sales_report_paths = _normalize_pdf_paths(sales_report_pdf_paths)

    if not os.path.isfile(master_path):
        raise FileNotFoundError(f"Excel de Lottery no encontrado: {master_path}")
    if not department_paths and not sales_report_paths:
        raise ValueError("No se proporcionaron PDFs de Lottery.")

    extension = os.path.splitext(master_path)[1].lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise ValueError("El Excel de Lottery debe ser .xlsx o .xlsm.")

    for pdf_path in department_paths + sales_report_paths:
        if not os.path.isfile(pdf_path):
            raise FileNotFoundError(f"PDF no encontrado: {pdf_path}")

    # Aislado por PDF: antes, un PDF ilegible de cualquiera de los dos
    # conjuntos tiraba TODO el lote (ambas fuentes), perdiendo el trabajo ya
    # bueno del resto -- ver _parse_pdfs_concurrently.
    department_fields_by_path, department_parse_errors = (
        _parse_pdfs_concurrently(department_paths, extract_lottery_department_fields_from_pdf)
        if department_paths
        else ({}, {})
    )
    sales_report_fields_by_path, sales_report_parse_errors = (
        _parse_pdfs_concurrently(sales_report_paths, extract_lottery_receipt_fields_from_sales_report)
        if sales_report_paths
        else ({}, {})
    )
    files_failed_to_parse = len(department_parse_errors) + len(sales_report_parse_errors)

    department_by_date = {
        fields["report_date"]: (path, fields) for path, fields in department_fields_by_path.items()
    }
    sales_report_by_date = {
        fields["report_date"]: (path, fields) for path, fields in sales_report_fields_by_path.items()
    }

    keep_vba = extension == ".xlsm"
    workbook = load_workbook(master_path, data_only=False, keep_vba=keep_vba)
    sheet = _find_lottery_sheet(workbook)

    all_dates = sorted(set(department_by_date) | set(sales_report_by_date))

    batch_results = []
    dates_failed_to_write = 0
    for report_date in all_dates:
        dept_path, dept_fields = department_by_date.get(report_date, (None, None))
        sales_path, sales_fields = sales_report_by_date.get(report_date, (None, None))

        merged = {"report_date": report_date}
        warnings = []
        if dept_fields is not None:
            merged.update(
                {
                    "online_count": dept_fields["online_count"],
                    "online_net_sales": dept_fields["online_net_sales"],
                    "skoff_count": dept_fields["skoff_count"],
                    "skoff_net_sales": dept_fields["skoff_net_sales"],
                }
            )
        elif department_paths and sales_report_paths:
            # Only worth flagging when this run intentionally supplied both
            # sources and one of them simply has no PDF for this particular
            # date — a single-source run (the normal case: the GUI's two
            # Lottery flows each call this with the other list empty) isn't
            # missing anything, it just never touches the other's columns.
            warnings.append(
                "No se subió el PDF diario con el Department Sales Report para esta fecha "
                "— columnas D/E/N/O sin actualizar."
            )
        if sales_fields is not None:
            for key in (
                "sales", "pagos", "cash_balance", "comis", "prize_free_plays",
                "pays_units", "pays_amount", "skoff_sales_amount", "sales_comm",
            ):
                merged[key] = sales_fields[key]
            if sales_fields.get("warning"):
                warnings.append(sales_fields["warning"])
        elif department_paths and sales_report_paths:
            warnings.append(
                "No se subió el Daily Sales Report de Lottery para esta fecha "
                "— columnas F/G/H/I/K/P/Q/R/S sin actualizar."
            )

        # Aislado por fecha: antes, una fecha que no matcheaba ninguna fila
        # de la hoja de Lottery (ej. pertenece a otro mes) tiraba TODO el
        # lote, perdiendo las filas de las demás fechas ya escritas en
        # memoria antes de que workbook.save() llegara a correr.
        try:
            row = write_lottery_row(sheet, merged)
        except (ValueError, TypeError) as exc:
            dates_failed_to_write += 1
            continue
        batch_results.append(
            {
                "filename": os.path.basename(sales_path or dept_path),
                "target_row": row,
                "report_date": report_date.strftime("%d/%m/%Y"),
                "warning": " | ".join(warnings) if warnings else None,
            }
        )

    if not batch_results:
        workbook.close()
        raise ValueError(
            "No se pudo cargar ninguna fecha: "
            f"{files_failed_to_parse} archivo(s) no se pudieron leer, "
            f"{dates_failed_to_write} fecha(s) no matchearon ninguna fila."
        )

    temp_path = _create_temp_workbook_path()
    workbook.save(os.path.abspath(temp_path))
    workbook.close()

    summary = {
        "files_processed": len(batch_results),
        "files_failed_to_parse": files_failed_to_parse,
        "dates_failed_to_write": dates_failed_to_write,
        "batch_results": batch_results,
    }
    return temp_path, summary
