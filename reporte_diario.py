"""
Reporte Diario — Elistar daily PDF extraction into Bradenton C-Store master.

Reads department totals from a user-selected PDF page of an Elistar daily closure
report, maps headers dynamically on sheet \"CARGA AQUI\" (row 3, column C onward),
and injects count/amount pairs on the first eligible operational row.
"""

import os
import re
import sys
import tempfile
import unicodedata
from difflib import get_close_matches
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    pdfplumber = None  # type: ignore[assignment]

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

DEFAULT_PDF_PAGE_INDEX = 3  # Fourth page (0-based) for full un-cropped daily PDF

# Fixed 1-based PDF column layout (Elistar daily closure)
PDF_DEPT_NAME_COL = 0  # 1st column — Dept.Name
PDF_NET_COUNT_COL = 4  # 5th column — Net Count
PDF_NET_SALES_COL = 7  # 8th column — Net Sales $
PDF_MIN_COLUMNS = 8

PDF_HEADER_DEPT_TOKENS = ("dept", "name")
PDF_HEADER_NET_COUNT_TOKENS = ("net", "count")
PDF_HEADER_NET_SALES_TOKENS = ("net", "sales")

DEPARTMENT_SALES_REPORT_ANCHOR = "Department Sales Report"

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

# Elistar department row layout (right to left): ... | Net Count | ... | Net Sales $ | % of sales
NET_COUNT_REVERSE_INDEX = -5
NET_SALES_REVERSE_INDEX = -2
MIN_ROW_SPLIT_PARTS = 5

SALES_ALERT_FONT_COLOR = "FF0000"
GETTEL_TOYOTA_LABEL = "gettel/toyota"
GROUP_500_DEPARTMENTS = frozenset(
    {
        "AUTO",
        "BOILED PEANUTS",
        "HBA",
        "ICECREAM",
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


def _ensure_pdfplumber():
    if pdfplumber is None:
        raise ImportError(
            "Reporte Diario requires pdfplumber. Install with: pip install pdfplumber"
        )


def _ensure_openpyxl():
    if not OPENPYXL_AVAILABLE:
        raise ImportError(
            "Reporte Diario requires openpyxl. Install with: pip install openpyxl"
        )


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
    """Strip phantom leading/trailing dots, hyphens, and spaces from Dept.Name."""
    text = _clean_dept_name(value)
    text = re.sub(r"^[.\-\s]+|[.\-\s]+$", "", text)
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
    if _is_protected_department(department):
        return None

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
    if key.startswith("dept "):
        return True
    if key in PROTECTED_DEPARTMENT_LABELS:
        return False
    return any(marker in key for marker in SUMMARY_STOP_MARKERS)


def _is_protected_department(label):
    key = _normalize_department_label(label)
    return key in PROTECTED_DEPARTMENT_LABELS


def _is_valid_date(value):
    if value is None:
        return False
    if isinstance(value, datetime):
        return True
    if isinstance(value, (int, float)):
        return value > 0
    text = _strip_cell(value)
    if not text:
        return False
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            datetime.strptime(text, fmt)
            return True
        except ValueError:
            continue
    if re.match(r"^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$", text):
        return True
    if re.match(r"^\d{1,2}-[a-zA-Z]{3}(-\d{2,4})?$", text):
        return True
    return False


def _cell_has_formula(cell):
    if getattr(cell, "data_type", None) == "f":
        return True
    value = cell.value
    return isinstance(value, str) and value.startswith("=")


def _create_temp_workbook_path():
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="reporte_diario_")
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


def _get_carga_aqui_sheet(workbook):
    target = SHEET_NAME.strip().lower()
    for name in workbook.sheetnames:
        if name.strip().lower() == target:
            return workbook[name]
    raise ValueError(
        f'Sheet "{SHEET_NAME}" not found. Available: {", ".join(workbook.sheetnames)}'
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


def _extract_page_triplet(cells):
    """Extract Dept.Name, Net Count, and Net Sales $ via reverse-index whitespace split."""
    line = _row_cells_to_line(cells)
    parsed = _parse_row_by_reverse_index(line)
    if parsed is not None:
        return parsed
    return {
        "department": "",
        "count": 0,
        "amount": 0.0,
    }


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
    Extract the calendar day-of-month from an Elistar daily PDF filename.

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
        f"Could not extract calendar day from filename: {os.path.basename(file_path)}. "
        "Expected a pattern like 'Close Store 01-05.pdf'."
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
        f"No row with calendar day {target_day} found in column A on sheet {SHEET_NAME} "
        f"(from row {start_row})."
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
            f'Anchor "{DEPARTMENT_SALES_REPORT_ANCHOR}" not found on PDF page.'
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


def _extract_page_tables(page):
    """Backward-compatible alias for table extraction."""
    return _extract_tables_from_page(page)


def parse_elistar_daily_pdf_page(pdf_path, page_index=DEFAULT_PDF_PAGE_INDEX):
    """
    Extract department records from the selected PDF page (0-based index).

    Only rows after the \"Department Sales Report\" anchor are parsed.
    Uses fixed columns: Dept.Name (1), Net Count (5), Net Sales $ (8).

    Returns:
        list[dict]: Each dict has keys department, count, amount.
    """
    _ensure_pdfplumber()
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    if page_index < 0:
        raise ValueError("PDF page index must be 0 or greater.")

    with pdfplumber.open(pdf_path) as pdf:
        if page_index >= len(pdf.pages):
            raise ValueError(
                f"PDF page {page_index + 1} not found; "
                f"document has {len(pdf.pages)} page(s)."
            )
        page = pdf.pages[page_index]
        records = _parse_tables_with_line_anchor(page)

    if not records:
        raise ValueError(
            f"No department records found on PDF page {page_index + 1} after "
            f'"{DEPARTMENT_SALES_REPORT_ANCHOR}". '
            "Expected Dept.Name (col 1), Net Count (col 5), Net Sales $ (col 8)."
        )

    return [dict(record) for record in records]


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
            f"No department headers found on row {HEADER_ROW} starting at column C."
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


def find_operational_row(sheet, column_map, start_row=DATA_START_ROW):
    """
    Locate the first row from start_row with a valid date in column A and empty tracking cells.
    Defaults to row 5 when the sheet is fresh for a new month.
    """
    max_row = max(sheet.max_row, start_row)
    for row in range(start_row, max_row + 1):
        date_found = _cell_calendar_day(
            sheet.cell(row=row, column=DATE_SCAN_COLUMN).value
        ) is not None
        if not date_found:
            continue

        tracking_empty = True
        for count_col, amount_col in column_map.values():
            for col in (count_col, amount_col):
                value = sheet.cell(row=row, column=col).value
                if value is not None and _strip_cell(value) != "":
                    tracking_empty = False
                    break
            if not tracking_empty:
                break

        if tracking_empty:
            return row

    return start_row


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


def _is_gettel_toyota_department(dept_name):
    return GETTEL_TOYOTA_LABEL in _department_label_keys(dept_name)


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


def process_reporte_diario(
    master_path, pdf_paths, page_index=DEFAULT_PDF_PAGE_INDEX
):
    """
    Parse one or more Elistar daily PDFs and inject each into its calendar day row.

    Day-of-month is read from each PDF filename (e.g. Close Store 07-05.pdf -> day 7).
    Workbook is saved once after the full batch completes.

    Returns:
        tuple: (temp_path, summary dict)
    """
    _ensure_openpyxl()
    master_path = os.path.abspath(str(master_path).strip())
    paths = _normalize_pdf_paths(pdf_paths)

    if not os.path.isfile(master_path):
        raise FileNotFoundError(f"Master workbook not found: {master_path}")
    if not paths:
        raise ValueError("No Elistar daily PDF files provided.")

    extension = os.path.splitext(master_path)[1].lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise ValueError("Master workbook must be .xlsx or .xlsm.")

    keep_vba = extension == ".xlsm"
    workbook = load_workbook(master_path, data_only=False, keep_vba=keep_vba)
    sheet = _get_carga_aqui_sheet(workbook)
    column_map, protected = build_department_column_map(sheet)

    batch_results = []
    total_written = 0
    total_skipped = 0
    total_departments = 0

    sorted_paths = sorted(paths, key=extract_day_from_filename)

    for pdf_path in sorted_paths:
        if not os.path.isfile(pdf_path):
            raise FileNotFoundError(f"PDF not found: {pdf_path}")

        target_day = int(extract_day_from_filename(pdf_path))
        pdf_records = parse_elistar_daily_pdf_page(pdf_path, page_index=page_index)
        target_row = find_row_for_calendar_day(sheet, target_day)

        written, skipped = inject_daily_sales(
            sheet, pdf_records, column_map, target_row
        )

        total_written += len(written)
        total_skipped += len(skipped)
        total_departments += len(pdf_records)
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
            }
        )
        pdf_records = None

    temp_path = _create_temp_workbook_path()
    workbook.save(os.path.abspath(temp_path))
    workbook.close()

    _launch_temp_workbook(temp_path)

    summary = {
        "files_processed": len(batch_results),
        "page_index": page_index,
        "page_number": page_index + 1,
        "departments_written": total_written,
        "departments_skipped": total_skipped,
        "pdf_departments": total_departments,
        "protected_headers": sorted(protected),
        "batch_results": batch_results,
    }
    return temp_path, summary
