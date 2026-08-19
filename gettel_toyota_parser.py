"""
GETTEL_TOYOTA_PARSER — Fuel coupon PDF extraction into a unified Excel workbook.

The source PDF holds an 8-column layout that is really two identical mirrored
blocks of 4 columns each (Date | Amount | Tracking | Gallons). The report rows
are mapped directly into a strict 8-column matrix — the left and right blocks
stay side by side (no vertical stacking) and the 6-digit tracking values are
preserved in their own columns for full mapping verification. Amount and gallons
are parsed into native floats with a '0.00' number format so a Spanish Excel
renders them with a decimal comma. Batch runs produce one workbook with one
worksheet per PDF (sheet title from the filename date range, e.g. 04-05 al 15-05).
"""

import os
import re
import sys
import time
from datetime import datetime

try:
    import pdfplumber
except ImportError:  # pragma: no cover - environment guard
    pdfplumber = None  # type: ignore[assignment]

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils.datetime import from_excel

    OPENPYXL_AVAILABLE = True
except ImportError:  # pragma: no cover - environment guard
    Workbook = None  # type: ignore[assignment,misc]
    load_workbook = None  # type: ignore[assignment,misc]
    Alignment = None  # type: ignore[assignment,misc]
    Font = None  # type: ignore[assignment,misc]
    from_excel = None  # type: ignore[assignment,misc]
    OPENPYXL_AVAILABLE = False

# Vendor detection keywords -> (vendor label, output sheet name).
VENDOR_GETTEL = ("Gettel", "Gettel Report")
VENDOR_TOYOTA = ("Toyota", "Toyota Report")
GETTEL_TOKENS = ("GETTEL", "GETTLE")
TOYOTA_TOKENS = ("TOYOTA",)

DECIMAL_NUMBER_FORMAT = "0.00"
DATE_OUTPUT_FORMAT = "%m/%d/%Y"

# A standalone 6-digit integer is a tracking number (Col 3 / Col 7) and is dropped.
TRACKING_PATTERN = re.compile(r"^\d{6}$")
# Date tokens such as 03/14/2025 or 3/14/25.
DATE_PATTERN = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{2,4})\b")
# Filename date-range e.g. "04-05 al 15-05" from "Gettel report 04-05 al 15-05.pdf"
FILENAME_DATE_RANGE_PATTERN = re.compile(
    r"(\d{1,2}\s*[-/.]\s*\d{1,2}\s+al\s+\d{1,2}\s*[-/.]\s*\d{1,2})",
    re.IGNORECASE,
)
_VENDOR_PREFIX_PATTERN = re.compile(
    r"^(?:(?:gettel|gettle|toyota|report|reporte|cupon|cupón|fuel)[\s_\-.]*)+",
    re.IGNORECASE,
)

_DATE_INPUT_FORMATS = (
    "%m/%d/%Y",
    "%m/%d/%y",
    "%d/%m/%Y",
    "%d/%m/%y",
)

LEFT_ALIGNMENT = Alignment(horizontal="left") if Alignment is not None else None
RIGHT_ALIGNMENT = Alignment(horizontal="right") if Alignment is not None else None
HEADER_FONT = Font(bold=True) if Font is not None else None


def _ensure_pdfplumber():
    if pdfplumber is None:
        raise ImportError(
            "GETTEL_TOYOTA_PARSER requires pdfplumber. Install with: pip install pdfplumber"
        )


def _ensure_openpyxl():
    if not OPENPYXL_AVAILABLE:
        raise ImportError(
            "GETTEL_TOYOTA_PARSER requires openpyxl. Install with: pip install openpyxl"
        )


def detect_vendor(page_texts):
    """
    Identify the report vendor from the header text of the pages.

    Returns:
        tuple[str, str]: (vendor label, output sheet name). Falls back to Gettel
        when no recognizable header token is present.
    """
    haystack = "\n".join(text for text in page_texts if text).upper()
    if any(token in haystack for token in TOYOTA_TOKENS):
        return VENDOR_TOYOTA
    if any(token in haystack for token in GETTEL_TOKENS):
        return VENDOR_GETTEL
    return VENDOR_GETTEL


def _parse_date(token):
    """Return a datetime for a MM/DD/YYYY-style token, or None when unparseable."""
    text = token.strip()
    for fmt in _DATE_INPUT_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _clean_number_token(token):
    """Strip currency noise ($, backslashes, commas, spaces) from a raw token."""
    return token.replace("$", "").replace("\\", "").replace(",", "").strip()


def _coerce_float(cleaned):
    """Return a float for a cleaned numeric token, or None when not numeric."""
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def _correct_amount(value):
    """
    Enforce at most 3 integer digits on a Coupon Amount (real value < 1000.00).

    A PDF extraction that drops the decimal point yields a 4+ integer-digit value
    (e.g., 3868.00); divide by 100.0 to reposition the decimal (-> 38.68).
    """
    if value is None:
        return None
    while abs(value) >= 1000:
        value /= 100.0
    return value


def _correct_gallons(value):
    """
    Enforce at most 2 integer digits on a Gallons value (real value < 100.00).

    A 3+ integer-digit value (e.g., 1555.00 or 150.00) means the decimal point was
    lost; divide by 100.0 to restore the true reading (-> 15.55 or 1.50).
    """
    if value is None:
        return None
    while abs(value) >= 100:
        value /= 100.0
    return value


def _parse_block(block_text):
    """
    Parse one 4-column block (Date | Amount | Tracking | Gallons) into a record.

    The 6-digit tracking integer is kept in its own column for full mapping
    verification. A valid block yields a date plus two numeric values (amount
    first, gallons last). Returns a dict with keys date (datetime), amount
    (float), tracking (str), gallons (float), or None for incomplete/total rows.
    """
    date_match = DATE_PATTERN.search(block_text)
    if not date_match:
        return None
    parsed_date = _parse_date(date_match.group(1))
    if parsed_date is None:
        return None

    remainder = block_text[date_match.end():]
    numeric_values = []
    tracking = ""
    for raw_token in remainder.split():
        cleaned = _clean_number_token(raw_token)
        if not cleaned:
            continue
        if TRACKING_PATTERN.match(cleaned):
            tracking = cleaned  # 6-digit tracking — keep in its own column.
            continue
        value = _coerce_float(cleaned)
        if value is not None:
            numeric_values.append(value)

    if len(numeric_values) < 2:
        return None  # Incomplete trailing text / total row — discard.

    return {
        "date": parsed_date,
        "amount": _correct_amount(numeric_values[0]),
        "tracking": tracking,
        "gallons": _correct_gallons(numeric_values[-1]),
    }


def _split_mirrored_blocks(line):
    """
    Split a row into its left and right blocks using the two date anchors.

    The mirrored layout places one date at the start of each 4-column block, so
    the second date marks the boundary between the left and right halves.
    """
    matches = list(DATE_PATTERN.finditer(line))
    if not matches:
        return []
    if len(matches) == 1:
        return [line[matches[0].start():]]
    second = matches[1]
    left = line[matches[0].start():second.start()]
    right = line[second.start():]
    return [left, right]


def _parse_row(line):
    """
    Parse one PDF line into an 8-column row, keeping the left and right blocks
    side by side (no vertical stacking).

    Returns a dict with keys left and right, each either a parsed block dict or
    None. Returns None when neither block is valid.
    """
    blocks = _split_mirrored_blocks(line)
    if not blocks:
        return None
    left = _parse_block(blocks[0]) if len(blocks) >= 1 else None
    right = _parse_block(blocks[1]) if len(blocks) >= 2 else None
    if left is None and right is None:
        return None
    return {"left": left, "right": right}


def parse_fuel_coupons_pdf(pdf_path):
    """
    Extract fuel coupon rows from a Gettel/Toyota PDF report, grouped by page.

    Each row preserves the report's mirrored 8-column layout (left block and
    right block side by side) in original document order. Rows are grouped per
    PDF page so the writer can insert a blank separator row between pages.

    Returns:
        tuple[str, str, list[list[dict]]]: (vendor label, sheet name, pages).
        Each page is a list of row dicts with keys left and right; each block
        (or None) carries date, amount, tracking, and gallons.
    """
    _ensure_pdfplumber()
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    page_texts = []
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            page_texts.append(text)
            page_rows = []
            for line in text.splitlines():
                if not line.strip():
                    continue
                row = _parse_row(line)
                if row is not None:
                    page_rows.append(row)
            if page_rows:  # Skip empty pages so no spurious separator is added.
                pages.append(page_rows)

    vendor_label, sheet_name = detect_vendor(page_texts)
    return vendor_label, sheet_name, pages, page_texts


def _sanitize_excel_sheet_name(name):
    """Excel sheet titles: max 31 chars; no : \\ / ? * [ ]."""
    cleaned = re.sub(r"[:\\/?*\[\]]", "-", str(name).strip())
    cleaned = re.sub(r"\s+", " ", cleaned)
    return (cleaned[:31] if cleaned else "Report").strip()


def extract_sheet_title_from_filename(pdf_path):
    """
    Derive a worksheet title from the PDF filename.

    Example: "Gettel report 04-05 al 15-05.pdf" -> "04-05 al 15-05"
    """
    base = os.path.splitext(os.path.basename(pdf_path))[0].strip()
    range_match = FILENAME_DATE_RANGE_PATTERN.search(base)
    if range_match:
        title = range_match.group(1)
        title = re.sub(r"\s*[-/.]\s*", "-", title)
        title = re.sub(r"\s+al\s+", " al ", title, flags=re.IGNORECASE)
        return _sanitize_excel_sheet_name(title.strip())

    title = base
    while True:
        stripped = _VENDOR_PREFIX_PATTERN.sub("", title).strip(" _-.")
        if stripped == title:
            break
        title = stripped
    return _sanitize_excel_sheet_name(title or base)


def _unique_sheet_title(workbook, desired_title, used_titles):
    """Return a sheet title unique within the workbook (Excel + used set)."""
    base = _sanitize_excel_sheet_name(desired_title)
    candidate = base
    suffix = 2
    while candidate in used_titles or candidate in workbook.sheetnames:
        tail = f"_{suffix}"
        candidate = f"{base[: 31 - len(tail)]}{tail}"
        suffix += 1
    used_titles.add(candidate)
    return candidate


def _normalize_pdf_paths(pdf_paths):
    normalized = []
    seen = set()
    for raw in pdf_paths:
        path = os.path.abspath(str(raw).strip())
        if not path or path in seen:
            continue
        seen.add(path)
        normalized.append(path)
    return normalized


def _build_output_path(pdf_paths, vendor_label):
    base_dir = os.path.dirname(os.path.abspath(pdf_paths[0]))
    if len(pdf_paths) == 1:
        filename = f"{vendor_label} Fuel Coupons.xlsx"
    else:
        filename = f"{vendor_label} Fuel Coupons Batch.xlsx"
    return os.path.join(base_dir, filename)


def _launch_workbook(path):
    """Open the saved workbook with a deferred, platform-aware launch."""
    abs_path = os.path.abspath(path)
    if sys.platform == "win32":
        time.sleep(0.35)
        os.startfile(abs_path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        os.system(f'open "{abs_path}"')
    else:
        os.system(f'xdg-open "{abs_path}"')


def _write_date_cell(worksheet, row, column, block):
    """Write a left-aligned MM/DD/YYYY date string (blank when the block is empty)."""
    cell = worksheet.cell(row=row, column=column)
    if block is not None and block.get("date") is not None:
        cell.value = block["date"].strftime(DATE_OUTPUT_FORMAT)
    if LEFT_ALIGNMENT is not None:
        cell.alignment = LEFT_ALIGNMENT


def _write_number_cell(worksheet, row, column, value):
    """Write a right-aligned native float with the '0.00' decimal format."""
    cell = worksheet.cell(row=row, column=column)
    if value is not None:
        cell.value = float(value)
        cell.number_format = DECIMAL_NUMBER_FORMAT
    if RIGHT_ALIGNMENT is not None:
        cell.alignment = RIGHT_ALIGNMENT


def _write_tracking_cell(worksheet, row, column, block):
    """Write the right-aligned 6-digit tracking value (kept for verification)."""
    cell = worksheet.cell(row=row, column=column)
    tracking = block.get("tracking") if block is not None else ""
    if tracking:
        cell.value = tracking
    if RIGHT_ALIGNMENT is not None:
        cell.alignment = RIGHT_ALIGNMENT


def _write_layout_row(worksheet, row, row_data):
    """Write one PDF line into its strict 8-column Excel row (A-H)."""
    left = row_data.get("left")
    right = row_data.get("right")

    # Left block: A Date | B Amount | C Tracking | D Gallons
    _write_date_cell(worksheet, row, 1, left)
    _write_number_cell(worksheet, row, 2, left.get("amount") if left else None)
    _write_tracking_cell(worksheet, row, 3, left)
    _write_number_cell(worksheet, row, 4, left.get("gallons") if left else None)

    # Right block: E Date | F Amount | G Tracking | H Gallons
    _write_date_cell(worksheet, row, 5, right)
    _write_number_cell(worksheet, row, 6, right.get("amount") if right else None)
    _write_tracking_cell(worksheet, row, 7, right)
    _write_number_cell(worksheet, row, 8, right.get("gallons") if right else None)


_LAYOUT_HEADERS = (
    "Date",
    "Amount",
    "Tracking",
    "Gallons",
    "Date",
    "Amount",
    "Tracking",
    "Gallons",
)


def _write_sheet_headers(worksheet):
    for column, label in enumerate(_LAYOUT_HEADERS, start=1):
        cell = worksheet.cell(row=1, column=column, value=label)
        if HEADER_FONT is not None:
            cell.font = HEADER_FONT


def _write_pages_to_worksheet(worksheet, pages):
    """
    Write parsed pages to one worksheet (8-column A-H, 1 blank row between pages).

    Returns:
        int: number of data rows written.
    """
    current_row = 2
    record_count = 0
    for page_index, page_rows in enumerate(pages):
        if page_index > 0:
            current_row += 1
        for row_data in page_rows:
            _write_layout_row(worksheet, current_row, row_data)
            current_row += 1
            record_count += 1
    return record_count


def write_multi_sheet_fuel_coupon_workbook(pdf_paths, output_path):
    """
    Parse each PDF into its own worksheet (hoja por archivo).

    Sheet names are derived from each file's date-range text in the filename.
    Returns:
        tuple[int, int]: (total data rows written, number of worksheets created)
    """
    _ensure_openpyxl()
    workbook = Workbook()
    used_titles = set()
    record_count = 0
    sheets_written = 0
    first_sheet = True
    all_page_texts = []

    for pdf_path in pdf_paths:
        _vendor, _sheet_name, pages, page_texts = parse_fuel_coupons_pdf(pdf_path)
        all_page_texts.extend(page_texts)
        if not pages:
            continue

        desired_title = extract_sheet_title_from_filename(pdf_path)
        sheet_title = _unique_sheet_title(workbook, desired_title, used_titles)

        if first_sheet:
            worksheet = workbook.active
            worksheet.title = sheet_title
            first_sheet = False
        else:
            worksheet = workbook.create_sheet(title=sheet_title)

        _write_sheet_headers(worksheet)
        record_count += _write_pages_to_worksheet(worksheet, pages)
        sheets_written += 1

    if record_count == 0:
        workbook.close()
        raise ValueError("No fuel coupon records found in the selected PDF(s).")

    workbook.save(output_path)
    workbook.close()
    vendor_label, _ = detect_vendor(all_page_texts)
    return record_count, sheets_written, vendor_label


def generate_fuel_coupon_workbook_from_pdfs(pdf_paths, output_path=None, launch=True):
    """
    Parse multiple PDFs into one workbook (one sheet per file).

    Returns:
        tuple[str, int, int]: (saved path, total data rows, sheet count)
    """
    pdf_paths = _normalize_pdf_paths(pdf_paths)
    if not pdf_paths:
        raise ValueError("No PDF paths provided.")

    if output_path is None:
        probe_texts = []
        for pdf_path in pdf_paths:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    probe_texts.append(page.extract_text() or "")
        vendor_label, _ = detect_vendor(probe_texts)
        output_path = _build_output_path(pdf_paths, vendor_label)

    record_count, sheet_count, _vendor_label = write_multi_sheet_fuel_coupon_workbook(
        pdf_paths, output_path
    )
    if launch:
        _launch_workbook(output_path)
    return output_path, record_count, sheet_count


_SOURCE_FIRST_DATA_ROW = 3
_GETTEL_TOYOTA_DEST_SHEET_PATTERN = re.compile(
    r"^Gettel-Toyota\s+\d{2}\.\d{4}$",
    re.IGNORECASE,
)
_DAILY_SHEET_HEADER_SCAN_ROWS = 5
_DAILY_SHEET_COLUMN_TOKENS = {
    "date": ("date",),
    "amount": ("amount",),
    "gallons": ("gallon",),
}


def _parse_excel_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.replace(hour=0, minute=0, second=0, microsecond=0)
    if isinstance(value, (int, float)) and from_excel is not None:
        try:
            converted = from_excel(value)
            if isinstance(converted, datetime):
                return converted.replace(hour=0, minute=0, second=0, microsecond=0)
        except (ValueError, OSError, TypeError):
            pass
    text = str(value).strip()
    if not text:
        return None
    return _parse_date(text)


def _date_dict_key(value):
    parsed = _parse_excel_date(value)
    if parsed is None:
        return None
    return parsed.date()


def _identify_vendor_sheets(workbook):
    """
    Find the Gettel and Toyota daily fuel-log sheets by name.

    The real source file (copy/pasted by hand from the station manager's
    reports) names its tabs "GETTLE"/"GETTEL"/"Gettel..." and "Toyota...",
    not the Excel-default "Sheet1"/"Sheet2" — so this matches by substring,
    tolerant of the "GETTLE" typo, instead of requiring exact default names.
    """
    gettel_sheet = None
    toyota_sheet = None
    for worksheet in workbook.worksheets:
        title = worksheet.title.strip().upper()
        if gettel_sheet is None and any(token in title for token in GETTEL_TOKENS):
            gettel_sheet = worksheet
        elif toyota_sheet is None and any(token in title for token in TOYOTA_TOKENS):
            toyota_sheet = worksheet

    if gettel_sheet is None or toyota_sheet is None:
        available = ", ".join(workbook.sheetnames) or "(none)"
        raise ValueError(
            "Could not find both a Gettel and a Toyota sheet in the source "
            f"workbook. Available sheets: {available}"
        )
    if gettel_sheet.title == toyota_sheet.title:
        raise ValueError("Gettel and Toyota resolved to the same worksheet.")
    return gettel_sheet, toyota_sheet


def _find_daily_sheet_header_row(worksheet):
    """Locate the header row carrying Date/Amount/Gallons labels."""
    max_col = max(worksheet.max_column, 4)
    for row in range(1, _DAILY_SHEET_HEADER_SCAN_ROWS + 1):
        labels = [
            str(worksheet.cell(row=row, column=col).value or "").strip().lower()
            for col in range(1, max_col + 1)
        ]
        if any("date" in label for label in labels) and any(
            "amount" in label for label in labels
        ):
            return row
    raise ValueError(
        f"Could not find a Date/Amount header row in sheet '{worksheet.title}'."
    )


def _map_daily_sheet_columns(worksheet, header_row):
    """Map Date/Amount/Gallons to column indices (Tracking is dropped)."""
    columns = {}
    max_col = max(worksheet.max_column, 4)
    for col in range(1, max_col + 1):
        label = str(worksheet.cell(row=header_row, column=col).value or "").strip().lower()
        for key, tokens in _DAILY_SHEET_COLUMN_TOKENS.items():
            if key in columns:
                continue
            if any(token in label for token in tokens):
                columns[key] = col
                break
    missing = [key for key in ("date", "amount", "gallons") if key not in columns]
    if missing:
        raise ValueError(
            f"Sheet '{worksheet.title}' is missing column(s): {', '.join(missing)}."
        )
    return columns


def _parse_daily_sheet_number(value):
    """
    Parse an Amount/Gallons cell from the daily fuel-log sheet.

    Amounts are pasted as Latin-formatted text like "$42,52" (comma is the
    decimal separator); Gallons are plain native numbers. Neither needs the
    PDF-OCR lost-decimal correction used elsewhere in this module — this
    data comes from a direct copy/paste into Excel, not text extraction.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("$", "").replace(" ", "")
    if not text:
        return None
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _summarize_daily_sheet(worksheet):
    """Sum Amount/Gallons per calendar day from a Gettel/Toyota daily-entry sheet."""
    header_row = _find_daily_sheet_header_row(worksheet)
    columns = _map_daily_sheet_columns(worksheet, header_row)
    totals_by_date = {}
    max_row = worksheet.max_row or header_row
    for row in range(header_row + 1, max_row + 1):
        day = _date_dict_key(worksheet.cell(row=row, column=columns["date"]).value)
        if day is None:
            continue
        amount = _parse_daily_sheet_number(
            worksheet.cell(row=row, column=columns["amount"]).value
        )
        gallons = _parse_daily_sheet_number(
            worksheet.cell(row=row, column=columns["gallons"]).value
        )
        if amount is None and gallons is None:
            continue
        bucket = totals_by_date.setdefault(day, {"amount": 0.0, "gallons": 0.0})
        bucket["amount"] += amount or 0.0
        bucket["gallons"] += gallons or 0.0
    return totals_by_date


def _summarize_origin_workbook(source_workbook):
    """Summarize the Gettel and Toyota daily fuel-log sheets independently."""
    gettel_sheet, toyota_sheet = _identify_vendor_sheets(source_workbook)
    gettel_totals = _summarize_daily_sheet(gettel_sheet)
    toyota_totals = _summarize_daily_sheet(toyota_sheet)
    return gettel_totals, toyota_totals


def _find_gettel_toyota_destination_sheet(workbook):
    if len(workbook.worksheets) >= 4:
        fourth = workbook.worksheets[3]
        if _GETTEL_TOYOTA_DEST_SHEET_PATTERN.match(fourth.title.strip()):
            return fourth
    for worksheet in workbook.worksheets:
        if _GETTEL_TOYOTA_DEST_SHEET_PATTERN.match(worksheet.title.strip()):
            return worksheet
    raise ValueError(
        "Destination workbook has no sheet matching 'Gettel-Toyota MM.YYYY'."
    )


def _write_summary_number_cell(worksheet, row, column, value):
    """
    Write only the numeric value — the master template already carries the
    correct number format and alignment for this cell, and must not be
    overridden.
    """
    worksheet.cell(row=row, column=column).value = float(value)


def _populate_gettel_toyota_destination_sheet(worksheet, gettel_totals, toyota_totals):
    """
    Match both Gettel and Toyota daily totals against Column A's date.

    The DIF formula on this sheet (=C{row}-E{row}-G{row}) compares Local
    Account, Gettel, and Toyota all on the same row, so both vendors must
    land on the row identified by Column A — Column B is a separate,
    unrelated date and is never used for matching.
    """
    rows_matched = 0
    max_row = worksheet.max_row or _SOURCE_FIRST_DATA_ROW
    for row in range(_SOURCE_FIRST_DATA_ROW, max_row + 1):
        day = _date_dict_key(worksheet.cell(row=row, column=1).value)
        if day is None:
            continue
        row_matched = False

        if day in gettel_totals:
            totals = gettel_totals[day]
            _write_summary_number_cell(worksheet, row, 5, totals["amount"])
            _write_summary_number_cell(worksheet, row, 6, totals["gallons"])
            row_matched = True

        if day in toyota_totals:
            totals = toyota_totals[day]
            _write_summary_number_cell(worksheet, row, 7, totals["amount"])
            _write_summary_number_cell(worksheet, row, 8, totals["gallons"])
            row_matched = True

        if row_matched:
            rows_matched += 1

    return rows_matched


def merge_gettel_toyota_into_master(source_path, destination_path, launch=True):
    """
    Summarize the origin Excel's Gettel and Toyota daily fuel-log sheets and
    merge the daily totals into the destination sheet (Gettel-Toyota
    MM.YYYY). Opens a temp preview copy.

    Returns:
        tuple[str, int, int, int]: (preview path, rows matched, gettel days, toyota days)
    """
    import shutil
    import tempfile

    _ensure_openpyxl()
    source_path = os.path.abspath(source_path)
    destination_path = os.path.abspath(destination_path)
    if not os.path.isfile(source_path):
        raise FileNotFoundError(f"Source Excel not found: {source_path}")
    if not os.path.isfile(destination_path):
        raise FileNotFoundError(f"Destination Excel not found: {destination_path}")

    source_workbook = load_workbook(source_path, data_only=True)
    try:
        gettel_totals, toyota_totals = _summarize_origin_workbook(source_workbook)
    finally:
        source_workbook.close()

    if not gettel_totals and not toyota_totals:
        raise ValueError("No coupon data found in the origin Excel file.")

    preview_handle = tempfile.NamedTemporaryFile(
        suffix=".xlsx",
        prefix="GettelToyotaMasterPreview_",
        delete=False,
    )
    preview_path = preview_handle.name
    preview_handle.close()
    shutil.copy2(destination_path, preview_path)

    destination_workbook = load_workbook(preview_path)
    try:
        target_sheet = _find_gettel_toyota_destination_sheet(destination_workbook)
        rows_matched = _populate_gettel_toyota_destination_sheet(
            target_sheet, gettel_totals, toyota_totals
        )
        destination_workbook.save(preview_path)
    finally:
        destination_workbook.close()

    if launch:
        _launch_workbook(preview_path)

    return preview_path, rows_matched, len(gettel_totals), len(toyota_totals)


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv:
        pdf_paths = argv
    else:
        try:
            import tkinter as tk
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()
            pdf_paths = filedialog.askopenfilenames(
                title="Select fuel coupon PDF report(s)",
                filetypes=[("PDF Files", "*.pdf")],
            )
            root.destroy()
        except Exception:
            return 1

    if not pdf_paths:
        return 1

    generate_fuel_coupon_workbook_from_pdfs(list(pdf_paths))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
