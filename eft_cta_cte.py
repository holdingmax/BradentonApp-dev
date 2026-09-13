"""
EFT PDF -> Cta Cte J.H.Williams engine: parses the bank EFT PDF (credit-card
coupons and paid-invoice references). Pura -- ya no escribe ningún Excel
(2026-09-12, ver CLAUDE.md "EFT y Cupones -- convertido a Carga de Datos"):
lo que antes era update_excel_workbook/eft_already_loaded_in_workbook (y
todos sus helpers de estilo/formato de openpyxl) se eliminó del todo, junto
con las constantes de layout de columnas que solo esas funciones usaban.
extract_eft_data sigue siendo el único punto de entrada, ahora alimentando
eft_db.py en vez de un Ledger.
"""

import re
from datetime import datetime

import pdfplumber

EFT_DUPLICATE_ALERT = (
    "Alerta: Este EFT ya fue cargado anteriormente con los mismos datos."
)

CREDIT_COUPON_PATTERN = re.compile(
    r"\b(SI-\d+)/(DDC-\d+)\b", re.IGNORECASE
)
SI_INVOICE_PATTERN = re.compile(r"\b(SI-\d+)\b", re.IGNORECASE)
DDC_COUPON_PATTERN = re.compile(r"\b(DDC-\d+)\b", re.IGNORECASE)
DRAFT_NO_PATTERN = re.compile(r"\b(RCV-\d+)\b", re.IGNORECASE)
DATE_PATTERN = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{4})\b")
EFT_DATE_LABEL_PATTERN = re.compile(
    r"EFT\s*Date\s*[:\-]?\s*(\d{1,2}/\d{1,2}/\d{4})",
    re.IGNORECASE,
)
CURRENCY_DECIMAL_PATTERN = re.compile(r"\.\d{1,2}\b|,\d{2}\b")
CREDIT_TABLE_HEADER_WORDS = {"gross", "fees", "net", "paid"}


def parse_amount(value):
    """Parse a currency string into a positive float using abs()."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        try:
            return abs(float(value))
        except (TypeError, ValueError):
            return None

    text = str(value).strip()
    if not text or text.upper() in {"-", "--", "N/A", "NA"}:
        return None

    if text.startswith("(") and text.endswith(")"):
        text = text[1:-1].strip()

    # European decimal comma (14,33) -> 14.33
    if re.search(r",\d{2}\b", text) and "." not in text.split(",")[-1][:3]:
        text = text.replace(".", "").replace(",", ".")

    text = text.replace("$", "").replace(" ", "")
    if text.count(",") > 0 and "." not in text:
        parts = text.rsplit(",", 1)
        if len(parts) == 2 and parts[1].isdigit() and len(parts[1]) <= 2:
            text = parts[0].replace(",", "") + "." + parts[1]
        else:
            text = text.replace(",", "")
    else:
        text = text.replace(",", "")

    text = text.strip()
    if text.startswith("-"):
        text = text[1:].strip()
    elif text.endswith("-"):
        text = text[:-1].strip()

    if not text:
        return None

    try:
        return abs(float(text))
    except ValueError:
        return None


def parse_date_to_datetime(date_str):
    """Parse a date string into datetime (tries common PDF formats)."""
    if not date_str:
        return None
    text = str(date_str).strip()
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def format_coupon_date_us(date_str):
    """Column A: strict US format MM/DD/YYYY (e.g. 05/08/2026)."""
    parsed = parse_date_to_datetime(date_str)
    if parsed:
        return parsed.strftime("%m/%d/%Y")
    return date_str


def parse_eft_pdf_date_us(date_str):
    """Parse an EFT PDF date token strictly as MM/DD/YYYY (American)."""
    if not date_str:
        return None
    text = str(date_str).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def format_eft_date_us(date_str):
    """Header EFT date as MM/DD/YYYY string, matching the source PDF."""
    parsed = parse_eft_pdf_date_us(date_str)
    if parsed:
        return parsed.strftime("%m/%d/%Y")
    return date_str


def parse_date_cell(cell):
    """Extract date from a PDF cell; returns MM/DD/YYYY for coupon rows."""
    if not cell:
        return None
    match = DATE_PATTERN.search(str(cell).strip())
    if match:
        return format_coupon_date_us(match.group(1))
    return None


def normalize_row_cells(cells):
    """Return a safe list of stripped string cell values."""
    try:
        return [str(cell).strip() if cell else "" for cell in cells]
    except Exception:
        return []


def expand_row_components(cells):
    """
    Flatten row cells into scan components.

    Splits comma-separated reference blobs (e.g. 0764,0765,0766,9133)
    so they do not shift currency column positions.
    """
    components = []
    for cell in cells:
        if not cell:
            continue
        text = str(cell).strip()
        if not text:
            continue
        if is_reference_blob(text):
            for part in text.split(","):
                part = part.strip()
                if part:
                    components.append(part)
        else:
            components.append(text)
    return components


def is_reference_blob(text):
    """True for comma-separated tracking refs without currency markers."""
    if not text or "$" in text or "(" in text or ")" in text:
        return False
    if CURRENCY_DECIMAL_PATTERN.search(text):
        return False
    if "," not in text:
        return False
    parts = [p.strip() for p in text.split(",")]
    if len(parts) < 2:
        return False
    return all(
        p.replace(" ", "").isdigit() and 1 <= len(p.replace(" ", "")) <= 8
        for p in parts
    )


def is_tracking_reference(value):
    """Detect single tracking numbers (not currency)."""
    if value is None:
        return False
    text = str(value).strip()
    if not text:
        return False
    if is_reference_blob(text):
        return True
    if "$" in text or "(" in text or ")" in text:
        return False
    if CURRENCY_DECIMAL_PATTERN.search(text):
        return False
    cleaned = text.replace(",", "").replace(" ", "")
    if not cleaned.isdigit():
        return False
    return len(cleaned) >= 4


def is_currency_component(value):
    """
    True when a token looks like Gross/Fees/Paid money.

    Uses $, parentheses, or decimal patterns (dot or comma cents).
    """
    if value is None:
        return False
    text = str(value).strip()
    if not text:
        return False
    if is_reference_blob(text) or is_tracking_reference(text):
        return False
    if "$" in text or "(" in text or ")" in text:
        return True
    if re.search(r"-\s*\$|\$\s*-", text):
        return True
    if text.startswith("-") and re.search(r"\d", text):
        return True
    if CURRENCY_DECIMAL_PATTERN.search(text):
        return True
    if re.fullmatch(r"-?\d+[.,]\d{2}", text.replace("$", "").strip()):
        return True
    return False


def extract_last_three_financial_amounts(components):
    """
    Extract Gross, Fees, and Paid from currency tokens at the row tail.

    Variable reference columns are already removed via expand_row_components().
    When four amounts exist (Gross, Fees, Net, Paid), Net is skipped by taking
    the 4th-from-last, 3rd-from-last, and last currency tokens.
    """
    currency_tokens = [c for c in components if is_currency_component(c)]
    if len(currency_tokens) < 3:
        return None, None, None

    if len(currency_tokens) >= 4:
        selected = [currency_tokens[-4], currency_tokens[-3], currency_tokens[-1]]
    else:
        selected = currency_tokens[-3:]

    amounts = [parse_amount(token) for token in selected]
    if any(amount is None for amount in amounts):
        return None, None, None
    return amounts[0], amounts[1], amounts[2]


def row_is_credit_table_header(cells):
    """True for the 'Gross Fees Net Paid' header row that opens the credit table."""
    try:
        words = {cell.strip().lower() for cell in cells if cell}
        return CREDIT_TABLE_HEADER_WORDS.issubset(words)
    except Exception:
        return False


def row_contains_credit_coupon(cells):
    """True when the row is a coupon table row."""
    try:
        row_text = " ".join(cells)
        if CREDIT_COUPON_PATTERN.search(row_text):
            return True
        has_si = any(SI_INVOICE_PATTERN.search(c) for c in cells)
        has_ddc = any(DDC_COUPON_PATTERN.search(c) for c in cells)
        return has_si and has_ddc
    except Exception:
        return False


def locate_coupon_combo(cells):
    """Find invoice, coupon, and combo cell index."""
    for index, cell in enumerate(cells):
        match = CREDIT_COUPON_PATTERN.search(cell)
        if match:
            return match.group(1).upper(), match.group(2).upper(), index

    invoice = coupon = None
    si_index = ddc_index = None
    for index, cell in enumerate(cells):
        si_match = SI_INVOICE_PATTERN.search(cell)
        if si_match and "/" not in cell:
            invoice = si_match.group(1).upper()
            si_index = index
        ddc_match = DDC_COUPON_PATTERN.search(cell)
        if ddc_match:
            coupon = ddc_match.group(0).upper()
            ddc_index = index

    if invoice and coupon:
        return invoice, coupon, max(si_index or 0, ddc_index or 0)

    row_text = " ".join(cells)
    match = CREDIT_COUPON_PATTERN.search(row_text)
    if match:
        return match.group(1).upper(), match.group(2).upper(), 1

    # Fila dentro de la tabla de créditos sin DDC- (no todos los pagos
    # traen coupon combo) — se conserva igual, con coupon vacío.
    if invoice:
        return invoice, None, si_index or 0

    return None, None, None


def extract_coupon_columns(cells, fallback_date=None):
    """Parse coupon row using last-three currency detection."""
    cells = normalize_row_cells(cells)
    if not cells:
        return None

    try:
        invoice, coupon, combo_index = locate_coupon_combo(cells)
        if not invoice:
            return None

        components = expand_row_components(cells)
        gross, fees, paid = extract_last_three_financial_amounts(components)
        if gross is None:
            return None

        row_date = None
        if cells:
            row_date = parse_date_cell(cells[0])
        if not row_date and combo_index > 0:
            row_date = parse_date_cell(cells[combo_index - 1])
        if not row_date:
            for cell in cells:
                row_date = parse_date_cell(cell)
                if row_date:
                    break
        if not row_date:
            match = DATE_PATTERN.search(" ".join(cells))
            if match:
                row_date = format_coupon_date_us(match.group(1))
        if not row_date and fallback_date:
            row_date = format_coupon_date_us(fallback_date)

        fees_value = fees if fees is not None else 0.0

        return {
            "date": row_date,
            "invoice": invoice,
            "coupon": coupon,
            "gross_amount": gross,
            "fees_amount": fees_value,
            "paid_amount": paid if paid is not None else 0.0,
        }
    except Exception:
        return None


def _si_match_inside_credit_combo(text, match):
    """True when an SI token belongs to an SI/DDC combo reference."""
    span_start = max(0, match.start() - 12)
    span_end = min(len(text), match.end() + 24)
    return bool(CREDIT_COUPON_PATTERN.search(text[span_start:span_end]))


def extract_paid_invoice_entries(cells):
    """
    Parse one PDF row for all paid invoices at the EFT header (ordered list).

    Collects every SI- reference on the row (not part of an SI/DDC coupon combo)
    and pairs each with a currency amount when available.
    """
    cells = normalize_row_cells(cells)
    if not cells or row_contains_credit_coupon(cells):
        return []

    try:
        invoices_ordered = []
        seen_invoices = set()
        for cell in cells:
            if CREDIT_COUPON_PATTERN.search(cell):
                continue
            for match in SI_INVOICE_PATTERN.finditer(cell):
                if _si_match_inside_credit_combo(cell, match):
                    continue
                invoice = match.group(1).upper()
                if invoice in seen_invoices:
                    continue
                seen_invoices.add(invoice)
                invoices_ordered.append(invoice)

        if not invoices_ordered:
            row_text = " ".join(cells)
            if not SI_INVOICE_PATTERN.search(row_text):
                return []
            for match in SI_INVOICE_PATTERN.finditer(row_text):
                if _si_match_inside_credit_combo(row_text, match):
                    continue
                invoice = match.group(1).upper()
                if invoice in seen_invoices:
                    continue
                seen_invoices.add(invoice)
                invoices_ordered.append(invoice)

        if not invoices_ordered:
            return []

        components = expand_row_components(cells)
        amounts = []
        for token in components:
            if not is_currency_component(token):
                continue
            paid_amount = parse_amount(token)
            if paid_amount is not None:
                amounts.append(paid_amount)
        if not amounts:
            return []

        if len(amounts) >= len(invoices_ordered):
            paired_amounts = amounts[-len(invoices_ordered) :]
        elif len(invoices_ordered) == 1:
            paired_amounts = amounts[-1:]
        else:
            return []

        return [
            {"invoice": invoice, "paid_amount": paid_amount}
            for invoice, paid_amount in zip(invoices_ordered, paired_amounts)
        ]
    except Exception:
        return []


def collect_pdf_rows(pdf_path):
    """Extract table rows and text lines from every PDF page."""
    full_text_parts = []
    rows = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            if page_text:
                full_text_parts.append(page_text)

            for table in page.extract_tables() or []:
                for row in table:
                    if not row or not any(cell for cell in row):
                        continue
                    rows.append(normalize_row_cells(row))

            for line in page_text.splitlines():
                line = line.strip()
                if not line:
                    continue
                try:
                    if "\t" in line:
                        cells = [p.strip() for p in line.split("\t") if p.strip()]
                    else:
                        cells = re.split(r"\s{2,}", line)
                        if len(cells) <= 1:
                            cells = line.split()
                    if cells:
                        rows.append(cells)
                except Exception:
                    continue

    return "\n".join(full_text_parts), rows


def extract_header(full_text):
    """Extract EFT header date (MM/DD/YYYY) and draft number."""
    header_data = {"eft_date": None, "draft_no": None}
    try:
        draft_match = DRAFT_NO_PATTERN.search(full_text)
        if draft_match:
            header_data["draft_no"] = draft_match.group(1).upper()

        date_match = EFT_DATE_LABEL_PATTERN.search(full_text)
        if not date_match:
            date_match = DATE_PATTERN.search(full_text)
        if date_match:
            header_data["eft_date"] = format_eft_date_us(date_match.group(1))
    except Exception:
        pass
    return header_data


def extract_eft_data(pdf_path):
    """
    Read EFT PDF and return (header_data, paid_invoices, credit_coupons, skipped_coupon_rows).

    `skipped_coupon_rows` counts rows that visibly look like a coupon line
    (an SI-/DDC- combo, per row_contains_credit_coupon) but that
    extract_coupon_columns still couldn't turn into a usable row (typically
    because it couldn't find 3+ legible currency tokens) -- previously these
    were dropped with zero tracking, so a partially-corrupted table silently
    lost coupons with no way to know without recounting the PDF by hand.
    Rows that simply aren't coupon rows at all (footer text, boilerplate)
    are NOT counted here, only ones that positively matched the coupon
    pattern and still failed to parse.
    """
    full_text, rows = collect_pdf_rows(pdf_path)
    header_data = extract_header(full_text)
    fallback_date = header_data.get("eft_date")

    paid_invoices = []
    credit_coupons = []
    seen_invoices = set()
    seen_coupons = set()
    coupon_section_started = False
    skipped_coupon_rows = 0

    for cells in rows:
        try:
            if row_is_credit_table_header(cells):
                coupon_section_started = True
                continue

            looks_like_coupon_row = row_contains_credit_coupon(cells)
            if looks_like_coupon_row:
                coupon_section_started = True

            if coupon_section_started:
                coupon_row = extract_coupon_columns(cells, fallback_date=fallback_date)
                if coupon_row:
                    key = (
                        coupon_row["date"],
                        coupon_row["invoice"],
                        coupon_row["coupon"],
                        coupon_row["gross_amount"],
                        coupon_row["fees_amount"],
                        coupon_row["paid_amount"],
                    )
                    if key not in seen_coupons:
                        seen_coupons.add(key)
                        credit_coupons.append(coupon_row)
                elif looks_like_coupon_row:
                    skipped_coupon_rows += 1
                continue

            for paid_entry in extract_paid_invoice_entries(cells):
                key = (paid_entry["invoice"], paid_entry["paid_amount"])
                if key not in seen_invoices:
                    seen_invoices.add(key)
                    paid_invoices.append(paid_entry)
        except Exception:
            continue

    return header_data, paid_invoices, credit_coupons, skipped_coupon_rows


