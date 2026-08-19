"""Append monthly J.H. Williams coupon reports to the Cupones worksheet."""

import gc
import os
import re
import sys
import tempfile
import time
from datetime import datetime

import pandas as pd

try:
    from openpyxl import load_workbook
    from openpyxl.cell.cell import MergedCell
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    OPENPYXL_AVAILABLE = True
except ImportError:
    load_workbook = None  # type: ignore[assignment,misc]
    MergedCell = None  # type: ignore[assignment,misc]
    Alignment = None  # type: ignore[assignment,misc]
    Font = None  # type: ignore[assignment,misc]
    get_column_letter = None  # type: ignore[assignment,misc]
    OPENPYXL_AVAILABLE = False

EFT_MASTER_FILENAME = "01.12.2022 al 31.05.2026 Aplicacion TC y EFT PRUEBA.xlsx"
MONTHLY_HEADER_ROW = 2
MONTHLY_DATA_START_ROW = 3

CUPONES_SHEET = "Cupones"
CTA_CTE_SHEET = "Cta Cte J.H.Wiliams"

CUPONES_COL_DATE = 1
CUPONES_COL_COUPON = 2
CUPONES_COL_GROSS = 3
CUPONES_COL_FEES = 4
CUPONES_COL_NET = 5
CUPONES_COL_EFT_FORMULA = 6
CUPONES_COL_DIF = 7
CUPONES_COL_EFT_NO = 8
CUPONES_COL_MES_EFT = 9

CTA_COL_INVOICE = 3
CTA_COL_GROSS = 4
CTA_COL_FEES = 5
CTA_COL_NET = 6
CTA_COL_NRO = 7
CTA_COL_NETO_FINAL = 14

CTA_SCAN_START_ROW = 5
CUPONES_SCAN_START_ROW = 2
DECIMAL_NUMBER_FORMAT = "0.00"
EFT_DATE_NUMBER_FORMAT = "dd-mmm"
COUPON_DATE_TEXT_PATTERN = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")
DDC_COUPON_PATTERN = re.compile(r"\b(DDC-\d+)\b", re.IGNORECASE)
RCV_PATTERN = re.compile(r"\b(RCV-\d+)\b", re.IGNORECASE)

LEFT_ALIGNMENT = Alignment(horizontal="left") if Alignment is not None else None
RIGHT_ALIGNMENT = Alignment(horizontal="right") if Alignment is not None else None
PRIMARY_SPLIT_RED_FONT = Font(color="FF0000") if Font is not None else None
CUPONES_LEFT_COLUMNS = (
    CUPONES_COL_DATE,
    CUPONES_COL_COUPON,
    CUPONES_COL_EFT_NO,
)
CUPONES_RIGHT_COLUMNS = (
    CUPONES_COL_GROSS,
    CUPONES_COL_FEES,
    CUPONES_COL_NET,
    CUPONES_COL_EFT_FORMULA,
    CUPONES_COL_DIF,
    CUPONES_COL_MES_EFT,
)

MONTHLY_REPORT_FULL_DUPLICATE_ALERT = (
    "Alerta: Este reporte mensual ya fue cargado en su totalidad anteriormente."
)
AMOUNT_MATCH_TOLERANCE = 0.01


class MonthlyReportFullyDuplicateError(Exception):
    """Raised when every incoming coupon row already exists on Cupones."""


def _ensure_openpyxl():
    if not OPENPYXL_AVAILABLE:
        raise ImportError(
            "Cupones append requires openpyxl. Install with: pip install openpyxl"
        )


def resolve_eft_master_workbook_path():
    """Resolve the fixed EFT/Cupones master workbook path on disk."""
    app_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(app_dir, EFT_MASTER_FILENAME),
        os.path.join(os.getcwd(), EFT_MASTER_FILENAME),
    ]
    for path in candidates:
        if os.path.isfile(path):
            return os.path.abspath(path)
    return os.path.abspath(candidates[0])


def _cell_for_write(worksheet, row, column):
    """Return the writable top-left cell when the target is inside a merge."""
    cell = worksheet.cell(row=row, column=column)
    if MergedCell is not None and isinstance(cell, MergedCell):
        for merged in worksheet.merged_cells.ranges:
            if (
                merged.min_row <= row <= merged.max_row
                and merged.min_col <= column <= merged.max_col
            ):
                return worksheet.cell(row=merged.min_row, column=merged.min_col)
    return cell


def _coerce_primitive_cell_value(value):
    """Return the underlying cell value, never an openpyxl cell object."""
    if value is None:
        return None
    type_name = type(value).__name__
    if type_name in {"ReadOnlyCell", "Cell", "MergedCell"}:
        return value.value
    if hasattr(value, "value") and not isinstance(
        value, (str, bytes, int, float, bool, datetime)
    ):
        try:
            return value.value
        except Exception:
            return value
    return value


def _safe_string_from_value(value):
    value = _coerce_primitive_cell_value(value)
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")
    return str(value).strip()


def _strip_cell(value):
    value = _coerce_primitive_cell_value(value)
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")
    return str(value).strip()


def _parse_amount(value):
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = _strip_cell(value).replace("$", "").replace(",", "").strip()
    if not text:
        return 0.0
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    try:
        return float(text)
    except ValueError:
        return 0.0


def _parse_date_to_datetime(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    text = _strip_cell(value)
    if not text:
        return None
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    try:
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.isna(parsed):
            return None
        return parsed.to_pydatetime()
    except Exception:
        return None


def _format_coupon_date_us(value):
    parsed = _parse_date_to_datetime(value)
    if parsed:
        return parsed.strftime("%m/%d/%Y")
    text = _strip_cell(value)
    return text if text else None


def _get_sheet(workbook, target_name):
    if target_name in workbook.sheetnames:
        return workbook[target_name]
    lowered = target_name.strip().lower()
    for name in workbook.sheetnames:
        if name.strip().lower() == lowered:
            return workbook[name]
    raise ValueError(
        f'Sheet "{target_name}" not found. Available: {", ".join(workbook.sheetnames)}'
    )


def _create_temp_workbook_path(suffix=".xlsx"):
    fd, temp_path = tempfile.mkstemp(suffix=suffix, prefix="cupones_append_")
    os.close(fd)
    return temp_path


def _launch_temp_workbook(temp_path):
    abs_path = os.path.abspath(temp_path)
    if sys.platform == "win32":
        time.sleep(0.35)
        os.startfile(abs_path)
    elif sys.platform == "darwin":
        os.system(f'open "{abs_path}"')
    else:
        os.system(f'xdg-open "{abs_path}"')


def _apply_decimal_format(worksheet, row, column, value):
    cell = _cell_for_write(worksheet, row, column)
    cell.value = float(value)
    cell.number_format = DECIMAL_NUMBER_FORMAT


def _record_sort_key(record):
    parsed = _parse_date_to_datetime(record.get("date"))
    if parsed:
        return parsed
    return datetime.max


def _sort_records_by_date(records):
    return sorted(records, key=_record_sort_key)


def _expand_records_by_coupon_split(records):
    """
    Split comma-separated Coupon IDs into one record per DDC code.

    First split row keeps report amounts; subsequent split rows are zeroed.
    """
    expanded = []
    split_group_counter = 0
    for record in records:
        coupon_ids = _split_coupon_ids(record.get("coupon"))
        if not coupon_ids:
            coupon_text = _strip_cell(record.get("coupon"))
            if coupon_text:
                expanded.append(
                    {
                        "date": record["date"],
                        "coupon": coupon_text,
                        "gross": 0.0,
                        "fees": 0.0,
                        "net": 0.0,
                        "is_primary_split": False,
                        "split_group_id": None,
                        "split_index": 0,
                    }
                )
            continue
        split_group_id = None
        if len(coupon_ids) > 1:
            split_group_counter += 1
            split_group_id = f"group_{split_group_counter}"
        is_multi_coupon = len(coupon_ids) > 1
        for idx, coupon_id in enumerate(coupon_ids):
            expanded.append(
                {
                    "date": record["date"],
                    "coupon": coupon_id,
                    "gross": _parse_amount(record.get("gross", 0)) if idx == 0 else 0.0,
                    "fees": _parse_amount(record.get("fees", 0)) if idx == 0 else 0.0,
                    "net": _parse_amount(record.get("net", 0)) if idx == 0 else 0.0,
                    "is_primary_split": is_multi_coupon and idx == 0,
                    "split_group_id": split_group_id,
                    "split_index": idx,
                }
            )
    return expanded


def _read_cta_financial_breakdown(worksheet, row):
    """
    Read Gross (D), Disc/Fees (E), and Net Amt (F) from one Cta Cte coupon row.
    """
    return (
        _parse_amount(worksheet.cell(row=row, column=CTA_COL_GROSS).value),
        _parse_amount(worksheet.cell(row=row, column=CTA_COL_FEES).value),
        _parse_amount(worksheet.cell(row=row, column=CTA_COL_NET).value),
    )


def _eft_cell_is_blank(cell):
    value = cell.value
    if value is None:
        return True
    if isinstance(value, str) and not _strip_cell(value):
        return True
    return False


def _cupones_row_has_coupon(worksheet, row):
    return bool(_strip_cell(worksheet.cell(row=row, column=CUPONES_COL_COUPON).value))


def _resolve_single_coupon_id(coupon_text):
    """Return the one Coupon ID for the current Cupones row (Column B)."""
    coupon_ids = _split_coupon_ids(coupon_text)
    if coupon_ids:
        return coupon_ids[0]
    text = _strip_cell(coupon_text).upper()
    return text if text else None


def _resolve_coupon_id_for_sync(coupon_text, preferred_ids=None):
    """
    Resolve which coupon ID to sync for a Cupones row.

    If the cell contains multiple DDC IDs (historical fused rows), prefer one that
    exists in preferred_ids; otherwise return the first parsed ID.
    """
    preferred_ids = preferred_ids or set()
    coupon_ids = _split_coupon_ids(coupon_text)
    for coupon_id in coupon_ids:
        if coupon_id in preferred_ids:
            return coupon_id
    return coupon_ids[0] if coupon_ids else _resolve_single_coupon_id(coupon_text)


def _find_single_cta_row(coupon_text, coupon_index):
    """Locate the single Cta Cte row whose Column C matches any coupon ID on this row."""
    coupon_ids = _split_coupon_ids(coupon_text)
    if not coupon_ids:
        single = _resolve_single_coupon_id(coupon_text)
        if single:
            coupon_ids = [single]
    for coupon_id in coupon_ids:
        rows = coupon_index.get(coupon_id, [])
        if rows:
            return rows[0]
    return None


def _write_cupones_date_cell(worksheet, row, record):
    """Write Column A as Excel datetime (mm/dd/yyyy) or normalized US date text."""
    cell_a = _cell_for_write(worksheet, row, CUPONES_COL_DATE)
    raw = record.get("date")
    parsed = raw if isinstance(raw, datetime) else _parse_date_to_datetime(raw)
    if parsed:
        cell_a.value = parsed
        cell_a.number_format = "mm/dd/yyyy"
        return

    formatted = _format_coupon_date_us(raw)
    if formatted:
        retry = _parse_date_to_datetime(formatted)
        if retry:
            cell_a.value = retry
        else:
            cell_a.value = formatted
        cell_a.number_format = "mm/dd/yyyy"


def _cupones_append_start_row(worksheet):
    """First append row strictly after the sheet's absolute max_row."""
    return max(worksheet.max_row, 0) + 1


def _write_cupones_base_columns(worksheet, row, record):
    _write_cupones_date_cell(worksheet, row, record)

    coupon_text = _strip_cell(record["coupon"])
    _cell_for_write(worksheet, row, CUPONES_COL_COUPON).value = coupon_text

    _apply_decimal_format(worksheet, row, CUPONES_COL_GROSS, record["gross"])
    _apply_decimal_format(worksheet, row, CUPONES_COL_FEES, record["fees"])
    _apply_decimal_format(worksheet, row, CUPONES_COL_NET, record["net"])
    return coupon_text


def _clear_cupones_control_columns(worksheet, row):
    """Initialize F-I control columns before Cta Cte cross-reference."""
    cell_f = _cell_for_write(worksheet, row, CUPONES_COL_EFT_FORMULA)
    cell_f.value = ""
    cell_f.alignment = RIGHT_ALIGNMENT
    cell_f.number_format = DECIMAL_NUMBER_FORMAT

    cell_g = _cell_for_write(worksheet, row, CUPONES_COL_DIF)
    cell_g.value = f"=E{row}-F{row}"
    cell_g.alignment = RIGHT_ALIGNMENT
    cell_g.number_format = DECIMAL_NUMBER_FORMAT

    cell_h = _cell_for_write(worksheet, row, CUPONES_COL_EFT_NO)
    cell_h.value = ""
    cell_h.alignment = LEFT_ALIGNMENT

    cell_i = _cell_for_write(worksheet, row, CUPONES_COL_MES_EFT)
    cell_i.value = ""
    cell_i.alignment = RIGHT_ALIGNMENT


def _amounts_close(left, right, tolerance=AMOUNT_MATCH_TOLERANCE):
    try:
        return abs(float(left) - float(right)) <= tolerance
    except (TypeError, ValueError):
        return False


def _record_amount_tuple(record):
    return (
        _parse_amount(record.get("gross")),
        _parse_amount(record.get("fees")),
        _parse_amount(record.get("net")),
    )


def _build_existing_cupones_entries(worksheet):
    """Index existing Cupones coupon IDs with their C/D/E metrics."""
    entries = []
    max_row = max(worksheet.max_row, CUPONES_SCAN_START_ROW)
    for row in range(CUPONES_SCAN_START_ROW, max_row + 1):
        coupon_text = _strip_cell(worksheet.cell(row=row, column=CUPONES_COL_COUPON).value)
        if not coupon_text:
            continue
        coupon_ids = _split_coupon_ids(coupon_text)
        if not coupon_ids:
            single = _strip_cell(coupon_text).upper()
            if single:
                coupon_ids = [single]
        metrics = (
            _parse_amount(worksheet.cell(row=row, column=CUPONES_COL_GROSS).value),
            _parse_amount(worksheet.cell(row=row, column=CUPONES_COL_FEES).value),
            _parse_amount(worksheet.cell(row=row, column=CUPONES_COL_NET).value),
        )
        for coupon_id in coupon_ids:
            entries.append({"coupon_id": coupon_id, "metrics": metrics})
    return entries


def _cupones_record_is_duplicate(existing_entries, record):
    """
    True when coupon ID already exists on Cupones with matching metrics.

    Split child rows that arrive as 0.00 are treated as duplicates when the
    coupon ID is already present on the sheet.
    """
    coupon_id = _strip_cell(record.get("coupon")).upper()
    if not coupon_id:
        return False

    incoming_metrics = _record_amount_tuple(record)
    incoming_all_zero = all(abs(value) < 1e-9 for value in incoming_metrics)

    for entry in existing_entries:
        if entry["coupon_id"] != coupon_id:
            continue
        if incoming_all_zero:
            return True
        if all(
            _amounts_close(incoming_value, existing_value)
            for incoming_value, existing_value in zip(
                incoming_metrics, entry["metrics"]
            )
        ):
            return True
    return False


def _filter_new_cupones_records(cupones_sheet, monthly_rows):
    """Return only incoming rows that are not already present on Cupones."""
    existing_entries = _build_existing_cupones_entries(cupones_sheet)
    filtered = []
    skipped = 0
    for record in monthly_rows:
        if _cupones_record_is_duplicate(existing_entries, record):
            skipped += 1
            continue
        filtered.append(record)
    return filtered, skipped


def _is_zero_amount(value):
    try:
        return abs(float(value)) < 1e-9
    except (TypeError, ValueError):
        return False


def _row_amounts_are_zero(cupones_sheet, row):
    gross = cupones_sheet.cell(row=row, column=CUPONES_COL_GROSS).value
    fees = cupones_sheet.cell(row=row, column=CUPONES_COL_FEES).value
    net = cupones_sheet.cell(row=row, column=CUPONES_COL_NET).value
    return _is_zero_amount(gross) and _is_zero_amount(fees) and _is_zero_amount(net)


def _aggregate_cta_totals_for_coupon_id(cta_sheet, coupon_id):
    """Sum all Cta Cte D/E/F amounts for one coupon ID (scan row 2..max)."""
    target = _strip_cell(coupon_id).upper()
    if not target:
        return 0.0, 0.0, 0.0

    total_gross = 0.0
    total_fees = 0.0
    total_net = 0.0
    max_row = max(cta_sheet.max_row, 2)

    for scan_row in range(2, max_row + 1):
        invoice_text = cta_sheet.cell(row=scan_row, column=CTA_COL_INVOICE).value
        row_coupon_ids = _extract_ddc_ids(invoice_text)
        if target not in row_coupon_ids:
            continue
        gross, fees, net = _read_cta_financial_breakdown(cta_sheet, scan_row)
        total_gross += float(gross or 0.0)
        total_fees += float(fees or 0.0)
        total_net += float(net or 0.0)

    return total_gross, total_fees, total_net


def _aggregate_cta_totals_for_coupon_text(cta_sheet, coupon_text):
    """Aggregate totals for all coupon IDs present in one Cupones coupon cell."""
    coupon_ids = _split_coupon_ids(coupon_text)
    if not coupon_ids:
        single = _resolve_single_coupon_id(coupon_text)
        coupon_ids = [single] if single else []

    total_gross = 0.0
    total_fees = 0.0
    total_net = 0.0
    for coupon_id in dict.fromkeys(coupon_ids):
        gross, fees, net = _aggregate_cta_totals_for_coupon_id(cta_sheet, coupon_id)
        total_gross += gross
        total_fees += fees
        total_net += net
    return total_gross, total_fees, total_net


def _fill_zero_amounts_from_cta(cupones_sheet, cta_sheet, row, coupon_text):
    """Inject static C/D/E floats from aggregated Cta Cte matches when C/D/E are zero."""
    if not _row_amounts_are_zero(cupones_sheet, row):
        return
    gross, fees, net = _aggregate_cta_totals_for_coupon_text(cta_sheet, coupon_text)
    _apply_decimal_format(cupones_sheet, row, CUPONES_COL_GROSS, gross)
    _apply_decimal_format(cupones_sheet, row, CUPONES_COL_FEES, fees)
    _apply_decimal_format(cupones_sheet, row, CUPONES_COL_NET, net)


def _highlight_primary_split_amounts_red(cupones_sheet, row):
    if PRIMARY_SPLIT_RED_FONT is None:
        return
    for col in (CUPONES_COL_GROSS, CUPONES_COL_FEES, CUPONES_COL_NET):
        _cell_for_write(cupones_sheet, row, col).font = PRIMARY_SPLIT_RED_FONT


def _apply_child_rows_standard_black(cupones_sheet, row):
    if Font is None:
        return
    black_font = Font(color="000000")
    for col in (CUPONES_COL_GROSS, CUPONES_COL_FEES, CUPONES_COL_NET):
        _cell_for_write(cupones_sheet, row, col).font = black_font


def _read_float_cell(cupones_sheet, row, column):
    return _parse_amount(cupones_sheet.cell(row=row, column=column).value)


def _apply_parent_child_subtractions(cupones_sheet, split_group_rows):
    """
    Phase 2: subtract all child C/D/E totals from each split-group parent row.
    """
    for rows in split_group_rows.values():
        if len(rows) <= 1:
            continue
        parent_row = rows[0]
        child_rows = rows[1:]

        parent_gross = _read_float_cell(cupones_sheet, parent_row, CUPONES_COL_GROSS)
        parent_fees = _read_float_cell(cupones_sheet, parent_row, CUPONES_COL_FEES)
        parent_net = _read_float_cell(cupones_sheet, parent_row, CUPONES_COL_NET)

        child_gross = sum(
            _read_float_cell(cupones_sheet, child_row, CUPONES_COL_GROSS)
            for child_row in child_rows
        )
        child_fees = sum(
            _read_float_cell(cupones_sheet, child_row, CUPONES_COL_FEES)
            for child_row in child_rows
        )
        child_net = sum(
            _read_float_cell(cupones_sheet, child_row, CUPONES_COL_NET)
            for child_row in child_rows
        )

        _apply_decimal_format(
            cupones_sheet, parent_row, CUPONES_COL_GROSS, parent_gross - child_gross
        )
        _apply_decimal_format(
            cupones_sheet, parent_row, CUPONES_COL_FEES, parent_fees - child_fees
        )
        _apply_decimal_format(cupones_sheet, parent_row, CUPONES_COL_NET, parent_net - child_net)
        _highlight_primary_split_amounts_red(cupones_sheet, parent_row)


def _apply_cta_truth_to_cupones_row(cupones_sheet, cta_sheet, row, coupon_text, coupon_index):
    """
    Apply cross-reference controls for one Cupones row from matched Cta Cte row.

    F/G formulas; H <- Cta G (Nro.); I <- Cta N (MES EFT).
    """
    matched_rows = _find_matching_cta_rows(coupon_text, coupon_index)
    if not matched_rows:
        return False

    formula = _build_eft_formula(CTA_CTE_SHEET, matched_rows)
    if formula:
        _cell_for_write(cupones_sheet, row, CUPONES_COL_EFT_FORMULA).value = formula
    matched_cta_row = matched_rows[0]
    _fill_zero_amounts_from_cta(cupones_sheet, cta_sheet, row, coupon_text)

    nro_value = _strip_cell(
        cta_sheet.cell(row=matched_cta_row, column=CTA_COL_NRO).value
    )
    if not nro_value:
        nro_value = _find_rcv_in_column_g(cta_sheet, matched_cta_row)
    if not nro_value:
        nro_value = "RCV-MISSING"
    _cell_for_write(cupones_sheet, row, CUPONES_COL_EFT_NO).value = nro_value

    mes_eft = _find_mes_eft_date(cta_sheet, matched_cta_row)
    if mes_eft:
        cell_i = _cell_for_write(cupones_sheet, row, CUPONES_COL_MES_EFT)
        cell_i.value = mes_eft
        cell_i.number_format = EFT_DATE_NUMBER_FORMAT
    else:
        raw_mes = _strip_cell(
            cta_sheet.cell(row=matched_cta_row, column=CTA_COL_NETO_FINAL).value
        )
        _cell_for_write(cupones_sheet, row, CUPONES_COL_MES_EFT).value = (
            raw_mes or "MES-MISSING"
        )

    return True


def _apply_control_columns(cupones_sheet, cta_sheet, row, coupon_text, coupon_index):
    """Backward-compatible alias for full Cta Cte row synchronization."""
    return _apply_cta_truth_to_cupones_row(
        cupones_sheet, cta_sheet, row, coupon_text, coupon_index
    )


def sync_cupones_for_cta_coupon_ids(workbook, coupon_ids):
    """
    Re-sync Cupones F-I by coupon IDs using dynamic Cta Cte matches.
    """
    cupones_sheet = _get_sheet(workbook, CUPONES_SHEET)
    cta_sheet = _get_sheet(workbook, CTA_CTE_SHEET)
    coupon_index = build_cta_coupon_index(cta_sheet)
    touched = 0

    for row in range(CUPONES_SCAN_START_ROW, cupones_sheet.max_row + 1):
        coupon_text = _strip_cell(
            cupones_sheet.cell(row=row, column=CUPONES_COL_COUPON).value
        )
        if not coupon_text:
            continue
        row_ids = _split_coupon_ids(coupon_text)
        if coupon_ids and not any(coupon_id in coupon_ids for coupon_id in row_ids):
            continue
        if _apply_control_columns(cupones_sheet, cta_sheet, row, coupon_text, coupon_index):
            touched += 1
        _apply_cupones_row_alignment(cupones_sheet, row)

    return touched


def _sort_cupones_sheet_by_date(worksheet, start_row=CUPONES_SCAN_START_ROW):
    """
    Sort Cupones rows chronologically by Column A (date) in ascending order.

    Preserves values/formulas in A–I by rewriting the row blocks.
    """
    max_row = max(worksheet.max_row, start_row)
    rows = []
    for row in range(start_row, max_row + 1):
        coupon = _strip_cell(worksheet.cell(row=row, column=CUPONES_COL_COUPON).value)
        if not coupon:
            continue
        date_value = worksheet.cell(row=row, column=CUPONES_COL_DATE).value
        sort_key = _parse_date_to_datetime(date_value) or datetime.max
        payload = []
        for col in range(1, CUPONES_COL_MES_EFT + 1):
            payload.append(worksheet.cell(row=row, column=col).value)
        rows.append((sort_key, payload))

    rows.sort(key=lambda item: item[0])
    for idx, (_key, payload) in enumerate(rows):
        dest_row = start_row + idx
        for col, value in enumerate(payload, start=1):
            cell = _cell_for_write(worksheet, dest_row, col)
            cell.value = value
            if col in (CUPONES_COL_GROSS, CUPONES_COL_FEES, CUPONES_COL_NET) and isinstance(
                value, (int, float)
            ):
                cell.number_format = DECIMAL_NUMBER_FORMAT
            if col == CUPONES_COL_MES_EFT and isinstance(value, datetime):
                cell.number_format = EFT_DATE_NUMBER_FORMAT
        _apply_cupones_row_alignment(worksheet, dest_row)


def _scan_max_column_b_length(worksheet, start_row=CUPONES_SCAN_START_ROW):
    max_len = 0
    for row in range(start_row, max(worksheet.max_row, start_row) + 1):
        text = _strip_cell(worksheet.cell(row=row, column=CUPONES_COL_COUPON).value)
        if text:
            max_len = max(max_len, len(text))
    return max_len


def _apply_cupones_row_alignment(worksheet, row):
    if Alignment is None:
        return
    for col in CUPONES_LEFT_COLUMNS:
        _cell_for_write(worksheet, row, col).alignment = LEFT_ALIGNMENT
    for col in CUPONES_RIGHT_COLUMNS:
        _cell_for_write(worksheet, row, col).alignment = RIGHT_ALIGNMENT


def _stretch_column_b_width(worksheet, max_text_length):
    if get_column_letter is None or max_text_length <= 0:
        return
    letter = get_column_letter(CUPONES_COL_COUPON)
    target_width = float(max_text_length + 3)
    current = worksheet.column_dimensions[letter].width
    if current is None or target_width > current:
        worksheet.column_dimensions[letter].width = target_width


def _row_has_content(worksheet, row, columns=(1, 2, 3, 4, 5)):
    for col in columns:
        value = worksheet.cell(row=row, column=col).value
        if value is not None and _strip_cell(value) != "":
            return True
    return False


def find_last_cupones_row(worksheet, start_row=CUPONES_SCAN_START_ROW):
    """Return the true last occupied row scanning column A bottom-up."""
    max_row = max(worksheet.max_row, start_row)
    for row in range(max_row, start_row - 1, -1):
        value_a = worksheet.cell(row=row, column=CUPONES_COL_DATE).value
        if value_a is not None and _strip_cell(value_a) != "":
            return row
    return start_row - 1


def _extract_ddc_ids(text):
    if not text:
        return []
    return [match.group(1).upper() for match in DDC_COUPON_PATTERN.finditer(str(text))]


def _split_coupon_ids(value):
    text = _strip_cell(value)
    if not text:
        return []
    parts = [part.strip().upper() for part in text.split(",")]
    coupons = []
    for part in parts:
        if not part:
            continue
        if DDC_COUPON_PATTERN.fullmatch(part):
            coupons.append(part.upper())
            continue
        coupons.extend(_extract_ddc_ids(part))
    unique = []
    seen = set()
    for coupon in coupons:
        if coupon and coupon not in seen:
            seen.add(coupon)
            unique.append(coupon)
    return unique


def build_cta_coupon_index(worksheet, start_row=CTA_SCAN_START_ROW):
    """
    Map DDC coupon IDs to Cta Cte row numbers by scanning column C (Invoice/Coupon).
    """
    index = {}
    max_row = max(worksheet.max_row, start_row)
    for row in range(start_row, max_row + 1):
        cell_value = worksheet.cell(row=row, column=CTA_COL_INVOICE).value
        for coupon_id in _extract_ddc_ids(cell_value):
            index.setdefault(coupon_id, []).append(row)
    for coupon_id in index:
        index[coupon_id] = sorted(set(index[coupon_id]))
    return index


def _find_rcv_in_column_g(worksheet, row, search_up=80):
    cell = worksheet.cell(row=row, column=CTA_COL_NRO)
    text = _strip_cell(cell.value)
    match = RCV_PATTERN.search(text)
    if match:
        return match.group(1).upper()
    for scan_row in range(row, max(CTA_SCAN_START_ROW, row - search_up) - 1, -1):
        text = _strip_cell(worksheet.cell(row=scan_row, column=CTA_COL_NRO).value)
        match = RCV_PATTERN.search(text)
        if match:
            return match.group(1).upper()
    return text if text else None


def _find_mes_eft_date(worksheet, row, search_up=80):
    cell = worksheet.cell(row=row, column=CTA_COL_NETO_FINAL)
    if isinstance(cell.value, datetime):
        return cell.value
    for scan_row in range(row, max(CTA_SCAN_START_ROW, row - search_up) - 1, -1):
        value = worksheet.cell(row=scan_row, column=CTA_COL_NETO_FINAL).value
        if value is None or _strip_cell(value) == "":
            continue
        if isinstance(value, datetime):
            return value
        parsed = _parse_date_to_datetime(value)
        if parsed:
            return parsed
    return None


def _build_eft_formula(sheet_name, cta_rows):
    if not cta_rows:
        return None
    parts = [f"+'{sheet_name}'!F{cta_row}" for cta_row in cta_rows]
    return "=" + "".join(parts)


def _find_matching_cta_rows(coupon_text, coupon_index):
    """Collect all matching Cta Cte rows for one Cupones coupon cell."""
    coupon_ids = _split_coupon_ids(coupon_text)
    if not coupon_ids:
        single = _resolve_single_coupon_id(coupon_text)
        coupon_ids = [single] if single else []

    seen = set()
    matches = []
    for coupon_id in coupon_ids:
        for row in coupon_index.get(coupon_id, []):
            if row in seen:
                continue
            seen.add(row)
            matches.append(row)
    return matches


def _is_batch_header(label):
    token = re.sub(r"\s+", " ", _strip_cell(label)).upper()
    return "BATCH" in token


def _read_monthly_excel_row_values(worksheet, row_number):
    """Read one report row via ws.cell(...).value (never raw cell objects)."""
    max_col = max(worksheet.max_column or 0, 12)
    values = []
    for col in range(1, max_col + 1):
        raw = _coerce_primitive_cell_value(
            worksheet.cell(row=row_number, column=col).value
        )
        if isinstance(raw, datetime):
            values.append(raw)
        elif isinstance(raw, (int, float)):
            values.append(raw)
        else:
            values.append(_safe_string_from_value(raw))

    while values:
        tail = values[-1]
        if isinstance(tail, (datetime, int, float)):
            break
        if _strip_cell(tail):
            break
        values.pop()
    return values


def _find_batch_column_index(header_cells):
    batch_idx = None
    try:
        for idx, label in enumerate(header_cells):
            if _is_batch_header(label):
                batch_idx = idx
                break
    except Exception:
        batch_idx = None
    return batch_idx


def _resolve_batch_column_index(header_cells):
    """
    Drop J.H. Williams \"Batch No(s)\" column (typically source column B / index 1).
    """
    batch_idx = _find_batch_column_index(header_cells)
    if batch_idx is None and len(header_cells) > 1:
        batch_idx = 1
    return batch_idx


def _build_keep_indices(header_cells, batch_idx):
    try:
        return [idx for idx in range(len(header_cells)) if idx != batch_idx]
    except Exception:
        return list(range(len(header_cells)))


def _record_from_filtered_row(filtered):
    if len(filtered) < 2:
        return None
    coupon_raw = _strip_cell(filtered[1] if len(filtered) > 1 else "")
    if not coupon_raw:
        return None
    if coupon_raw.lower() in {"coupon", "cupon", "cupón", "coupon id", "id"}:
        return None
    return {
        "date": filtered[0] if filtered else "",
        "coupon": coupon_raw,
        "gross": _parse_amount(filtered[2] if len(filtered) > 2 else 0),
        "fees": _parse_amount(filtered[3] if len(filtered) > 3 else 0),
        "net": _parse_amount(filtered[4] if len(filtered) > 4 else 0),
    }


def _read_monthly_coupon_rows_from_table(header_cells, data_rows):
    """
    Row 1 skipped upstream; header_cells are from row 2; data_rows start at row 3.
    """
    batch_idx = _resolve_batch_column_index(header_cells)
    keep_indices = _build_keep_indices(header_cells, batch_idx)
    if not keep_indices:
        return []

    records = []
    for row_values in data_rows:
        filtered = [
            row_values[idx] if idx < len(row_values) else ""
            for idx in keep_indices
        ]
        record = _record_from_filtered_row(filtered)
        if record:
            records.append(record)
    return records


def _read_monthly_coupon_rows_excel(monthly_path):
    _ensure_openpyxl()
    workbook = load_workbook(monthly_path, data_only=True)
    try:
        worksheet = workbook.active
        header_cells = _read_monthly_excel_row_values(worksheet, MONTHLY_HEADER_ROW)

        data_rows = []
        max_row = worksheet.max_row or MONTHLY_DATA_START_ROW
        for row_number in range(MONTHLY_DATA_START_ROW, max_row + 1):
            row_values = _read_monthly_excel_row_values(worksheet, row_number)
            if not row_values:
                continue
            has_content = any(
                isinstance(item, (datetime, int, float)) or _strip_cell(item)
                for item in row_values
            )
            if has_content:
                data_rows.append(row_values)
    finally:
        workbook.close()

    return _read_monthly_coupon_rows_from_table(header_cells, data_rows)


def _read_monthly_coupon_rows_csv(monthly_path):
    raw = pd.read_csv(monthly_path, header=None, dtype=str, keep_default_na=False)
    if len(raw) < MONTHLY_DATA_START_ROW:
        return []

    header_cells = [_strip_cell(value) for value in raw.iloc[MONTHLY_HEADER_ROW - 1].tolist()]
    data_rows = [
        [_strip_cell(value) for value in row.tolist()]
        for _, row in raw.iloc[MONTHLY_DATA_START_ROW - 1 :].iterrows()
    ]
    return _read_monthly_coupon_rows_from_table(header_cells, data_rows)


def _read_monthly_coupon_rows(monthly_path):
    """
    Parse J.H. Williams monthly coupon export.

    Row 1 (merged title block) is skipped. Row 2 provides headers; any column
    containing \"BATCH\" is dropped before mapping Date/Coupon/Gross/Fees/Net.
    """
    extension = os.path.splitext(monthly_path)[1].lower()
    if extension == ".csv":
        return _read_monthly_coupon_rows_csv(monthly_path)
    if extension in {".xlsx", ".xlsm", ".xls"}:
        return _read_monthly_coupon_rows_excel(monthly_path)
    raise ValueError(
        f"Unsupported monthly report type '{extension}'. Use CSV or Excel."
    )


def append_monthly_cupones(master_path, monthly_path):
    """
    Append monthly coupon rows to Cupones and populate F-I from Cta Cte.

    Returns:
        tuple: (saved master path, summary dict)
    """
    _ensure_openpyxl()
    master_path = os.path.abspath(str(master_path).strip())
    monthly_path = os.path.abspath(str(monthly_path).strip())

    if not os.path.isfile(master_path):
        raise FileNotFoundError(f"Master workbook not found: {master_path}")
    if not os.path.isfile(monthly_path):
        raise FileNotFoundError(f"Monthly coupon report not found: {monthly_path}")

    raw_monthly_rows = _read_monthly_coupon_rows(monthly_path)
    if not raw_monthly_rows:
        raise ValueError("No coupon rows found in the monthly report.")

    # Enforce chronological append order (earliest date first).
    monthly_rows = _sort_records_by_date(_expand_records_by_coupon_split(raw_monthly_rows))
    del raw_monthly_rows

    extension = os.path.splitext(master_path)[1].lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise ValueError("Master workbook must be .xlsx or .xlsm.")

    keep_vba = extension == ".xlsm"
    rows_skipped_duplicates = 0
    start_row = 0
    appended = 0
    workbook = None
    try:
        workbook = load_workbook(master_path, data_only=False, keep_vba=keep_vba)
        cupones_sheet = _get_sheet(workbook, CUPONES_SHEET)
        monthly_rows, rows_skipped_duplicates = _filter_new_cupones_records(
            cupones_sheet, monthly_rows
        )
        if not monthly_rows:
            raise MonthlyReportFullyDuplicateError(MONTHLY_REPORT_FULL_DUPLICATE_ALERT)

        cta_sheet = _get_sheet(workbook, CTA_CTE_SHEET)
        coupon_index = build_cta_coupon_index(cta_sheet)

        last_row = find_last_cupones_row(cupones_sheet)
        start_row = last_row + 1

        save_path = _create_temp_workbook_path(suffix=extension or ".xlsx")
        appended = 0
        rows_matched = 0
        unmatched_coupons = []
        max_coupon_text_len = 0
        split_group_rows = {}

        for offset, record in enumerate(monthly_rows):
            row = start_row + offset
            coupon_text = _write_cupones_base_columns(cupones_sheet, row, record)
            max_coupon_text_len = max(max_coupon_text_len, len(coupon_text or ""))
            _clear_cupones_control_columns(cupones_sheet, row)
            if _apply_control_columns(
                cupones_sheet, cta_sheet, row, coupon_text, coupon_index
            ):
                rows_matched += 1
            else:
                unmatched_coupons.append(coupon_text)
            split_group_id = record.get("split_group_id")
            if split_group_id:
                split_group_rows.setdefault(split_group_id, []).append(row)
                if not record.get("is_primary_split"):
                    _apply_child_rows_standard_black(cupones_sheet, row)
            _apply_cupones_row_alignment(cupones_sheet, row)
            appended += 1

        _apply_parent_child_subtractions(cupones_sheet, split_group_rows)
        for rows in split_group_rows.values():
            for r in rows:
                _apply_cupones_row_alignment(cupones_sheet, r)
        _stretch_column_b_width(cupones_sheet, max_coupon_text_len)
        workbook.save(save_path)
    except PermissionError as exc:
        raise PermissionError("ERROR: Cierra el archivo Excel antes de continuar") from exc
    finally:
        if workbook is not None:
            workbook.close()
        if "monthly_rows" in locals():
            monthly_rows = None
        gc.collect()

    _launch_temp_workbook(save_path)

    summary = {
        "rows_appended": appended,
        "start_row": start_row,
        "end_row": start_row + appended - 1 if appended else start_row - 1,
        "rows_matched": rows_matched,
        "rows_skipped_duplicates": rows_skipped_duplicates,
        "historical_repaired": 0,
        "unmatched_coupons": unmatched_coupons,
    }
    return save_path, summary
