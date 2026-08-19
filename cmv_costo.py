"""
CMV / Elistars department cost automation for the master COSTO.TODOS sheet.

Parses department files with pandas, updates the master workbook with openpyxl
(style-preserving), and saves a temp preview only.
"""

import logging
import os
import sys
import tempfile

import pandas as pd

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Border, Side

    OPENPYXL_AVAILABLE = True
except ImportError:
    load_workbook = None  # type: ignore[assignment,misc]
    Border = None  # type: ignore[assignment,misc]
    Side = None  # type: ignore[assignment,misc]
    OPENPYXL_AVAILABLE = False

try:
    import xlrd  # noqa: F401

    XLRD_AVAILABLE = True
except ImportError:
    XLRD_AVAILABLE = False

logger = logging.getLogger(__name__)

COSTO_TODOS_SHEET_NAME = "COSTO.TODOS"
COSTO_TODOS_SHEET_INDEX = 3  # 1-based sheet position in workbook
DATA_START_ROW = 2  # Row 1 is header; product grid overwrite begins at A2
COSTO_DATA_COLUMNS = 7  # A through G (UPC … DeptName)

if OPENPYXL_AVAILABLE and Side is not None and Border is not None:
    _THIN_SIDE = Side(style="thin", color="000000")
    COSTO_GRID_BORDER = Border(
        left=_THIN_SIDE,
        right=_THIN_SIDE,
        top=_THIN_SIDE,
        bottom=_THIN_SIDE,
    )
else:
    COSTO_GRID_BORDER = None
DEFAULT_UPC_LENGTHS = (13, 12, 14, 8)

RAW_COLUMN_NAMES = [
    "UPC",
    "UPCMod",
    "Name",
    "Cost",
    "BuyDown",
    "Price",
    "DeptID",
    "DeptName",
    "SellUnit",
]
DROP_COLUMNS = frozenset({"BuyDown", "SellUnit"})
EXPECTED_FIELD_COUNT = len(RAW_COLUMN_NAMES)

PATH_SEPARATORS = (";", "|")


def join_paths(paths):
    """Join absolute department paths for the Entry widget (multi-select display)."""
    normalized = [
        os.path.abspath(str(path).strip())
        for path in paths
        if path is not None and str(path).strip()
    ]
    return "; ".join(normalized)


def split_paths(text):
    if not text or not str(text).strip():
        return []
    raw = str(text).strip()
    for separator in PATH_SEPARATORS:
        if separator in raw:
            return [part.strip() for part in raw.split(separator) if part.strip()]
    return [raw]


def _ensure_openpyxl_available():
    if not OPENPYXL_AVAILABLE:
        raise ImportError(
            "CMV merge requires openpyxl. Install with: pip install openpyxl"
        )


def clean_upc(value):
    """Cast to string, strip whitespace, and remove trailing '.0' artifacts."""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    if isinstance(value, bool):
        return ""

    if isinstance(value, str):
        text = value.strip()
    elif isinstance(value, int):
        text = str(value)
    elif isinstance(value, float):
        if value == int(value):
            text = str(int(value))
        else:
            text = format(value, "f").rstrip("0").rstrip(".")
    else:
        text = str(value).strip()

    if not text or text.lower() in {"nan", "none"}:
        return ""

    if "." in text:
        whole, fractional = text.split(".", 1)
        if fractional == "" or set(fractional) <= {"0"}:
            text = whole

    return text.strip()


def collect_upc_padding_lengths(price_map):
    lengths = {len(key) for key in price_map if key.isdigit()}
    for fallback in DEFAULT_UPC_LENGTHS:
        lengths.add(fallback)
    return sorted(lengths, reverse=True)


def upc_match_candidates(value, padding_lengths):
    base = clean_upc(value)
    if not base:
        return []
    if not base.isdigit():
        return [base]

    candidates = []
    seen = set()
    for candidate in [base] + [
        base.zfill(length) for length in padding_lengths if len(base) <= length
    ]:
        if candidate not in seen:
            seen.add(candidate)
            candidates.append(candidate)
    return candidates


def resolve_upc_key(raw_value, price_map, padding_lengths):
    for candidate in upc_match_candidates(raw_value, padding_lengths):
        if candidate in price_map:
            return candidate
    return None


def _strip_cell(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _split_raw_fields(raw_text):
    parts = [_strip_cell(part) for part in str(raw_text).split(",")]
    if len(parts) < EXPECTED_FIELD_COUNT:
        parts.extend([""] * (EXPECTED_FIELD_COUNT - len(parts)))
    elif len(parts) > EXPECTED_FIELD_COUNT:
        name = ",".join(parts[2 : len(parts) - 6])
        parts = parts[:2] + [name] + parts[-6:]
    return parts[:EXPECTED_FIELD_COUNT]


def _row_to_fields(row):
    cells = [_strip_cell(value) for value in row]
    non_empty = [cell for cell in cells if cell]
    if not non_empty:
        return None

    if len(non_empty) == 1:
        return _split_raw_fields(non_empty[0])

    if len(cells) >= EXPECTED_FIELD_COUNT and any(cells[2:]):
        return cells[:EXPECTED_FIELD_COUNT]

    return _split_raw_fields(",".join(cells))


def _read_raw_table(file_path):
    extension = os.path.splitext(file_path)[1].lower()
    if extension == ".csv":
        return pd.read_csv(file_path, header=None, dtype=str, keep_default_na=False)
    if extension in {".xlsx", ".xlsm"}:
        return pd.read_excel(file_path, header=None, dtype=str, keep_default_na=False)
    if extension == ".xls":
        if not XLRD_AVAILABLE:
            raise ImportError(
                "Reading .xls department files requires xlrd. "
                "Install dependencies: pip install -r requirements.txt"
            )
        return pd.read_excel(file_path, header=None, dtype=str, engine="xlrd")
    raise ValueError(
        f"Unsupported file type '{extension}'. Use .csv, .xlsx, .xlsm, or .xls."
    )


def read_elistars_department_file(file_path):
    df = _read_raw_table(file_path)
    records = []
    for _, row in df.iterrows():
        fields = _row_to_fields(row.tolist())
        if fields is None or all(not field for field in fields):
            continue
        joined = ",".join(fields).lower()
        if joined.startswith("upc,") and "buydown" in joined:
            continue
        records.append(dict(zip(RAW_COLUMN_NAMES, fields)))

    if not records:
        raise ValueError(
            f"No product rows in {os.path.basename(file_path)}. "
            "Expected comma-delimited data in one column or columns A–I."
        )

    result = pd.DataFrame(records)
    result = result.drop(columns=[c for c in DROP_COLUMNS if c in result.columns])
    return _finalize_department_frame(result)


def _parse_decimal(value):
    text = _strip_cell(value)
    if not text:
        return None
    text = text.replace("$", "").replace(" ", "")
    if text.count(",") == 1 and text.count(".") == 0:
        text = text.replace(",", ".")
    else:
        text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def _finalize_department_frame(df):
    if "UPC" in df.columns:
        df["UPC"] = df["UPC"].astype(str).str.strip().apply(clean_upc)
        df = df[df["UPC"] != ""]

    for col in ("Cost", "Price"):
        if col in df.columns:
            df[col] = df[col].apply(_parse_decimal)

    for col in ("UPCMod", "Name", "DeptName"):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    return df


def _department_sort_column(df):
    """Resolve department label column ('dept name' / DeptName) for grouping."""
    for col in df.columns:
        normalized = str(col).strip().lower().replace("_", " ")
        if normalized in {"dept name", "deptname"}:
            return col
    if "DeptName" in df.columns:
        return "DeptName"
    return None


def _sort_by_department(df):
    """Group products by department before writing to COSTO.TODOS."""
    dept_col = _department_sort_column(df)
    if dept_col is None:
        return df

    sorted_df = df.copy()
    sorted_df[dept_col] = sorted_df[dept_col].astype(str).str.strip()
    return sorted_df.sort_values(by=dept_col, ascending=True, kind="stable")


def _consolidate_department_files(department_paths):
    """
    Read all Elistars department exports, concatenate, clean, and sort by dept.
    """
    frames = []
    file_stats = []
    ordered_paths = [os.path.abspath(path) for path in department_paths]

    for department_path in ordered_paths:
        dept_df = read_elistars_department_file(department_path)
        label = infer_department_label(dept_df, department_path)
        frames.append(dept_df)
        file_stats.append(
            {
                "path": department_path,
                "label": label,
                "parsed": len(dept_df),
                "mapped": len(dept_df),
            }
        )

    if not frames:
        raise ValueError("No department rows could be read from the selected files.")

    combined = pd.concat(frames, ignore_index=True)
    combined = _finalize_department_frame(combined)
    combined = _sort_by_department(combined)
    return combined, file_stats


def infer_department_label(df, source_path):
    if "DeptName" in df.columns:
        names = df["DeptName"].replace("", pd.NA).dropna().astype(str).str.strip()
        names = names[names != ""]
        if not names.empty:
            return names.mode().iloc[0]
    return os.path.splitext(os.path.basename(source_path))[0]


def _cell_is_empty(value):
    if value is None:
        return True
    text = str(value).strip()
    return text == "" or text.lower() in {"nan", "none"}


def _coerce_upc_integer(value, default=0):
    """Convert a cleaned UPC string into a plain integer for Excel column A/B."""
    text = clean_upc(value) if value is not None else ""
    if not text:
        return default
    digits = "".join(character for character in text if character.isdigit())
    if not digits:
        return default
    return int(digits)


def _coerce_cost_price_float(value):
    """Cast Cost/Price to native float for Excel numeric cells."""
    if value is None:
        return None
    try:
        if isinstance(value, float) and pd.isna(value):
            return None
        if isinstance(value, (int, float)):
            return float(value)
        parsed = _parse_decimal(value)
        if parsed is None:
            return None
        return float(parsed)
    except (TypeError, ValueError):
        return None


def _validate_master_path(master_path):
    extension = os.path.splitext(master_path)[1].lower()
    if extension not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError("Master CMV file must be .xls, .xlsx, or .xlsm.")
    if extension == ".xls":
        raise ValueError(
            "Legacy .xls masters cannot be updated in place with openpyxl. "
            "Open the master in Excel and Save As .xlsx or .xlsm, then retry."
        )
    return extension


def _create_temp_xlsx_path():
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="cmv_preview_")
    os.close(fd)
    return temp_path


def _get_costo_todos_sheet(workbook):
    """Target COSTO.TODOS by name; fall back to the third worksheet."""
    names = list(workbook.sheetnames)
    target = COSTO_TODOS_SHEET_NAME.strip().lower()
    for name in names:
        if name.strip().lower() == target:
            return workbook[name]

    if len(workbook.worksheets) >= COSTO_TODOS_SHEET_INDEX:
        return workbook.worksheets[COSTO_TODOS_SHEET_INDEX - 1]

    raise ValueError(
        f'Sheet "{COSTO_TODOS_SHEET_NAME}" not found. Available: {", ".join(names)}'
    )


def _load_master_workbook(master_path):
    _ensure_openpyxl_available()
    extension = os.path.splitext(master_path)[1].lower()
    keep_vba = extension == ".xlsm"
    return load_workbook(master_path, data_only=False, keep_vba=keep_vba)


def _hide_column_b(sheet):
    sheet.column_dimensions["B"].hidden = True


def _apply_costo_row_grid_border(sheet, row_idx):
    """Thin black border on all four sides for every cell in columns A–G."""
    if COSTO_GRID_BORDER is None:
        return
    for col in range(1, COSTO_DATA_COLUMNS + 1):
        sheet.cell(row=row_idx, column=col).border = COSTO_GRID_BORDER


def _write_upc_integer_cell(sheet, row_idx, column, value, default=0):
    try:
        amount = _coerce_upc_integer(value, default=default)
    except (TypeError, ValueError):
        return

    cell = sheet.cell(row=row_idx, column=column, value=int(amount))
    cell.number_format = "0"


def _write_currency_cell(sheet, row_idx, column, value):
    try:
        if isinstance(value, (int, float)):
            amount = float(value)
        else:
            coerced = _coerce_cost_price_float(value)
            if coerced is None:
                return
            amount = float(coerced)
    except (TypeError, ValueError):
        return

    cell = sheet.cell(row=row_idx, column=column, value=amount)
    cell.number_format = "0.00"


def _write_department_rows(sheet, dept_df, start_row):
    """Write consolidated department export starting at start_row; return next empty row."""
    row_idx = start_row
    rows_appended = 0

    for _, row in dept_df.iterrows():
        upc = clean_upc(row.get("UPC"))
        if not upc:
            continue

        _write_upc_integer_cell(sheet, row_idx, 1, upc)

        upc_mod = row.get("UPCMod")
        if upc_mod is not None and not _cell_is_empty(upc_mod):
            _write_upc_integer_cell(sheet, row_idx, 2, upc_mod, default=0)
        else:
            _write_upc_integer_cell(sheet, row_idx, 2, "0", default=0)

        name = row.get("Name")
        sheet.cell(
            row=row_idx,
            column=3,
            value="" if name is None else str(name).strip(),
        )

        cost = row.get("Cost")
        if cost is not None:
            try:
                cost_float = (
                    float(cost)
                    if isinstance(cost, (int, float))
                    else _coerce_cost_price_float(cost)
                )
                if cost_float is not None:
                    _write_currency_cell(sheet, row_idx, 4, cost_float)
            except (TypeError, ValueError):
                pass

        price = row.get("Price")
        if price is not None:
            try:
                price_float = (
                    float(price)
                    if isinstance(price, (int, float))
                    else _coerce_cost_price_float(price)
                )
                if price_float is not None:
                    _write_currency_cell(sheet, row_idx, 5, price_float)
            except (TypeError, ValueError):
                pass

        dept_id = row.get("DeptID")
        if dept_id is not None and not _cell_is_empty(dept_id):
            try:
                sheet.cell(row=row_idx, column=6, value=int(float(str(dept_id).strip())))
            except (TypeError, ValueError):
                sheet.cell(row=row_idx, column=6, value=str(dept_id).strip())

        dept_name = row.get("DeptName")
        sheet.cell(
            row=row_idx,
            column=7,
            value="" if dept_name is None else str(dept_name).strip(),
        )

        _apply_costo_row_grid_border(sheet, row_idx)

        row_idx += 1
        rows_appended += 1

    return row_idx, rows_appended


def _row_has_costo_data(sheet, row_idx):
    """True when any product column (A–G) on the row contains a value."""
    return any(
        not _cell_is_empty(sheet.cell(row=row_idx, column=col).value)
        for col in range(1, COSTO_DATA_COLUMNS + 1)
    )


def _find_last_costo_data_row(sheet):
    """Last populated product row at or below DATA_START_ROW."""
    last_used = DATA_START_ROW - 1
    max_scan = max(sheet.max_row or 0, DATA_START_ROW)
    for row_idx in range(DATA_START_ROW, max_scan + 1):
        if _row_has_costo_data(sheet, row_idx):
            last_used = row_idx
    return last_used


def _clear_costo_data_grid(sheet):
    """
    Wipe existing product rows from DATA_START_ROW through the last used row.

    Uses delete_rows so stale UPCs, costs, prices, and cell styles are removed
    before the fresh import overwrites the grid from the top entry row.
    """
    last_row = _find_last_costo_data_row(sheet)
    if last_row < DATA_START_ROW:
        return 0

    rows_to_delete = last_row - DATA_START_ROW + 1
    sheet.delete_rows(DATA_START_ROW, amount=rows_to_delete)
    return rows_to_delete


def _replace_costo_departments(sheet, department_paths):
    """
    Consolidate department exports, clear the COSTO.TODOS data grid, and write
    fresh rows starting at DATA_START_ROW (overwrite, never append).
    """
    _hide_column_b(sheet)
    combined_df, file_stats = _consolidate_department_files(department_paths)
    _clear_costo_data_grid(sheet)
    _, rows_written = _write_department_rows(sheet, combined_df, DATA_START_ROW)
    _hide_column_b(sheet)
    return file_stats, rows_written


def _openpyxl_merge_and_save(master_path, department_paths, temp_xlsx_path):
    workbook = None
    try:
        workbook = _load_master_workbook(master_path)
        sheet = _get_costo_todos_sheet(workbook)
        file_stats, rows_appended = _replace_costo_departments(sheet, department_paths)
        workbook.save(os.path.abspath(temp_xlsx_path))
        return file_stats, rows_appended
    finally:
        if workbook is not None:
            workbook.close()


def update_master_costo_todos(master_path, department_path):
    return update_master_costo_todos_bulk(master_path, [department_path])


def _launch_temp_workbook(temp_path):
    """Open the transient preview workbook without UI alerts."""
    abs_path = os.path.abspath(temp_path)
    if sys.platform == "win32":
        os.startfile(abs_path)
    elif sys.platform == "darwin":
        os.system(f'open "{abs_path}"')
    else:
        os.system(f'xdg-open "{abs_path}"')


def update_master_costo_todos_bulk(master_path, department_paths):
    """
    Overwrite COSTO.TODOS product rows via openpyxl (clean-then-write).

    Reads all selected department files into one pandas DataFrame, cleans UPC
    and Cost/Price fields, sorts by department name, deletes existing rows from
    row 2 downward (columns A–G), then writes the consolidated import from A2.
    Saves a transient .xlsx preview and opens it silently.

    Returns:
        tuple: (temp_xlsx_path, file_stats, total_parsed, rows_appended,
                upcs_not_in_master, master_row_count)
    """
    if not department_paths:
        raise ValueError("No department files provided.")

    master_path = os.path.abspath(master_path)
    _validate_master_path(master_path)

    temp_xlsx_path = _create_temp_xlsx_path()

    file_stats, rows_appended = _openpyxl_merge_and_save(
        master_path, department_paths, temp_xlsx_path
    )

    if rows_appended == 0:
        raise ValueError("No department rows could be written to COSTO.TODOS.")

    _launch_temp_workbook(temp_xlsx_path)

    total_parsed = sum(item["parsed"] for item in file_stats)
    return (
        temp_xlsx_path,
        file_stats,
        total_parsed,
        rows_appended,
        0,
        rows_appended,
    )


def workbook_has_costo_todos_sheet(file_path):
    extension = os.path.splitext(file_path)[1].lower()
    if extension not in {".xlsx", ".xlsm", ".xls"}:
        return False

    if extension in {".xlsx", ".xlsm"} and OPENPYXL_AVAILABLE:
        workbook = None
        try:
            workbook = load_workbook(
                os.path.abspath(file_path), read_only=True, data_only=False
            )
            _get_costo_todos_sheet(workbook)
            return True
        except ValueError:
            return False
        except Exception:
            pass
        finally:
            if workbook is not None:
                workbook.close()

    try:
        if extension == ".xls":
            if not XLRD_AVAILABLE:
                return False
            book = pd.ExcelFile(file_path, engine="xlrd")
        else:
            book = pd.ExcelFile(file_path)
        try:
            target = COSTO_TODOS_SHEET_NAME.lower()
            names = [name.strip().lower() for name in book.sheet_names]
            if target in names:
                return True
            return len(book.sheet_names) > COSTO_TODOS_SHEET_INDEX - 1
        finally:
            book.close()
    except Exception:
        return False
