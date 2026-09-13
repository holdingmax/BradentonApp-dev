"""
Lee el reporte mensual de Cupones de J.H. Williams. Pura -- ya no escribe
ningún Excel (2026-09-12, ver CLAUDE.md "EFT y Cupones -- convertido a
Carga de Datos"): lo que antes era append_monthly_cupones/
resync_cupones_only (y todo el motor de cruce contra Cta Cte --
build_cta_coupon_index, build_cta_eft_boxes, el manejo de grupos partidos
para no romper una fórmula SUM de Excel, los avisos de posible duplicado
pintados en la hoja, etc.) se eliminó del todo. read_monthly_coupon_rows +
expand_monthly_records_for_storage son el único punto de entrada ahora,
alimentando eft_db.py -- el cruce EFT<->Cupón se resuelve ahí con un JOIN
al leer, no con una fórmula guardada que pueda quedar desactualizada.
"""

import os
import re
from datetime import datetime

import pandas as pd

try:
    from openpyxl import load_workbook

    OPENPYXL_AVAILABLE = True
except ImportError:
    load_workbook = None  # type: ignore[assignment,misc]
    OPENPYXL_AVAILABLE = False

MONTHLY_HEADER_ROW = 2
MONTHLY_DATA_START_ROW = 3

DDC_COUPON_PATTERN = re.compile(r"\b(DDC-\d+)\b", re.IGNORECASE)

# Estructura de la hoja Cupones del Excel "Aplicacion TC y EFT -" -- estas
# constantes y find_last_cupones_row (más abajo) son de solo LECTURA
# (nunca escriben nada) y siguen haciendo falta acá porque
# controles_cupones.py (Controles -> Cupones, que no se toca en esta
# conversión) las importa para ubicar filas en el Excel real que sigue
# leyendo -- no se eliminaron junto con el resto del motor de escritura.
CUPONES_SHEET = "Cupones"
CUPONES_COL_DATE = 1
CUPONES_COL_COUPON = 2
CUPONES_COL_GROSS = 3
CUPONES_COL_FEES = 4
CUPONES_COL_NET = 5
CUPONES_COL_EFT_FORMULA = 6
CUPONES_SCAN_START_ROW = 2


def _ensure_openpyxl():
    if not OPENPYXL_AVAILABLE:
        raise ImportError(
            "Leer el reporte mensual de Cupones requiere openpyxl. Instale con: pip install openpyxl"
        )


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


def find_last_cupones_row(worksheet, start_row=CUPONES_SCAN_START_ROW):
    """
    Return the true last occupied row scanning bottom-up.

    Checks column A (date) first, but falls back to column B (coupon) so a
    row whose date failed to parse and was left blank still counts as
    occupied. Solo lectura -- usada por controles_cupones.py.
    """
    max_row = max(worksheet.max_row, start_row)
    for row in range(max_row, start_row - 1, -1):
        value_a = worksheet.cell(row=row, column=CUPONES_COL_DATE).value
        value_b = worksheet.cell(row=row, column=CUPONES_COL_COUPON).value
        if (value_a is not None and _strip_cell(value_a) != "") or (
            value_b is not None and _strip_cell(value_b) != ""
        ):
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
    """Drop J.H. Williams "Batch No(s)" column (typically source column B / index 1)."""
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
    """Row 1 skipped upstream; header_cells are from row 2; data_rows start at row 3."""
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
    containing "BATCH" is dropped before mapping Date/Coupon/Gross/Fees/Net.
    """
    extension = os.path.splitext(monthly_path)[1].lower()
    if extension == ".csv":
        return _read_monthly_coupon_rows_csv(monthly_path)
    if extension in {".xlsx", ".xlsm", ".xls"}:
        return _read_monthly_coupon_rows_excel(monthly_path)
    raise ValueError(
        f"Tipo de reporte mensual no soportado '{extension}'. Use CSV o Excel."
    )


def read_monthly_coupon_rows(monthly_path):
    """Wrapper público de _read_monthly_coupon_rows -- pura, no toca ningún Excel."""
    return _read_monthly_coupon_rows(monthly_path)


def expand_monthly_records_for_storage(records):
    """
    Divide una fila del reporte mensual en una fila por cada DDC que trae
    (una fila puede combinar varios cupones liquidados juntos, con un solo
    total reportado). Para guardar en eft_db.py cada DDC queda como su
    propia fila independiente -- no hace falta la danza de "restarle al
    padre lo que van resolviendo los hijos" que sí necesitaba el Excel para
    mantener consistente una fórmula SUM; acá el gross/fees/net individual
    de un DDC agrupado queda en 0 hasta que algún EFT confirme el monto
    real de ESE DDC puntual (ver eft_db.get_cupones_with_status), y el
    total/texto reportado del grupo se guarda aparte como metadata
    informativa en cada miembro.
    """
    expanded = []
    for record in records:
        coupon_ids = _split_coupon_ids(record.get("coupon"))
        if not coupon_ids:
            continue
        is_group = len(coupon_ids) > 1
        group_text = _strip_cell(record.get("coupon")) if is_group else None
        group_total = _parse_amount(record.get("net", 0)) if is_group else None
        for coupon_id in coupon_ids:
            expanded.append(
                {
                    "coupon": coupon_id,
                    "date": record.get("date"),
                    "gross": 0.0 if is_group else _parse_amount(record.get("gross", 0)),
                    "fees": 0.0 if is_group else _parse_amount(record.get("fees", 0)),
                    "net": 0.0 if is_group else _parse_amount(record.get("net", 0)),
                    "reported_group_text": group_text,
                    "reported_group_total": group_total,
                }
            )
    return expanded
