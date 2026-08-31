"""
Chase bank activity engine: keyword-to-Detalle categorization rules
(persisted + hardcoded) and the full read/categorize/write pipeline.
"""

import json
import os
import re
import unicodedata
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

RULES_FILENAME = "chase_rules.json"


def _rules_file_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), RULES_FILENAME)


def normalize_rule_text(value):
    """Lowercase, strip accents, collapse whitespace; CHECK/CHEQUE -> CHEK."""
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"\b(check|cheque)\b", "chek", text)
    return re.sub(r"\s+", " ", text).strip()


def _normalize_rule_entry(keyword, detail):
    keyword_text = str(keyword).strip()
    detail_text = str(detail).strip()
    if not keyword_text or not detail_text:
        raise ValueError("La palabra clave y el detalle son obligatorios.")
    return {"keyword": keyword_text, "detail": detail_text}


def load_dynamic_rules():
    """Return all persisted rule dicts: {"keyword": str, "detail": str}."""
    path = _rules_file_path()
    if not os.path.isfile(path):
        return []

    try:
        with open(path, encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return []

    if isinstance(payload, dict):
        rules = payload.get("rules", [])
    elif isinstance(payload, list):
        rules = payload
    else:
        return []

    cleaned = []
    for item in rules:
        if not isinstance(item, dict):
            continue
        keyword = str(item.get("keyword", "")).strip()
        detail = str(item.get("detail", "")).strip()
        if keyword and detail:
            cleaned.append({"keyword": keyword, "detail": detail})
    return cleaned


def save_dynamic_rules(rules):
    """Persist the full rules list atomically without dropping unrelated entries."""
    if not isinstance(rules, list):
        raise TypeError("Rules must be a list.")

    path = _rules_file_path()
    directory = os.path.dirname(path)
    os.makedirs(directory, exist_ok=True)

    payload = {"rules": rules}
    temp_path = f"{path}.tmp"
    with open(temp_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, ensure_ascii=False)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp_path, path)


def add_dynamic_rule(keyword, detail):
    """Append or update one rule while preserving every other saved rule."""
    entry = _normalize_rule_entry(keyword, detail)
    rules = load_dynamic_rules()
    keyword_key = entry["keyword"].lower()
    rules = [rule for rule in rules if rule["keyword"].lower() != keyword_key]
    rules.append(entry)
    save_dynamic_rules(rules)
    return entry


def delete_dynamic_rule(keyword, detail):
    """Remove one matching rule and leave all other saved rules untouched."""
    keyword_key = str(keyword).strip().lower()
    detail_key = str(detail).strip()
    rules = load_dynamic_rules()
    filtered = [
        rule
        for rule in rules
        if not (
            rule["keyword"].lower() == keyword_key and rule["detail"] == detail_key
        )
    ]
    if len(filtered) == len(rules):
        raise ValueError("La regla seleccionada no se encontró en el almacenamiento.")
    save_dynamic_rules(filtered)


def delete_dynamic_rule_by_index(index):
    """Remove one persisted rule by index in chase_rules.json."""
    rules = load_dynamic_rules()
    try:
        idx = int(index)
    except (TypeError, ValueError):
        raise ValueError("Índice de regla inválido.") from None
    if idx < 0 or idx >= len(rules):
        raise ValueError("La regla seleccionada no se encontró en el almacenamiento.")
    removed = rules.pop(idx)
    save_dynamic_rules(rules)
    return removed


def match_dynamic_detalle(description):
    """
    Case-insensitive substring match against saved keyword rules.

    Among every rule whose keyword appears in the description, the LONGEST
    keyword wins -- rules are stored in add order (add_dynamic_rule always
    appends), so matching by insertion order instead would let a generic
    rule added early ("KOOLER") permanently shadow a more specific one
    added later ("KOOLER FARMS LLC").

    Returns Target Detail text or None.
    """
    desc = normalize_rule_text(description)
    if not desc:
        return None

    best_rule = None
    best_keyword = ""
    for rule in load_dynamic_rules():
        keyword = normalize_rule_text(rule.get("keyword", ""))
        if keyword and keyword in desc and len(keyword) > len(best_keyword):
            best_rule = rule
            best_keyword = keyword
    return best_rule.get("detail") if best_rule else None


# ---------------------------------------------------------------------------
# Categorization engine: hardcoded keyword groups (Alfonso's baseline rules)
# ---------------------------------------------------------------------------

EFT_RCV_EXACT_LABEL = "EFT RCV-"
EFT_RCV_RED_FONT = Font(color="FF0000", bold=True)

CHASE_CHECK_WORDS = ("chek",)


def _desc_has_check_keyword(desc):
    """True when description contains CHECK, CHEQUE, or CHEK (normalized to chek)."""
    return "chek" in desc


def _desc_contains_any(desc, keywords):
    """Return True if any keyword appears in normalized description text."""
    return any(keyword in desc for keyword in keywords)


# Alfonso master keyword groups (case-insensitive via normalize_rule_text).
CHASE_PROVEEDORES_KEYWORDS = (
    "frito-la",
    "hackneyrectampa",
    "cec distributing",
    "gold coast eagle",
    "jj taylor distri",
    "pbg",
    "colonial",
    "redbull",
    "airgas",
    "johnson brothers",
)
CHASE_GASTOS_BANCARIOS_KEYWORDS = (
    "low value",
    "initial fee",
    "cash deposit immediate",
    "monthly service fee",
)
CHASE_REBATE_KEYWORDS = (
    "helix ucp",
    "ussmokless",
    "njoy",
    "itg brands",
    "john middleton",
)
CHASE_AGUA_KEYWORDS = ("mucs", "manatee")
CHASE_ALARMA_KEYWORDS = ("slomin's", "slomins", "slomin")
CHASE_ENERGIA_KEYWORDS = ("fpl",)
CHASE_TELEFONO_KEYWORDS = ("text me",)
CHASE_ELISTAR_KEYWORDS = ("innov",)
CHASE_CHECK_PROVEEDORES_KEYWORDS = (
    "coca",
    "coke",
    "midtown",
    "king",
    "liu",
    "icecream",
    "ice cream",
)
CHASE_SEGURO_KEYWORDS = ("ipf",)
CHASE_SALE_TAX_KEYWORDS = ("fla dept",)
CHASE_ALG_DISTR_KEYWORDS = ("alg distr",)
CHASE_ADMIN_FEE_KEYWORDS = ("finova",)

# Alfonso single-string master rules (keyword substring -> Detalle).
ALFONSO_MASTER_SINGLE_RULES = (
    ("orig co name:mvnt", "REBATE"),
    ("orig co name:fla lottery", "LOTTERY"),
    ("orig co name:cantaloupe", "VENTA ICE"),
    ("online realtime vendor payment", "SUELDOS"),
    ("online realtime payroll payment", "SUELDOS"),
    ("deposit  id number", "DEPOSITO"),
    ("spectrum", "INTERNET"),
    ("jeffrey's lawn", "REPARACION Y MANTENIMIENTO"),
    ("jeffreys lawn", "REPARACION Y MANTENIMIENTO"),
    ("merchant bank", "COMISIONES Y GASTOS BANCARIOS"),
    ("reynolds", "REBATE"),
    ("fla lottery", "LOTTERY"),
    ("cantaloupe", "VENTA ICE"),
    ("mvnt", "REBATE"),
)

# Alfonso compound rules for UI display (read-only baseline).
ALFONSO_COMPOUND_DISPLAY_RULES = (
    ("OPERATING ACCT + PAYMENT", "EFT RCV-"),
    ("OPERATING ACCT + CHEVRON", "REBATE COMBUSTIBLE"),
    ("OPERATING ACCT + MONTHLY", "REBATE COMBUSTIBLE"),
    (
        "CHECK/CHEQUE/CHEK + COCA|COKE|MIDTOWN|KING|LIU|ICECREAM",
        "PROVEEDORES",
    ),
    ("CHECK/CHEQUE/CHEK + FLORIDA", "ADMINISTRATION FEE"),
    ("CHECK - FLORI", "PROVEEDORES"),
)


def builtin_display_rules():
    """Hardcoded Alfonso master rules mirrored from categorize_chase_description, for UI display only."""
    rules = []

    def add_keywords(keywords, detail):
        for keyword in keywords:
            rules.append({"keyword": keyword, "detail": detail, "source": "Maestra"})

    add_keywords(CHASE_PROVEEDORES_KEYWORDS, "PROVEEDORES")
    add_keywords(CHASE_GASTOS_BANCARIOS_KEYWORDS, "GASTOS BANCARIOS")
    add_keywords(CHASE_REBATE_KEYWORDS, "REBATE")
    add_keywords(CHASE_AGUA_KEYWORDS, "AGUA")
    add_keywords(CHASE_ALARMA_KEYWORDS, "ALARMA")
    add_keywords(CHASE_ENERGIA_KEYWORDS, "ENERGIA ELECTRICA")
    add_keywords(CHASE_TELEFONO_KEYWORDS, "TELEFONO")
    add_keywords(CHASE_ELISTAR_KEYWORDS, "ELISTAR")
    add_keywords(CHASE_SEGURO_KEYWORDS, "SEGURO")
    add_keywords(CHASE_SALE_TAX_KEYWORDS, "SALE TAX")
    add_keywords(CHASE_ALG_DISTR_KEYWORDS, "REBATE")
    add_keywords(CHASE_ADMIN_FEE_KEYWORDS, "ADMINISTRATION FEE")

    for keyword, detail in ALFONSO_MASTER_SINGLE_RULES:
        rules.append({"keyword": keyword, "detail": detail, "source": "Maestra"})

    for keyword, detail in ALFONSO_COMPOUND_DISPLAY_RULES:
        rules.append({"keyword": keyword, "detail": detail, "source": "Maestra"})

    return rules


def list_display_rules():
    """
    Merge built-in Alfonso master rules and persisted JSON rules for the UI
    grid -- "Maestra" rules are read-only/protected, "Personalizada" rules
    carry a dynamic_index so the caller can delete them via
    delete_dynamic_rule_by_index.
    """
    merged = []
    seen = set()

    for rule in builtin_display_rules():
        key = (rule["keyword"].lower(), rule["detail"])
        if key not in seen:
            seen.add(key)
            merged.append(rule)

    for idx, rule in enumerate(load_dynamic_rules()):
        entry = {
            "keyword": rule["keyword"],
            "detail": rule["detail"],
            "source": "Personalizada",
            "dynamic_index": idx,
        }
        key = (entry["keyword"].lower(), entry["detail"])
        if key not in seen:
            seen.add(key)
            merged.append(entry)

    return merged


def _is_check_proveedores(desc):
    """CHECK/CHEQUE/CHEK combined with vendor keywords -> PROVEEDORES."""
    return _desc_has_check_keyword(desc) and _desc_contains_any(
        desc, CHASE_CHECK_PROVEEDORES_KEYWORDS
    )


def _is_deposit_id_number(desc):
    """DEPOSIT ID NUMBER with flexible spacing."""
    return "deposit  id number" in desc or (
        "deposit" in desc and "id number" in desc
    )


def categorize_chase_description(description):
    """
    Map Chase Description text to Detalle category using Alfonso's keyword rules.

    Returns category string or None when no rule matches.
    """
    desc = normalize_rule_text(description)
    if not desc:
        return None

    # Double Rule 1: OPERATING ACCT + PAYMENT (after Chevron/Monthly checks).
    if "operating acct" in desc and "chevron" in desc:
        return "REBATE COMBUSTIBLE"
    if "operating acct" in desc and "monthly" in desc:
        return "REBATE COMBUSTIBLE"
    if "operating acct" in desc and "payment" in desc:
        return "EFT RCV-"

    # Double Rule 5: CHECK/CHEQUE/CHEK + FLORIDA.
    if _desc_has_check_keyword(desc) and "florida" in desc:
        return "ADMINISTRATION FEE"

    # Single Rule: CHECK - FLORI.
    if "chek - flori" in desc:
        return "PROVEEDORES"

    # Double Rule 4: CHECK/CHEQUE/CHEK + vendor keywords.
    if _is_check_proveedores(desc):
        return "PROVEEDORES"

    if _desc_contains_any(desc, CHASE_GASTOS_BANCARIOS_KEYWORDS):
        return "GASTOS BANCARIOS"
    if _desc_contains_any(desc, CHASE_PROVEEDORES_KEYWORDS):
        return "PROVEEDORES"
    if _desc_contains_any(desc, CHASE_REBATE_KEYWORDS):
        return "REBATE"
    if _desc_contains_any(desc, CHASE_AGUA_KEYWORDS):
        return "AGUA"
    if _desc_contains_any(desc, CHASE_ALARMA_KEYWORDS):
        return "ALARMA"
    if _desc_contains_any(desc, CHASE_ENERGIA_KEYWORDS):
        return "ENERGIA ELECTRICA"
    if _desc_contains_any(desc, CHASE_TELEFONO_KEYWORDS):
        return "TELEFONO"
    if _desc_contains_any(desc, CHASE_ELISTAR_KEYWORDS):
        return "ELISTAR"
    if _desc_contains_any(desc, CHASE_SEGURO_KEYWORDS):
        return "SEGURO"
    if _desc_contains_any(desc, CHASE_SALE_TAX_KEYWORDS):
        return "SALE TAX"
    if _desc_contains_any(desc, CHASE_ALG_DISTR_KEYWORDS):
        return "REBATE"
    if _desc_contains_any(desc, CHASE_ADMIN_FEE_KEYWORDS):
        return "ADMINISTRATION FEE"

    for keyword, detail in ALFONSO_MASTER_SINGLE_RULES:
        if keyword in desc:
            return detail
    if _is_deposit_id_number(desc):
        return "DEPOSITO"

    dynamic_detail = match_dynamic_detalle(description)
    if dynamic_detail:
        return dynamic_detail

    return None


# ---------------------------------------------------------------------------
# Read / categorize / write pipeline
# ---------------------------------------------------------------------------

CHASE_COL_POSTING_DATE = 2  # B
CHASE_COL_AMOUNT = 4  # D
CHASE_COL_BALANCE = 6  # F
CHASE_DATE_NUMBER_FORMAT = "dd/mm/yyyy"
CHASE_INTEGER_NUMBER_FORMAT = "0"
CHASE_DECIMAL_NUMBER_FORMAT = "0.00"
CHASE_DATA_START_ROW = 2


def detalle_cell_is_empty(value):
    """True when Chase Detalle (Column H) has no pre-existing content."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return True
    return str(value).strip() == ""


def find_chase_column(df, name_hint, fallback_index):
    """Locate a column by header name or fixed Excel column index."""
    target = name_hint.strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == target:
            return col
    for col in df.columns:
        if target in str(col).strip().lower():
            return col
    if 0 <= fallback_index < len(df.columns):
        return df.columns[fallback_index]
    return None


def ensure_detalle_column(df):
    """Ensure Detalle exists as column H (8th column); create if missing."""
    detalle_col = find_chase_column(df, "Detalle", 7)
    if detalle_col is not None:
        return detalle_col

    if len(df.columns) >= 8:
        return df.columns[7]

    while len(df.columns) < 7:
        df[f"__col_{len(df.columns)}__"] = ""
    df["Detalle"] = ""
    return "Detalle"


def apply_chase_amount_cell(cell, amount):
    """Apply integer or Spanish-decimal Excel format based on the amount value."""
    amount_float = float(amount)
    if amount_float.is_integer():
        cell.value = int(amount_float)
        cell.number_format = CHASE_INTEGER_NUMBER_FORMAT
    else:
        cell.value = amount_float
        cell.number_format = CHASE_DECIMAL_NUMBER_FORMAT


def read_chase_activity_file(file_path):
    """Read Chase bank activity from CSV or Excel."""
    extension = os.path.splitext(file_path)[1].lower()
    if extension == ".csv":
        return pd.read_csv(file_path, dtype=str, keep_default_na=False)
    if extension in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(file_path, dtype=str, keep_default_na=False)
    raise ValueError(
        f"Tipo de archivo no soportado '{extension}'. Use CSV o Excel (.csv, .xlsx, .xlsm)."
    )


def parse_amount_signed(value):
    """Parse amount text into a signed native float."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()
    if text.startswith("-"):
        negative = True
        text = text[1:].strip()

    text = text.replace("$", "").replace(" ", "")

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.rsplit(",", 1)
        if len(parts) == 2 and len(parts[1]) <= 2:
            text = parts[0].replace(".", "") + "." + parts[1]
        else:
            text = text.replace(",", "")

    try:
        amount = float(text)
    except ValueError:
        return 0.0

    return -abs(amount) if negative else amount


def parse_posting_date_value(value):
    """Return datetime for Column B or None when parsing fails."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    if isinstance(value, datetime):
        return value

    text = str(value).strip()
    if not text:
        return None

    for fmt in (
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%m-%d-%Y",
        "%Y/%m/%d",
    ):
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


def sync_detalle_from_excel_workbook(df, file_path, detalle_col):
    """Load pre-existing Column H values from Excel before categorization."""
    extension = os.path.splitext(file_path)[1].lower()
    if extension not in {".xlsx", ".xlsm"}:
        return

    workbook = load_workbook(file_path, read_only=True, data_only=True)
    worksheet = workbook.active
    try:
        for row_offset, index in enumerate(df.index):
            excel_row = CHASE_DATA_START_ROW + row_offset
            cell_value = worksheet.cell(row=excel_row, column=8).value
            if not detalle_cell_is_empty(cell_value):
                df.at[index, detalle_col] = cell_value
    finally:
        workbook.close()


def apply_chase_detalle_categorization(df, description_col, detalle_col):
    """
    Apply keyword rules to Detalle (Column H) only for empty cells.

    Rows with any pre-existing Detalle value are left untouched.
    """
    updated_count = 0
    for index in df.index:
        current = df.at[index, detalle_col]
        if not detalle_cell_is_empty(current):
            continue

        description = df.at[index, description_col]
        category = categorize_chase_description(description)
        if category:
            df.at[index, detalle_col] = category
            updated_count += 1
    return updated_count


def write_chase_excel_file(df, file_path, column_map):
    """Write Chase data to Excel with dates, amounts, balance formulas, and formats."""
    workbook = load_workbook(file_path)
    worksheet = workbook.active

    posting_idx = column_map["posting"]
    amount_idx = column_map["amount"]

    posting_col_name = df.columns[posting_idx]
    amount_col_name = df.columns[amount_idx]
    detalle_col_name = df.columns[column_map["detalle"]]

    protected_detalle = {}
    for row_offset, index in enumerate(df.index):
        excel_row = CHASE_DATA_START_ROW + row_offset
        existing_h = worksheet.cell(row=excel_row, column=8).value
        if not detalle_cell_is_empty(existing_h):
            protected_detalle[excel_row] = existing_h

    for row_offset, index in enumerate(df.index):
        excel_row = CHASE_DATA_START_ROW + row_offset

        for col_num, col_name in enumerate(df.columns, start=1):
            if col_num == 8 and excel_row in protected_detalle:
                continue
            worksheet.cell(row=excel_row, column=col_num, value=df.at[index, col_name])

        posting_dt = parse_posting_date_value(df.at[index, posting_col_name])
        cell_b = worksheet.cell(row=excel_row, column=CHASE_COL_POSTING_DATE)
        if posting_dt:
            cell_b.value = posting_dt
            cell_b.number_format = CHASE_DATE_NUMBER_FORMAT
        else:
            cell_b.value = df.at[index, posting_col_name]

        cell_d = worksheet.cell(row=excel_row, column=CHASE_COL_AMOUNT)
        apply_chase_amount_cell(
            cell_d, parse_amount_signed(df.at[index, amount_col_name])
        )

        cell_f = worksheet.cell(row=excel_row, column=CHASE_COL_BALANCE)
        if excel_row == CHASE_DATA_START_ROW:
            cell_f.value = f"=+F1+D{excel_row}"
        else:
            cell_f.value = f"=+F{excel_row - 1}+D{excel_row}"
        cell_f.number_format = CHASE_DECIMAL_NUMBER_FORMAT

        if excel_row not in protected_detalle:
            detalle_value = str(df.at[index, detalle_col_name]).strip()
            cell_h = worksheet.cell(row=excel_row, column=8)
            cell_h.value = df.at[index, detalle_col_name]
            if detalle_value == EFT_RCV_EXACT_LABEL:
                cell_h.font = EFT_RCV_RED_FONT

    workbook.save(file_path)


def write_chase_csv_file(df, file_path, column_map):
    """Write Chase CSV with formatted dates, floats, and balance formulas."""
    posting_idx = column_map["posting"]
    amount_idx = column_map["amount"]
    balance_idx = column_map["balance"]
    detalle_idx = column_map["detalle"]

    output = df.copy()
    for row_offset, index in enumerate(output.index):
        posting_dt = parse_posting_date_value(output.at[index, output.columns[posting_idx]])
        if posting_dt:
            output.at[index, output.columns[posting_idx]] = posting_dt.strftime("%d/%m/%Y")

        output.at[index, output.columns[amount_idx]] = parse_amount_signed(
            output.at[index, output.columns[amount_idx]]
        )

        excel_row = CHASE_DATA_START_ROW + row_offset
        if excel_row == CHASE_DATA_START_ROW:
            output.at[index, output.columns[balance_idx]] = f"=+F1+D{excel_row}"
        else:
            output.at[index, output.columns[balance_idx]] = (
                f"=+F{excel_row - 1}+D{excel_row}"
            )

    output.to_csv(file_path, index=False)


def process_chase_categorization(file_path):
    """
    Categorize Chase activity, format columns B/D/F, and save the workbook.

    Column B: DD/MM/YYYY dates
    Column D: signed floats with comma decimal display
    Column F: F1 static baseline; F2 =+F1+D2; row 3+ =+F{r-1}+D{r}
    Column H: Detalle keyword categorization
    """
    df = read_chase_activity_file(file_path)

    description_col = find_chase_column(df, "Description", 2)
    if description_col is None:
        raise ValueError(
            "No se encontró la columna Description (se esperaba en la Columna C)."
        )

    posting_col = find_chase_column(df, "Posting Date", 1)
    if posting_col is None:
        posting_col = find_chase_column(df, "Posting", 1)
    amount_col = find_chase_column(df, "Amount", 3)
    balance_col = find_chase_column(df, "Balance", 5)

    if posting_col is None or amount_col is None or balance_col is None:
        raise ValueError(
            "No se encontraron las columnas requeridas (B: Posting Date, D: Amount, F: Balance)."
        )

    detalle_col = ensure_detalle_column(df)
    if detalle_col not in df.columns:
        df[detalle_col] = ""

    sync_detalle_from_excel_workbook(df, file_path, detalle_col)

    updated_count = apply_chase_detalle_categorization(
        df, description_col, detalle_col
    )

    column_map = {
        "posting": list(df.columns).index(posting_col),
        "amount": list(df.columns).index(amount_col),
        "balance": list(df.columns).index(balance_col),
        "detalle": list(df.columns).index(detalle_col),
    }

    extension = os.path.splitext(file_path)[1].lower()
    if extension == ".csv":
        write_chase_csv_file(df, file_path, column_map)
    elif extension in {".xlsx", ".xlsm"}:
        write_chase_excel_file(df, file_path, column_map)
    elif extension == ".xls":
        raise ValueError(
            "El formato .xls antiguo no soporta fórmulas. Guarde como .xlsx e intente de nuevo."
        )
    else:
        raise ValueError(f"No se puede guardar un tipo de archivo no soportado '{extension}'.")

    return updated_count, len(df)
