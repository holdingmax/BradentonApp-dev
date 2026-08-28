"""
Persistent keyword-to-sheet-name rules for routing Chase bank payments to the
right supplier sheet in the Proveedores ledger. Mirrors chase_rules.py's
JSON-backed rule engine (same normalization, same longest-keyword-wins
matching), but maps a bank description keyword to a ledger sheet name
instead of a Detalle category.
"""

import json
import os
import re
import unicodedata

RULES_FILENAME = "proveedores_pago_rules.json"

# Seed rules for the 22 suppliers already automated in SUPPLIER_REGISTRY.
# Where the Chase Bank tab's own hardcoded categorization keywords
# (CHASE_PROVEEDORES_KEYWORDS / CHASE_CHECK_PROVEEDORES_KEYWORDS in app.py)
# already proved a real bank-description fragment for a supplier, that
# fragment is used here too -- it matches real Chase text better than the
# bare sheet name. The rest fall back to the sheet name itself. Suppliers
# outside these 22 (Coca-Cola, Pepsi, J.H.W., Manatee, Liu, Slush Puppies,
# J.J. Taylor) are deliberately NOT seeded -- their real ledger sheet names
# aren't confirmed in code, so the user adds those rules by hand the first
# time an unmatched payment surfaces them.
DEFAULT_RULES = [
    {"keyword": "hackneyrectampa", "sheet_name": "HT Hackney"},
    {"keyword": "cec distributing", "sheet_name": "Chinook CEC"},
    {"keyword": "colonial", "sheet_name": "Colonial"},
    {"keyword": "gold coast eagle", "sheet_name": "GOLDCE"},
    {"keyword": "frito-la", "sheet_name": "FRITO-LAY"},
    {"keyword": "king", "sheet_name": "KING'S"},
    {"keyword": "redbull", "sheet_name": "RED BULL"},
    {"keyword": "SWEETHEART-ICE CREAM", "sheet_name": "SWEETHEART-ICE CREAM"},
    {"keyword": "BIMBO", "sheet_name": "BIMBO"},
    {"keyword": "midtown", "sheet_name": "MIDTOWN"},
    {"keyword": "johnson brothers", "sheet_name": "JOHNSON"},
    {"keyword": "FLORI-GAS", "sheet_name": "FLORI-GAS"},
    {"keyword": "airgas", "sheet_name": "AIRGAS"},
    {"keyword": "AZ Sout", "sheet_name": "AZ Sout"},
    {"keyword": "EXPRESS", "sheet_name": "EXPRESS "},
    {"keyword": "KOOLER", "sheet_name": "KOOLER ICE"},
    {"keyword": "SAM'S", "sheet_name": "SAM'S"},
    {"keyword": "FS WHOLESALE", "sheet_name": "FS WHOLESALE"},
    {"keyword": "LMT", "sheet_name": "LMT"},
    {"keyword": "OVERFLOW", "sheet_name": "OVERFLOW"},
    {"keyword": "SWISHER", "sheet_name": "SWISHER"},
    {"keyword": "SIGNARAMA", "sheet_name": "SIGNARAMA"},
]


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


def _normalize_rule_entry(keyword, sheet_name):
    keyword_text = str(keyword).strip()
    sheet_name_text = str(sheet_name).strip()
    if not keyword_text or not sheet_name_text:
        raise ValueError("La palabra clave y la hoja de proveedor son obligatorias.")
    return {"keyword": keyword_text, "sheet_name": sheet_name_text}


def _seed_default_rules_file(path):
    directory = os.path.dirname(path)
    os.makedirs(directory, exist_ok=True)
    payload = {"rules": DEFAULT_RULES}
    temp_path = f"{path}.tmp"
    with open(temp_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, ensure_ascii=False)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp_path, path)


def load_dynamic_rules():
    """Return all persisted rule dicts: {"keyword": str, "sheet_name": str}."""
    path = _rules_file_path()
    if not os.path.isfile(path):
        _seed_default_rules_file(path)

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
        sheet_name = str(item.get("sheet_name", "")).strip()
        if keyword and sheet_name:
            cleaned.append({"keyword": keyword, "sheet_name": sheet_name})
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


def add_dynamic_rule(keyword, sheet_name):
    """Append or update one rule while preserving every other saved rule."""
    entry = _normalize_rule_entry(keyword, sheet_name)
    rules = load_dynamic_rules()
    keyword_key = entry["keyword"].lower()
    rules = [rule for rule in rules if rule["keyword"].lower() != keyword_key]
    rules.append(entry)
    save_dynamic_rules(rules)
    return entry


def delete_dynamic_rule_by_index(index):
    """Remove one persisted rule by index in proveedores_pago_rules.json."""
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


def match_supplier_sheet(description):
    """
    Case-insensitive substring match against saved keyword rules.

    Among every rule whose keyword appears in the description, the LONGEST
    keyword wins -- same criterion as chase_rules.match_dynamic_detalle, so a
    generic rule doesn't permanently shadow a more specific one added later.

    Returns the target sheet_name, or None.
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
    return best_rule.get("sheet_name") if best_rule else None
