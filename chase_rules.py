"""
Persistent keyword-to-Detalle rules for Chase bank activity categorization.
"""

import json
import os
import re
import unicodedata

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

    Returns Target Detail text or None.
    """
    desc = normalize_rule_text(description)
    if not desc:
        return None

    for rule in load_dynamic_rules():
        keyword = normalize_rule_text(rule.get("keyword", ""))
        if keyword and keyword in desc:
            return rule.get("detail")
    return None
