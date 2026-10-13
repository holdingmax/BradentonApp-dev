"""
Facturas varias de J.H. Williams (pedido del usuario, 2026-10-13): las que
no son de combustible y que también paga el EFT. Unas vienen todos los meses
(Mako Network $131, VPN $7, Chevron Network Fee $171) y otras de vez en
cuando (gift gas cards, charge backs, mystery shop, ajustes chicos, a veces
en negativo). Se suben en Carga de Datos -> EFT y Cupones -> Facturas varias
y Próximos EFT (eft_pronostico.py) las suma en el EFT que las va a pagar.

Lectura, sin UI; se guardan en jh_facturas_db.py. Dos formatos de PDF con
texto, los dos del portal de J.H.:
- Desde 2025: "Invoice No: SI-232163", "Invoice Date", "Invoice Due Date",
  renglones "1.0 NETWORKFEEMAKO MAKO NETWORK FEE 131.00000 131.00000 131.00"
  y dos totales que tienen que coincidir ("INVOICE TOTAL" y "TOTAL INVOICE
  AMOUNT DUE").
- Hasta 2024: "Invoice No: 2024405", "Due Date", el código del producto en
  un renglón ("NETWORKFEEMAKO EACH") y la descripción con cantidad, precio e
  importe en el siguiente; un solo total ("Total Amount").
En los dos, los renglones tienen que sumar el total impreso: si no, no se
guarda. Una factura de combustible, un EFT o un reporte mensual se rechazan
con un aviso que dice dónde se cargan.
"""

import os
import re
from datetime import datetime, timedelta

import pdfplumber

# Tipo de cada factura: clave -> nombre. Las tres primeras vienen todos los meses.
KINDS = {
    "mako": "Mako Network Fee",
    "vpn": "VPN",
    "network": "Chevron Network Fee",
    "gift": "Gift Gas Cards",
    "chargeback": "Credit Card Charge Back",
    "mystery": "Mystery Shop",
    "otra": "Otra",
}
MONTHLY_KINDS = ("mako", "vpn", "network")
TERMS_DAYS = 5  # "EFT NET 5": si la factura no trae el vencimiento

_MONEY = r"-?\$?-?[\d,]*\.\d{2}"
_NEW_LINE_RE = re.compile(r"^(-?\d+\.\d+)\s+(.*?)\s*(-?[\d,]*\.\d{5})\s+(-?[\d,]*\.\d{5}|inf)\s+(-?[\d,]+\.\d{2})$")
_OLD_LINE_RE = re.compile(r"^(.*?)\s+(-?[\d,]+\.\d+)\s+(-?[\d,]*\.\d{6})\s+(-?[\d,]+\.\d{2})$")
_REPORT_TITLES = ("Invoice History", "EFT History", "Credit Card Daily Summary")


def label(invoice):
    """Nombre para mostrar: el tipo, o lo que dice la factura si es de otro tipo."""
    if invoice["kind"] in KINDS and invoice["kind"] != "otra":
        return KINDS[invoice["kind"]]
    text = invoice.get("description") or invoice.get("detail") or "Factura de J.H."
    return text.split(" · ")[0].title()


def _money(text):
    value = text.replace("$", "").replace(",", "")
    negative = value.count("-") % 2 == 1
    return round(-float(value.replace("-", "")) if negative else float(value.replace("-", "")), 2)


def _date(text):
    return datetime.strptime(text, "%m/%d/%Y").date()


def _kind(text):
    upper = text.upper()
    if "MAKO" in upper:
        return "mako"
    if "VPN" in upper:
        return "vpn"
    if "NETWORK FEE" in upper or "NETWORKFEE" in upper or "CHEVNTWF" in upper:
        return "network"
    if "GIFT" in upper or "PPGAS" in upper:
        return "gift"
    if "CHARGE BACK" in upper or "CHGBK" in upper:
        return "chargeback"
    if "MYSTERY" in upper or "MYSTSHOP" in upper:
        return "mystery"
    return "otra"


def _without_code(description):
    """"NETWORKFEEMAKO MAKO NETWORK FEE" -> "MAKO NETWORK FEE": el código del producto va primero."""
    words = description.split()
    return " ".join(words[1:]) if len(words) > 1 else description.strip()


def _is_fuel(text):
    return bool(re.search(r"\bOCTANE\b", text) or re.search(r"^\S+ GAL$", text, re.M)
                or re.search(r"BOL Number:\s*\d", text))


def _new_format(text, lines):
    """Formato desde 2025: renglones y comentarios entre el encabezado y los totales."""
    items, comments = [], []
    for line in lines:
        if line.startswith("INVOICE TOTAL"):
            break
        match = _NEW_LINE_RE.match(line)
        if match:
            items.append({"description": _without_code(match.group(2)), "amount": _money(match.group(5))})
    # Comentarios: entre "Source Standard Ship Via" y la cabecera de los renglones.
    start = next((i for i, line in enumerate(lines) if line.startswith("Source Standard")), None)
    for line in lines[start + 1:] if start is not None else []:
        if line.startswith(("Unit Price", "Total Price", "Qty")):
            break
        if line.startswith("Terms"):
            # "Terms EFT NET 5 Origin BATCH 139 DATED 09-12-26": lo que sigue a Origin
            # es un comentario, salvo el código de origen (CHEVMAKO, CHEVNTWF).
            line = line.split("Origin", 1)[1].strip() if "Origin" in line else ""
            if re.fullmatch(r"CHEV\w*", line):
                line = ""
        if line:
            comments.append(line)
    invoice_total = re.search(r"INVOICE TOTAL:\s*(" + _MONEY + ")", text)
    amount_due = re.search(r"TOTAL INVOICE AMOUNT DUE:\s*(" + _MONEY + ")", text)
    if not invoice_total or not amount_due:
        raise ValueError("no se encontró el total")
    total = _money(amount_due.group(1))
    if total != _money(invoice_total.group(1)):
        raise ValueError("los dos totales impresos no coinciden")
    return items, comments, total


def _old_format(text, lines):
    """Formato hasta 2024: "CODIGO EACH", después "DESCRIPCION cant. precio importe" y comentarios."""
    items, comments = [], []
    started = False
    for line in lines:
        if line.startswith("Total Amount"):
            break
        if line.startswith("Description Shipped"):
            started = True
            continue
        if not started or line.endswith(" EACH"):
            continue
        match = _OLD_LINE_RE.match(line)
        if match:
            items.append({"description": match.group(1).strip(), "amount": _money(match.group(4))})
        elif items:
            comments.append(line)
    total = re.search(r"Total Amount:\s*(" + _MONEY + ")", text)
    if not total:
        raise ValueError("no se encontró el total")
    return items, comments, _money(total.group(1))


def extract_invoice(pdf_path):
    """
    Lee una factura varia de J.H. y devuelve {"invoice_number", "invoice_date",
    "due_date", "amount", "kind", "description", "detail"}. ValueError con un
    aviso corto si no es una factura varia de J.H. o si algo no se lee con
    confianza (nunca se guarda un importe dudoso).
    """
    name = os.path.basename(pdf_path)
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception:
        raise ValueError(f"{name}: no se pudo abrir el PDF.")
    if not text.strip():
        raise ValueError(f"{name}: el PDF no tiene texto (¿es un escaneo?). Bajalo de nuevo del portal de J.H.")
    if "EFT NOTICE" in text[:300] or re.search(r"(Draft No|Transaction ID):\s*RCV-", text):
        raise ValueError(f"{name}: es un EFT; va en \"EFT (PDF del banco)\".")
    if any(title in text[:400] for title in _REPORT_TITLES):
        raise ValueError(f"{name}: es un reporte mensual de J.H.; va en \"Reportes mensuales de J.H.\".")
    number = re.search(r"Invoice No:\s*(SI-\d+|\d+)", text)
    if not number:
        raise ValueError(f"{name}: no es una factura de J.H. Williams.")
    if _is_fuel(text):
        raise ValueError(f"{name}: es una factura de combustible; se carga en Combustible.")
    issued = re.search(r"Invoice Date:\s*(\d{1,2}/\d{1,2}/\d{4})", text)
    due = re.search(r"Due Date:\s*(\d{1,2}/\d{1,2}/\d{4})", text)
    if not issued:
        raise ValueError(f"{name}: no se pudo leer la fecha de la factura.")
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    try:
        items, comments, total = (_new_format if "INVOICE TOTAL" in text else _old_format)(text, lines)
    except ValueError as exc:
        raise ValueError(f"{name}: {exc}; cargala de nuevo o avisá.")
    if not items:
        raise ValueError(f"{name}: no se pudieron leer los renglones de la factura.")
    if round(sum(i["amount"] for i in items), 2) != total:
        raise ValueError(f"{name}: los renglones no suman el total impreso ({total:,.2f}).")
    invoice_date = _date(issued.group(1))
    descriptions = [i["description"] for i in items if i["description"]]
    return {
        "invoice_number": number.group(1),
        "invoice_date": invoice_date,
        "due_date": _date(due.group(1)) if due else invoice_date + timedelta(days=TERMS_DAYS),
        "amount": total,
        "kind": _kind(text),
        "description": " / ".join(dict.fromkeys(descriptions)) or None,
        "detail": " · ".join(comments) or None,
    }
