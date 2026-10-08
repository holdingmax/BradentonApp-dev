"""
Ice Machine (pedido del usuario, 2026-10-07): lectura de los Payment Summary
de Cantaloupe, el lector de tarjetas de la máquina de hielo. Sin UI; la base
es ice_machine_db.py y el control contra Chase, junto con todos los
depósitos, está en control_depositos.py.

Lo que se sube:
- Payment Summary de Cantaloupe (el lector de tarjetas de la máquina de
  hielo): PDF con texto, uno por semana. Trae # de ventas, Gross, el fee de
  proceso (5.95%), el "Terminal Service Fee" de cada terminal una vez por
  mes ($9.95 cada uno; hay dos terminales) y el Net. El Net entra a Chase
  como "ORIG CO NAME:Cantaloupe ... IND ID:<Reference #>" el mismo día del
  "To:" del resumen (validado con mayo y julio de 2026: 10 de 10).
- Recibos de depósito ("Transaccion #133 (Ice Machine)", "(Food Truck)",
  "(Vaccumms)"): el efectivo de la máquina de hielo, el alquiler de los Food
  Truck ($1,000 + 3%) y los Vaccumms. Cada recibo es un depósito de Chase con
  la misma fecha e importe.

Venta de hielo del mes = Net de Cantaloupe + efectivo depositado (lo mismo
que calculaba el Excel "Ice Machine MM-YY").
"""

import re
from datetime import date, datetime

TOLERANCE = 0.005

_MONEY = r"-?\$\s?-?[\d,]+\.\d{2}"


def _money(text):
    value = text.replace("$", "").replace(",", "").replace(" ", "")
    negative = value.count("-") % 2 == 1
    return -float(value.replace("-", "")) if negative else float(value.replace("-", ""))


def _undouble(line):
    """
    El encabezado sale con cada letra repetida ("FFrroomm::MMaayy"): se vuelve
    a "From:May". Solo si TODA la línea viene repetida: "Totals: 66" (66
    transacciones) se leía como 6 (revisión 2026-10-08).
    """
    def doubled(word):
        return len(word) >= 2 and len(word) % 2 == 0 and all(word[i] == word[i + 1] for i in range(0, len(word), 2))

    words = [word for word in line.split(" ") if word]
    if not words or not all(doubled(word) for word in words):
        return line
    return " ".join(word[::2] for word in words)


def _parse_date(text):
    return datetime.strptime(re.sub(r"\s+", " ", text.strip()), "%B %d, %Y").date()


def is_payment_summary_text(text):
    return "Payment Summary #" in (text or "") and "Cantaloupe" in (text or "")


def read_payment_summary_text(text):
    """
    Campos de un Payment Summary a partir del texto del PDF. ValueError si
    falta algo o si las cuentas no cierran (Gross + fees = Net).
    """
    lines = [_undouble(line) for line in (text or "").splitlines()]
    joined = "\n".join(lines)

    def find(pattern, label):
        match = re.search(pattern, joined)
        if not match:
            raise ValueError(f"no se encontró {label} en el Payment Summary.")
        return match.group(1)

    number = find(r"Payment Summary #\s*(\d+)", "el número")
    from_date = _parse_date(find(r"From:\s*([A-Za-z]+ \d{1,2}, \d{4})", "la fecha From"))
    to_date = _parse_date(find(r"To:\s*([A-Za-z]+ \d{1,2}, \d{4})", "la fecha To"))
    reference = find(r"Reference #:\s*(\d+)", "el Reference #")
    net = _money(find(r"Net Revenue:\s*(\$\s?-?[\d,]+\.\d{2})", "el Net Revenue"))

    totals = re.search(r"^Totals:\s*(\d+)\s+(.*)$", joined, re.M)
    if totals:
        transactions = int(totals.group(1))
        amounts = [_money(m) for m in re.findall(_MONEY, totals.group(2))]
    else:
        # Un resumen solo con el Terminal Service Fee de la terminal vieja (sin ventas).
        transactions, amounts = 0, []
    gross = amounts[0] if amounts and amounts[0] > 0 else 0.0
    deductions = [a for a in amounts if a < 0]

    # Service Fees: un cargo por renglón con importe negativo ("... Terminal
    # Service Fee 06/20/2026 -$9.95 $0.00 Monthly"; en feb-mar 2026 la tabla
    # sale partida en varios renglones y el importe queda junto a la terminal).
    service = 0.0
    section = re.search(r"^Service Fees\n(.*?)(?=^Adjustments\b|\Z)", joined, re.M | re.S)
    if section:
        for line in section.group(1).splitlines():
            fees = [_money(m) for m in re.findall(_MONEY, line) if _money(m) < 0]
            if fees:
                service += fees[0]
    process = round(sum(deductions) - service, 2)
    other = round(net - gross - sum(deductions), 2)
    if abs(other) > TOLERANCE:
        raise ValueError(
            f"Payment Summary #{number}: el Gross (${gross:,.2f}) menos los fees "
            f"(${-sum(deductions):,.2f}) no da el Net (${net:,.2f}). Revisalo a mano."
        )
    notes = [name for name in ("Chargebacks", "Refunds", "Adjustments")
             if re.search(rf"^{name}\n(?!No )", joined, re.M)]
    return {
        "summary_no": number,
        "reference": reference.lstrip("0") or "0",
        "from_date": from_date.isoformat(),
        "to_date": to_date.isoformat(),
        "transactions": transactions,
        "gross": round(gross, 2),
        "process_fees": process,
        "service_fees": round(service, 2),
        "net": round(net, 2),
        "notes": ", ".join(notes) or None,
    }


def read_payment_summary(pdf_path):
    import pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join((page.extract_text() or "") for page in pdf.pages)
    if not is_payment_summary_text(text):
        raise ValueError("no es un Payment Summary de Cantaloupe.")
    return read_payment_summary_text(text)


# ---------------------------------------------------------------------------
# Pagos de Cantaloupe en Chase
# ---------------------------------------------------------------------------

_IND_ID = re.compile(r"IND ID:\s*0*(\d+)")


def chase_cantaloupe(rows):
    """Pagos de Cantaloupe de Chase: [{date, amount, reference, row}]."""
    found = []
    for row in rows:
        description = row.get("description") or ""
        if "CANTALOUPE" in description.upper() and (row.get("amount") or 0) > 0:
            match = _IND_ID.search(description)
            found.append({"date": row["posting_date"], "amount": row["amount"],
                          "reference": match.group(1) if match else None})
    return found


def month_of(iso):
    d = date.fromisoformat(iso)
    return d.year, d.month
