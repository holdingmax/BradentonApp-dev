"""
Ice Machine y Food Truck (pedido del usuario, 2026-10-07): control del mes de
la máquina de hielo, los Food Truck y Vaccumms contra Chase. Lectura y
cálculo, sin UI; la base de los Payment Summary es ice_machine_db.py y los
recibos de depósito son los del módulo Depósitos (depositos.py/depositos_db.py).

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

ICE_MACHINE = "Ice Machine"
FOOD_TRUCK = "Food Truck"
VACCUMMS = "Vaccumms"
DEPOSIT_KINDS = (ICE_MACHINE, FOOD_TRUCK, VACCUMMS)
# Cómo queda categorizado en Chase cada tipo (detalle).
CHASE_DETALLE = {ICE_MACHINE: "DEPOSITO VENTA ICE", FOOD_TRUCK: "FOOD TRUCK", VACCUMMS: "VACCUMMS"}

_MONEY = r"-?\$\s?-?[\d,]+\.\d{2}"


def deposit_kind(kind):
    """El tipo de un recibo de Depósitos si es de este módulo (Ice/Food Truck/Vaccumms), o None."""
    lowered = (kind or "").lower()
    if re.search(r"\b(ice|hielo)\b", lowered):
        return ICE_MACHINE
    if re.search(r"\b(food|truck)\b", lowered):
        return FOOD_TRUCK
    if re.search(r"\bvac", lowered):
        return VACCUMMS
    return None


def _money(text):
    value = text.replace("$", "").replace(",", "").replace(" ", "")
    negative = value.count("-") % 2 == 1
    return -float(value.replace("-", "")) if negative else float(value.replace("-", ""))


def _undouble(line):
    """El encabezado sale con cada letra repetida ("FFrroomm::MMaayy"): se vuelve a "From:May"."""
    words = []
    for word in line.split(" "):
        if len(word) >= 2 and len(word) % 2 == 0 and all(word[i] == word[i + 1] for i in range(0, len(word), 2)):
            word = word[::2]
        words.append(word)
    return " ".join(words)


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
# Control del mes contra Chase
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


def chase_deposits(rows):
    """Depósitos de Chase (descripción DEPOSIT, importe positivo)."""
    return [{"date": r["posting_date"], "amount": r["amount"], "detalle": r.get("detalle")}
            for r in rows if (r.get("amount") or 0) > 0 and (r.get("description") or "").upper().startswith("DEPOSIT")]


def _detalle_kind(detalle):
    upper = (detalle or "").upper()
    for kind, name in CHASE_DETALLE.items():
        if upper == name or (kind == VACCUMMS and upper.startswith("VAC")):
            return kind
    return None


def build_month_control(year, month, summaries, deposits, chase_rows, chase_last_date, known_references=()):
    """
    summaries: los Payment Summary del mes (por fecha "To"); deposits: los
    recibos de Depósitos del mes de tipo Ice/Food Truck/Vaccumms
    ({id, deposit_date, amount, tx_number, kind}); chase_rows: los
    movimientos de Chase del mes; chase_last_date: último día cargado de
    Chase (ISO) o None. Lo posterior a ese día queda "todavía no está en
    Chase", fuera del control. known_references: los Reference de los
    resumenes del mes anterior (un pago que entra los primeros días del mes
    ya tiene su Payment Summary allá).
    """
    prefix = f"{year:04d}-{month:02d}-"
    covered = (lambda iso: chase_last_date is not None and iso <= chase_last_date)
    issues = []

    # Payment Summary <-> pago de Cantaloupe, por Reference (= IND ID) e importe.
    payments = chase_cantaloupe(chase_rows)
    used = set()
    summary_rows = []
    for s in sorted(summaries, key=lambda s: s["to_date"]):
        match = next((i for i, p in enumerate(payments) if i not in used and p["reference"] == s["reference"]), None)
        if match is None:  # pagos viejos sin IND ID: misma fecha e importe
            match = next((i for i, p in enumerate(payments) if i not in used and p["reference"] is None
                          and p["date"] == s["to_date"] and abs(p["amount"] - s["net"]) <= TOLERANCE), None)
        row = dict(s, chase=None, status=None)
        if match is not None:
            used.add(match)
            row["chase"] = payments[match]
            row["status"] = "ok" if abs(payments[match]["amount"] - s["net"]) <= TOLERANCE else "diff"
            if row["status"] == "diff":
                issues.append(f"Payment Summary #{s['summary_no']}: Chase acreditó ${payments[match]['amount']:,.2f} "
                              f"y el resumen dice ${s['net']:,.2f}.")
        elif covered(s["to_date"]):
            row["status"] = "missing"
            issues.append(f"Payment Summary #{s['summary_no']} (${s['net']:,.2f} del {s['to_date'][8:10]}/{s['to_date'][5:7]}) "
                          "no está en Chase.")
        else:
            row["status"] = "after"
        summary_rows.append(row)
    orphan_payments = [p for i, p in enumerate(payments) if i not in used and p["date"].startswith(prefix)
                       and p["reference"] not in set(known_references)]
    for p in orphan_payments:
        issues.append(f"Chase: pago de Cantaloupe de ${p['amount']:,.2f} del {p['date'][8:10]}/{p['date'][5:7]} "
                      "sin su Payment Summary.")

    # Recibo <-> depósito de Chase, misma fecha e importe.
    chase_deps = chase_deposits(chase_rows)
    taken = set()
    deposit_rows = []
    for d in sorted(deposits, key=lambda d: (d["deposit_date"] or "", d.get("tx_number") or 0)):
        row = dict(d, chase=None, status=None, uncategorized=False)
        if d["deposit_date"] is None or d["amount"] is None:
            row["status"] = "incomplete"
            issues.append(f"Recibo #{d.get('tx_number') or '?'}: falta la fecha o el importe (completalo en Depósitos).")
            deposit_rows.append(row)
            continue
        candidates = [i for i, c in enumerate(chase_deps) if i not in taken and c["date"] == d["deposit_date"]
                      and abs(c["amount"] - d["amount"]) <= TOLERANCE]
        # Si hay varios iguales, primero el que ya está categorizado como este tipo.
        candidates.sort(key=lambda i: _detalle_kind(chase_deps[i]["detalle"]) != d["kind"])
        if candidates:
            taken.add(candidates[0])
            row["chase"] = chase_deps[candidates[0]]
            row["status"] = "ok"
            row["uncategorized"] = _detalle_kind(row["chase"]["detalle"]) != d["kind"]
        elif covered(d["deposit_date"]):
            row["status"] = "missing"
            issues.append(f"{d['kind']}: el depósito de ${d['amount']:,.2f} del {d['deposit_date'][8:10]}/"
                          f"{d['deposit_date'][5:7]} (recibo #{d.get('tx_number') or '?'}) no está en Chase.")
        else:
            row["status"] = "after"
        deposit_rows.append(row)
    orphan_deposits = [dict(c, kind=_detalle_kind(c["detalle"])) for i, c in enumerate(chase_deps)
                       if i not in taken and _detalle_kind(c["detalle"]) and c["date"].startswith(prefix)]
    for c in orphan_deposits:
        issues.append(f"Chase: depósito de {c['kind']} de ${c['amount']:,.2f} del {c['date'][8:10]}/{c['date'][5:7]} "
                      "sin su recibo.")

    def total(items, key):
        return round(sum(i[key] or 0 for i in items), 2)

    ice_deposits = [d for d in deposit_rows if d["kind"] == ICE_MACHINE]
    groups = {kind: [d for d in deposit_rows if d["kind"] == kind] for kind in DEPOSIT_KINDS}
    ice = {
        "transactions": sum(s["transactions"] for s in summary_rows),
        "gross": total(summary_rows, "gross"),
        "process_fees": total(summary_rows, "process_fees"),
        "service_fees": total(summary_rows, "service_fees"),
        "net": total(summary_rows, "net"),
        "cash": total(ice_deposits, "amount"),
    }
    ice["total"] = round(ice["net"] + ice["cash"], 2)
    uncategorized = [d for d in deposit_rows if d["uncategorized"]]
    return {
        "summaries": summary_rows,
        "deposits": deposit_rows,
        "groups": {kind: {"rows": rows, "total": total(rows, "amount")} for kind, rows in groups.items()},
        "ice": ice,
        "orphan_payments": orphan_payments,
        "orphan_deposits": orphan_deposits,
        "uncategorized": uncategorized,
        "issues": issues,
        "chase_last_date": chase_last_date,
        "pending_after": [r for r in summary_rows + deposit_rows if r["status"] == "after"],
        "ok": not issues,
    }


def month_of(iso):
    d = date.fromisoformat(iso)
    return d.year, d.month
