"""
Reportes mensuales de J.H. Williams (pedido del usuario, 2026-10-06, chat 21):
los tres que se bajan del portal de J.H. para un mes, en PDF:

- "Invoice History": las facturas del mes (combustible, con BOL, y otras).
- "EFT History": los EFT del mes (RCV), con el importe del débito/crédito.
- "Credit Card Daily Summary": los cupones (DDC) del mes, por día.

Lectura y cruces, sin UI. Se guardan en jh_mensual_db.py, se suben en Carga
de Datos → EFT y Cupones y se controlan en /controles/tarjetas:

- EFT: el importe del reporte contra el EFT cargado (facturas pagadas menos
  cupones, por sus líneas) y contra el movimiento "EFT RCV-" de Chase.
- Facturas: cada factura del reporte contra lo que pagaron los EFT cargados.
- Cupones: cada día del reporte contra los cupones cargados (Gross, Neto y
  los DDC) y contra lo que aplicaron los EFT. Lo que queda sin aplicar al
  cierre son los últimos días (acreditación a 72 hs): por eso los cupones
  del mes no coinciden con los cupones de los EFT del mes.

Un PDF cuyos renglones no suman el total impreso se rechaza entero: nunca se
guarda un reporte leído a medias. La única excepción es el PDF que imprime
solo la primera página del portal: si lo que falta para el total son,
justo, renglones ya cargados de los días de al lado, se acepta.
"""

import calendar
import re
from datetime import date, datetime, timedelta

import pdfplumber

import chase_db
import eft_db
import jh_mensual_db

KINDS = ("eft", "invoices", "coupons")
KIND_LABELS = {
    "eft": "EFT History",
    "invoices": "Invoice History",
    "coupons": "Credit Card Daily Summary",
}
_TITLES = (
    ("Credit Card Daily Summary", "coupons"),
    ("Invoice History", "invoices"),
    ("EFT History", "eft"),
)

_MONEY = r"(-?\$[\d,]+\.\d{2})"
_DATE = r"(\d{1,2}/\d{1,2}/\d{4})"
_INVOICE_ROW_RE = re.compile(_DATE + r"\s+([A-Z]{1,4}-\d+)\s+(\S+)\s+(\S+)\s+" + _MONEY + r"\s+(-|" + _MONEY[1:-1] + r")\s*$")
_EFT_ROW_RE = re.compile(_DATE + r"\s+(RCV-\d+)\s+([A-Za-z]+)\s+" + _MONEY + r"\s*$")
_COUPON_ROW_RE = re.compile(r"^" + _DATE + r"\s+((?:DDC-\d+,?)+)\s+(.*?)\s*" + _MONEY + r"\s+" + _MONEY + r"\s+" + _MONEY + r"\s*$")
_PERIOD_RE = re.compile(r"Period:\s*" + _DATE + r"\s*-\s*" + _DATE)
_EFT_SIGN_RCV = "EFT RCV"

# Días de tolerancia entre la fecha de un EFT y su movimiento en Chase.
CHASE_WINDOW_DAYS = 3
# Un cupón sin aplicar cuando ya hay un EFT cargado de estos días después de
# su fecha es un problema, no la demora normal (en agosto-2026 el más lento
# tardó 10 días: un cupón de jueves que no entra en el EFT del lunes).
COUPON_LATE_DAYS = 14


def _money(text):
    return round(float(text.replace("$", "").replace(",", "")), 2)


def _date(text):
    return datetime.strptime(text, "%m/%d/%Y").date()


def _close(a, b):
    return a is not None and b is not None and abs(a - b) < 0.005


def _totals_line(lines, count):
    """Último renglón formado solo por `count` importes: el total impreso del reporte."""
    pattern = re.compile(r"^" + r"\s+".join([_MONEY] * count) + r"$")
    for line in reversed(lines):
        match = pattern.match(line)
        if match:
            return [_money(group) for group in match.groups()]
    return None


_ROW_NOUN = {"eft": ("EFT", "EFT"), "invoices": ("factura", "facturas"), "coupons": ("día", "días")}


def _rows_completing_total(kind, rows, printed, known_rows):
    """
    Los renglones ya cargados que, sumados a los del PDF, dan justo el total
    impreso, o None. El portal imprime la página que se ve, pero el total de
    todo el rango pedido (el Credit Card Daily Summary del 01-06/10 trajo el
    total desde el 06/09): lo que falta son los renglones más viejos (o más
    nuevos) que el PDF, seguidos, y tienen que cerrar todos los importes al
    centavo.
    """
    key, fields = jh_mensual_db.ROW_KEY[kind], jh_mensual_db.TOTAL_FIELDS[kind]
    seen = {r[key] for r in rows}
    first, last = min(r["date"] for r in rows), max(r["date"] for r in rows)
    others = [r for r in known_rows if r[key] not in seen]
    older = sorted((r for r in others if r["date"] < first), key=lambda r: r["date"], reverse=True)
    newer = sorted((r for r in others if r["date"] > last), key=lambda r: r["date"])
    for side in (older, newer):
        sums = [sum(r[f] for r in rows) for f in fields]
        for count, r in enumerate(side, start=1):
            sums = [s + (r.get(f) or 0.0) for s, f in zip(sums, fields)]
            if all(_close(round(s, 2), p) for s, p in zip(sums, printed)):
                return side[:count]
    return None


def extract_report(pdf_path, known_rows=None):
    """
    Lee uno de los tres reportes (se reconoce por el título), de cualquier
    extensión, y lo parte por mes: {"kind", "months": [{"year", "month",
    "from_date", "to_date", "rows"}], "notes"} (pedido del usuario,
    2026-10-06: "que cada mes pueda aceptar el reporte no importa lo extenso
    que sea"). ValueError si no es ninguno de los tres o si los renglones no
    suman el total impreso (se valida el reporte entero, antes de partirlo),
    salvo que lo que falte sean renglones ya cargados (`known_rows(kind)`,
    ver _rows_completing_total): eso queda en `notes`.
    """
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    kind = next((k for title, k in _TITLES if title in text[:400]), None)
    if kind is None:
        raise ValueError("Uno de los archivos no es un reporte de J.H. (Invoice History, EFT History o Credit Card Daily Summary).")
    label = KIND_LABELS[kind]

    rows = []
    if kind == "invoices":
        for line in lines:
            match = _INVOICE_ROW_RE.search(line)
            if match:
                inv_date, number, po, bol, amount, balance = match.groups()[:6]
                rows.append({
                    "date": _date(inv_date).isoformat(), "invoice": number,
                    "po": None if po == "-" else po, "bol": None if bol == "-" else bol,
                    "amount": _money(amount), "balance": 0.0 if balance == "-" else _money(balance),
                })
        printed = _totals_line(lines, 2)
        period = _PERIOD_RE.search(text)
        period = (_date(period.group(1)), _date(period.group(2))) if period else None
    elif kind == "eft":
        for line in lines:
            match = _EFT_ROW_RE.search(line)
            if match:
                eft_date, reference, status, amount = match.groups()
                rows.append({"date": _date(eft_date).isoformat(), "reference": reference, "status": status, "amount": _money(amount)})
        printed = _totals_line(lines, 1)
    else:
        for line in lines:
            match = _COUPON_ROW_RE.match(line)
            if match:
                day, references, batches, gross, fees, net = match.groups()
                rows.append({
                    "date": _date(day).isoformat(), "coupons": [c for c in references.split(",") if c],
                    "batches": batches, "gross": _money(gross), "fees": _money(fees), "net": _money(net),
                })
        printed = _totals_line(lines, 3)

    if printed is None:
        raise ValueError(f"No se encontró el total impreso del {label}: no se guardó.")
    if not rows:
        raise ValueError(f"El {label} no tiene renglones.")
    notes = []
    read = [round(sum(r[f] for r in rows), 2) for f in jh_mensual_db.TOTAL_FIELDS[kind]]
    if not all(_close(value, total) for value, total in zip(read, printed)):
        extra = _rows_completing_total(kind, rows, printed, known_rows(kind) if known_rows else [])
        if extra is None:
            raise ValueError(
                f"Los renglones del {label} suman {read[0]:,.2f} y el total impreso es {printed[0]:,.2f}: "
                "el PDF trae solo una parte (en el portal, mostrá todas las filas antes de imprimir). No se guardó."
            )
        days = sorted(r["date"] for r in extra)
        noun = _ROW_NOUN[kind][len(extra) != 1]
        span = f"{_ddmm(days[0])}" + (f" al {_ddmm(days[-1])}" if days[-1] != days[0] else "")
        notes.append(f"el total impreso también suma {len(extra)} {noun} ya cargados ({span}) y coincide al centavo")
    rows.sort(key=lambda r: r["date"])
    by_month = {}
    for r in rows:
        day = date.fromisoformat(r["date"])
        by_month.setdefault((day.year, day.month), []).append(r)
    months = []
    for (year, month), month_rows in sorted(by_month.items()):
        first = date(year, month, 1)
        last = date(year, month, calendar.monthrange(year, month)[1])
        if kind == "invoices" and period:
            # El Invoice History cubre su período entero, tenga o no facturas cada día.
            start, end = max(period[0], first), min(period[1], last)
        else:
            start, end = date.fromisoformat(month_rows[0]["date"]), date.fromisoformat(month_rows[-1]["date"])
        months.append({"year": year, "month": month, "from_date": start.isoformat(), "to_date": end.isoformat(), "rows": month_rows})
    return {"kind": kind, "months": months, "notes": notes}


def _ddmm(iso):
    return date.fromisoformat(iso).strftime("%d/%m")


# ---------------------------------------------------------------------------
# Cruces
# ---------------------------------------------------------------------------

def _next_month(year, month):
    return (year + 1, 1) if month == 12 else (year, month + 1)


def _deposits(year, month):
    """EFT cargados del mes y del siguiente (una factura o un cupón de fin de mes se paga en el siguiente)."""
    result = []
    for y, m in ((year, month), _next_month(year, month)):
        for item in eft_db.get_month_deposits(y, m):
            dep = item["deposit"]
            eft_day = datetime.strptime(dep["eft_date"], "%m/%d/%Y").date()
            invoices = round(sum(p["paid_amount"] or 0.0 for p in item["paid_invoices"]), 2)
            coupons = round(sum(c["paid_amount"] or 0.0 for c in item["coupons"]), 2)
            result.append({
                "rcv": dep["rcv_number"], "date": eft_day,
                "invoices": invoices, "coupons_net": coupons,
                "draft": round(invoices - coupons, 2),
                "paid_invoices": item["paid_invoices"], "coupon_lines": item["coupons"],
            })
    return result


def _chase_eft_movements(year, month):
    movements = []
    for y, m in ((year, month), _next_month(year, month)):
        for row in chase_db.get_month_transactions(y, m):
            if (row.get("detalle") or "").upper().startswith(_EFT_SIGN_RCV) and row.get("amount") is not None:
                movements.append({"date": date.fromisoformat(row["posting_date"]), "amount": round(row["amount"], 2)})
    return movements


def eft_check(report, year, month):
    """
    Cada EFT del reporte contra el EFT cargado (facturas pagadas menos
    cupones) y contra Chase. Los importes van con el signo del reporte: en
    negativo, lo que J.H. nos acredita (en Chase entra como depósito).
    """
    deposits = _deposits(year, month)
    by_rcv = {d["rcv"]: d for d in deposits}
    movements = _chase_eft_movements(year, month)
    last_chase = chase_db.get_last_posting_date()
    used = set()
    rows = []

    def match_chase(eft_day, draft):
        best = None
        for index, mov in enumerate(movements):
            if index in used or abs((mov["date"] - eft_day).days) > CHASE_WINDOW_DAYS:
                continue
            if _close(-mov["amount"], draft) and (best is None or abs((mov["date"] - eft_day).days) < abs((movements[best]["date"] - eft_day).days)):
                best = index
        if best is not None:
            used.add(best)
            return movements[best]
        return None

    for r in report["rows"]:
        eft_day = date.fromisoformat(r["date"])
        loaded = by_rcv.get(r["reference"])
        chase = match_chase(eft_day, r["amount"])
        chase_covered = last_chase is not None and last_chase >= eft_day + timedelta(days=CHASE_WINDOW_DAYS)
        if loaded is None:
            status = "no_eft"
        elif not _close(loaded["draft"], r["amount"]):
            status = "diff"
        elif chase is None:
            status = "no_chase" if chase_covered else "chase_not_loaded"
        else:
            status = "ok"
        rows.append({
            "date": r["date"], "rcv": r["reference"], "report": r["amount"],
            "loaded": loaded["draft"] if loaded else None,
            "loaded_date": loaded["date"].isoformat() if loaded else None,
            "invoices": loaded["invoices"] if loaded else None,
            "coupons_net": loaded["coupons_net"] if loaded else None,
            "chase_date": chase["date"].isoformat() if chase else None,
            "chase": -chase["amount"] if chase else None,
            "status": status,
        })
    in_report = {r["reference"] for r in report["rows"]}
    extra_loaded = [
        {"rcv": d["rcv"], "date": d["date"].isoformat(), "draft": d["draft"]}
        for d in deposits if d["date"].year == year and d["date"].month == month and d["rcv"] not in in_report
    ]
    extra_chase = [
        {"date": m["date"].isoformat(), "amount": -m["amount"]}
        for index, m in enumerate(movements)
        if index not in used and m["date"].year == year and m["date"].month == month
    ]
    return {
        "rows": rows, "extra_loaded": extra_loaded, "extra_chase": extra_chase,
        "ok": all(r["status"] == "ok" for r in rows) and not extra_loaded and not extra_chase,
        "total": report["totals"]["amount"],
    }


def invoice_check(report, year, month, known_invoices=()):
    """
    Cada factura del reporte contra lo que pagaron los EFT cargados (del mes
    y del siguiente). Una factura que todavía no está en ningún EFT cargado
    queda "Sin EFT" (normal para las últimas del mes si el EFT siguiente no
    se cargó). `known_invoices`: todas las facturas de los Invoice History
    guardados, para avisar las que pagó un EFT del mes y no figuran en
    ninguno (suelen ser del mes anterior, sin su reporte cargado).
    """
    deposits = _deposits(year, month)
    paid = {}
    for d in deposits:
        for p in d["paid_invoices"]:
            paid.setdefault(p["invoice"], []).append({"rcv": d["rcv"], "date": d["date"].isoformat(), "amount": p["paid_amount"] or 0.0})
    rows = []
    for r in report["rows"]:
        payments = paid.get(r["invoice"], [])
        paid_total = round(sum(p["amount"] for p in payments), 2) if payments else None
        if not payments:
            status = "no_eft"
        elif _close(paid_total, r["amount"]):
            status = "ok"
        else:
            status = "diff"
        rows.append({**r, "fuel": bool(r["bol"]), "payments": payments, "paid": paid_total, "status": status})
    known = set(known_invoices) | {r["invoice"] for r in report["rows"]}
    unknown = sorted({
        (p["invoice"], d["rcv"]) for d in deposits if d["date"].year == year and d["date"].month == month
        for p in d["paid_invoices"] if p["invoice"] not in known
    })
    fuel = [r for r in rows if r["fuel"]]
    others = [r for r in rows if not r["fuel"]]
    return {
        "rows": rows,
        "unknown": [{"invoice": inv, "rcv": rcv} for inv, rcv in unknown],
        "fuel_total": round(sum(r["amount"] for r in fuel), 2), "fuel_count": len(fuel),
        "other_total": round(sum(r["amount"] for r in others), 2), "other_count": len(others),
        "total": report["totals"]["amount"],
        "bad": [r for r in rows if r["status"] == "diff"],
        "pending": [r for r in rows if r["status"] == "no_eft"],
        "unpaid_balance": round(sum(r["balance"] for r in rows), 2),
        "ok": all(r["status"] == "ok" for r in rows) and not any(r["balance"] for r in rows),
    }


def _detail_for_row(row, detail_groups):
    """El grupo del detalle de cupones de esa fila del reporte: mismos importes y su último batch hasta 3 días antes."""
    day = date.fromisoformat(row["date"])
    matches = [
        g for g in detail_groups
        if all(abs((g[k] or 0.0) - row[k]) < 0.005 for k in ("gross", "fees", "net"))
        and day - timedelta(days=3) <= date.fromisoformat(g["last_date"]) <= day
    ]
    return matches[0] if len(matches) == 1 else None


def coupon_check(report, year, month, detail_groups=()):
    """
    Cada día del reporte contra los cupones cargados (Gross, Neto y los DDC)
    y contra lo que aplicaron los EFT (del mes y del siguiente). Al cierre,
    un cupón sin aplicar en un EFT del mes es lo que pasa al mes siguiente
    por la acreditación a 72 hs: se informa, no es un error. Con el detalle
    de cupones cargado (`detail_groups`, eft_db.get_detail_groups), cada
    fila muestra los días de venta de sus batches.
    """
    loaded_by_date = eft_db.get_cupones_detail_by_date()
    month_end = date(year, month, calendar.monthrange(year, month)[1])
    applied = {}  # DDC -> [(fecha del EFT, rcv, gross, pagado)]
    deposits = _deposits(year, month)
    last_eft = max((d["date"] for d in deposits), default=None)
    for d in deposits:
        for line in d["coupon_lines"]:
            if line.get("coupon"):
                applied.setdefault(line["coupon"], []).append(
                    (d["date"], d["rcv"], line.get("gross_amount") or 0.0, line.get("paid_amount") or 0.0)
                )
    rows = []
    for r in report["rows"]:
        loaded = loaded_by_date.get(r["date"]) or {"gross": 0.0, "fees": 0.0, "net": 0.0, "coupons": [], "unknown": 0}
        # Un DDC de un grupo todavía sin monto propio (en 0 hasta que un EFT
        # lo paga) no deja comparar el importe del día: alcanza con los DDC.
        loaded_ok = sorted(loaded["coupons"]) == sorted(r["coupons"]) and (
            loaded.get("unknown") or (_close(loaded["gross"], r["gross"]) and _close(loaded["net"], r["net"]))
        )
        lines = [line for ddc in r["coupons"] for line in applied.get(ddc, [])]
        unapplied = [ddc for ddc in r["coupons"] if ddc not in applied]
        applied_gross = round(sum(line[2] for line in lines), 2)
        applied_net = round(sum(line[3] for line in lines), 2)
        efts = sorted({(line[0], line[1]) for line in lines})
        if not loaded["coupons"]:
            status = "not_loaded"
        elif not loaded_ok:
            status = "diff"
        elif unapplied:
            late = last_eft is not None and last_eft >= date.fromisoformat(r["date"]) + timedelta(days=COUPON_LATE_DAYS)
            status = "late" if late else "pending"
        elif not (_close(applied_gross, r["gross"]) and _close(applied_net, r["net"])):
            status = "applied_diff"
        else:
            status = "ok"
        detail = _detail_for_row(r, detail_groups)
        rows.append({
            **r,
            "sale_from": detail["first_date"] if detail else None,
            "sale_to": detail["last_date"] if detail else None,
            "loaded_gross": loaded["gross"], "loaded_net": loaded["net"], "loaded_unknown": loaded.get("unknown", 0),
            "missing_coupons": sorted(set(r["coupons"]) - set(loaded["coupons"])),
            "extra_coupons": sorted(set(loaded["coupons"]) - set(r["coupons"])),
            "applied_gross": applied_gross, "applied_net": applied_net,
            "efts": [{"date": day.isoformat(), "rcv": rcv} for day, rcv in efts],
            "in_month_net": round(sum(line[3] for line in lines if line[0] <= month_end), 2),
            "ddcs_in_month": [ddc for ddc in r["coupons"] if any(line[0] <= month_end for line in applied.get(ddc, []))],
            "unapplied": unapplied,
            # Lo que todavía no aplicó ningún EFT cargado (un día puede tener
            # un DDC aplicado y otro no).
            "open_gross": round(r["gross"] - applied_gross, 2) if unapplied else 0.0,
            "open_net": round(r["net"] - applied_net, 2) if unapplied else 0.0,
            "status": status,
        })

    # Los cupones del mes no coinciden con los de los EFT del mes (72 hs):
    # los EFT del mes pagan cupones del mes anterior y los últimos días del
    # mes pasan a los EFT del siguiente.
    ddc_count = sum(len(r["coupons"]) for r in rows)
    in_month_count = sum(len(r["ddcs_in_month"]) for r in rows)
    in_month_net = round(sum(r["in_month_net"] for r in rows), 2)
    eft_month_net = round(sum(d["coupons_net"] for d in deposits if d["date"].year == year and d["date"].month == month), 2)
    carried = [r for r in rows if len(r["ddcs_in_month"]) < len(r["coupons"])]
    return {
        "rows": rows,
        "totals": report["totals"],
        "ddc_count": ddc_count,
        "applied_in_month": {"count": in_month_count, "amount": in_month_net},
        "carried": {"count": ddc_count - in_month_count, "amount": round(report["totals"]["net"] - in_month_net, 2)},
        "carried_from": carried[0]["date"] if carried else None,
        "eft_month_net": eft_month_net,
        "from_prev_month": round(eft_month_net - in_month_net, 2),
        "bad": [r for r in rows if r["status"] in ("diff", "applied_diff", "not_loaded", "late")],
        "pending": [r for r in rows if r["status"] == "pending"],
        # Cuánto queda sin aplicar, sumado (pedido del usuario, 2026-10-06:
        # "para saberlo a simple vista y no tener que sumarlo al ojo").
        "open": {
            "days": sum(1 for r in rows if r["unapplied"]),
            "gross": round(sum(r["open_gross"] for r in rows), 2),
            "net": round(sum(r["open_net"] for r in rows), 2),
        },
        "ok": all(r["status"] in ("ok", "pending") for r in rows),
    }
