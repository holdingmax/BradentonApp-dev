"""
Reporte mensual de Lottery (pedido del usuario, 2026-10-06, chat 21): el
"Monthly Sales Report" del portal de Florida Lottery, el mismo formato que el
Daily Sales Report de cada día pero con el mes entero. Lectura y cruces, sin
UI: se guarda en lottery_db (tabla lottery_monthly_reports), se sube en
Carga de Datos → Lottery y se controla en /controles/lottery.

- Cruce: cada importe del reporte mensual contra la suma de los reportes
  diarios del mes (lottery_days, por su fecha). Tiene que dar al centavo.
  La comisión de Terminal se compara entera (Comis + Premios FP): el diario
  la parte con la tasa fija del 6 % y el redondeo de cada día puede no
  sumar igual que la parte del mes.
- Pagos: el Debito de cada bloque semanal que se paga en el mes contra el
  pago real a Florida Lottery en Chase (Detalle "LOTTERY").
"""

import calendar
import re
from datetime import date, timedelta

import pdfplumber

import chase_db
import lottery_db
from reporte_diario import extract_lottery_receipt_fields_from_sales_report

_END_DATE_RE = re.compile(r"End Date:\s*(\d{4})-(\d{2})-(\d{2})")
# Un renglón "Concepto valor" del reporte, tal cual lo imprime el portal.
_LINE_RE = re.compile(r"^([A-Za-z][A-Za-z /&-]*?)\s+(-?\$?[\d,]+(?:\.\d+)?)$")
_SECTIONS = ("Online Sales Summary", "Instant Sales Summary")

# (clave, concepto como en el cuadro de Lottery, concepto del reporte,
#  campos diarios que suma, es cantidad)
CROSS_FIELDS = (
    ("sales", "Ventas (Terminal)", "Net Terminal Sales Amount", ("sales",), False),
    ("pagos", "Pagos (Terminal)", "Terminal Pay Amount", ("pagos",), False),
    ("total_comm", "Comis + Premios FP (Terminal)", "Terminal Sales Commission", ("comis", "prize_free_plays"), False),
    ("skoff_sales_amount", "Monto Ventas (SKOFF)", "Instant Sales Amount", ("skoff_sales_amount",), False),
    ("sales_comm", "Comm Ventas (SKOFF)", "Instant Sales Commission", ("sales_comm",), False),
    ("pays_units", "Pagos U (SKOFF)", "Instant Tickets Paid", ("pays_units",), True),
    ("pays_amount", "Pagos $ (SKOFF)", "Instant Pay Amount", ("pays_amount",), False),
)

# Días de tolerancia entre la fecha de pago de un bloque y el débito en Chase.
PAYMENT_WINDOW_DAYS = 3


def extract_monthly_report(pdf_path):
    """
    Lee el Monthly Sales Report. Los importes salen con el mismo lector del
    reporte diario (mismos conceptos y mismos signos que lottery_days); además
    se guardan todos los renglones tal cual, para mostrar el reporte entero.
    ValueError si no es un reporte mensual o no cubre un mes completo.
    """
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    if "Monthly Sales Report" not in text:
        raise ValueError("Uno de los archivos no es un Monthly Sales Report de Lottery (¿es un reporte diario?).")
    fields = extract_lottery_receipt_fields_from_sales_report(pdf_path)
    start = fields["report_date"]
    match = _END_DATE_RE.search(text)
    if not match:
        raise ValueError('No se encontró "End Date" en el Monthly Sales Report de Lottery.')
    end = date(*(int(part) for part in match.groups()))
    month_end = date(start.year, start.month, calendar.monthrange(start.year, start.month)[1])
    if start.day != 1 or end != month_end:
        raise ValueError(
            f"El Monthly Sales Report de Lottery va del {start:%d/%m/%Y} al {end:%d/%m/%Y}: tiene que ser un mes completo."
        )

    lines = []
    for raw in text.splitlines():
        line = raw.strip()
        if line in _SECTIONS:
            lines.append({"section": line})
            continue
        parsed = _LINE_RE.match(line)
        if parsed:
            lines.append({"label": parsed.group(1), "value": parsed.group(2)})

    values = {key: fields.get(key) for key in ("sales", "pagos", "skoff_sales_amount", "sales_comm", "pays_units", "pays_amount")}
    comis, prize = fields.get("comis"), fields.get("prize_free_plays")
    values["total_comm"] = round(comis + prize, 2) if comis is not None and prize is not None else None
    warnings = [
        f'No se pudo leer "{pdf_label}".' for key, _, pdf_label, _, _ in CROSS_FIELDS if values.get(key) is None
    ]
    return {
        "year": start.year,
        "month": start.month,
        "from_date": start.isoformat(),
        "to_date": end.isoformat(),
        "values": values,
        "lines": lines,
        "warnings": warnings,
    }


def _month_days(year, month):
    first = date(year, month, 1)
    return [first + timedelta(days=k) for k in range(calendar.monthrange(year, month)[1])]


def cross_check(report, year, month):
    """
    Cada importe del reporte mensual contra la suma de los reportes diarios
    del mes. `missing_days` son los días sin Daily Sales Report cargado.
    """
    days = _month_days(year, month)
    rows_by_date = lottery_db.get_days_between(days[0], days[-1])
    loaded = [rows_by_date[d.isoformat()] for d in days if rows_by_date.get(d.isoformat(), {}).get("sales") is not None]
    missing = [d.isoformat() for d in days if rows_by_date.get(d.isoformat(), {}).get("sales") is None]
    lines = []
    for key, label, pdf_label, fields, is_count in CROSS_FIELDS:
        values = [sum(row.get(f) or 0.0 for f in fields) for row in loaded if any(row.get(f) is not None for f in fields)]
        days_total = round(sum(values), 2) if values else None
        report_value = report["values"].get(key)
        diff = None if days_total is None or report_value is None else round(days_total - report_value, 2)
        lines.append({
            "key": key, "label": label, "pdf_label": pdf_label, "is_count": is_count,
            "days": days_total, "report": report_value, "diff": diff,
            "ok": diff is not None and abs(diff) < 0.005,
        })
    bad = [line for line in lines if not line["ok"]]
    return {
        "lines": lines,
        "ok": not bad,
        "differences": [line["label"] for line in bad if line["report"] is not None],
        "unread": [line["label"] for line in lines if line["report"] is None],
        "missing_days": missing,
        "days_loaded": len(loaded),
    }


def _chase_lottery_payments_around(year, month):
    """Pagos de Lottery en Chase del mes y de los meses vecinos (un bloque de fin de mes puede debitarse en el otro)."""
    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    next_year, next_month = (year + 1, 1) if month == 12 else (year, month + 1)
    payments = []
    for y, m in ((prev_year, prev_month), (year, month), (next_year, next_month)):
        payments.extend(lottery_db.monthly_chase_lottery_payments(y, m))
    return [{"date": date.fromisoformat(d), "amount": amount} for d, amount in sorted(payments)]


def payments_check(year, month):
    """
    El Debito de cada bloque que se paga en el mes (fecha confirmada o
    sugerida) contra el pago de Lottery en Chase más cercano a esa fecha
    (hasta PAYMENT_WINDOW_DAYS días; si hay varios, el del mismo importe).
    Un bloque al que le faltan días no puede coincidir: se avisa cuántos.
    """
    chase_payments = _chase_lottery_payments_around(year, month)
    last_chase = chase_db.get_last_posting_date()
    used = set()
    rows = []
    for block in lottery_db.blocks_paid_in_month(year, month):
        pay_date = date.fromisoformat(block["chase_bank_date"] or block["suggested_chase_date"])
        debit = block["debito"].get("net_debit")
        debit = round(debit, 2) if debit is not None else None
        missing = [d["date"] for d in block["days"] if d.get("sales") is None]
        candidates = [
            (index, p) for index, p in enumerate(chase_payments)
            if index not in used and abs((p["date"] - pay_date).days) <= PAYMENT_WINDOW_DAYS
        ]
        candidates.sort(key=lambda item: (
            debit is None or abs(item[1]["amount"] - debit) >= 0.005, abs((item[1]["date"] - pay_date).days),
        ))
        chase = None
        if candidates:
            index, chase = candidates[0]
            used.add(index)
        diff = round(debit - chase["amount"], 2) if chase and debit is not None else None
        if chase is None:
            covered = last_chase is not None and last_chase >= pay_date + timedelta(days=PAYMENT_WINDOW_DAYS)
            status = "no_chase" if covered else "chase_not_loaded"
        else:
            status = "ok" if diff is not None and abs(diff) < 0.005 else ("incomplete" if missing else "diff")
        rows.append({
            "start": block["start"], "end": block["end"],
            "pay_date": pay_date.isoformat(), "confirmed": bool(block["chase_bank_date"]),
            "debit": debit, "missing_days": missing, "has_data": block["has_data"],
            "chase_date": chase["date"].isoformat() if chase else None,
            "chase_amount": chase["amount"] if chase else None,
            "diff": diff, "status": status,
        })
    extra = [
        {"date": p["date"].isoformat(), "amount": p["amount"]}
        for index, p in enumerate(chase_payments)
        if index not in used and p["date"].year == year and p["date"].month == month
    ]
    return {
        "rows": rows,
        "extra": extra,
        "ok": all(r["status"] == "ok" for r in rows) and not extra,
        "last_chase": last_chase.isoformat() if last_chase else None,
    }
