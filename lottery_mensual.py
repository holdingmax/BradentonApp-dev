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
from datetime import date, datetime, timedelta

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


# Cada importe del cruce sale de un renglón del reporte, con el signo que usa
# lottery_days (lo que entra a la caja positivo, pagos y comisiones negativos).
_VALUE_SOURCES = (
    ("sales", "Net Terminal Sales Amount", 1),
    ("pagos", "Terminal Pay Amount", -1),
    ("total_comm", "Terminal Sales Commission", -1),
    ("skoff_sales_amount", "Instant Sales Amount", 1),
    ("sales_comm", "Instant Sales Commission", -1),
    ("pays_units", "Instant Tickets Paid", 1),
    ("pays_amount", "Instant Pay Amount", -1),
)
# Cuentas que el propio reporte tiene que cerrar: (resultado, [(renglón, signo)]).
_INTERNAL_CHECKS = (
    ("Net Terminal Sales Amount", (("Terminal Sales Amount", 1), ("Terminal Cancel Amount", -1))),
    ("Net Terminal Sales", (("Terminal Tickets", 1), ("Terminal Cancels", -1))),
    ("Instant Sales Amount", (("Instant Books Amount", 1), ("Instant Full Returns Amount", -1),
                              ("Instant Partial Returns Amount", -1))),
    ("Instant Pay Amount", (("Instant Low Tier Pay Amount", 1), ("Instant Mid Tier Pay Amount", 1))),
    ("Instant Tickets Paid", (("Instant Low Tier Tickets Paid", 1), ("Instant Mid Tier Tickets Paid", 1))),
)
# El Excel del portal nombra un par de renglones distinto que el PDF.
_EXCEL_LABELS = {"Terminal Cancels Amount": "Terminal Cancel Amount", "Terminal Cash Commission": "Terminal Cash Commission Amount"}


def _number(value):
    """'$8,907.50' / '5937' / '$ 6,490.00' / 8907.5 -> float; None si no es un número."""
    if isinstance(value, (int, float)):
        return float(value)
    text = re.sub(r"[$,\s]", "", str(value or ""))
    try:
        return float(text)
    except ValueError:
        return None


def _parse_date(text):
    """'2026-03-01' / '03/01/2026' / '08-01-2024' (o una fecha de Excel) -> date, o None."""
    if isinstance(text, datetime):
        return text.date()
    if isinstance(text, date):
        return text
    text = str(text or "").strip()
    match = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", text)
    if match:
        return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    match = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", text)
    if match:
        return date(int(match.group(3)), int(match.group(1)), int(match.group(2)))
    return None


def _pdf_report(path):
    """(inicio, fin, renglones) del PDF; además controla lo leído contra el lector del reporte diario."""
    with pdfplumber.open(path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    if "Monthly Sales Report" not in text:
        raise ValueError("Uno de los archivos no es el Monthly Sales Report (o Summary) de Lottery: "
                         "puede ser un reporte diario u otro reporte del portal.")
    lines = []
    for raw in text.splitlines():
        line = raw.strip()
        if line in _SECTIONS:
            lines.append({"section": line})
            continue
        parsed = _LINE_RE.match(line)
        if parsed:
            lines.append({"label": parsed.group(1), "value": parsed.group(2)})
    match = _END_DATE_RE.search(text)
    if not match:
        raise ValueError('No se encontró "End Date" en el Monthly Sales Report de Lottery.')
    end = date(*(int(part) for part in match.groups()))
    start = extract_lottery_receipt_fields_from_sales_report(path)["report_date"]
    return start, end, lines


def _excel_rows(path):
    """Filas (listas de valores) de la primera hoja del Excel (.xlsx con openpyxl, .xls con pandas)."""
    if path.lower().endswith(".xls"):
        import pandas as pd
        frame = pd.read_excel(path, header=None, dtype=object)
        return [[None if pd.isna(v) else v for v in row] for row in frame.itertuples(index=False)]
    import openpyxl
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return [list(row) for row in workbook.worksheets[0].iter_rows(values_only=True)]
    finally:
        workbook.close()


def _excel_report(path):
    """
    (inicio, fin, renglones) del Monthly Sales Summary en Excel (lo que baja el
    portal; de 2024-2025 solo existe así): fechas en sus filas "Starting Date"
    / "Ending Date", títulos de columna en la fila que empieza con "Terminal
    Tickets" y los valores en la fila de abajo.
    """
    rows = _excel_rows(path)
    found = {}
    header_at = None
    for index, row in enumerate(rows):
        cells = [c for c in row if c is not None and str(c).strip() != ""]
        if not cells:
            continue
        first = str(cells[0]).strip()
        if first in ("Starting Date", "Ending Date") and len(cells) > 1:
            found[first] = _parse_date(cells[1])
        if first == "Terminal Tickets":
            header_at = index
            break
    if "Monthly Sales Summary" not in " ".join(str(c) for row in rows[:3] for c in row if c is not None) or header_at is None:
        raise ValueError("Uno de los archivos no es el Monthly Sales Summary de Lottery.")
    if not found.get("Starting Date") or not found.get("Ending Date"):
        raise ValueError('No se encontraron "Starting Date" y "Ending Date" en el Monthly Sales Summary de Lottery.')
    header = rows[header_at]
    values = rows[header_at + 1] if header_at + 1 < len(rows) else []
    lines, section = [], None
    for column, title in enumerate(header):
        if title is None or str(title).strip() == "":
            continue
        label = _EXCEL_LABELS.get(str(title).strip(), str(title).strip())
        wanted = "Online Sales Summary" if label.startswith("Terminal") else (
            "Instant Sales Summary" if label.startswith("Instant") else None)
        if wanted != section and wanted is not None:
            lines.append({"section": wanted})
        section = wanted or section
        number = _number(values[column]) if column < len(values) else None
        if number is None:
            continue
        money = re.search(r"Amount|Commission", label) is not None
        lines.append({"label": label, "value": f"${number:,.2f}" if money else f"{int(round(number))}"})
    return found["Starting Date"], found["Ending Date"], lines


def extract_monthly_report(path):
    """
    Lee el Monthly Sales Report de Lottery, en PDF o en Excel (Monthly Sales
    Summary, como lo baja el portal). Los importes del cruce salen de los
    renglones del reporte con los signos de lottery_days; en el PDF además
    tienen que coincidir con el lector del reporte diario (dos lecturas
    independientes). Las cuentas del propio reporte (ventas menos bajas,
    libros menos devoluciones, premios Low + Mid Tier) se controlan y lo que
    no cierra va a `warnings`. ValueError si no es un reporte mensual o no
    cubre un mes completo.
    """
    if path.lower().endswith((".xlsx", ".xls")):
        start, end, lines = _excel_report(path)
        daily_fields = None
    else:
        start, end, lines = _pdf_report(path)
        daily_fields = extract_lottery_receipt_fields_from_sales_report(path)
    month_end = date(start.year, start.month, calendar.monthrange(start.year, start.month)[1])
    if start.day != 1 or end != month_end:
        raise ValueError(
            f"El Monthly Sales Report de Lottery va del {start:%d/%m/%Y} al {end:%d/%m/%Y}: tiene que ser un mes completo."
        )

    by_label = {}
    for line in lines:
        if "label" in line:
            by_label.setdefault(line["label"].strip(), _number(line["value"]))
    values = {}
    for key, label, sign in _VALUE_SOURCES:
        number = by_label.get(label)
        values[key] = None if number is None else (int(number) if key == "pays_units" else round(sign * number, 2))
    warnings = [
        f'No se pudo leer "{pdf_label}".' for key, _, pdf_label, _, _ in CROSS_FIELDS if values.get(key) is None
    ]

    if daily_fields is not None:
        comis, prize = daily_fields.get("comis"), daily_fields.get("prize_free_plays")
        other = {key: daily_fields.get(key) for key in ("sales", "pagos", "skoff_sales_amount", "sales_comm",
                                                         "pays_units", "pays_amount")}
        other["total_comm"] = round(comis + prize, 2) if comis is not None and prize is not None else None
        for key, label, _ in _VALUE_SOURCES:
            if values.get(key) is not None and other.get(key) is not None and abs(values[key] - other[key]) >= 0.005:
                raise ValueError(f'"{label}" del Monthly Sales Report de Lottery se leyó de dos formas distintas '
                                 f"({values[key]:,.2f} y {other[key]:,.2f}): revisalo a mano.")

    for result, parts in _INTERNAL_CHECKS:
        if by_label.get(result) is None or any(by_label.get(label) is None for label, _ in parts):
            continue
        total = round(sum(sign * by_label[label] for label, sign in parts), 2)
        if abs(total - by_label[result]) >= 0.005:
            warnings.append(f'El reporte no cierra: "{result}" dice {by_label[result]:,.2f} y la cuenta da {total:,.2f}.')

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


def _chase_month_loaded(year, month):
    return bool(chase_db.get_month_transactions(year, month))


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
            # Cubierto = Chase cargado hasta después del pago y con movimientos
            # en el mes del pago (revisión 2026-10-08: con meses de Chase sin
            # cargar en el medio salía "sin pago" en rojo).
            covered = (last_chase is not None and last_chase >= pay_date + timedelta(days=PAYMENT_WINDOW_DAYS)
                       and _chase_month_loaded(pay_date.year, pay_date.month))
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
