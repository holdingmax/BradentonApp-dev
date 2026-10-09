"""
Control Cierre (pedido del usuario, 2026-10-06): los dos asientos de cierre
del mes que arma el Excel Cierre en la hoja Store Info (O47:T64), con los
totales del mes ya guardados en la app. Cálculo, PDF y Excel, sin UI (la
pantalla es /controles/cierre en webapp.py).

Fórmulas del Excel real (fila 33 = totales del mes de la hoja Store Info):
  Caja                                  Debe  = Cash − LOTTO + Other      (S33−L33+U33)
  J.H. WILIAMS-RECAUD A LIQ             Debe  = TC − VS                   (T33−M33)
  Cuenta Corriente a Cobrar-Gettel KIA  Debe  = Local Account             (V33)
    a Venta Combustible                 Haber = Total combustible         (H33)
    a Venta Car Wash / a Venta ICE      Haber = 0
    a Venta C-Store                     Haber = Non Fuel + desc otros + Tax − VS − LOTTO
Debajo, los dos totales sumando LOTTO (el asiento no considera Lottery) y la
diferencia Debe − Haber. Segundo asiento: J.H. Williams a Cuenta Corriente a
Cobrar-Gettel KIA por VS, "por lo cobrado" al último día.

Tercer asiento (pedido del usuario, 2026-10-09): el devengamiento de Sale
Tax del mes, al último día: 5.13.00.00 TAX BGS a 2.02.01.00 SALE TAX A PAGAR
por el Tax Collect del mes (julio 2026: 4,671.10, igual que el Excel Cierre).
Más que un control es un recordatorio: hay meses que se olvidaba registrarlo,
así que el aviso queda hasta que se marca como registrado (sale_tax_record).

LOTTO y VS no son de Store Info: salen de Ventas por Departamento (las
categorías "LOTERY/LOTTO" y "Gettel"). Un día sin departamentos los deja
cortos (y con ellos Caja, J.H. Williams y Venta C-Store); la pantalla y el
PDF lo avisan. La diferencia no depende de LOTTO ni de VS (se cancelan): es
Total Revenue contra Total Ventas + VS, el mismo chequeo que el Excel.
"""

import calendar
import json
import os
from datetime import datetime

import app_paths

ENTRY_TITLE = "Asiento NºZXXX"
ACCOUNT_CAJA = "Caja"
ACCOUNT_JH = "J.H. WILIAMS-RECAUD A LIQ"
ACCOUNT_GETTEL = "Cuenta Corriente a Cobrar-Gettel KIA"
ACCOUNT_TAX_BGS = "5.13.00.00 - TAX BGS"
ACCOUNT_SALE_TAX = "2.02.01.00 - SALE TAX A PAGAR"
LOTTERY_NOTE = "**No se considera la incidencia de Lottery en este asiento"
LOTTO_LABEL = "LOTERY/LOTTO"
VS_LABEL = "Gettel"

GRAY = "D9D9D9"
RED = "FF0000"


def _total(rows, getter):
    return round(sum(getter(r) or 0.0 for r in rows), 2)


def _category(label):
    return lambda r: (r.get("category_amounts") or {}).get(label)


def _ddmm(iso):
    return f"{iso[8:10]}/{iso[5:7]}"


# Campos de Store Info que usan el asiento y el Excel: uno vacío (no
# cargado) cuenta como 0, y el control avisa que falta.
_STORE_INFO_REQUIRED = (
    ("sales_fuel", "Sales Fuel"), ("desc_comb", "Desc. Comb"), ("non_fuel_total", "Non Fuel"),
    ("desc_otros", "Desc. Otros"), ("tax_collect", "Tax Collect"), ("cash", "Cash"),
    ("credit_terms", "Tarjetas"), ("local_accounts", "Local Acc."), ("other_amount", "Other"),
)


def _is_empty(row, key):
    """Un campo de Store Info sin cargar. Las tarjetas se guardan como lista: vacía = sin cargar."""
    return not row.get(key) if key == "credit_terms" else row.get(key) is None


def month_totals(rows):
    """Los totales del mes que usa el asiento, sumando los días."""
    return {
        "total_fuel": _total(rows, lambda r: r.get("total_fuel")),
        "non_fuel": _total(rows, lambda r: r.get("non_fuel_total")),
        "desc_otros": _total(rows, lambda r: r.get("desc_otros")),
        "tax_collect": _total(rows, lambda r: r.get("tax_collect")),
        "cash": _total(rows, lambda r: r.get("cash")),
        "tc": _total(rows, lambda r: sum(v for v in (r.get("credit_terms") or []) if v is not None)),
        "other": _total(rows, lambda r: r.get("other_amount")),
        "local_accounts": _total(rows, lambda r: r.get("local_accounts")),
        "lotto": _total(rows, _category(LOTTO_LABEL)),
        "vs": _total(rows, _category(VS_LABEL)),
    }


def _calc(*terms):
    """Suma con signo; None si falta algún término (un dato que no se pudo leer)."""
    return None if any(t is None for t in terms) else round(sum(terms), 2)


def _neg(value):
    return None if value is None else -value


def entry_lines(t):
    """
    Los renglones del asiento a partir de los totales del mes (`t` como lo
    devuelve month_totals o reporte_mensual.report_totals).
    """
    # `breakdown`: de dónde sale cada importe (el cuadro que se abre al tocarlo).
    return [
        {"side": "debit", "account": ACCOUNT_CAJA, "amount": _calc(t["cash"], _neg(t["lotto"]), t["other"]),
         "uses_departments": True,
         "breakdown": [("Cash", t["cash"]), ("LOTTO (resta)", _neg(t["lotto"])), ("Other", t["other"])]},
        {"side": "debit", "account": ACCOUNT_JH, "amount": _calc(t["tc"], _neg(t["vs"])), "uses_departments": True,
         "breakdown": [("Tarjeta (TC)", t["tc"]), ("VS (Gettel, resta)", _neg(t["vs"]))]},
        {"side": "debit", "account": ACCOUNT_GETTEL, "amount": t["local_accounts"], "red": True,
         "breakdown": [("Local Account", t["local_accounts"])]},
        {"side": "credit", "account": "Venta Combustible", "amount": t["total_fuel"],
         "breakdown": [("Total combustible (Sales Fuel + Desc. Comb)", t["total_fuel"])]},
        {"side": "credit", "account": "Venta Car Wash", "amount": 0.0, "breakdown": None},
        {"side": "credit", "account": "Venta ICE", "amount": 0.0, "breakdown": None},
        {"side": "credit", "account": "Venta C-Store",
         "amount": _calc(t["non_fuel"], t["desc_otros"], t["tax_collect"], _neg(t["vs"]), _neg(t["lotto"])),
         "uses_departments": True,
         "breakdown": [("Non Fuel", t["non_fuel"]), ("Desc. Otros", t["desc_otros"]),
                       ("Tax Collect", t["tax_collect"]), ("VS (Gettel, resta)", _neg(t["vs"])),
                       ("LOTTO (resta)", _neg(t["lotto"]))]},
    ]


def build_month_entries(rows, year, month):
    """
    `rows`: Store Info del mes tal cual lo arma webapp._build_store_info_rows
    (un renglón por día, con `category_amounts` y `has_departments`). None si
    el mes no tiene ningún día de Store Info cargado.
    """
    loaded = [r for r in rows if r.get("store_info_source")]
    if not loaded:
        return None
    last_day = max(r["date"] for r in loaded)
    month_last_day = f"{year:04d}-{month:02d}-{calendar.monthrange(year, month)[1]:02d}"

    t = month_totals(rows)
    lines = entry_lines(t)
    caja = lines[0]["amount"]
    debit_total = round(sum(l["amount"] for l in lines if l["side"] == "debit"), 2)
    credit_total = round(sum(l["amount"] for l in lines if l["side"] == "credit"), 2)
    return {
        "year": year,
        "month": month,
        "first_day": f"{year:04d}-{month:02d}-01",
        "last_day": last_day,
        "complete_month": last_day == month_last_day,
        "totals": t,
        "lines": lines,
        "debit_total": debit_total,
        "credit_total": credit_total,
        # Fila de abajo del Excel (S57/T57): los mismos totales sumando LOTTO.
        "debit_with_lottery": round(caja + t["lotto"], 2),
        "credit_with_lottery": round(credit_total + t["lotto"], 2),
        "difference": round(debit_total - credit_total, 2),
        "collected": t["vs"],
        "collected_label": f"por lo cobrado al {_ddmm(last_day)}",
        "sale_tax": {"date": month_last_day, "amount": t["tax_collect"],
                     "empty_days": [r["date"] for r in loaded if _is_empty(r, "tax_collect")]},
        "missing_store_info": [r["date"] for r in rows if not r.get("store_info_source") and r["date"] <= last_day],
        "missing_departments": [r["date"] for r in loaded if not r.get("has_departments")],
        "empty_fields": [
            (r["date"], [label for key, label in _STORE_INFO_REQUIRED if _is_empty(r, key)])
            for r in loaded if any(_is_empty(r, key) for key, _label in _STORE_INFO_REQUIRED)
        ],
    }


# Qué meses ya tienen registrado el devengamiento de Sale Tax:
# {"2026-07": {"asiento": "11457", "marked_at": "2026-10-09 10:15"}}.
_SALE_TAX_PATH = os.path.join(app_paths.DATA_DIR, "control_cierre_sale_tax.json")


def _load_sale_tax_records():
    try:
        with open(_SALE_TAX_PATH, encoding="utf-8") as handle:
            data = json.load(handle)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def sale_tax_record(year, month):
    """El registro del devengamiento del mes, o None si todavía no se marcó."""
    return _load_sale_tax_records().get(f"{year:04d}-{month:02d}")


def set_sale_tax_record(year, month, registered, asiento=""):
    """Marca (o desmarca) el devengamiento de Sale Tax del mes como registrado."""
    records = _load_sale_tax_records()
    key = f"{year:04d}-{month:02d}"
    if registered:
        records[key] = {"asiento": (asiento or "").strip()[:30],
                        "marked_at": datetime.now().strftime("%Y-%m-%d %H:%M")}
    else:
        records.pop(key, None)
    os.makedirs(os.path.dirname(_SALE_TAX_PATH), exist_ok=True)
    temp_path = f"{_SALE_TAX_PATH}.tmp"
    with open(temp_path, "w", encoding="utf-8") as handle:
        json.dump(records, handle, indent=2, ensure_ascii=False)
    os.replace(temp_path, _SALE_TAX_PATH)


def missing_notice(entries):
    """Aviso de datos que faltan (pantalla y PDF), o None si el mes está completo."""
    parts = []
    if entries["missing_departments"]:
        days = entries["missing_departments"]
        parts.append(
            f"Faltan las Ventas por Departamento de {len(days)} día{'s' if len(days) != 1 else ''} "
            f"({', '.join(_ddmm(d) for d in days)}): sin ellas LOTTO y VS quedan cortos, y con ellos "
            "Caja, J.H. Williams, Venta C-Store y lo cobrado a Gettel KIA."
        )
    if entries["missing_store_info"]:
        days = entries["missing_store_info"]
        parts.append(f"Sin Store Info: {', '.join(_ddmm(d) for d in days)}.")
    if entries.get("empty_fields"):
        days = "; ".join(f"{_ddmm(d)} ({', '.join(labels)})" for d, labels in entries["empty_fields"])
        parts.append(f"Store Info con datos vacíos (no cargados): {days}. Completalos en el día: "
                     "mientras tanto cuentan como 0 (también en el Excel).")
    return " ".join(parts) or None


# Datos de base del asiento, en el orden en que se muestran en el cruce.
BASE_LABELS = (
    ("cash", "Cash"), ("tc", "Tarjeta (TC)"), ("other", "Other"), ("local_accounts", "Local Account"),
    ("total_fuel", "Total combustible"), ("non_fuel", "Non Fuel"), ("desc_otros", "Desc. Otros"),
    ("tax_collect", "Tax Collect"), ("lotto", "LOTTO (ONLINE + SKOFF)"), ("vs", "VS (LOCAL ACCT)"),
)


def cross_check(entries, report_t):
    """
    Cruce del asiento armado con los días contra el mismo asiento armado con
    el reporte mensual (`report_t` de reporte_mensual.report_totals): renglón
    por renglón y dato por dato. `ok` solo si todo coincide al centavo.
    """
    def row(label, days, report):
        diff = None if days is None or report is None else round(days - report, 2)
        return {"label": label, "days": days, "report": report, "diff": diff,
                "ok": diff is not None and abs(diff) < 0.005}

    lines = [dict(row(line["account"], line["amount"], other["amount"]), side=line["side"])
             for line, other in zip(entries["lines"], entry_lines(report_t))]
    lines.append(dict(row("Cobrado a Gettel KIA", entries["collected"], report_t["vs"]), side="debit"))
    bases = [row(label, entries["totals"][key], report_t[key]) for key, label in BASE_LABELS]
    bad = [r for r in lines + bases if not r["ok"]]
    return {
        "lines": lines,
        "bases": bases,
        "ok": not bad,
        "differences": [r["label"] for r in bad if r["report"] is not None],
        "unread": [r["label"] for r in bases if r["report"] is None],
    }


def jh_cards_check(entries, detail, coupons):
    """
    Las tarjetas del asiento (TC de Store Info, de donde sale J.H. Williams)
    contra lo que informa J.H. (pedido del usuario, 2026-10-06: "que el
    control de las tarjetas se hace en contra de lo que sale de la página de
    J.H."). Dos fuentes de J.H., las mismas de Controles → Tarjetas y Cupones:

    - `detail` (control_tarjetas.build_detail_by_day): los batches del POS
      del detalle de cupones del portal, por día de venta, contra lo vendido
      con tarjeta ese día. Lo que el detalle todavía no cubre queda "sin
      controlar" (no es una diferencia).
    - `coupons` (jh_mensual.coupon_check): el Credit Card Daily Summary del
      mes contra los cupones cargados y los EFT que los aplicaron.
    """
    tc = entries["totals"]["tc"]
    result = {"tc": tc, "detail": None, "coupons": None}
    if detail:
        days = [r["date"] for r in detail["rows"] if r["sold"] is not None]
        result["detail"] = {
            "sold": detail["sold_total"],
            "jh": detail["pos_total"],
            "diff": detail["diff_total"],
            "unchecked": round(tc - detail["sold_total"], 2),
            "from": min(days, default=None),
            "until": max(days, default=None),
            "bad": [r["date"] for r in detail["bad"]],
            "pending": [r["date"] for r in detail["pending"]],
            "edge": [r["date"] for r in detail["edge"]],
            # De qué se compone la diferencia: lo que es error de verdad y lo
            # que todavía no se depositó (los últimos días, 72 hs).
            "diff_bad": round(sum(r["diff"] for r in detail["bad"]), 2),
            "diff_pending": round(sum(r["diff"] for r in detail["pending"]), 2),
            "diff_edge": round(sum(r["diff"] for r in detail["edge"]), 2),
            # Lo sin depositar que son pagos de Kia/Toyota (suben las tarjetas del día).
            "kia_pending": [(r["date"], r["kia"]) for r in detail["pending"] if r.get("kia")],
            "ok": detail["ok"],
        }
    if coupons:
        result["coupons"] = {
            "gross": coupons["totals"]["gross"],
            "net": coupons["totals"]["net"],
            "ddc_count": coupons["ddc_count"],
            "bad": [r["date"] for r in coupons["bad"]],
            "open": coupons["open"],
            "ok": coupons["ok"],
        }
    parts = [part for part in (result["detail"], result["coupons"]) if part]
    # Rojo solo para lo real (pedido del usuario, 2026-10-07): los días que
    # todavía no se depositaron o que no se pueden comparar no son diferencia.
    real_bad = bool(result["detail"] and result["detail"]["bad"]) or bool(result["coupons"] and not result["coupons"]["ok"])
    # Revisión 2026-10-08: "ok" del detalle no mira lo pendiente ni lo que el
    # detalle no cubre, así que antes esos meses decían "Bien" igual.
    detail_part = result["detail"]
    if not parts:
        result["status"] = "missing"
    elif real_bad:
        result["status"] = "bad"
    elif detail_part and (detail_part["pending"] or detail_part["edge"]):
        result["status"] = "pending"
    elif detail_part and abs(detail_part["unchecked"]) >= 0.005:
        result["status"] = "unchecked"
    elif all(part["ok"] for part in parts):
        result["status"] = "ok" if len(parts) == 2 else "partial"
    else:
        result["status"] = "bad"
    return result


def _money(value):
    if value is None:
        return "—"
    return f"-${-value:,.2f}" if value < 0 else f"${value:,.2f}"


def build_entries_pdf(entries, title, period_label, dest_path, cross=None):
    """PDF de los dos asientos, con el membrete de la empresa y, si está, el cruce con el reporte mensual."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    from pdf_export import _letterhead_table

    gray = colors.HexColor("#" + GRAY)
    red = colors.HexColor("#" + RED)
    doc = SimpleDocTemplate(
        dest_path, pagesize=letter, leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=14 * mm, bottomMargin=14 * mm, title=title,
    )
    styles = getSampleStyleSheet()
    note_style = ParagraphStyle("CierreNote", parent=styles["Normal"], fontName="Helvetica-BoldOblique", fontSize=9.5)
    warn_style = ParagraphStyle("CierreWarn", parent=styles["Normal"], textColor=colors.HexColor("#B45309"), fontSize=9)
    col_widths = [10 * mm, 100 * mm, 34 * mm, 34 * mm]

    def entry_table(data, gray_rows, red_cells, bold_rows=(), plain_rows=()):
        table = Table(data, colWidths=col_widths)
        style = [
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica-BoldOblique"),
            ("FONTNAME", (2, 0), (3, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9.5),
            ("ALIGN", (0, 0), (0, -1), "RIGHT"),
            ("ALIGN", (2, 0), (3, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]
        for r in gray_rows:
            style += [("BACKGROUND", (0, r), (-1, r), gray), ("GRID", (0, r), (-1, r), 0.6, colors.black)]
        for r in plain_rows:
            style += [("FONTNAME", (0, r), (-1, r), "Helvetica")]
        for r in bold_rows:
            style += [("FONTNAME", (2, r), (3, r), "Helvetica-Bold")]
        for (c0, r0, c1, r1) in red_cells:
            style += [("TEXTCOLOR", (c0, r0), (c1, r1), red)]
        for r, row in enumerate(data):
            if row[0] and row[1] == "":
                style += [("SPAN", (0, r), (1, r)), ("ALIGN", (0, r), (1, r), "LEFT")]
        table.setStyle(TableStyle(style))
        return table

    elements = list(_letterhead_table(doc.width))
    elements.append(Paragraph(title, styles["Heading2"]))
    elements.append(Paragraph(period_label, styles["Normal"]))
    elements.append(Spacer(1, 10))
    notice = missing_notice(entries)
    if notice:
        elements.append(Paragraph("Incompleto. " + notice, warn_style))
        elements.append(Spacer(1, 8))

    # Asiento de ventas: título, renglones, un renglón vacío y los totales.
    data = [[ENTRY_TITLE, "", "Debe", "Haber"]]
    red_cells = [(0, 0, 1, 0)]
    for line in entries["lines"]:
        if line["side"] == "debit":
            data.append([line["account"], "", _money(line["amount"]), ""])
        else:
            data.append(["a", line["account"], "", _money(line["amount"])])
        if line.get("red"):
            red_cells.append((0, len(data) - 1, 3, len(data) - 1))
    data.append(["", "", "", ""])
    data.append(["", "", _money(entries["debit_total"]), _money(entries["credit_total"])])
    totals_row = len(data) - 1
    data.append(["Sumando Lottery", "", _money(entries["debit_with_lottery"]), _money(entries["credit_with_lottery"])])
    lottery_row = len(data) - 1
    elements.append(entry_table(
        data, gray_rows=range(1, totals_row + 1), red_cells=red_cells,
        bold_rows=(totals_row,), plain_rows=(lottery_row,),
    ))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph(LOTTERY_NOTE, note_style))
    diff_style = ParagraphStyle("CierreDiff", parent=styles["Normal"], fontName="Helvetica-Bold", textColor=red)
    elements.append(Paragraph(f"Diferencia Debe − Haber: {_money(entries['difference'])}", diff_style))
    elements.append(Spacer(1, 18))

    # Asiento de lo cobrado a Gettel KIA.
    data = [
        [ACCOUNT_JH, "", _money(entries["collected"]), ""],
        ["a", ACCOUNT_GETTEL, "", _money(entries["collected"])],
        [entries["collected_label"], "", "", ""],
    ]
    elements.append(entry_table(data, gray_rows=range(0, 3), red_cells=[(1, 1, 1, 1)]))

    # Cruce con el reporte mensual del POS, si está cargado.
    if cross:
        elements.append(PageBreak())
        elements.append(Paragraph(f"Cruce con el reporte mensual — {period_label}", styles["Heading3"]))
        if cross["ok"]:
            verdict = "La suma de los reportes diarios del mes coincide al centavo con el reporte mensual."
        else:
            verdict = "No coincide: " + ", ".join(cross["differences"] + [f"{u} (no se leyó)" for u in cross["unread"]]) + "."
        result_style = ParagraphStyle(
            "CierreCross", parent=styles["Normal"], fontName="Helvetica-Bold",
            textColor=colors.HexColor("#166534" if cross["ok"] else "#B91C1C"),
        )
        elements.append(Paragraph(verdict, result_style))
        elements.append(Spacer(1, 6))
        data = [["Concepto", "Reportes por Día", "Reporte mensual", "Diferencia"]]
        rows = cross["lines"] + [None] + cross["bases"]
        bad_rows, section_row = [], None
        for row in rows:
            if row is None:
                data.append(["Datos de los que sale cada importe", "", "", ""])
                section_row = len(data) - 1
                continue
            data.append([row["label"], _money(row["days"]), _money(row["report"]), _money(row["diff"])])
            if not row["ok"]:
                bad_rows.append(len(data) - 1)
        table = Table(data, colWidths=[78 * mm, 33 * mm, 33 * mm, 34 * mm], repeatRows=1)
        style = [
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("BACKGROUND", (0, 0), (-1, 0), gray),
            ("SPAN", (0, section_row), (-1, section_row)),
            ("FONTNAME", (0, section_row), (-1, section_row), "Helvetica-Bold"),
            ("BACKGROUND", (0, section_row), (-1, section_row), gray),
        ]
        for r in bad_rows:
            style.append(("TEXTCOLOR", (3, r), (3, r), red))
        table.setStyle(TableStyle(style))
        elements.append(table)

    doc.build(elements)
    return dest_path


def add_entries_to_store_info_workbook(path, entries):
    """
    Al Excel de Store Info ya armado (reporte_diario.build_store_info_export_
    workbook) le suma la fila de totales del mes ("Ventas", como la fila 33
    del Excel real) y, debajo, los dos asientos en las columnas O-T, con las
    mismas fórmulas y el mismo formato que la hoja real.
    """
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.load_workbook(path)
    ws = wb["Store Info"]
    last = ws.max_row
    t = last + 1
    thin = Side(style="thin", color="FF000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    money = '"$"\\ #,##0.00'
    gray = PatternFill("solid", fgColor="FF" + GRAY)

    # Fila de totales del mes, como la fila 33 del Excel Cierre.
    ws.cell(t, 3, "Ventas").font = Font(bold=True)
    for col in range(5, 25):
        letter = get_column_letter(col)
        if col == 18:
            formula = f"=+H{t}+O{t}+P{t}+Q{t}-M{t}"
        elif col == 23:
            formula = f"=SUM(S{t}:V{t})"
        else:
            formula = f"=SUM({letter}2:{letter}{last})"
        cell = ws.cell(t, col, formula)
        cell.font = Font(bold=True)
        cell.border = border
        cell.number_format = ws.cell(last, col).number_format

    def put(row, col, value, bold=True, italic=True, red=False, fmt=None, align=None, fill=True, sides=True, size=11):
        cell = ws[f"{col}{row}"]
        cell.value = value
        cell.font = Font(bold=bold, italic=italic, color=("FF" + RED) if red else None, size=size)
        if fmt:
            cell.number_format = fmt
        if align:
            cell.alignment = Alignment(horizontal=align)
        if fill:
            cell.fill = gray
        if sides:
            cell.border = border
        return cell

    def box_row(row):
        for col in "OPQRST":
            put(row, col, None)

    a = t + 3
    put(a, "O", ENTRY_TITLE, red=True, fill=False, sides=False).font = Font(bold=True, italic=True, underline="single", color="FF" + RED)
    for r in range(a + 1, a + 10):
        box_row(r)
    put(a + 1, "O", ACCOUNT_CAJA)
    put(a + 1, "S", f"=+S{t}-L{t}+U{t}", italic=False, fmt=money)
    put(a + 2, "O", ACCOUNT_JH)
    put(a + 2, "S", f"=+T{t}-M{t}", italic=False, fmt=money)
    put(a + 3, "O", ACCOUNT_GETTEL, red=True, align="left")
    put(a + 3, "S", f"=+V{t}", italic=False, red=True, fmt=money)
    for offset, (account, formula) in enumerate((
        ("Venta Combustible", f"=+H{t}"),
        ("Venta Car Wash", 0),
        ("Venta ICE", 0),
        ("Venta C-Store", f"=+O{t}+P{t}+Q{t}-M{t}-L{t}"),
    ), start=4):
        put(a + offset, "O", "a", align="right")
        put(a + offset, "P", account)
        put(a + offset, "T", formula, italic=False, fmt=money)
    put(a + 9, "S", f"=SUM(S{a + 1}:S{a + 7})", italic=False, fmt=money)
    put(a + 9, "T", f"=SUM(T{a + 1}:T{a + 8})", italic=False, fmt=money)
    put(a + 10, "S", f"=+S{a + 1}+L{t}", bold=False, italic=False, fmt="#,##0.00", fill=False, sides=False)
    put(a + 10, "T", f"=+T{a + 9}+L{t}", bold=False, italic=False, fmt=money, fill=False, sides=False)
    put(a + 11, "O", LOTTERY_NOTE, fill=False, sides=False, size=12)
    diff = put(a + 12, "O", f"=+S{a + 9}-T{a + 9}", italic=False, red=True, fmt="#,##0.00", fill=False, sides=False)
    diff.border = Border(left=thin, right=thin, bottom=thin)

    b = a + 15
    for r in range(b, b + 3):
        box_row(r)
    put(b, "O", ACCOUNT_JH)
    put(b, "S", f"=+M{t}", italic=False, fmt=money)
    put(b + 1, "O", "a", align="right")
    put(b + 1, "P", ACCOUNT_GETTEL, red=True, align="left")
    put(b + 1, "T", f"=+S{b}", italic=False, fmt=money)
    put(b + 2, "O", entries["collected_label"], align="left")

    wb.save(path)
    return path
