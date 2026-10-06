"""
Reporte Mensual del POS (pedido del usuario, 2026-10-06): el "Resumen Ventas"
del mes, escaneado, con los mismos reportes que el cierre diario pero del mes
entero: Store Sales Summary (págs. 1-2), Department Sales (pág. 3), Method
of Payment y el ticket de inventario de los tanques (estos dos no se usan).
Se lee con los mismos lectores del Reporte Diario (reporte_diario.py), y se
cruza contra lo cargado día por día: Store Info, Ventas por Departamento y
el asiento de cierre (control_cierre.py). Lectura y cruce, sin UI; la base
es reporte_mensual_db.py.
"""

import calendar
import re

from reporte_diario import extract_store_info_from_pdf, group_department_sales, parse_elistar_daily_pdf_page

TOLERANCE = 0.005
# Los diarios imprimen el Volume redondeado a 2 decimales y el mensual suma
# los galones con 3: la suma de los días puede correrse medio centavo por día.
VOLUME_ROUNDING_PER_DAY = 0.005

STORE_INFO_LABELS = (
    ("volume", "Volume"),
    ("sales_fuel", "Sales Fuel"),
    ("desc_comb", "Desc. Comb"),
    ("total_fuel", "Total Fuel"),
    ("non_fuel_total", "Non Fuel"),
    ("desc_otros", "Desc. Otros"),
    ("tax_collect", "Tax Collect"),
    ("total_sales", "Total Sales"),
    ("cash", "Cash"),
    ("tc", "Tarjeta/Créd."),
    ("local_accounts", "Local Acc."),
    ("other_amount", "Other"),
    ("network_revenue", "Network Rev."),
    ("total_revenue", "Total Rev."),
)


def _ddmmyyyy(value):
    return value.strftime("%d/%m/%Y")


def _sum(*values):
    return None if any(v is None for v in values) else round(sum(values), 2)


def extract_monthly_report(pdf_path):
    """
    Lee el reporte mensual. Falla (ValueError) solo si no es un reporte de un
    mes completo; lo que no se pudo leer con seguridad queda en None y va a
    `warnings`, nunca se adivina.
    """
    fields = extract_store_info_from_pdf(pdf_path, start_page_index=0, strict=False)
    from_date, to_date = fields["from_date"], fields["to_date"]
    last_day = calendar.monthrange(from_date.year, from_date.month)[1]
    if from_date.day != 1 or (to_date.year, to_date.month, to_date.day) != (from_date.year, from_date.month, last_day):
        raise ValueError(
            f"No es un reporte mensual: el PDF cubre del {_ddmmyyyy(from_date)} al {_ddmmyyyy(to_date)}, "
            "no un mes completo."
        )

    warnings = []
    if fields.get("missing_fields"):
        warnings.append("Store Info: no se pudo leer " + ", ".join(fields["missing_fields"]) + ".")
    if fields.get("total_sales_mismatch"):
        warnings.append("Store Info: lo leído no cierra contra el Total Sales impreso.")

    # Department Sales arranca en la página que sigue a Store Info.
    departments, printed = None, None
    try:
        records, diagnostics = parse_elistar_daily_pdf_page(pdf_path, page_index=max(fields["pages_used"]))
        period = diagnostics.get("period")
        if period and (period["from_date"], period["to_date"]) != (from_date, to_date):
            raise ValueError("el reporte de departamentos es de otro período")
        for record in records:
            # Mismo nombre real que guarda el Reporte Diario (ver extract_department_sales_for_day).
            if record.get("department") == "GETTEL/TOYOTA":
                record["department"] = "LOCAL ACCT"
        departments = [{"department": r["department"], "count": r["count"], "amount": r["amount"]} for r in records]
        printed = diagnostics.get("printed_totals")
        if diagnostics.get("doubtful_departments"):
            warnings.append("Departamentos dudosos (quedaron vacíos): " + ", ".join(diagnostics["doubtful_departments"]) + ".")
        if diagnostics.get("total_unverified"):
            warnings.append("Departamentos: no se pudo controlar la suma contra el total impreso.")
        elif diagnostics.get("subtotal_mismatch"):
            warnings.append("Departamentos: la suma no cierra contra el total impreso.")
    except Exception as exc:
        warnings.append(f"No se pudieron leer los departamentos ({exc}).")

    store_info = {key: fields.get(key) for key in (
        "volume", "sales_fuel", "desc_comb", "non_fuel_total", "desc_otros", "tax_collect", "total_sales",
        "cash", "local_accounts", "other_amount", "network_revenue", "total_revenue", "credit_terms",
    )}
    return {
        "year": from_date.year,
        "month": from_date.month,
        "from_date": from_date.isoformat(),
        "to_date": to_date.isoformat(),
        "store_info": store_info,
        "departments": departments,
        "printed_department_total": printed,
        "warnings": warnings,
    }


def _report_groups(report):
    """{categoría: monto} de los departamentos del reporte, o None si no se leyeron."""
    if not report.get("departments"):
        return None
    groups, _ = group_department_sales(report["departments"])
    return {g["label"]: g["amount"] for g in groups}


def report_totals(report):
    """
    Los mismos totales que control_cierre.month_totals arma con los días,
    sacados del reporte mensual (None lo que no se pudo leer).
    """
    info = report["store_info"]
    groups = _report_groups(report) or {}
    credit_terms = info.get("credit_terms")
    return {
        "total_fuel": _sum(info.get("sales_fuel"), info.get("desc_comb")),
        "non_fuel": info.get("non_fuel_total"),
        "desc_otros": info.get("desc_otros"),
        "tax_collect": info.get("tax_collect"),
        "cash": info.get("cash"),
        "tc": round(sum(credit_terms), 2) if credit_terms else None,
        "other": info.get("other_amount"),
        "local_accounts": info.get("local_accounts"),
        "lotto": groups.get("LOTERY/LOTTO"),
        "vs": groups.get("Gettel"),
    }


def _row(days, report, tolerance=TOLERANCE):
    diff = None if days is None or report is None else round(days - report, 2)
    return {"days": days, "report": report, "diff": diff,
            "ok": diff is not None and abs(diff) <= tolerance,
            "rounding": diff is not None and TOLERANCE < abs(diff) <= tolerance}


def store_info_comparison(days_totals, report, days_loaded):
    """
    Fila "Reporte mensual" + "Diferencia" de la página de Store Info:
    {campo: {days, report, diff, ok, rounding}} con los mismos campos que
    webapp._store_info_totals. Total Sales se compara sin VS, como lo muestra
    la página (la fórmula del Excel Cierre).
    """
    info = report["store_info"]
    vs = (_report_groups(report) or {}).get("Gettel")
    credit_terms = info.get("credit_terms")
    values = {key: info.get(key) for key, _label in STORE_INFO_LABELS}
    values["total_fuel"] = _sum(info.get("sales_fuel"), info.get("desc_comb"))
    values["tc"] = round(sum(credit_terms), 2) if credit_terms else None
    values["total_sales"] = _sum(info.get("total_sales"), -vs if vs is not None else None)
    result = {}
    for key, label in STORE_INFO_LABELS:
        tolerance = max(TOLERANCE, VOLUME_ROUNDING_PER_DAY * days_loaded) if key == "volume" else TOLERANCE
        result[key] = dict(_row(days_totals.get(key), values[key], tolerance), label=label)
    return result


def category_comparison(days_groups, report):
    """Las 6 categorías de Ventas por Departamento: suma de los días contra el reporte."""
    report_groups = _report_groups(report)
    if report_groups is None:
        return None
    return {g["label"]: _row(round(g["amount"], 2), report_groups.get(g["label"], 0.0)) for g in days_groups}


def department_comparison(days_departments, report):
    """
    Cada departamento: suma de los días contra el reporte. Un departamento
    que está en uno solo cuenta como 0 en el otro.
    """
    if not report.get("departments"):
        return None
    # Mismo departamento aunque el OCR se coma un signo ("HOT DOGS & SANDWICH").
    def key(name):
        return re.sub(r"[^A-Z0-9]", "", (name or "").upper())

    days = {key(d["department"]): d for d in days_departments}
    month = {key(d["department"]): d for d in report["departments"]}
    keys = sorted(set(days) | set(month), key=lambda k: -max((days.get(k) or {}).get("amount") or 0,
                                                             (month.get(k) or {}).get("amount") or 0))
    rows = []
    for k in keys:
        d, m = days.get(k) or {}, month.get(k) or {}
        row = _row(round(d.get("amount") or 0.0, 2), m.get("amount") if m else 0.0)
        row.update(department=(d or m)["department"], days_count=d.get("count"), report_count=m.get("count"),
                   count_ok=(d.get("count") or 0) == (m.get("count") or 0))
        rows.append(row)
    return rows
