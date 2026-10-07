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


# Lo que no se puede leer casi siempre es una hoja borrosa o mal escaneada
# (septiembre 2026): se avisa para revisarla y pedirla de nuevo, o
# completarlo a mano en "Editar datos del reporte".
BLURRY_SHEET = ("La hoja puede estar borrosa o mal escaneada: revisala y pedila de nuevo, "
                "o completalo a mano en \"Editar datos del reporte\".")


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
        warnings.append("Store Info: no se pudo leer " + ", ".join(fields["missing_fields"]) + ". " + BLURRY_SHEET)
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
            names = [d["department"] or "sin nombre" for d in diagnostics["doubtful_departments"]]
            warnings.append("Departamentos dudosos (quedaron vacíos): " + ", ".join(names) + ". " + BLURRY_SHEET)
        if diagnostics.get("total_unverified"):
            warnings.append("Departamentos: no se pudo controlar la suma contra el total impreso.")
        elif diagnostics.get("subtotal_mismatch"):
            warnings.append("Departamentos: la suma no cierra contra el total impreso.")
    except Exception as exc:
        warnings.append(f"No se pudieron leer los departamentos ({exc}). {BLURRY_SHEET}")

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


# Campos de Store Info que se pueden corregir a mano, en el orden del PDF
# ("tc" es la suma de Credit Terms: se edita como un solo importe).
EDITABLE_STORE_INFO = tuple(item for item in STORE_INFO_LABELS if item[0] != "total_fuel")
_RECALCULATED_WARNINGS = ("Store Info: no se pudo leer", "Departamentos dudosos", "Departamentos: ",
                          "No se pudieron leer los departamentos")


def parse_amount(text):
    """'$1,234.56' / '-12.5' / '(12.50)' -> float; vacío -> None; ValueError si no es un número."""
    text = (text or "").strip().replace("$", "").replace(",", "").replace(" ", "")
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    value = float(text.strip("()"))
    return -value if negative else value


def edited_report(report, form):
    """
    Corrección a mano del reporte guardado (pedido del usuario, 2026-10-07:
    el de septiembre no traía Network/Total Revenue por una hoja escaneada al
    revés). `form` es el formulario de la página: si_<campo>, dept_count_<i>,
    dept_amount_<i>, dept_name_<i> (el renglón nuevo), printed_amount y
    printed_count. Lo que se deja vacío queda sin dato. Los avisos de "no se
    pudo leer" se rehacen con lo que siga vacío; los demás se mantienen.
    Devuelve (store_info, departments, printed_amount, printed_count,
    warnings); ValueError con el campo que no es un número.
    """
    def number(name, label, integer=False):
        try:
            value = parse_amount(form.get(name))
        except ValueError:
            raise ValueError(f"{label}: no es un número.") from None
        if value is not None and integer:
            if value != int(value):
                raise ValueError(f"{label}: la cantidad tiene que ser entera.")
            value = int(value)
        return value

    store_info = dict(report["store_info"])
    for key, label in EDITABLE_STORE_INFO:
        value = number(f"si_{key}", label)
        if key == "tc":
            old = store_info.get("credit_terms")
            old_sum = round(sum(old), 2) if old else None
            if value != old_sum:
                store_info["credit_terms"] = [value] if value is not None else None
        else:
            store_info[key] = value

    departments = []
    for index, record in enumerate(report.get("departments") or []):
        departments.append({
            "department": record["department"],
            "count": number(f"dept_count_{index}", f"{record['department']} (cantidad)", integer=True),
            "amount": number(f"dept_amount_{index}", f"{record['department']} (importe)"),
        })
    new_name = (form.get("dept_name_new") or "").strip().upper()
    if new_name:
        if any(d["department"] == new_name for d in departments):
            raise ValueError(f"El departamento {new_name} ya está en el reporte.")
        departments.append({
            "department": new_name,
            "count": number("dept_count_new", f"{new_name} (cantidad)", integer=True),
            "amount": number("dept_amount_new", f"{new_name} (importe)"),
        })

    printed_amount = number("printed_amount", "Total impreso de departamentos (importe)")
    printed_count = number("printed_count", "Total impreso de departamentos (cantidad)", integer=True)

    warnings = [w for w in report.get("warnings") or [] if not w.startswith(_RECALCULATED_WARNINGS)]
    missing = [label for key, label in EDITABLE_STORE_INFO
               if (store_info.get("credit_terms") if key == "tc" else store_info.get(key)) is None]
    if missing:
        warnings.insert(0, "Store Info: no se pudo leer " + ", ".join(missing).rstrip(".") + ". " + BLURRY_SHEET)
    doubtful = [d["department"] for d in departments if d["count"] is None or d["amount"] is None]
    if doubtful:
        warnings.append("Departamentos dudosos (quedaron vacíos): " + ", ".join(doubtful) + ". " + BLURRY_SHEET)
    if not departments:
        warnings.append("No se pudieron leer los departamentos. " + BLURRY_SHEET)
    elif printed_amount is None:
        warnings.append("Departamentos: no se pudo controlar la suma contra el total impreso.")
    elif not doubtful and (
        abs(round(sum(d["amount"] for d in departments), 2) - printed_amount) > TOLERANCE
        or (printed_count is not None and sum(d["count"] for d in departments) != printed_count)
    ):
        warnings.append("Departamentos: la suma no cierra contra el total impreso.")
    return store_info, departments, printed_amount, printed_count, warnings


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
