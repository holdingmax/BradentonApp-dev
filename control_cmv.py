"""
Control CMV (pedido del usuario, 2026-10-07): las ventas de cada
departamento del mes, verificadas tres veces:

- Reportes diarios del POS (Ventas por Departamento, reportes_db).
- Reporte mensual del POS (reporte_mensual_db).
- La página de Elistar: el "Depts Report" (ventas por día y por categoría de
  Elistar) o, si no, el "P & L Report" (solo el total del mes). Los dos se
  bajan como .xls, que en realidad son páginas HTML.

Más el reporte de ventas que se carga en CMV (cmv_db, el Top-Selling por
UPC), por departamento del POS con cantidades.

Elistar no usa los departamentos del POS sino categorías que juntan varios
(Cigarette = MAJ PAK + GEN-PAK, Lotto = SKOFF + ONLINE, ...): ELISTAR_GROUPS,
sacado de septiembre 2026, donde cada categoría cierra al centavo. Cálculo
puro, sin UI ni base.
"""

import re
from html.parser import HTMLParser

TOLERANCE = 0.005

# Categoría de Elistar -> departamentos del POS (por clave, ver dept_key).
ELISTAR_GROUPS = {
    "Auto": ("AUTO",),
    "Beer/ Wine": ("BEER/WINE",),
    "Cigarette": ("MAJ PAK", "GEN-PAK"),
    "Cigarette Carton": ("GEN-CTN", "MAJ CR"),
    "E cig": ("E-GIGARETTE",),
    "Fountain/Coffee": ("FOUTAIN", "COFFE"),
    "HBA": ("HBA",),
    "Lotto": ("SKOFF", "ONLINE"),
    "Non-Tax Sales": ("LOCAL ACCT", "WATER", "JUICE", "MILK", "NONTAX"),
    "Propane": ("PROPANE",),
    "Soda": ("SODA",),
    "Taxable Sales": ("TAXABLE", "SNACK", "FLOWERS", "CANDY", "ICECREAM", "BOILED PEANUTS", "GROCERIES",
                      "HOT DOGS & SANDWICH"),
    "Tobacco": ("CIGARS", "SNUFF"),
}
# Elistar trae el combustible como una categoría más: en el Depts Report es
# Sales Fuel + Desc. Comb de Store Info (la venta de combustible del asiento),
# al centavo cada día de septiembre 2026. El Fuel del P & L es otro número
# ($177,133.60 contra $175,365.42) y no se compara.
FUEL = "Fuel"
# Departamentos del POS que Elistar no cuenta en ninguna categoría.
NOT_IN_ELISTAR = ("CHEVRON GIFT CARD", "STORE COUPON")


def dept_key(name):
    """El mismo departamento aunque cambien signos o espacios ("MAJPAK", "FLOWER")."""
    key = re.sub(r"[^A-Z0-9]", "", (name or "").upper())
    return {"FLOWER": "FLOWERS"}.get(key, key)


_GROUP_OF = {dept_key(d): group for group, depts in ELISTAR_GROUPS.items() for d in depts}


def elistar_group(department):
    """La categoría de Elistar de un departamento del POS (None si Elistar no lo cuenta o no se conoce)."""
    return _GROUP_OF.get(dept_key(department))


# ---------------------------------------------------------------------------
# Lectura de los .xls de Elistar (HTML)
# ---------------------------------------------------------------------------

class _Tables(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows, self._row, self._cell, self._span = [], None, None, 1
        self.text = []

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []
            self._span = int(dict(attrs).get("colspan") or 1)

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cell is not None:
            value = " ".join("".join(self._cell).split())
            self._row.extend([value] + [None] * (self._span - 1))
            self._cell = None
        elif tag == "tr" and self._row is not None:
            self.rows.append(self._row)
            self._row = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)
        self.text.append(data)


def _money(text):
    text = (text or "").strip()
    if not text:
        return None
    negative = text.startswith("(") or text.startswith("-")
    value = float(re.sub(r"[^0-9.]", "", text) or "nan")
    if value != value:
        raise ValueError(f"importe que no se entiende: {text!r}")
    return -value if negative else value


_MONTHS = {m: i for i, m in enumerate(
    ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"), start=1)}


def _period(text):
    """'01 Sep 2026 - 30 Sep 2026' -> (año, mes); tiene que ser un mes entero."""
    import calendar
    match = re.search(r"(\d{1,2}) (\w{3}) (\d{4}) - (\d{1,2}) (\w{3}) (\d{4})", text)
    if not match:
        raise ValueError("No se encontró el período del reporte de Elistar.")
    d1, m1, y1, d2, m2, y2 = match.groups()
    year, month = int(y1), _MONTHS.get(m1[:3].title())
    if not month or (m2[:3].title(), y2) != (m1[:3].title(), y1) or int(d1) != 1 \
            or int(d2) != calendar.monthrange(year, month)[1]:
        raise ValueError(f"El reporte de Elistar tiene que ser de un mes entero (dice {match.group(0)}).")
    return year, month


def read_elistar_report(path):
    """
    Lee un "Depts Report" o un "P & L Report" de Elistar. Devuelve
    {"kind": "depts"|"pl", "year", "month", "totals": {categoría: monto},
    "days": {fecha ISO: {categoría: monto}} (vacío en el P & L)}.
    """
    with open(path, encoding="utf-8", errors="replace") as handle:
        html = handle.read()
    parser = _Tables()
    parser.feed(html)
    text = " ".join(" ".join(parser.text).split())
    year, month = _period(text)
    rows = parser.rows
    if "Depts Report" in text:
        # Encabezado: "Date" y cada categoría con 1 o 2 columnas (Sale y, a
        # veces, Purchase); la segunda fila dice cuál es cuál. Nos quedamos con Sale.
        header, subheader = rows[0][1:], rows[1]
        sale_index, position = {}, 0
        for i, name in enumerate(header):
            if name is None:
                continue
            span = 1
            while i + span < len(header) and header[i + span] is None:
                span += 1
            labels = subheader[position:position + span]
            if "Sale" in labels:
                sale_index[name] = 1 + position + labels.index("Sale")
            position += span
        days, totals = {}, {}
        for row in rows[2:]:
            match = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", row[0] or "")
            if match:
                iso = f"{match.group(3)}-{match.group(1)}-{match.group(2)}"
                days[iso] = {name: _money(row[i]) for name, i in sale_index.items() if _money(row[i])}
            elif not (row[0] or "").strip() and len(row) > 1:
                # Fila de totales: sin rótulo (pie de la tabla).
                totals = {name: _money(row[i]) or 0.0 for name, i in sale_index.items()}
        summed = {}
        for values in days.values():
            for name, amount in values.items():
                summed[name] = round(summed.get(name, 0.0) + amount, 2)
        for name, amount in totals.items():
            if abs(summed.get(name, 0.0) - amount) > TOLERANCE:
                raise ValueError(f"El Depts Report de Elistar no cierra: {name} suma ${summed.get(name, 0):,.2f} "
                                 f"por día y el total dice ${amount:,.2f}.")
        return {"kind": "depts", "year": year, "month": month,
                "totals": {k: v for k, v in summed.items() if v}, "days": days}
    if "P & L Report" in text or "P &amp; L Report" in html:
        totals = {}
        for row in rows:
            if len(row) >= 2 and row[0] and row[0] not in ("Dept", "Total", "Gross Profit", "Net Profit", "Name",
                                                            "Sales Tax", "Gallons", "Inventory +/-"):
                amount = _money(row[1])
                if amount is not None:
                    totals[row[0]] = amount
        if not totals:
            raise ValueError("No se encontraron ventas por categoría en el P & L de Elistar.")
        return {"kind": "pl", "year": year, "month": month, "totals": totals, "days": {}}
    raise ValueError("No es un Depts Report ni un P & L Report de Elistar.")


# ---------------------------------------------------------------------------
# Control
# ---------------------------------------------------------------------------

def _cell(values, key):
    return values.get(key) if values else None


def _diff(a, b):
    if a is None or b is None:
        return None
    return round(a - b, 2)


def _group_totals(departments):
    """{categoría de Elistar: monto} con los departamentos del POS."""
    totals = {}
    for d in departments:
        group = elistar_group(d["department"])
        if group:
            totals[group] = round(totals.get(group, 0.0) + (d["amount"] or 0.0), 2)
    return totals


def build_control(days_departments, monthly_departments, cmv_departments, elistar,
                  days_by_date=None, fuel_by_date=None, fuel_monthly=None):
    """
    days_departments: reportes_db.get_month_department_totals; monthly_departments:
    los departamentos del reporte mensual (o None); cmv_departments:
    cmv_db.get_month_department_totals (o []); elistar: read_elistar_report
    guardado (o None); days_by_date: {fecha: [{department, amount}]} de los
    reportes diarios; fuel_by_date: {fecha: Sales Fuel + Desc. Comb de Store
    Info}; fuel_monthly: lo mismo del reporte mensual (o None).
    """
    issues = []

    # 1) Por departamento del POS: días, mensual y CMV, con cantidades.
    by_key = {}

    def put(rows, source, name_field="department"):
        for r in rows or []:
            key = dept_key(r[name_field])
            entry = by_key.setdefault(key, {"department": r[name_field]})
            entry[source] = {"count": r.get("count"), "amount": round(r.get("amount") or 0.0, 2)}

    put(days_departments, "days")
    put(monthly_departments, "monthly")
    put(cmv_departments, "cmv", "dept_name")
    departments = []
    for key, entry in by_key.items():
        days, monthly, cmv = entry.get("days"), entry.get("monthly"), entry.get("cmv")
        row = {"department": entry["department"], "group": elistar_group(entry["department"]),
               "days": days, "monthly": monthly, "cmv": cmv, "problems": []}
        base = days or {"count": 0, "amount": 0.0}
        for source, label in (("monthly", "reporte mensual"), ("cmv", "CMV")):
            other = entry.get(source)
            if source == "monthly" and monthly_departments is None:
                continue
            if source == "cmv" and not cmv_departments:
                continue
            # El CMV no trae lo que no tiene UPC (Lotto, Local Acct, ...): un
            # departamento que falta en el CMV solo cuenta si el POS lo vendió con UPC.
            if other is None:
                if source == "cmv":
                    row[source + "_missing"] = True
                    continue
                other = {"count": 0, "amount": 0.0}
            amount_diff = _diff(base["amount"], other["amount"])
            count_diff = (base["count"] or 0) - (other["count"] or 0)
            row[source + "_diff"] = {"amount": amount_diff, "count": count_diff}
            if abs(amount_diff) > TOLERANCE or count_diff:
                row["problems"].append(label)
        departments.append(row)
    departments.sort(key=lambda r: -max((r["days"] or {}).get("amount") or 0, (r["monthly"] or {}).get("amount") or 0))
    for row in departments:
        if "reporte mensual" in row["problems"]:
            d = row["monthly_diff"]
            issues.append(f"{row['department']}: reportes diarios contra el reporte mensual "
                          f"{_fmt_count(d['count'])}{_fmt_amount(d['amount'])}.")

    # 2) Por categoría de Elistar: Elistar, días y mensual.
    groups = []
    if elistar:
        days_groups = _group_totals(days_departments)
        monthly_groups = _group_totals(monthly_departments) if monthly_departments is not None else None
        names = list(ELISTAR_GROUPS) + sorted(set(elistar["totals"]) - set(ELISTAR_GROUPS) - {FUEL})
        for name in names:
            el = elistar["totals"].get(name, 0.0)
            if name not in ELISTAR_GROUPS:
                groups.append({"group": name, "departments": [], "elistar": el, "days": None, "monthly": None,
                               "diff_days": None, "diff_monthly": None, "ok": abs(el) <= TOLERANCE, "unknown": True})
                if abs(el) > TOLERANCE:
                    issues.append(f"Elistar: la categoría {name} (${el:,.2f}) no está en el control; "
                                  f"hay que decir a qué departamentos del POS corresponde.")
                continue
            days = days_groups.get(name, 0.0)
            monthly = monthly_groups.get(name, 0.0) if monthly_groups is not None else None
            row = {"group": name, "departments": list(ELISTAR_GROUPS[name]), "elistar": el, "days": days,
                   "monthly": monthly, "diff_days": _diff(days, el), "diff_monthly": _diff(monthly, el),
                   "unknown": False}
            row["ok"] = abs(row["diff_days"]) <= TOLERANCE and (monthly is None or abs(row["diff_monthly"]) <= TOLERANCE)
            if abs(row["diff_days"]) > TOLERANCE:
                issues.append(f"{name}: reportes diarios ${days:,.2f} y Elistar ${el:,.2f}.")
            if monthly is not None and abs(row["diff_monthly"]) > TOLERANCE:
                issues.append(f"{name}: reporte mensual ${monthly:,.2f} y Elistar ${el:,.2f}.")
            groups.append(row)
        if elistar["kind"] == "depts" and FUEL in elistar["totals"] and fuel_by_date is not None:
            el = elistar["totals"][FUEL]
            days = round(sum(fuel_by_date.values()), 2)
            row = {"group": "Combustible (Sales Fuel + Desc. Comb)", "departments": [], "elistar": el, "days": days,
                   "monthly": fuel_monthly, "diff_days": _diff(days, el), "diff_monthly": _diff(fuel_monthly, el),
                   "unknown": False, "fuel": True}
            row["ok"] = abs(row["diff_days"]) <= TOLERANCE and (fuel_monthly is None or abs(row["diff_monthly"]) <= TOLERANCE)
            if not row["ok"]:
                issues.append(f"Combustible: reportes diarios ${days:,.2f}"
                              + (f", reporte mensual ${fuel_monthly:,.2f}" if fuel_monthly is not None else "")
                              + f" y Elistar ${el:,.2f}.")
            groups.append(row)
        unknown = sorted({d["department"] for d in departments
                          if d["group"] is None and dept_key(d["department"]) not in {dept_key(n) for n in NOT_IN_ELISTAR}
                          and ((d["days"] or {}).get("amount") or 0)})
        for name in unknown:
            issues.append(f"El departamento {name} no tiene categoría de Elistar en el control.")

    # 3) Por día (solo con el Depts Report): cada categoría de Elistar contra
    # los departamentos de ese día. Y el combustible contra Sales Fuel.
    day_rows = []
    if elistar and elistar.get("days") and days_by_date is not None:
        for iso in sorted(set(elistar["days"]) | set(days_by_date)):
            el = elistar["days"].get(iso, {})
            mine = _group_totals(days_by_date.get(iso, []))
            loaded = iso in days_by_date
            diffs = []
            for name in ELISTAR_GROUPS:
                a, b = mine.get(name, 0.0), el.get(name, 0.0)
                if loaded and abs(a - b) > TOLERANCE:
                    diffs.append({"group": name, "days": a, "elistar": b, "diff": round(a - b, 2)})
            # Un día que Elistar trae sin combustible es $0 (la lectura saca los
            # ceros); si falta de un lado y del otro hay venta, no está bien
            # (revisión 2026-10-08: antes pasaba como "ok").
            fuel_mine = (fuel_by_date or {}).get(iso)
            fuel_el = el.get(FUEL, 0.0 if iso in elistar["days"] else None)
            fuel_diff = None if fuel_el is None or fuel_mine is None else round(fuel_mine - fuel_el, 2)
            fuel_missing = fuel_diff is None and abs(fuel_el or fuel_mine or 0.0) > TOLERANCE
            day_rows.append({"date": iso, "loaded": loaded, "diffs": diffs, "fuel_elistar": fuel_el,
                             "fuel_days": fuel_mine, "fuel_diff": fuel_diff, "fuel_missing": fuel_missing,
                             "ok": loaded and not diffs and not fuel_missing
                                   and (fuel_diff is None or abs(fuel_diff) <= TOLERANCE)})
        bad_days = [r for r in day_rows if r["loaded"] and not r["ok"]]
        if bad_days:
            issues.append(f"{len(bad_days)} día(s) con diferencias contra Elistar: "
                          + ", ".join(f"{r['date'][8:10]}/{r['date'][5:7]}" for r in bad_days) + ".")
        missing = [r for r in day_rows if not r["loaded"]]
        if missing:
            issues.append(f"{len(missing)} día(s) sin Ventas por Departamento cargadas: "
                          + ", ".join(f"{r['date'][8:10]}/{r['date'][5:7]}" for r in missing) + ".")

    cmv_problems = [d for d in departments if "CMV" in d["problems"]]
    # El "Bien" de arriba dice que coinciden diarios, mensual y Elistar: sin
    # mensual o con el CMV distinto no se puede decir (revisión 2026-10-08).
    if cmv_problems:
        issues.append(f"{len(cmv_problems)} departamento(s) no coinciden con las ventas del CMV.")
    if elistar and monthly_departments is None:
        issues.append("Falta el reporte mensual del POS de este mes para cruzarlo.")
    return {
        "departments": departments,
        "groups": groups,
        "days": day_rows,
        "has_monthly": monthly_departments is not None,
        "has_cmv": bool(cmv_departments),
        "cmv_problems": cmv_problems,
        "issues": issues,
        "ok": not issues and bool(elistar),
    }


def _fmt_count(diff):
    return f"{diff:+d} ítem(s), " if diff else ""


def _fmt_amount(diff):
    if diff is None or abs(diff) <= TOLERANCE:
        return "mismo importe"
    return f"{'-' if diff < 0 else '+'}${abs(diff):,.2f}"
