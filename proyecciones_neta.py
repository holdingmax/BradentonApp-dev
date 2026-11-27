"""
Utilidad neta proyectada (pedido del usuario, 2026-11-02): la utilidad bruta
de Proyecciones menos los gastos del mes, con los mismos renglones que el
"ESTADO RESULTADO" del BALANCE HISTORICO del Bce.

Todo sale de lo que se carga día a día:
  - Utilidad bruta: combustible y C-Store por rubro como en Proyecciones, la
    comisión real de Lottery (proyecciones.lottery_commissions), los
    ingresos de Chase (rebates, hielo, Food Truck, Vaccumms) y, restando, los
    fees de tarjeta de los Cupones (por fecha de venta).
  - Gastos: Chase por categoría y los gastos de efectivo de Caja. Lo de Chase
    sin categoría va aparte, a la vista, para que no se pierda.

A fin de mes se proyecta así:
  - Lo que se paga una vez por mes (MONTHLY): lo real si ya se pagó; si
    todavía no, lo del mes anterior.
  - Lo que entra o sale de golpe (UNTIL_COMPLETE: sueldos quincenales,
    rebates, el Food Truck): el mayor entre lo del mes y el mes anterior,
    hasta que el mes esté completo.
  - Lo demás: al ritmo de los días cargados, pero recién con MIN_DAYS días;
    antes, el mayor entre lo del mes y el mes anterior (2026-11-24: con 6
    días de Chase, un pago grande multiplicado por 5 tiraba el mes a pérdida).
  - Lo de Chase sin categoría no entra en la utilidad neta: queda aparte,
    como pendiente de categorizar (podría ser un pago de combustible que ya
    entra por el PPP).

Chase es lo que se paga y el Bce lo que corresponde al mes, así que algunos
renglones no coinciden (en agosto 2026: sueldos $12.288 contra $16.738). Por
eso, si el mes tiene Bce, cada grupo muestra al lado lo real del Bce.
"""

import calendar
import datetime as dt

import caja_db
import chase_db
import eft_db
import proyecciones

# Categorías de Chase de cada renglón.
REBATE_FUEL = ("REBATE COMBUSTIBLE",)
REBATE_CSTORE = ("REBATE",)
ICE = ("VENTA ICE", "DEPOSITO VENTA ICE")
FOOD_TRUCK = ("FOOD TRUCK",)
VACCUMMS = ("DEPOSITO VACCUMMS",)
# (campo del Bce, renglón, categorías de Chase). "Otros Gastos Varios" suma
# además los gastos de Caja.
EXPENSES = (
    ("sueldos", "Sueldos y Jornales", ("SUELDOS", "SUELDO")),
    ("gastos_bancarios", "Comisiones y Gastos Bancarios", ("GASTOS BANCARIOS", "COMISIONES Y GASTOS BANCARIOS")),
    ("energia", "Energía Eléctrica", ("ENERGIA ELECTRICA",)),
    ("agua", "Water and Waste Services", ("AGUA",)),
    ("intereses", "Intereses y Recargos", ()),
    ("seguro", "Seguro", ("SEGURO",)),
    ("reparaciones", "Reparación y Mantenimiento", ("REPARACION Y MANTENIMIENTO",)),
    ("adm_fees", "Administration Fees", ("ADMINISTRATION FEE",)),
    ("otros_gastos_varios", "Otros Gastos Varios", ("INTERNET", "ALARMA", "TELEFONO", "ELISTAR")),
    ("otros_gastos", "Otros (Movilidad y Viáticos)", ("MOVILIDAD Y VIATICOS",)),
)
# Las facturas de una vez por mes.
MONTHLY = {"ENERGIA ELECTRICA", "AGUA", "SEGURO", "INTERNET", "ALARMA", "TELEFONO", "ELISTAR"}
# Lo que entra o sale de golpe: el alquiler del Food Truck (fijo, en varios
# depósitos), los sueldos (quincenales) y los rebates. Hasta que el mes está
# completo, el mayor entre lo del mes y el mes anterior.
UNTIL_COMPLETE = {"FOOD TRUCK", "SUELDOS", "SUELDO", "REBATE", "REBATE COMBUSTIBLE"}
# Con menos días cargados que esto no se proyecta al ritmo (2026-11-24).
MIN_DAYS = 15


def _prev_month(year, month):
    return (year - 1, 12) if month == 1 else (year, month - 1)


def _elapsed(last, year, month):
    """Días del mes cubiertos hasta `last` (una fecha o None)."""
    if last is None:
        return 0
    first = dt.date(year, month, 1)
    if last < first:
        return 0
    return min((last - first).days + 1, calendar.monthrange(year, month)[1])


def _at_pace(actual, elapsed, days_in_month, previous):
    """Al ritmo de los días cargados, con MIN_DAYS como mínimo; antes, el mayor con el mes anterior."""
    if elapsed >= days_in_month:
        return actual
    if elapsed >= MIN_DAYS:
        return actual * days_in_month / elapsed
    return max(actual, previous)


def _chase_by_category(year, month):
    """{categoría: neto} del mes; "" son los débitos sin categoría."""
    totals = {}
    for tx in chase_db.get_month_transactions(year, month):
        key = tx.get("detalle") or ""
        amount = tx.get("amount") or 0.0
        if not key and amount >= 0:
            continue
        totals[key] = totals.get(key, 0.0) + amount
    return totals


def _card_fees(year, month):
    """Fees de los Cupones con fecha de venta en el mes, y la última fecha."""
    total, last = 0.0, None
    for row in eft_db.list_cupones_for_forecast():
        day = row["date"].date() if isinstance(row["date"], dt.datetime) else row["date"]
        if (day.year, day.month) != (year, month):
            continue
        total += row.get("fees") or 0.0
        last = day if last is None or day > last else last
    return total, last


def _caja_expenses(year, month):
    items = caja_db.get_month_expense_items(year, month)
    return sum(item["amount"] or 0.0 for day_items in items.values() for item in day_items)


def _money(value):
    """$1,234.56 (−$1,234.56 en negativo), para los textos de las fórmulas."""
    if value is None:
        return "—"
    return ("−$" if value < 0 else "$") + f"{abs(value):,.2f}"


def _pace_text(actual, elapsed, days_in_month, previous, what=""):
    """Cómo sale `_at_pace`, en palabras y con los números."""
    if elapsed >= days_in_month:
        return f"{_money(actual)}{what} (el mes completo)"
    if elapsed >= MIN_DAYS:
        return (f"{_money(actual)}{what} ÷ {elapsed} días cargados"
                f" × {days_in_month} días del mes = {_money(actual * days_in_month / elapsed)}")
    return (f"{elapsed} día{'' if elapsed == 1 else 's'} cargado{'' if elapsed == 1 else 's'} (menos de {MIN_DAYS}): el mayor entre"
            f" lo del mes {_money(actual)}{what} y el mes anterior {_money(previous)} = {_money(max(actual, previous))}")


def freshness(last, year, month, today=None):
    """
    Semáforo de una fuente (pedido del usuario, 2026-11-24): "ok" si está al
    día (2 días o menos de atraso, o el mes completo), "late" hasta 6 y "bad"
    con más o sin nada del mes. None para un mes que todavía no empezó.
    """
    today = today or dt.date.today()
    first = dt.date(year, month, 1)
    last_of_month = dt.date(year, month, calendar.monthrange(year, month)[1])
    if first > today:
        return None
    if last is None or last < first:
        return "bad"
    if last >= last_of_month:
        return "ok"
    lag = (min(today, last_of_month) - last).days
    return "ok" if lag <= 2 else ("late" if lag <= 6 else "bad")


def _line(label, source, actual, proj, bce=None, formula=None):
    return {"label": label, "source": source, "actual": actual, "proj": proj, "bce": bce, "formula": formula}


def _group(title, lines, bce=None, key=None):
    def total(field):
        values = [ln[field] for ln in lines]
        return None if any(v is None for v in values) else sum(values)
    group = {"key": key, "title": title, "lines": lines, "actual": total("actual"), "proj": total("proj"), "bce": bce}
    group["formula"] = {
        field: (" + ".join(f"{ln['label']} {_money(ln[field])}" for ln in lines) + f" = {_money(group[field])}")
        if len(lines) > 1 else None
        for field in ("actual", "proj")
    }
    return group


def _var(now, before):
    if now is None or not before:
        return None
    return (now - before) / abs(before)


def _attach_previous(net, prev):
    """
    Lo del mes anterior al lado de cada renglón, y la variación de lo
    proyectado contra eso (pedido del usuario, 2026-11-24).
    """
    def put(obj, before):
        obj["prev"] = before
        obj["var"] = _var(obj.get("proj"), before)

    prev_groups = {g["title"]: g for g in prev["gross_groups"]}
    for g in net["gross_groups"]:
        pg = prev_groups.get(g["title"])
        put(g, pg["actual"] if pg else None)
        prev_lines = {ln["label"]: ln for ln in (pg["lines"] if pg else [])}
        for ln in g["lines"]:
            put(ln, prev_lines[ln["label"]]["actual"] if ln["label"] in prev_lines else None)
    prev_expenses = {ln["label"]: ln for ln in prev["expenses"]["lines"]}
    for ln in net["expenses"]["lines"]:
        put(ln, prev_expenses[ln["label"]]["actual"] if ln["label"] in prev_expenses else None)
    for key in ("gross", "expenses", "net"):
        put(net[key], prev[key]["actual"])
    net["prev_label"] = f"{prev['month']:02d}/{prev['year']}"


def build_net(year, month, projection=None, with_previous=False):
    """
    El desglose de la utilidad neta del mes: lo de hasta hoy, lo proyectado y
    lo real del Bce. Con `with_previous`, también el mes anterior.
    """
    p = projection or proyecciones.build_projection(year, month)
    days_in_month = calendar.monthrange(year, month)[1]
    prev_year, prev_month = _prev_month(year, month)

    last_chase = chase_db.get_last_posting_date()
    last_chase = dt.date.fromisoformat(str(last_chase)[:10]) if last_chase else None
    chase_elapsed = _elapsed(last_chase, year, month)
    chase = _chase_by_category(year, month)
    chase_prev = _chase_by_category(prev_year, prev_month)
    chase_until = last_chase if chase_elapsed and chase_elapsed < days_in_month else None

    def chase_amount(categories, sign, totals):
        # + 0.0: que un gasto en cero no salga "−0.00".
        return sign * sum(totals.get(c, 0.0) for c in categories) + 0.0

    def chase_line(label, categories, sign=1):
        """Un renglón de Chase: ingresos con sign=1, gastos con sign=-1 (quedan en positivo)."""
        actual = chase_amount(categories, sign, chase)
        proj = 0.0
        steps = []
        for category in categories:
            now = chase_amount((category,), sign, chase)
            previous = chase_amount((category,), sign, chase_prev)
            name = category or "sin categoría"
            if chase_elapsed >= days_in_month:
                proj += now
                steps.append(f"{name}: {_money(now)} (el mes completo)")
            elif category in MONTHLY:
                proj += now if now else previous
                steps.append(f"{name}: lo pagado en el mes, {_money(now)}" if now
                             else f"{name}: todavía no se pagó, lo del mes anterior {_money(previous)}")
            elif category in UNTIL_COMPLETE:
                proj += max(now, previous)
                steps.append(f"{name}: hasta completar el mes, el mayor entre lo del mes ({_money(now)}) y el mes anterior ({_money(previous)})")
            else:
                proj += _at_pace(now, chase_elapsed, days_in_month, previous)
                steps.append(f"{name}: " + _pace_text(now, chase_elapsed, days_in_month, previous))
        formula = {
            "actual": f"Chase del mes: {', '.join(categories) if any(categories) else 'débitos sin categoría'}",
            "proj": "\n".join(steps) + (f"\nTotal: {_money(proj)}" if len(steps) > 1 else ""),
        }
        names = ", ".join(categories) if categories else "sin categoría"
        if chase_elapsed >= days_in_month:
            how = ""
        elif any(c in MONTHLY for c in categories):
            how = "si todavía no se pagó, la del mes anterior"
        elif any(c in UNTIL_COMPLETE for c in categories):
            how = "hasta completar el mes, el mayor con el mes anterior"
        elif chase_elapsed < MIN_DAYS:
            how = f"menos de {MIN_DAYS} días cargados: el mayor con el mes anterior"
        elif chase_until:
            how = f"al ritmo hasta el {chase_until.strftime('%d/%m')}"
        else:
            how = ""
        return _line(label, f"Chase: {names}" + (f" · {how}" if how else ""), actual, proj, formula=formula)

    real = proyecciones.balance_for_month(year, month)

    def bce(*fields, sign=1):
        if not real:
            return None
        values = [real.get(f) for f in fields]
        if any(v is None for v in values):
            return None
        return sign * sum(values)

    # ---------------- Utilidad bruta ----------------
    def fuel_text(block):
        if block["utilidad"] is None:
            return "Sin PPP no se puede calcular: falta el inventario inicial o las facturas del mes en Físico."
        return (f"Venta {_money(block['sales'])} − {block['gallons']:,.2f} gal × PPP {_money(p['ppp'])}"
                f" = {_money(block['utilidad'])}")

    fuel = _group("Combustible", [
        _line("Venta − costo (PPP)", "Close Store y Físico", p["fuel"]["utilidad"], p["fuel_proj"]["utilidad"],
              formula={"actual": fuel_text(p["fuel"]), "proj": fuel_text(p["fuel_proj"])}),
        chase_line("Rebate de combustible", REBATE_FUEL),
    ], bce=(bce("ventas_combustible", "otros_ingresos_combustible") - real["costo_combustible"])
        if bce("ventas_combustible", "otros_ingresos_combustible", "costo_combustible") is not None else None)

    rubros = [r for r in p["cstore"]["rubros"] if r["in_balance"]]
    cstore = _group("C-Store", [
        _line(r["display"], f"Close Store × {r['margin'] * 100:.0f} % (supuesto)", r["utilidad"], r["utilidad_proj"], formula={
            "actual": f"Vendido {_money(r['amount'])} × {r['margin'] * 100:.2f} % = {_money(r['utilidad'])}",
            "proj": (f"Proyectado {_money(r['proj'])} × {r['margin'] * 100:.2f} % = {_money(r['utilidad_proj'])}"
                     if r["proj"] is not None else None),
        }) for r in rubros
    ] + [chase_line("Rebates", REBATE_CSTORE)],
        bce=(real["ventas_cstore"] - real["costo_cstore"]) if bce("ventas_cstore", "costo_cstore") is not None else None,
        key="cstore")

    # Lottery por su % teórico, como en el Panel (pedido del usuario, 2026-11-24).
    lotto = next(r for r in p["cstore"]["rubros"] if r.get("lottery"))
    lottery = _group("Lottery", [_line(
        "Venta de LOTTO", f"Close Store × {lotto['margin'] * 100:.2f} % (supuesto)", lotto["utilidad"], lotto["utilidad_proj"],
        formula={"actual": f"Vendido {_money(lotto['amount'])} × {lotto['margin'] * 100:.2f} % = {_money(lotto['utilidad'])}",
                 "proj": (f"Proyectado {_money(lotto['proj'])} × {lotto['margin'] * 100:.2f} % = {_money(lotto['utilidad_proj'])}"
                          if lotto["proj"] is not None else None)})],
        bce=bce("comisiones_lottery"))

    others = _group("Otros ingresos", [
        chase_line("Máquina de hielo", ICE),
        chase_line("Alquiler Food Truck", FOOD_TRUCK),
        chase_line("Vaccumms", VACCUMMS),
    ], bce=bce("ventas_ice", "alquiler_food_truck", "ventas_vaccum", "otros_ingresos"))

    fees, last_coupon = _card_fees(year, month)
    fees_prev, _ = _card_fees(prev_year, prev_month)
    coupon_elapsed = _elapsed(last_coupon, year, month)
    fees_source = "Cupones: fees por fecha de venta"
    if MIN_DAYS <= coupon_elapsed < days_in_month:
        fees_source += f" · al ritmo hasta el {last_coupon.strftime('%d/%m')}"
    elif coupon_elapsed < MIN_DAYS:
        fees_source += f" · menos de {MIN_DAYS} días cargados: el mayor con el mes anterior"
    cards = _group("Gastos de tarjeta", [
        _line("Fees de tarjeta", fees_source, -fees, -_at_pace(fees, coupon_elapsed, days_in_month, fees_prev),
              formula={"actual": "Fees de los cupones con fecha de venta en el mes",
                       "proj": "Fees " + _pace_text(fees, coupon_elapsed, days_in_month, fees_prev)}),
    ], bce=bce("gastos_tarjeta", sign=-1))

    gross_groups = [fuel, cstore, lottery, others, cards]

    # ---------------- Gastos ----------------
    store_elapsed = _elapsed(p["last_day"], year, month)
    caja = _caja_expenses(year, month)
    caja_prev = _caja_expenses(prev_year, prev_month)
    caja_proj = _at_pace(caja, store_elapsed, days_in_month, caja_prev)
    expense_lines = []
    for field, label, categories in EXPENSES:
        line = chase_line(label, categories, sign=-1)
        if field == "otros_gastos_varios":
            line["actual"] += caja
            line["proj"] += caja_proj
            line["source"] += " · Caja: gastos de efectivo"
            line["formula"] = {
                "actual": line["formula"]["actual"] + f" + gastos de Caja {_money(caja)}",
                "proj": line["formula"]["proj"] + "\nCaja: " + _pace_text(caja, store_elapsed, days_in_month, caja_prev)
                        + f"\nTotal: {_money(line['proj'])}",
            }
        if not categories:
            # Intereses y Recargos no tiene categoría en Chase: solo aparece si el Bce lo trae.
            if not bce(field):
                continue
            line = _line(label, "No tiene categoría en Chase", 0.0, 0.0)
        line["bce"] = bce(field, sign=-1)
        expense_lines.append(line)
    # Lo de Chase sin categoría queda aparte: no entra en la utilidad neta
    # hasta que se categorice (2026-11-24).
    uncategorized = [tx for tx in chase_db.get_month_transactions(year, month)
                     if not (tx.get("detalle") or "") and (tx.get("amount") or 0.0) < 0]
    pending = {"amount": -sum(tx.get("amount") or 0.0 for tx in uncategorized) + 0.0, "count": len(uncategorized),
               "items": [{"date": str(tx.get("posting_date") or tx.get("date") or "")[:10],
                          "description": tx.get("description") or "", "amount": tx.get("amount") or 0.0}
                         for tx in uncategorized]}
    expenses = _group("Gastos", expense_lines, bce=bce("gastos_directos", "otros_gastos", sign=-1))

    def add(groups, field):
        values = [g[field] for g in groups]
        return None if any(v is None for v in values) else sum(values)

    gross = {"actual": add(gross_groups, "actual"), "proj": add(gross_groups, "proj"), "bce": bce("ganancia_bruta")}
    net = {
        field: (gross[field] - expenses[field]) if (gross[field] is not None and expenses[field] is not None) else None
        for field in ("actual", "proj")
    }
    net["bce"] = bce("resultado_neto")
    gross["formula"] = {field: " + ".join(f"{g['title']} {_money(g[field])}" for g in gross_groups) + f" = {_money(gross[field])}"
                        for field in ("actual", "proj")}

    # El puente de la venta a la utilidad neta (2026-11-24): el margen de
    # ventas (combustible + C-Store con Lottery, lo del gráfico del Panel),
    # más los otros ingresos, menos los fees = utilidad bruta, menos gastos.
    sales_lines = [fuel["lines"][0]] + cstore["lines"][:-1] + lottery["lines"]
    income_lines = [fuel["lines"][1], cstore["lines"][-1]] + others["lines"]

    def part(title, lines):
        return {"title": title, **{f: add(lines, f) for f in ("actual", "proj")},
                "formula": {f: " + ".join(f"{ln['label']} {_money(ln[f])}" for ln in lines) + f" = {_money(add(lines, f))}"
                            for f in ("actual", "proj")}}

    bridge = {
        "sales_margin": part("Margen de ventas", sales_lines),
        "other_income": part("Otros ingresos", income_lines),
        "fees": part("Fees de tarjeta", cards["lines"]),
    }
    net["formula"] = {field: f"Utilidad bruta {_money(gross[field])} − gastos {_money(expenses[field])} = {_money(net[field])}"
                      for field in ("actual", "proj")}
    result = {
        "year": year,
        "month": month,
        "days_in_month": days_in_month,
        "last_store_day": p["last_day"],
        "last_chase_day": last_chase,
        "last_coupon_day": last_coupon,
        "sources": [
            {"name": "Close Store", "last": p["last_day"], "status": freshness(p["last_day"], year, month)},
            {"name": "Chase", "last": last_chase, "status": freshness(last_chase, year, month)},
            {"name": "Cupones", "last": last_coupon, "status": freshness(last_coupon, year, month)},
        ],
        "pending": pending,
        "bridge": bridge,
        "min_days": MIN_DAYS,
        "gross_groups": gross_groups,
        "gross": gross,
        "expenses": expenses,
        "net": net,
        "has_bce": bool(real and real.get("resultado_neto") is not None),
        "bce_source": real.get("archivo") if real else None,
        "ppp_missing": p["ppp"] is None,
    }
    if with_previous:
        _attach_previous(result, build_net(prev_year, prev_month))
    return result


def balance_comparison(projection, net):
    """
    La proyección de fin de mes contra el Bce, rubro por rubro (pedido del
    usuario, 2026-11-26: los dueños comparan solo los números finales). Usa
    los mismos números que el desglose de la utilidad neta, así las dos
    páginas coinciden. Cada renglón: {"label", "proj", "bce", "diff", "pct",
    "kind" ("$" o "gal"), "level" ("head", "line" o "total")}; diff es
    proyección − Bce y pct, diff sobre el Bce. None si el mes no tiene Bce.
    """
    real = proyecciones.balance_for_month(projection["year"], projection["month"])
    if not real:
        return None
    groups = {g["title"]: g for g in net["gross_groups"]}
    fuel, cstore, lottery = groups["Combustible"], groups["C-Store"], groups["Lottery"]
    others, cards = groups["Otros ingresos"], groups["Gastos de tarjeta"]
    fuel_proj = projection["fuel_proj"]
    cs = projection["cstore"]["balance"]

    def minus(a, b):
        return (a - b) if (a is not None and b is not None) else None

    rows = []

    def row(label, proj, bce, level="line", kind="$"):
        diff = minus(proj, bce)
        # Al centavo y + 0.0: que una coincidencia no salga "−$0.00".
        diff = round(diff, 2) + 0.0 if diff is not None else None
        pct = (round(diff / abs(bce), 4) + 0.0) if (diff is not None and bce) else None
        rows.append({"label": label, "proj": proj, "bce": bce, "diff": diff, "pct": pct, "kind": kind, "level": level})

    def head(label):
        rows.append({"label": label, "level": "head"})

    head("Combustible")
    row("Galones vendidos", fuel_proj["gallons"], real.get("galones"), kind="gal")
    row("Venta de combustible (4.01)", fuel_proj["sales"], real.get("ventas_combustible"))
    row("Costo de venta (5.02)", minus(fuel_proj["sales"], fuel_proj["utilidad"]), real.get("costo_combustible"))
    row("Rebate de combustible (4.07)", fuel["lines"][1]["proj"], real.get("otros_ingresos_combustible"))
    row("Utilidad de combustible", fuel["proj"], fuel["bce"], "total")

    head("C-Store (sin Lotto ni Kia/Toyota)")
    row("Venta C-Store (4.09)", cs["proj"], real.get("ventas_cstore"))
    row("Costo de venta (5.15)", minus(cs["proj"], cs["utilidad_proj"]), real.get("costo_cstore"))
    row("Rebates", cstore["lines"][-1]["proj"], None)
    row("Utilidad C-Store", cstore["proj"], cstore["bce"], "total")

    head("Otros ingresos")
    row("Comisiones Lottery (4.10)", lottery["proj"], lottery["bce"])
    for line, field in zip(others["lines"], ("ventas_ice", "alquiler_food_truck", "ventas_vaccum")):
        row(line["label"], line["proj"], real.get(field))
    row("Otros ingresos y descuentos (4.03)", None, real.get("otros_ingresos"))
    row("Fees de tarjeta (5.04)", cards["proj"], cards["bce"])

    row("Utilidad bruta", net["gross"]["proj"], net["gross"]["bce"], "total")

    # Como el Estado de Resultado del Bce: gastos directos, resultado
    # operativo, expensas y "Otros" (en el Bce es un ingreso: +6,535.51 en
    # septiembre 2026, no los viáticos de la app), y el resultado neto. La app
    # no proyecta expensas ni esos otros: van solo con el número del Bce.
    other_expense = next(label for field, label, _ in EXPENSES if field == "otros_gastos")
    head("Gastos")
    for line in net["expenses"]["lines"]:
        row(line["label"], line["proj"], None if line["label"] == other_expense else line["bce"])
    directos = real.get("gastos_directos")
    row("Total de gastos directos", net["expenses"]["proj"], -directos + 0.0 if directos is not None else None, "total")

    row("Resultado operativo", net["net"]["proj"], real.get("resultado_operativo"), "total")
    row("Expensas", None, real.get("expensas"))
    row("Otros (ingresos)", None, real.get("otros_gastos"))
    row("Utilidad neta", net["net"]["proj"], net["net"]["bce"], "total")
    return {"rows": rows, "source": real.get("archivo"), "loaded_at": real.get("cargado")}
