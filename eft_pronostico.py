"""
Próximos EFT de J.H. Williams (pedido del usuario, 2026-10-13): qué días
llegan los EFT siguientes y cuánto van a dar, para tener fondos en Chase si
viene un débito (los cupones no alcanzan para la factura) o saber cuánto se
acredita. Solo cálculo; la pantalla es /carga-datos/eft/proximos (barra
lateral de EFT y Cupones) y el cuadro de la botonera.

Cómo arma J.H. un EFT (medido con los 19 EFT de junio a septiembre 2026):
- Cae el día que vence una factura de combustible (la entrega + 5 días; si
  cae sábado, domingo o feriado bancario, el día hábil siguiente): uno por
  semana, ~8,600 galones, $30,000-$36,000. La factura cargada en Carga de
  Datos -> Combustible da la fecha y el importe exactos; sin ella se estima
  una semana después del EFT anterior, con el importe de la última.
- Paga la factura con los cupones (DDC) que todavía no usó: lo que sobra se
  acredita en Chase y lo que falta lo debita.
- Qué cupones entran: los principales (las ventas del POS de un día, más de
  $1,000) depositados hasta 2 días hábiles antes del EFT, y los chicos
  ("slri…", de $5 a $220) creados hasta el último de esos principales o,
  si son de hasta 2 días antes del corte, antes del principal siguiente
  (J.H. arma el EFT entre esos dos). El número del DDC dice el orden en que
  se crearon. Con los 34 EFT de marzo a septiembre 2026 acierta los DDC de
  todos menos uno ($9.89 de más).
- Con los cupones cargados (Carga de Datos -> EFT y Cupones, el Credit
  Card Daily Summary) el importe sale de los DDC mismos. Para lo que no está
  cargado, de Store Info: las ventas con tarjeta de los días que J.H.
  deposita hasta ese corte (control_tarjetas.jh_deposit_day) menos la
  comisión de los cupones principales (~2.33%), más lo que vienen dando
  por día los cupones chicos. Un día sin Store Info se estima con el
  promedio de ese día de la semana (sin los pagos de Kia/Toyota).
- Cargos fijos de J.H. que también paga el EFT: los que se repiten todos
  los meses con el mismo importe (Chevron network $171 a fin de mes, VPN $7
  y Mako $131 a mitad de mes) se suman en el EFT del mes que corresponde.
  Uno suelto (gift cards $100) no se puede prever.
"""

from datetime import date, datetime, timedelta

import chase_db
import eft_db
import fisico_db
import reportes_db
from control_tarjetas import bank_holidays, jh_deposit_day

MAIN_COUPON_MIN = 1000.0  # un DDC de más es el de las ventas del POS de un día
FUEL_INVOICE_MIN = 1000.0  # una factura de menos es un cargo chico
DEFAULT_MAIN_FEE = 0.0233
WEEK_DAYS = 7
HORIZON_DAYS = 10  # se muestran los EFT hasta hoy + 10 días
MAX_EVENTS = 8
FEE_SAMPLE = 4  # EFT con los que se mide la comisión
SMALL_SAMPLE = 2  # EFT con los que se mide lo que dan por día los cupones chicos (vienen en aumento)
AVERAGE_WEEKS = 4
SMALL_LOOKBACK_DAYS = 10  # un cupón chico de antes del corte anterior sin usar: ya lo pagó un EFT no cargado
RECURRING_LOOKBACK_DAYS = 120
RECURRING_GAP = (24, 38)  # un cargo que vuelve cada 24-38 días es mensual
EFT_DETALLE = "EFT RCV-"


def is_business_day(day):
    return day.weekday() < 5 and day not in bank_holidays(day.year)


def next_business_day(day):
    while not is_business_day(day):
        day += timedelta(days=1)
    return day


def business_days_before(day, count):
    while count:
        day -= timedelta(days=1)
        if is_business_day(day):
            count -= 1
    return day


def cutoff(eft_day):
    """Último día de depósito cuyos cupones entran en el EFT: 2 días hábiles antes."""
    return business_days_before(eft_day, 2)


def _parse(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(value), fmt).date()
        except (TypeError, ValueError):
            pass
    return None


def _digits(text):
    return "".join(ch for ch in str(text or "") if ch.isdigit())


def _ddc_number(coupon_id):
    digits = _digits(coupon_id)
    return int(digits) if digits else None


# ---------------------------------------------------------------------------
# Lo que ya se sabe: EFT cargados, cupones cargados, facturas.
# ---------------------------------------------------------------------------

def _history(deposits):
    """Un renglón por EFT cargado, del más viejo al más nuevo."""
    rows = []
    for dep in deposits:
        by_coupon = {}
        for line in dep["coupons"]:
            by_coupon.setdefault(line["coupon"], []).append(line)
        main_gross = main_fees = small_net = 0.0
        for lines in by_coupon.values():
            gross = sum(l["gross_amount"] or 0.0 for l in lines)
            if gross >= MAIN_COUPON_MIN:
                main_gross += gross
                main_fees += sum(l["fees_amount"] or 0.0 for l in lines)
            else:
                small_net += sum(l["paid_amount"] or 0.0 for l in lines)
        net = sum(l["paid_amount"] or 0.0 for l in dep["coupons"])
        invoices = [(i["invoice"], i["paid_amount"] or 0.0) for i in dep["invoices"]]
        invoice_total = sum(amount for _, amount in invoices)
        rows.append({
            "date": dep["eft_date"],
            "rcv": dep["rcv_number"],
            "cut": cutoff(dep["eft_date"]),
            "coupons": set(by_coupon),
            "main_gross": main_gross,
            "main_fees": main_fees,
            "small_net": small_net,
            "net": round(net, 2),
            "invoices": invoices,
            "fuel": [amount for _, amount in invoices if amount >= FUEL_INVOICE_MIN],
            "invoice_total": round(invoice_total, 2),
            "result": round(net - invoice_total, 2),
        })
    return rows


def _rates(history):
    """Comisión de los cupones principales y neto por día de los chicos, de los EFT anteriores."""
    sample = history[-FEE_SAMPLE:]
    gross = sum(h["main_gross"] for h in sample)
    main_fee = sum(h["main_fees"] for h in sample) / gross if gross else DEFAULT_MAIN_FEE
    recent = history[-(SMALL_SAMPLE + 1):]
    span = (recent[-1]["date"] - recent[0]["date"]).days if len(recent) > 1 else 0
    small_per_day = sum(h["small_net"] for h in recent[1:]) / span if span else 0.0
    return main_fee, small_per_day


def _coupon_items(applied):
    """
    DDC cargados con número, fecha, neto, si es un cupón principal y si ya lo
    usó un EFT cargado. De un grupo sin monto propio, el de número más bajo
    toma lo que le falta al grupo y los demás quedan en 0.
    """
    rows = eft_db.list_cupones_for_forecast()
    lowest_pending = {}
    for row in rows:
        if not (row["net"] or 0.0) and row["group_net"] is not None:
            number = _ddc_number(row["coupon_id"])
            group = row["reported_group_text"]
            if number is not None and (group not in lowest_pending or number < lowest_pending[group]):
                lowest_pending[group] = number
    items = []
    for row in rows:
        number = _ddc_number(row["coupon_id"])
        if number is None:
            continue
        net, gross = row["net"] or 0.0, row["gross"] or 0.0
        if not net and row["group_net"] is not None:
            net = gross = row["group_net"] if lowest_pending.get(row["reported_group_text"]) == number else 0.0
        items.append({
            "id": row["coupon_id"],
            "number": number,
            "date": _parse(row["date"]),
            "net": round(net, 2),
            "main": gross >= MAIN_COUPON_MIN,
            "applied": row["coupon_id"] in applied,
        })
    return items


def _fuel_invoices(paid_numbers):
    """Facturas de combustible cargadas que ningún EFT cargado pagó todavía."""
    result = []
    for inv in fisico_db.list_all_invoices():
        due = _parse(inv.get("due_date"))
        if due is None or _digits(inv["invoice_number"]) in paid_numbers:
            continue
        result.append({"number": inv["invoice_number"], "delivered": _parse(inv["invoice_date"]),
                       "due": due, "eft_day": next_business_day(due), "amount": inv["amount"] or 0.0,
                       "gallons": inv.get("gallons")})
    return result


def _recurring_charges(history, today):
    """Cargos chicos que se repiten cada mes con el mismo importe: [{amount, last, next}]."""
    since = today - timedelta(days=RECURRING_LOOKBACK_DAYS)
    by_amount = {}
    for h in history:
        if h["date"] < since:
            continue
        for _, amount in h["invoices"]:
            if 0 < amount < FUEL_INVOICE_MIN:
                by_amount.setdefault(round(amount, 2), []).append(h["date"])
    charges = []
    for amount, days in sorted(by_amount.items()):
        days = sorted(set(days))
        gaps = [(b - a).days for a, b in zip(days, days[1:])]
        if len(days) < 2 or not all(RECURRING_GAP[0] <= g <= RECURRING_GAP[1] for g in gaps):
            continue
        step = round(sum(gaps) / len(gaps))
        charges.append({"amount": amount, "last": days[-1], "next": days[-1] + timedelta(days=step)})
    return charges


def _weekday_averages(card_sales, kia_sales):
    """Promedio de lo vendido con tarjeta (sin Kia/Toyota) de cada día de la semana, últimas 4 semanas cargadas."""
    by_weekday = {}
    for key in sorted(card_sales)[-AVERAGE_WEEKS * WEEK_DAYS:]:
        sale = card_sales[key] or 0.0
        by_weekday.setdefault(date.fromisoformat(key).weekday(), []).append(sale - min(kia_sales.get(key, 0.0), sale))
    values = [v for vs in by_weekday.values() for v in vs]
    fallback = sum(values) / len(values) if values else 0.0
    return {wd: (sum(by_weekday[wd]) / len(by_weekday[wd]) if by_weekday.get(wd) else fallback) for wd in range(7)}


# ---------------------------------------------------------------------------
# Un EFT.
# ---------------------------------------------------------------------------

def _sale_days(deposit_from, deposit_to):
    """Días de venta que J.H. deposita después de `deposit_from` y hasta `deposit_to` (inclusive)."""
    days = []
    day = deposit_from - timedelta(days=WEEK_DAYS)
    while day <= deposit_to:
        if deposit_from < jh_deposit_day(day) <= deposit_to:
            days.append(day)
        day += timedelta(days=1)
    return days


def _event_coupons(eft_day, covered_to, items, consumed, data_to, ctx):
    """
    Los cupones que entran en el EFT de `eft_day`: los DDC cargados (sin
    usar, hasta el último principal depositado al corte) y, para los días de
    depósito sin cupones cargados, lo vendido con tarjeta de Store Info.
    `covered_to`: último día de depósito que ya pagó el EFT anterior.
    """
    cut = cutoff(eft_day)
    mains = [i["number"] for i in items if i["main"] and i["date"] <= cut]
    limit = max(mains) if mains else None
    later = [i["number"] for i in items if i["main"] and i["date"] > cut]
    next_main = min(later) if later else float("inf")

    def enters(item):
        if item["id"] in consumed or item["applied"]:
            return False
        if item["main"]:
            return covered_to < item["date"] <= cut
        if item["date"] <= covered_to - timedelta(days=SMALL_LOOKBACK_DAYS):
            return False
        # Un chico entra si se creó antes del último principal del corte, o
        # si es de hasta 2 días antes del corte y se creó antes del principal
        # siguiente (J.H. arma el EFT entre esos dos). Verificado con los 34
        # EFT de marzo a septiembre 2026: 1 DDC de más ($9.89).
        return (limit is not None and item["number"] <= limit) or \
            (item["date"] <= cut - timedelta(days=2) and item["number"] < next_main)

    loaded = [i for i in items if enters(i)]
    store_from = max(covered_to, data_to) if data_to else covered_to
    days = _sale_days(store_from, cut) if store_from < cut else []
    gross = 0.0
    estimated = []
    for day in days:
        value = ctx["card_sales"].get(day.isoformat())
        if value is None:
            value = ctx["averages"][day.weekday()]
            estimated.append(day)
        gross += value
    loaded_net = sum(i["net"] for i in loaded)
    # Los chicos de esos días: lo que vienen dando por día de corte a corte.
    small_days = (cut - store_from).days if days else 0
    store_net = gross * (1 - ctx["main_fee"]) + ctx["small_per_day"] * small_days
    return {
        "cut": cut,
        "coupons": loaded,
        "coupons_net": round(loaded_net, 2),
        "store_days": days,
        "store_gross": round(gross, 2),
        "store_net": round(store_net, 2),
        "estimated_days": estimated,
        "net": round(loaded_net + store_net, 2),
    }


def _charges_for(eft_day, charges, taken):
    """Cargos mensuales que caen en este EFT: el primero desde 3 días antes de la fecha esperada."""
    due = [c for c in charges if id(c) not in taken and c["next"] - timedelta(days=3) <= eft_day]
    taken.update(id(c) for c in due)
    return due


def _context(history):
    card_sales = reportes_db.get_card_sales_by_date()
    main_fee, small_per_day = _rates(history)
    return {
        "card_sales": card_sales,
        "averages": _weekday_averages(card_sales, reportes_db.get_local_account_sales_by_date()),
        "main_fee": main_fee,
        "small_per_day": small_per_day,
    }


def backtest(history, count=6):
    """
    Los últimos EFT cargados calculados como si no se supieran, de dos
    formas: solo con Store Info, y con los cupones cargados hasta el corte.
    Contra lo que fue de verdad a Chase. Las tasas salen de los EFT de antes.
    """
    all_items = _coupon_items(set())
    card_sales = reportes_db.get_card_sales_by_date()
    averages = _weekday_averages(card_sales, reportes_db.get_local_account_sales_by_date())
    applied_before = set()
    rows = []
    for i in range(1, len(history)):
        prev, current = history[i - 1], history[i]
        applied_before |= prev["coupons"]
        if not current["invoice_total"]:
            continue
        main_fee, small_per_day = _rates(history[:i])
        ctx = {"card_sales": card_sales, "averages": averages, "main_fee": main_fee, "small_per_day": small_per_day}
        covered_to = cutoff(prev["date"])
        store_only = _event_coupons(current["date"], covered_to, [], set(), None, ctx)
        if store_only["estimated_days"]:
            continue  # sin Store Info de esos días no dice nada del cálculo
        # Lo que se sabría con los cupones bajados el día hábil anterior al EFT.
        known_to = business_days_before(current["date"], 1)
        known = [dict(it, applied=it["id"] in applied_before) for it in all_items if it["date"] <= known_to]
        data_to = max((it["date"] for it in known if it["main"]), default=None)
        with_coupons = _event_coupons(current["date"], covered_to, known, set(), data_to, ctx)
        rows.append({
            "date": current["date"],
            "rcv": current["rcv"],
            "real": current["result"],
            "store": round(store_only["net"] - current["invoice_total"], 2),
            "coupons": round(with_coupons["net"] - current["invoice_total"], 2),
        })
    return rows[-count:]


def build_forecast(today=None):
    """
    Los EFT que vienen desde el último cargado: fecha, facturas, cupones que
    entran y lo que daría en Chase (positivo = crédito, negativo = débito).
    None si todavía no hay ningún EFT cargado.
    """
    today = today or date.today()
    history = _history(eft_db.list_deposits_with_lines())
    if not history:
        return None
    ctx = _context(history)
    last = history[-1]
    applied = set().union(*(h["coupons"] for h in history))
    items = _coupon_items(applied)
    # Hasta dónde están cargados los cupones: el último día con un cupón
    # principal (un día con solo chicos cargados todavía no está completo).
    data_to = max((i["date"] for i in items if i["main"]), default=None)
    paid = {_digits(number) for h in history for number, _ in h["invoices"]}
    loaded_fuel = _fuel_invoices(paid)
    fuel = [f for f in loaded_fuel if f["eft_day"] > last["date"]]
    # Una factura sin cargar se estima con la última conocida: el precio del
    # combustible cambia de semana a semana y con junio-septiembre 2026 la
    # última erró en promedio $1,069 (el promedio de 4, $1,719).
    known = [(h["date"], h["fuel"][-1]) for h in history if h["fuel"]] +             [(f["eft_day"], f["amount"]) for f in loaded_fuel]
    estimated_invoice = round(max(known)[1], 2) if known else 0.0
    charges = _recurring_charges(history, today)
    chase_last = chase_db.get_last_posting_date()
    chase_moves = chase_db.get_detalle_transactions(EFT_DETALLE, last["date"] + timedelta(days=1),
                                                    today + timedelta(days=HORIZON_DAYS)) if chase_last else []

    events = []
    used_moves, consumed, taken = [], set(), set()
    covered_to = cutoff(last["date"])
    cursor = last["date"]
    pending = sorted(fuel, key=lambda f: f["eft_day"])
    while len(events) < MAX_EVENTS:
        guess = next_business_day(cursor + timedelta(days=WEEK_DAYS))
        if pending and pending[0]["eft_day"] <= guess + timedelta(days=3):
            day = pending[0]["eft_day"]
            invoices = [f for f in pending if f["eft_day"] == day]
            pending = [f for f in pending if f["eft_day"] != day]
        else:
            day, invoices = guess, []
        if day > today + timedelta(days=HORIZON_DAYS) and any(e["date"] >= today for e in events):
            break
        # El EFT en Chase: el mismo día (o el siguiente) si la fecha sale de
        # una factura; si es estimada, el más cercano en ±3 días.
        window = 1 if invoices else 3
        near = [m for m in chase_moves if m not in used_moves and -window <= (m[0] - day).days <= window]
        move = min(near, key=lambda m: abs((m[0] - day).days)) if near else None
        if move is not None:
            used_moves.append(move)
            day = move[0]
        coupons = _event_coupons(day, covered_to, items, consumed, data_to, ctx)
        extra = _charges_for(day, charges, taken)
        fuel_total = sum(f["amount"] for f in invoices) if invoices else estimated_invoice
        invoice_total = round(fuel_total + sum(c["amount"] for c in extra), 2)
        if move is None and day < today:
            status = "overdue" if chase_last and chase_last >= day else "unknown"
        else:
            status = "arrived" if move is not None else "upcoming"
        events.append(dict(
            coupons, date=day, invoices=invoices, invoice_estimated=not invoices, fuel_total=round(fuel_total, 2),
            charges=extra, invoice_total=invoice_total, result=round(coupons["net"] - invoice_total, 2),
            arrived=move[1] if move else None, status=status,
        ))
        consumed |= {i["id"] for i in coupons["coupons"]}
        covered_to = coupons["cut"]
        cursor = day

    upcoming = [e for e in events if e["date"] >= today]
    tests = backtest(history)
    store_errors = [abs(t["store"] - t["real"]) for t in tests]
    return {
        "today": today,
        "last": last,
        "events": events,
        # Cuánto se suele errar con solo Store Info (los últimos EFT).
        "typical_error": round(sum(store_errors) / len(store_errors), 2) if store_errors else None,
        "next": upcoming[0] if upcoming else None,
        "main_fee": ctx["main_fee"],
        "small_per_day": ctx["small_per_day"],
        "estimated_invoice": estimated_invoice,
        "coupons_to": data_to,
        "store_info_last": date.fromisoformat(max(ctx["card_sales"])) if ctx["card_sales"] else None,
        "chase_last": chase_last,
        "estimated_past": sorted({d for e in events for d in e["estimated_days"] if d < today}),
        "backtest": tests,
    }
