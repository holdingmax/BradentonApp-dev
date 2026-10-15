"""
Control Tarjetas y Cupones (pedido del usuario, 2026-10-02): lo que se cobra
con tarjeta en el C-store es lo que JH después nos acredita como cupones
(DDC) y usa para pagar sus facturas en los EFT. La acreditación tarda unas
72 hs, así que en cualquier momento lo vendido con tarjeta que todavía no
llegó como cupón ("pendiente de acreditar") son, normalmente, los últimos ~3
días: hasta unos $15,000. Más que eso es una alerta.

Cálculo, sin UI (la pantalla es /controles/tarjetas en webapp.py):
- Ventas con tarjeta = suma de los términos de crédito de Store Info (la
  columna TC de Caja), día por día.
- Cupones = Gross de los DDC por su fecha.
- Pendiente acumulado = ventas con tarjeta desde el inicio del tramo menos
  los batches del POS de esas ventas (con detalle de cupones, ver abajo) o,
  sin detalle, los cupones que llegaron desde 2 días después de ese inicio
  (los de los 2 primeros días pagan ventas de antes del tramo). Un tramo son días seguidos
  con Store Info: si falta un día, el pendiente vuelve a arrancar en el
  siguiente día cargado, porque si no quedaría corrido para siempre en lo que
  se vendió ese día.

Pagos de Kia y Toyota (pedido del usuario, 2026-10-09): pagan los vales con
una Amex (LOCAL ACCT = VS en Ventas por Departamento) y ese día las ventas
con tarjeta suben en ese monto (01/09/2026: $16,901.39 con $11,455.29 de
Kia). Un pago grande hacía pasar el límite sin nada raro: la alerta mira el
pendiente sin los pagos que J.H. todavía no depositó, que se muestran aparte
(desde 2026-10-13, hasta el día exacto en que deposita, ver jh_deposit_day;
antes eran los últimos 3 días fijos y un pago de jueves volvía a contar el
domingo: la alerta del 30/08/2026 era ese pago; desde 2026-10-15 la alerta
mira lo atrasado, que deja afuera todo lo que J.H. todavía no tenía que
depositar, Kia/Toyota incluidos: ver abajo). Un
día con VS mayor que todo lo vendido con tarjeta (07/08/2026) es un pago que
no entró como tarjeta: se avisa (kia_not_in_cards).

No se compara día contra día: cada DDC es un lote de tarjetas que no siempre
cierra con el día del POS (solo 12 de 45 días coinciden al centavo).

Con el detalle de cupones (pedido del usuario, 2026-10-15): de cada depósito
se restan solo los batches del POS (in_store_info) de ventas desde el inicio
del tramo. Antes se restaba el depósito entero y el pendiente se corría en
negativo: los batches "slri…"/7 dígitos (unos $70 por día, $5,351.59 del
01/08 al 13/10/2026) no son ventas de Store Info, y el primer depósito del
tramo del 01/08 pagaba también las ventas del 30 y 31/07 ($10,098.39). Al
13/10 daba −$5,387.37 y ahora $10,003.81 (lo vendido el 12 y 13/10,
$10,277.28, menos $273.47 de batches cruzados de día y de dos DDC sin
detalle). Los DDC sin detalle se restan enteros, como antes.

Con el pendiente real, los domingos y los fines de semana largos son 4 o 5
días de ventas ($20,000 a $30,000) y el límite ya no se puede mirar sobre el
pendiente: se mira sobre lo atrasado, el pendiente menos lo que J.H. todavía
tiene tiempo de depositar (las ventas cuyo jh_deposit_day no llegó, "en
camino"; incluye los pagos de Kia/Toyota). Ago-oct 2026: entre −$362 y
+$13,014.80 (30/09: J.H. depositó el 01/10 las ventas del 28 y 29/09).
"""

import calendar
import re
from datetime import date, timedelta

ALERT_THRESHOLD = 15000.0
COUPON_START_LAG_DAYS = 2
JH_DEPOSIT_LAG_DAYS = 2
# Hasta cuántos días atrás puede haber ventas que J.H. todavía no tenía que
# depositar (jueves a lunes feriado son 5).
TRANSIT_LOOKBACK_DAYS = 10


def _nth_weekday(year, month, weekday, n):
    """El n-ésimo `weekday` (0 = lunes) del mes; n = -1 es el último."""
    if n > 0:
        first = date(year, month, 1)
        return first + timedelta(days=(weekday - first.weekday()) % 7 + 7 * (n - 1))
    last = date(year, month, calendar.monthrange(year, month)[1])
    return last - timedelta(days=(last.weekday() - weekday) % 7)


def bank_holidays(year):
    """
    Feriados de la Reserva Federal (sin ACH: J.H. no deposita). Uno que cae
    domingo se pasa al lunes; uno que cae sábado no cierra (la Fed abre el
    viernes).
    """
    fixed = [date(year, 1, 1), date(year, 6, 19), date(year, 7, 4), date(year, 11, 11), date(year, 12, 25)]
    days = {d + timedelta(days=1) if d.weekday() == 6 else d for d in fixed if d.weekday() != 5}
    days.update({
        _nth_weekday(year, 1, 0, 3),   # Martin Luther King
        _nth_weekday(year, 2, 0, 3),   # Presidents' Day
        _nth_weekday(year, 5, 0, -1),  # Memorial Day
        _nth_weekday(year, 9, 0, 1),   # Labor Day
        _nth_weekday(year, 10, 0, 2),  # Columbus Day
        _nth_weekday(year, 11, 3, 4),  # Thanksgiving
    })
    return days


def jh_deposit_day(sale_day):
    """
    Día en que J.H. deposita lo vendido con tarjeta el `sale_day`: dos días
    después, corrido al día hábil siguiente si cae sábado, domingo o feriado
    bancario. Medido con el detalle de cupones de septiembre 2026: lunes a
    miércoles a los 2 días; jueves, viernes y sábado juntos el lunes; domingo
    el martes; con el Labor Day (07/09) jueves a sábado pasaron al martes.
    """
    day = sale_day + timedelta(days=JH_DEPOSIT_LAG_DAYS)
    while day.weekday() >= 5 or day in bank_holidays(day.year):
        day += timedelta(days=1)
    return day


def in_store_info(batch):
    """Batch del POS (número de 4 dígitos, 0xxx/82xx/92xx): está en lo vendido con tarjeta de Store Info."""
    return bool(re.fullmatch(r"\d{4}", batch or ""))


def _running(card_sales, coupons, until, kia_sales=None):
    """{fecha: fila} desde el primer día con ventas con tarjeta hasta `until`, con el pendiente acumulado."""
    kia_sales = kia_sales or {}
    if not card_sales:
        return {}
    day = date.fromisoformat(min(card_sales))
    last_loaded = date.fromisoformat(max(card_sales))
    rows = {}
    pending = None
    segment_start = None
    while day <= until:
        key = day.isoformat()
        sale = card_sales.get(key)
        coupon = coupons.get(key) or {"gross": 0.0, "count": 0, "unknown": 0}
        batches = coupon.get("batches") or []
        pos_batches = [b for b in batches if in_store_info(b["batch"])]
        row = {
            "date": key,
            "sale": sale,
            "coupon_gross": coupon["gross"],
            "coupon_count": coupon["count"],
            "coupon_unknown": coupon["unknown"],
            # Batches del depósito que no son ventas de Store Info: no se restan.
            "coupon_other": round(sum(b["gross"] for b in batches if not in_store_info(b["batch"])), 2),
            # Lo del depósito que paga ventas de antes del tramo: no se resta.
            "coupon_previous": 0.0,
            "pending": None,
            # Un pago que no entró como tarjeta (kia_not_in_cards) cuenta
            # solo hasta lo vendido con tarjeta ese día.
            "kia": min(kia_sales.get(key, 0.0), sale) if sale is not None else 0.0,
            "kia_recent": 0.0,
            "in_transit": None,
            "overdue": None,
        }
        if sale is None:
            pending = segment_start = None
            row["status"] = "later" if day > last_loaded else "missing"
        else:
            if segment_start is None:
                segment_start, pending = day, 0.0
                row["segment_start"] = True
            start = segment_start.isoformat()
            credited = sum(b["gross"] for b in pos_batches if b["batch_date"] >= start)
            previous = sum(b["gross"] for b in pos_batches if b["batch_date"] < start)
            # Sin detalle no se sabe de qué días son: los DDC de los primeros
            # días del tramo pagan ventas de antes.
            no_detail = coupon.get("no_detail", coupon["gross"])
            if day >= segment_start + timedelta(days=COUPON_START_LAG_DAYS):
                credited += no_detail
            else:
                previous += no_detail
            pending = round(pending + sale - credited, 2)
            row["pending"] = pending
            row["coupon_previous"] = round(previous, 2)
            # Lo vendido que J.H. todavía no tenía que depositar (los días
            # del tramo seguidos tienen todos Store Info).
            recent = [day - timedelta(days=k) for k in range(TRANSIT_LOOKBACK_DAYS)]
            recent = [d for d in recent if d >= segment_start and jh_deposit_day(d) > day]
            row["in_transit"] = round(sum(card_sales[d.isoformat()] for d in recent), 2)
            row["overdue"] = round(pending - row["in_transit"], 2)
            # De lo en camino, los pagos de Kia/Toyota (nunca más que el pendiente).
            kia_recent = sum(rows[d.isoformat()]["kia"] for d in recent if d != day) + row["kia"]
            row["kia_recent"] = round(min(kia_recent, max(pending, 0.0)), 2)
            row["status"] = "alert" if abs(row["overdue"]) > ALERT_THRESHOLD else "ok"
        rows[key] = row
        day += timedelta(days=1)
    return rows


def kia_not_in_cards(card_sales, kia_sales, days):
    """Días (de `days`) con un pago de Kia/Toyota mayor que todo lo vendido con tarjeta: no entró como tarjeta."""
    return [{"date": d, "kia": kia_sales[d], "sale": card_sales[d]} for d in days
            if kia_sales.get(d) and card_sales.get(d) is not None and card_sales[d] < kia_sales[d] - 0.005]


def build_month_control(year, month, card_sales, coupons, today=None, kia_sales=None):
    """
    Control de un mes: una fila por día (hasta hoy) y el resumen del último
    día con Store Info. `card_sales` y `coupons` vienen de
    reportes_db.get_card_sales_by_date y eft_db.get_coupon_gross_by_date.
    """
    today = today or date.today()
    month_end = date(year, month, calendar.monthrange(year, month)[1])
    until = min(month_end, today)
    running = _running(card_sales, coupons, until, kia_sales) if until >= date(year, month, 1) else {}
    rows = [row for key, row in sorted(running.items()) if key[:7] == f"{year:04d}-{month:02d}"]
    evaluated = [row for row in rows if row["pending"] is not None]
    last = evaluated[-1] if evaluated else None
    return {
        "rows": rows,
        "last": last,
        "alert_days": [row for row in evaluated if row["status"] == "alert"],
        # Lo que mira la alerta: lo atrasado.
        "max_overdue": max((row["overdue"] for row in evaluated), default=None),
        "kia_total": round(sum(row["kia"] for row in rows), 2),
        "kia_not_in_cards": kia_not_in_cards(card_sales, kia_sales or {}, [row["date"] for row in rows]),
        "missing_days": [row["date"] for row in rows if row["status"] == "missing"],
        "unknown_coupons": sum(row["coupon_unknown"] for row in rows),
        "sales_total": round(sum(row["sale"] or 0.0 for row in rows), 2),
        "coupon_total": round(sum(row["coupon_gross"] for row in rows), 2),
        "coupon_other_total": round(sum(row["coupon_other"] for row in rows), 2),
        "threshold": ALERT_THRESHOLD,
    }


def latest_status(card_sales, coupons, today=None, kia_sales=None):
    """Fila del último día con Store Info (con su pendiente y estado), o None si no hay datos."""
    if not card_sales:
        return None
    last_day = date.fromisoformat(max(card_sales))
    return _running(card_sales, coupons, last_day, kia_sales).get(last_day.isoformat())


# ---------------------------------------------------------------------------
# Por día de venta, con el detalle de cupones (pedido del usuario, 2026-10-06,
# chat 21; lectura en cupones_detalle.py). Validado con agosto-septiembre
# 2026: los batches del POS (número de 4 dígitos) de un día suman lo vendido
# con tarjeta de Store Info ese día, salvo un batch que cruza la medianoche y
# deja la misma diferencia, al revés, en el día de al lado. Los demás batches
# ("slri…", de 7 dígitos, entre $5 y $150 por día) no están en Store Info:
# se muestran aparte.
# ---------------------------------------------------------------------------

# Un día cuyos batches pueden no haberse depositado todavía: los cupones de
# un día llegan en depósitos de hasta 4 días después (fin de semana).
DETAIL_SETTLE_DAYS = 4
# Días siguientes con los que una diferencia puede compensarse (en
# septiembre 2026, una del 07/09 se compensó recién el 10/09).
COMPENSATE_DAYS = 3


def build_detail_by_day(year, month, card_sales, batches, covered_from, covered_to, today=None, kia_sales=None):
    """
    Una fila por día del mes (hasta hoy): vendido con tarjeta (Store Info)
    contra los batches del POS de ese día, y los otros batches aparte.
    `batches`: eft_db.get_detail_batches_between, del mes con unos días de
    margen; `covered_from`/`covered_to`: el primer y el último depósito con
    detalle cargado (un día antes del primero no tiene todos sus batches).

    Una diferencia que vuelve a cero sumándola con los días siguientes (el
    batch que cruza la medianoche) queda "compensated"; los últimos días,
    cuyos cupones pueden no haberse depositado todavía, "pending"; el resto,
    "diff". Devuelve None si no hay detalle cargado que toque el mes.
    """
    if not covered_from or not covered_to:
        return None
    today = today or date.today()
    first = date(year, month, 1)
    month_end = date(year, month, calendar.monthrange(year, month)[1])
    covered_from, covered_to = date.fromisoformat(covered_from), date.fromisoformat(covered_to)
    if covered_to < first or covered_from > month_end:
        return None
    pos, other = {}, {}
    for b in batches:
        target = pos if in_store_info(b["batch"]) else other
        target[b["batch_date"]] = round(target.get(b["batch_date"], 0.0) + (b["gross"] or 0.0), 2)

    window_start = max(covered_from, first - timedelta(days=DETAIL_SETTLE_DAYS))
    window_end = min(covered_to, month_end + timedelta(days=DETAIL_SETTLE_DAYS))
    rows = []
    day = window_start
    while day <= window_end:
        key = day.isoformat()
        sold = card_sales.get(key)
        row = {"date": key, "sold": sold, "pos": pos.get(key, 0.0), "other": other.get(key, 0.0),
               "kia": (kia_sales or {}).get(key, 0.0) if sold is not None else 0.0,
               "diff": None, "status": "no_sales" if sold is None else None}
        if sold is not None:
            row["diff"] = round(row["pos"] - sold, 2)
            if abs(row["diff"]) < 0.005:
                row["status"] = "ok"
        rows.append(row)
        day += timedelta(days=1)

    # Una diferencia se compensa si suma cero con los días siguientes (hasta
    # COMPENSATE_DAYS): el batch que cruzó la medianoche. Si no, y el día de
    # al lado no se puede comparar (sin Store Info o antes del detalle), no se
    # puede saber ("edge"); los últimos días con menos batches que ventas
    # todavía no se depositaron enteros ("pending").
    recent = covered_to - timedelta(days=DETAIL_SETTLE_DAYS)
    for i, row in enumerate(rows):
        if row["status"] is not None:
            continue
        total = 0.0
        for j in range(i, min(i + COMPENSATE_DAYS + 1, len(rows))):
            if rows[j]["diff"] is None:
                break
            total = round(total + rows[j]["diff"], 2)
            if j > i and abs(total) < 0.005:
                for r in rows[i:j + 1]:
                    if r["status"] is None:
                        r["status"] = "compensated"
                break
        if row["status"] is not None:
            continue
        before = rows[i - 1]["diff"] if i > 0 else None
        after = rows[i + 1]["diff"] if i + 1 < len(rows) else None
        if date.fromisoformat(row["date"]) > recent and row["diff"] < 0:
            row["status"] = "pending"
        elif before is None or after is None:
            row["status"] = "edge"
        else:
            row["status"] = "diff"

    until = min(month_end, today, covered_to)
    shown = [r for r in rows if first.isoformat() <= r["date"] <= until.isoformat()]
    if covered_from > first:
        missing_from = first.isoformat()
    else:
        missing_from = None
    with_sales = [r for r in shown if r["sold"] is not None]
    return {
        "rows": shown,
        "covered_from": covered_from.isoformat(),
        "covered_to": covered_to.isoformat(),
        "missing_before": missing_from,
        "sold_total": round(sum(r["sold"] for r in with_sales), 2),
        "pos_total": round(sum(r["pos"] for r in with_sales), 2),
        "other_total": round(sum(r["other"] for r in shown), 2),
        "diff_total": round(sum(r["diff"] for r in with_sales), 2),
        "bad": [r for r in shown if r["status"] == "diff"],
        "pending": [r for r in shown if r["status"] == "pending"],
        "edge": [r for r in shown if r["status"] == "edge"],
        "with_sales": len(with_sales),
        "ok": not any(r["status"] == "diff" for r in shown),
    }
