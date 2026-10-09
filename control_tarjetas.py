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
  los cupones que llegaron desde 2 días después de ese inicio (los de los 2
  primeros días pagan ventas de antes del tramo). Un tramo son días seguidos
  con Store Info: si falta un día, el pendiente vuelve a arrancar en el
  siguiente día cargado, porque si no quedaría corrido para siempre en lo que
  se vendió ese día.

Pagos de Kia y Toyota (pedido del usuario, 2026-10-09): pagan los vales con
una Amex (LOCAL ACCT = VS en Ventas por Departamento) y ese día las ventas
con tarjeta suben en ese monto (01/09/2026: $16,901.39 con $11,455.29 de
Kia). Un pago grande hacía pasar el límite sin nada raro: la alerta mira el
pendiente sin los pagos de los últimos 3 días, que se muestran aparte. Un
día con VS mayor que todo lo vendido con tarjeta (07/08/2026) es un pago que
no entró como tarjeta: se avisa (kia_not_in_cards).

No se compara día contra día: cada DDC es un lote de tarjetas que no siempre
cierra con el día del POS (solo 12 de 45 días coinciden al centavo). En
agosto-septiembre 2026 el pendiente estuvo entre -$3,000 y +$15,000, con los
picos en los fines de semana (el lunes llegan juntos los cupones de jueves a
sábado).
"""

import calendar
import re
from datetime import date, timedelta

ALERT_THRESHOLD = 15000.0
COUPON_START_LAG_DAYS = 2
TRANSIT_DAYS = 3


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
        last3 = [card_sales.get((day - timedelta(days=k)).isoformat()) for k in range(TRANSIT_DAYS)]
        row = {
            "date": key,
            "sale": sale,
            "coupon_gross": coupon["gross"],
            "coupon_count": coupon["count"],
            "coupon_unknown": coupon["unknown"],
            "last3": round(sum(v for v in last3 if v is not None), 2),
            "pending": None,
            "counted": False,
            "kia": kia_sales.get(key, 0.0) if sale is not None else 0.0,
            "kia_recent": 0.0,
            "pending_without_kia": None,
        }
        if sale is None:
            pending = segment_start = None
            row["status"] = "later" if day > last_loaded else "missing"
        else:
            if segment_start is None:
                segment_start, pending = day, 0.0
                row["segment_start"] = True
            pending += sale
            if day >= segment_start + timedelta(days=COUPON_START_LAG_DAYS):
                pending -= coupon["gross"]
                row["counted"] = True
            pending = round(pending, 2)
            row["pending"] = pending
            # Lo pendiente que son pagos de Kia/Toyota de los últimos días
            # (nunca más que el propio pendiente).
            kia_recent = sum(kia_sales.get((day - timedelta(days=k)).isoformat(), 0.0)
                             for k in range(TRANSIT_DAYS) if (day - timedelta(days=k)) >= segment_start)
            row["kia_recent"] = round(min(kia_recent, max(pending, 0.0)), 2)
            row["pending_without_kia"] = round(pending - row["kia_recent"], 2)
            row["status"] = "alert" if abs(row["pending_without_kia"]) > ALERT_THRESHOLD else "ok"
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
        "max_pending": max((row["pending"] for row in evaluated), default=None),
        "kia_total": round(sum(row["kia"] for row in rows), 2),
        "kia_not_in_cards": kia_not_in_cards(card_sales, kia_sales or {}, [row["date"] for row in rows]),
        "missing_days": [row["date"] for row in rows if row["status"] == "missing"],
        "unknown_coupons": sum(row["coupon_unknown"] for row in rows),
        "sales_total": round(sum(row["sale"] or 0.0 for row in rows), 2),
        "coupon_total": round(sum(row["coupon_gross"] for row in rows), 2),
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


def in_store_info(batch):
    """Batch del POS (número de 4 dígitos, 0xxx/82xx/92xx): está en lo vendido con tarjeta de Store Info."""
    return bool(re.fullmatch(r"\d{4}", batch or ""))


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
