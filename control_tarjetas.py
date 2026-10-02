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

No se compara día contra día: cada DDC es un lote de tarjetas que no siempre
cierra con el día del POS (solo 12 de 45 días coinciden al centavo). En
agosto-septiembre 2026 el pendiente estuvo entre -$3,000 y +$15,000, con los
picos en los fines de semana (el lunes llegan juntos los cupones de jueves a
sábado).
"""

import calendar
from datetime import date, timedelta

ALERT_THRESHOLD = 15000.0
COUPON_START_LAG_DAYS = 2
TRANSIT_DAYS = 3


def _running(card_sales, coupons, until):
    """{fecha: fila} desde el primer día con ventas con tarjeta hasta `until`, con el pendiente acumulado."""
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
            row["status"] = "alert" if abs(pending) > ALERT_THRESHOLD else "ok"
        rows[key] = row
        day += timedelta(days=1)
    return rows


def build_month_control(year, month, card_sales, coupons, today=None):
    """
    Control de un mes: una fila por día (hasta hoy) y el resumen del último
    día con Store Info. `card_sales` y `coupons` vienen de
    reportes_db.get_card_sales_by_date y eft_db.get_coupon_gross_by_date.
    """
    today = today or date.today()
    month_end = date(year, month, calendar.monthrange(year, month)[1])
    until = min(month_end, today)
    running = _running(card_sales, coupons, until) if until >= date(year, month, 1) else {}
    rows = [row for key, row in sorted(running.items()) if key[:7] == f"{year:04d}-{month:02d}"]
    evaluated = [row for row in rows if row["pending"] is not None]
    last = evaluated[-1] if evaluated else None
    return {
        "rows": rows,
        "last": last,
        "alert_days": [row for row in evaluated if row["status"] == "alert"],
        "max_pending": max((row["pending"] for row in evaluated), default=None),
        "missing_days": [row["date"] for row in rows if row["status"] == "missing"],
        "unknown_coupons": sum(row["coupon_unknown"] for row in rows),
        "sales_total": round(sum(row["sale"] or 0.0 for row in rows), 2),
        "coupon_total": round(sum(row["coupon_gross"] for row in rows), 2),
        "threshold": ALERT_THRESHOLD,
    }


def latest_status(card_sales, coupons, today=None):
    """Fila del último día con Store Info (con su pendiente y estado), o None si no hay datos."""
    if not card_sales:
        return None
    last_day = date.fromisoformat(max(card_sales))
    return _running(card_sales, coupons, last_day).get(last_day.isoformat())
