"""
Controles rápidos (pedido del jefe, 2026-10-05): la botonera chica de abajo a
la derecha (base.html) muestra en un cuadro cómo viene cada cosa, sin entrar a
cada control. Acá solo el cálculo, con los datos ya cargados; la pantalla
completa de cada uno sigue en Controles.

Arranca en el MES ACTUAL (pedido del usuario, el mismo día) y desde el
2026-10-07 se puede pasar a otro mes desde el propio cuadro (el cálculo se
hace "al último día" de ese mes). Si el mes todavía no tiene nada cargado, el
cuadro lo dice en vez de mostrar el último mes que sí tiene. Precios a
revisar salió de la botonera ese día ("de momento", pedido del usuario): el
cálculo queda acá y en /controles/rapidos/precios.

- Tarjetas: lo vendido con tarjeta que todavía no acreditó JH, al último día
  del mes con Store Info (el mismo cálculo y la misma alerta de
  control_tarjetas).
- Caja: el efectivo que debería haber en la estación (el Saldo de Caja) al
  último día del mes con Reporte Diario, los días de antes para ver si se
  viene depositando, y desde cuándo falta cargar. Alerta (pedido del
  usuario, 2026-10-07) si el Saldo da negativo o pasa del límite, que se
  cambia desde el propio cuadro.
- Precios: los productos de la última factura del mes de cada proveedor cuyo
  UPC está en el CMV (el Costo que se exporta de eListar) con otro costo: son
  los que hay que revisar de precio. Solo cruza por UPC (sin los ceros de
  adelante, igual que Productos de proveedores), así que J.J. Taylor y
  Midtown, que no traen UPC, no entran.
"""

import calendar
import json
import os
from datetime import date

import caja
import control_tarjetas
from proveedores_productos import _pct, normalize_upc

_RECENT_DAYS = 7  # cuántos días se muestran de Tarjetas y de Caja

# Un costo por unidad que da más del doble o menos de la mitad del POS casi
# siempre es un pack mal leído (ej. los cigarros que H.T. factura por caja y
# el POS vende de a uno), no un aumento: se marca para revisar el pack.
_PACK_SUSPECT_RATIO = 2.0


# Límite de la alerta de Caja: se cambia desde el cuadro y se guarda en
# reportes_data (gitignored, como las bases).
_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data", "controles_rapidos.json")
DEFAULT_CAJA_LIMIT = 10500.0  # lo más alto que se ve un fin de semana normal


def caja_limit():
    try:
        with open(_CONFIG_PATH, encoding="utf-8") as handle:
            return float(json.load(handle)["caja_limit"])
    except (OSError, ValueError, KeyError, TypeError):
        return DEFAULT_CAJA_LIMIT


def set_caja_limit(value):
    if value is None or value <= 0:
        raise ValueError("El límite tiene que ser un monto mayor que cero.")
    os.makedirs(os.path.dirname(_CONFIG_PATH), exist_ok=True)
    try:
        with open(_CONFIG_PATH, encoding="utf-8") as handle:
            config = json.load(handle)
    except (OSError, ValueError):
        config = {}
    config["caja_limit"] = round(float(value), 2)
    with open(_CONFIG_PATH, "w", encoding="utf-8") as handle:
        json.dump(config, handle, indent=2)


def caja_alert(saldo, limit):
    """'negative' / 'high' / None."""
    if saldo is None:
        return None
    if saldo < -0.005:
        return "negative"
    return "high" if saldo > limit else None


def reference_day(year, month, today=None):
    """
    El "hoy" con el que se calcula un mes elegido en el cuadro: hoy si es el
    mes actual, el último día si es un mes anterior. None si es un mes futuro.
    """
    today = today or date.today()
    if (year, month) == (today.year, today.month):
        return today
    if (year, month) > (today.year, today.month):
        return None
    return date(year, month, calendar.monthrange(year, month)[1])


def _month_prefix(today):
    return f"{today.year:04d}-{today.month:02d}-"


def tarjetas_status(card_sales, coupons, today=None):
    """
    Pendiente de acreditar al último día del mes actual con Store Info y los
    días de antes de ese mes; None si el mes todavía no tiene ventas con
    tarjeta cargadas.
    """
    today = today or date.today()
    month_sales = {day: sale for day, sale in card_sales.items() if day.startswith(_month_prefix(today))}
    if not month_sales:
        return None
    # El pendiente es un acumulado corrido: se calcula con todo lo cargado
    # (puede arrancar a fin del mes anterior) y se muestra solo este mes.
    month = control_tarjetas.build_month_control(today.year, today.month, card_sales, coupons, today=today)
    last_date = max(month_sales)
    rows = [row for row in month["rows"] if row["date"] <= last_date]
    last = next((row for row in reversed(rows) if row["pending"] is not None), None)
    if last is None:
        return None
    return {
        "last": last,
        "recent": rows[-_RECENT_DAYS:],
        "threshold": control_tarjetas.ALERT_THRESHOLD,
        "days_behind": (today - date.fromisoformat(last["date"])).days,
    }


def caja_status(today=None, build_month=None):
    """
    Saldo de Caja al último día del mes actual con Reporte Diario (Store Info)
    y los días de antes, para ver si el efectivo se viene depositando o se
    junta. None si el mes todavía no tiene ningún día cargado.
    """
    today = today or date.today()
    build_month = build_month or caja.build_month_report_from_db
    rows = [row for row in build_month(today.year, today.month)["rows"] if row["date"] <= today.isoformat()]
    loaded = [row for row in rows if row["cash"] is not None]
    if not loaded:
        return None
    last = loaded[-1]
    deposits = [row for row in rows if row["deposit"]]
    upto_last = [row for row in rows if row["date"] <= last["date"]]
    limit = caja_limit()
    return {
        "date": last["date"],
        "saldo": last["saldo"],
        "limit": limit,
        "alert": caja_alert(last["saldo"], limit),
        "alert_days": {row["date"]: caja_alert(row["saldo"], limit) for row in upto_last},
        "days_behind": (today - date.fromisoformat(last["date"])).days,
        "recent": upto_last[-_RECENT_DAYS:],
        "last_deposit": {"date": deposits[-1]["date"], "amount": deposits[-1]["deposit"]} if deposits else None,
    }


def _last_invoice_lines(supplier_lines):
    """Renglones de la última factura del proveedor (por fecha y N°, igual que Productos de proveedores)."""
    last = max((line["invoice_date"], str(line["invoice_no"])) for line in supplier_lines)
    return last, [line for line in supplier_lines if (line["invoice_date"], str(line["invoice_no"])) == last]


def price_review(lines, pos_costs, supplier_labels, today=None):
    """
    Por proveedor, los productos de su última factura del mes actual con el
    mismo UPC en el CMV y otro costo por unidad: primero los que más
    subieron. Un proveedor sin facturas este mes no aparece. Medio centavo o
    menos es redondeo del POS y no cuenta.
    """
    today = today or date.today()
    pos_by_upc = {normalize_upc(row.get("upc")): row for row in pos_costs if normalize_upc(row.get("upc"))}
    by_supplier = {}
    for line in lines:
        if line["invoice_date"].startswith(_month_prefix(today)):
            by_supplier.setdefault(line["supplier_key"], []).append(line)

    suppliers = []
    for key, supplier_lines in by_supplier.items():
        (invoice_date, invoice_no), invoice_lines = _last_invoice_lines(supplier_lines)
        by_upc = {}
        for line in invoice_lines:
            upc = normalize_upc(line["upc"])
            if upc in pos_by_upc:
                by_upc[upc] = line  # el mismo producto dos veces en la factura: vale el último renglón
        rows = []
        for upc, line in by_upc.items():
            pos = pos_by_upc[upc]
            pos_cost = pos.get("cost")
            if pos_cost is None:
                continue
            diff = round(line["unit_cost"] - pos_cost, 4)
            if abs(diff) <= 0.0051:
                continue
            pos_price = pos.get("price")
            ratio = line["unit_cost"] / pos_cost if pos_cost else None
            rows.append({
                "upc": upc,
                "description": line["description"] or pos.get("name"),
                "pos_name": pos.get("name"),
                "unit_cost": line["unit_cost"],
                "pos_cost": pos_cost,
                "diff": diff,
                "diff_pct": _pct(diff, pos_cost),
                "pos_price": pos_price,
                "margin_pct": _pct(pos_price - line["unit_cost"], pos_price) if pos_price else None,
                "pack_suspect": ratio is not None and not (1 / _PACK_SUSPECT_RATIO <= ratio <= _PACK_SUSPECT_RATIO),
            })
        rows.sort(key=lambda row: (row["pack_suspect"], -(row["diff_pct"] or 0)))
        suppliers.append({
            "key": key,
            "label": supplier_labels.get(key, key),
            "invoice_date": invoice_date,
            "invoice_no": invoice_no,
            "lines": len(invoice_lines),
            "matched": len(by_upc),
            "rows": rows,
        })
    suppliers.sort(key=lambda s: s["label"].lower())
    return suppliers
