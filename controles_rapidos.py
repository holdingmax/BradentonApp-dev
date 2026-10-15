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
- Portales (2026-10-15): cuándo entrar a FPL, Manatee, Airgas y Elistar a
  bajar lo que no llega solo, y si ya está cargado. No depende del mes.
"""

import calendar
import json
import math
import app_paths
import os
from datetime import date, timedelta

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
_CONFIG_PATH = os.path.join(app_paths.DATA_DIR, "controles_rapidos.json")
DEFAULT_CAJA_LIMIT = 10500.0  # lo más alto que se ve un fin de semana normal


def caja_limit():
    try:
        with open(_CONFIG_PATH, encoding="utf-8") as handle:
            value = float(json.load(handle)["caja_limit"])
    except (OSError, ValueError, KeyError, TypeError):
        return DEFAULT_CAJA_LIMIT
    return value if math.isfinite(value) and value > 0 else DEFAULT_CAJA_LIMIT


def set_caja_limit(value):
    if value is None or not math.isfinite(value) or value <= 0:
        raise ValueError("El límite tiene que ser un monto mayor que cero.")
    os.makedirs(os.path.dirname(_CONFIG_PATH), exist_ok=True)
    try:
        with open(_CONFIG_PATH, encoding="utf-8") as handle:
            config = json.load(handle)
    except (OSError, ValueError):
        config = {}
    config["caja_limit"] = round(float(value), 2)
    # Se escribe aparte y se reemplaza: un corte a mitad no deja el archivo roto.
    temp_path = _CONFIG_PATH + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as handle:
        json.dump(config, handle, indent=2)
    os.replace(temp_path, _CONFIG_PATH)


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


def tarjetas_status(card_sales, coupons, today=None, kia_sales=None):
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
    month = control_tarjetas.build_month_control(today.year, today.month, card_sales, coupons, today=today,
                                                 kia_sales=kia_sales)
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


# ---------------------------------------------------------------------------
# Portales (pedido del usuario, 2026-10-15): cuándo entrar a cada portal a
# bajar las facturas que no llegan solas. Las fechas salen del historial de
# cada portal (2025-2026, revisado ese día):
# - FPL: la factura sale del 19 al 24 (casi siempre el 22-23); se avisa desde el 23.
# - Manatee (agua): sale del 9 al 13; se avisa desde el 13. La app la guarda
#   con la fecha de la lectura (2 al 4 del mismo mes).
# - Airgas, alquiler del cilindro (RENT): el último día de cada mes; se avisa
#   desde el 1 del mes siguiente. Las entregas de CO2 no tienen fecha: una
#   cada 6 a 10 semanas, se avisa a las 8 semanas de la última.
# - Elistar: el Depts Report y el P & L de un mes están completos el 1 del
#   mes siguiente (Control CMV).
# Se mira solo desde la primera factura cargada de cada uno y los últimos
# 3 meses: lo de antes no se avisa.
# ---------------------------------------------------------------------------

_MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
          "septiembre", "octubre", "noviembre", "diciembre"]
_PORTALES_LOOKBACK_DAYS = 100
_CO2_WARN_DAYS = 56  # 8 semanas: lo que suele pasar entre dos entregas de CO2

PORTAL_URLS = {
    "fpl": "https://www.fpl.com/my-account/account-landing.html",
    "manatee": "https://css-prod.mymanatee.org",
    "airgas": "https://www.airgas.com/login",
    "elistar": "https://www.elistars.com",
}


def _add_months(year, month, delta):
    index = year * 12 + (month - 1) + delta
    return index // 12, index % 12 + 1


def _month_end(year, month):
    return date(year, month, calendar.monthrange(year, month)[1])


def _dm(day):
    return f"{day.day:02d}/{day.month:02d}"


def _monthly_portal(label, url, today, first_month, warn_from, expected, loaded, what, pattern):
    """
    Un renglón de un portal que publica una vez por mes. warn_from(año, mes):
    desde qué día tendría que estar; expected(año, mes): el día típico (para
    mostrar); loaded(año, mes): el monto (o True) si está cargado, o None.
    """
    year, month = today.year, today.month
    # El último mes que ya tendría que estar (este o uno anterior).
    while warn_from(year, month) > today:
        year, month = _add_months(year, month, -1)
    months = []
    y, m = year, month
    while (today - warn_from(y, m)).days <= _PORTALES_LOOKBACK_DAYS and (first_month is None or (y, m) >= first_month):
        months.append((y, m))
        y, m = _add_months(y, m, -1)
    if not months:  # nada cargado desde ese mes: se mira solo el último
        months = [(year, month)]
    missing = [ym for ym in months if loaded(*ym) is None]
    last_ok = next((ym for ym in months if loaded(*ym) is not None), None)
    gender = "o" if what.startswith("el ") else "a"  # "el reporte" / "la factura"
    if missing:
        names = [_MESES[m - 1] for _, m in reversed(missing)]
        if len(names) == 1:
            text = f"Falta {what} de {names[0]}: entrá a bajarl{gender}."
        else:
            plural = ("los " if gender == "o" else "las ") + what.split(" ", 1)[1] + "s"
            text = f"Faltan {plural} de {', '.join(names[:-1])} y {names[-1]}: entrá a bajarl{gender}s."
        return {"label": label, "url": url, "state": "warn", "text": text, "sub": pattern}
    # La próxima: el primer mes después del último que ya tendría que estar
    # que todavía no está cargado (la factura puede cargarse antes del aviso).
    ny, nm = _add_months(year, month, 1)
    while loaded(ny, nm) is not None:
        ny, nm = _add_months(ny, nm, 1)
    amount = loaded(*last_ok)
    amount_text = f" (${amount:,.2f})" if isinstance(amount, (int, float)) and not isinstance(amount, bool) else ""
    return {"label": label, "url": url, "state": "ok",
            "text": f"Cargad{gender} {what} de {_MESES[last_ok[1] - 1]}{amount_text}.",
            "sub": f"Próxima: ~{_dm(expected(ny, nm))} ({_MESES[nm - 1]}). {pattern}"}


def _first_month(dates, to_month):
    months = [to_month(d) for d in dates]
    return min(months) if months else None


def portales_status(invoices, elistar_months, today=None):
    """
    invoices: {"fpl": [...], "manatee_county": [...], "airgas": [...]}, las
    facturas de Proveedores (invoice_date ISO, amount); elistar_months: los
    (año, mes) con Elistar cargado en Control CMV.
    """
    today = today or date.today()

    def dates_of(key):
        result = []
        for inv in invoices.get(key) or []:
            try:
                result.append((date.fromisoformat(str(inv["invoice_date"])[:10]), inv.get("amount")))
            except (TypeError, ValueError):
                continue
        return result

    def by_month(items, to_month):
        result = {}
        for day, amount in items:
            result.setdefault(to_month(day), amount)
        return result

    rows = []

    # FPL: la factura lleva la fecha en que sale (fin del período).
    fpl = by_month(dates_of("fpl"), lambda d: (d.year, d.month))
    rows.append(_monthly_portal(
        "FPL (luz)", PORTAL_URLS["fpl"], today, min(fpl) if fpl else None,
        warn_from=lambda y, m: date(y, m, 23),
        expected=lambda y, m: date(y, m, 22),
        loaded=lambda y, m: fpl.get((y, m)),
        what="la factura", pattern="Sale del 19 al 24 de cada mes.",
    ))

    # Manatee: la app guarda la fecha de la lectura (2 al 4); una lectura que
    # cae a fin del mes anterior cuenta para el mes siguiente.
    manatee = by_month(dates_of("manatee_county"),
                       lambda d: _add_months(d.year, d.month, 1) if d.day > 20 else (d.year, d.month))
    rows.append(_monthly_portal(
        "Manatee (agua)", PORTAL_URLS["manatee"], today, min(manatee) if manatee else None,
        warn_from=lambda y, m: date(y, m, 13),
        expected=lambda y, m: date(y, m, 11),
        loaded=lambda y, m: manatee.get((y, m)),
        what="la factura", pattern="Sale del 9 al 13 de cada mes.",
    ))

    # Airgas: el alquiler es la factura del último día del mes; lo demás, CO2.
    airgas = dates_of("airgas")
    rent = by_month([(d, a) for d, a in airgas if d == _month_end(d.year, d.month)], lambda d: (d.year, d.month))
    rows.append(_monthly_portal(
        "Airgas (alquiler)", PORTAL_URLS["airgas"], today, min(rent) if rent else None,
        warn_from=lambda y, m: _month_end(y, m) + timedelta(days=1),
        expected=lambda y, m: _month_end(y, m),
        loaded=lambda y, m: rent.get((y, m)),
        what="la factura", pattern="Sale el último día de cada mes.",
    ))
    co2 = sorted(d for d, _ in airgas if d != _month_end(d.year, d.month))
    if co2:
        last = co2[-1]
        days = (today - last).days
        pattern = "No tiene fecha fija: suele haber una cada 6 a 10 semanas."
        if days >= _CO2_WARN_DAYS:
            rows.append({"label": "Airgas (CO2)", "url": PORTAL_URLS["airgas"], "state": "warn",
                         "text": f"Pasaron {days} días desde la última entrega ({_dm(last)}): fijate si hay una nueva.",
                         "sub": pattern})
        else:
            rows.append({"label": "Airgas (CO2)", "url": PORTAL_URLS["airgas"], "state": "ok",
                         "text": f"Última entrega cargada: {_dm(last)} (hace {days} días).",
                         "sub": f"Próxima: ~{_dm(last + timedelta(days=_CO2_WARN_DAYS))}. {pattern}"})

    # Elistar: los reportes de un mes, desde el 1 del mes siguiente.
    rows.append(_monthly_portal(
        "Elistar (Depts y P & L)", PORTAL_URLS["elistar"], today, min(elistar_months) if elistar_months else None,
        warn_from=lambda y, m: _month_end(y, m) + timedelta(days=1),
        expected=lambda y, m: _month_end(y, m) + timedelta(days=1),
        loaded=lambda y, m: True if (y, m) in elistar_months else None,
        what="el reporte", pattern="Completos el 1 del mes siguiente; se cargan en Controles → Control CMV.",
    ))
    return {"rows": rows, "pending": sum(1 for row in rows if row["state"] == "warn")}
