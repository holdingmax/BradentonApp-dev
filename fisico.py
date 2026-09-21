"""
Módulo "Físico" -- reconciliación mensual de combustible: Inventario
Teórico (lo que DEBERÍA quedar en los tanques según compras y ventas)
contra la lectura física real (lo que un gauge/medición mide de verdad).
Reproduce la lógica real de la hoja "Fisico" del Excel de cierre --
decodificada con openpyxl leyendo un ejemplo real (`hoja_fisico.xlsx`),
no adivinada.

Fórmula real por mes (columnas Gal/$'):
    Inventario Inicial Teórico   = Inventario Final Teórico del mes anterior
                                    (encadenado, mismo criterio que
                                    caja._resolve_opening_balance -- un
                                    override manual para ESTE mes manda
                                    siempre; si no hay, se encadena UN
                                    solo mes atrás, nunca más)
    Total Compras                = Inventario Inicial + SUM(facturas del mes)
    Precio promedio ($/gal)      = Total Compras $ / Total Compras Gal
    Ventas (gal)                 = -Volume del mes (Store Info, mismo dato
                                    que ya usa Proyecciones) -- negativo
                                    porque es una salida de combustible
    Ventas ($)                   = Ventas (gal) * Precio promedio
                                    (valuado al COSTO de compra, no al
                                    precio de venta -- así lo hace la
                                    hoja real: `=+D12*F10`)
    Inventario Final Teórico     = Total Compras + Ventas (gal y $)
    Inventario Final Real        = lectura física de los tanques, SIEMPRE
                                    un dato externo cargado a mano (ver
                                    fisico_db.set_month_real_ending) --
                                    "pendiente" hasta que se cargue (la
                                    hoja real la toma recién el día 1 del
                                    mes siguiente, así que un mes recién
                                    cerrado normalmente todavía no la
                                    tiene)
    Diferencia (gal)             = Inventario Final Teórico - Inventario
                                    Final Real (positivo = falta
                                    combustible/merma; negativo = sobra)
    Diferencia ($)                = Diferencia (gal) * Precio promedio
    Diferencia (%)                 = Diferencia (gal) / Inventario Final
                                    Teórico (gal)

Cada factura puede traer el galonaje partido en dos productos que se
suman (ej. `=7600+1199` en la hoja real, típicamente Regular + Diesel) --
`fisico_db.add_invoice` ya recibe un solo número de galones, así que esa
suma (si aplica) se hace en el formulario de carga (ver carga_datos_
combustible en webapp.py), replicando el mismo criterio.
"""

import calendar

import fisico_db
import reportes_db


def _month_volume_actual(year, month):
    """Suma de Volume (galones vendidos) del mes, directo de Store Info -- mismo dato que proyecciones.build_projection."""
    rows = reportes_db.get_month_store_info(year, month)
    return sum(row.get("volume") or 0.0 for row in rows if row.get("volume") is not None)


def _resolve_initial(year, month, _chain=True):
    """
    (gallons, amount, source) del Inventario Inicial Teórico -- mismo
    patrón que caja._resolve_opening_balance: un override manual de ESTE
    mes manda siempre; si no hay y `_chain` es True, se encadena UN mes
    atrás (su Inventario Final Teórico, calculado); si tampoco hay nada
    ahí, 0.0/0.0 (mes realmente inicial, sin nada cargado antes).
    """
    settings = fisico_db.get_month_settings(year, month)
    if settings and settings.get("initial_gallons_override") is not None:
        return settings["initial_gallons_override"], settings.get("initial_amount_override") or 0.0, "manual"
    if not _chain:
        return 0.0, 0.0, "default"

    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    prev_report = build_month_report(prev_year, prev_month, _chain=False)
    if prev_report["has_any_data"]:
        return (
            prev_report["final_teorico_gallons"],
            prev_report["final_teorico_amount"],
            "prev_month_computed",
        )
    return 0.0, 0.0, "default"


def build_month_report(year, month, _chain=True):
    invoices = fisico_db.get_month_invoices(year, month)
    settings = fisico_db.get_month_settings(year, month)

    initial_gallons, initial_amount, initial_source = _resolve_initial(year, month, _chain=_chain)

    invoices_gallons = round(sum(inv["gallons"] or 0.0 for inv in invoices), 2)
    invoices_amount = round(sum(inv["amount"] or 0.0 for inv in invoices), 2)

    total_gallons = round(initial_gallons + invoices_gallons, 2)
    total_amount = round(initial_amount + invoices_amount, 2)
    avg_price = (total_amount / total_gallons) if total_gallons else None

    ventas_gallons = round(-_month_volume_actual(year, month), 2)
    ventas_amount = round(ventas_gallons * avg_price, 2) if avg_price is not None else None

    final_teorico_gallons = round(total_gallons + ventas_gallons, 2)
    final_teorico_amount = (
        round(total_amount + ventas_amount, 2) if ventas_amount is not None else total_amount
    )

    real_gallons = settings.get("real_ending_gallons") if settings else None
    real_reading_date = settings.get("real_reading_date") if settings else None

    if real_gallons is not None:
        diferencia_gallons = round(final_teorico_gallons - real_gallons, 2)
        diferencia_amount = round(diferencia_gallons * avg_price, 2) if avg_price is not None else None
        diferencia_pct = (
            (diferencia_gallons / final_teorico_gallons) if final_teorico_gallons else None
        )
    else:
        diferencia_gallons = None
        diferencia_amount = None
        diferencia_pct = None

    has_any_data = bool(invoices) or (settings is not None) or initial_gallons or initial_amount

    return {
        "year": year,
        "month": month,
        "has_any_data": has_any_data,
        "invoices": invoices,
        "initial_gallons": initial_gallons,
        "initial_amount": round(initial_amount, 2),
        "initial_source": initial_source,
        "invoices_gallons": invoices_gallons,
        "invoices_amount": invoices_amount,
        "total_gallons": total_gallons,
        "total_amount": total_amount,
        "avg_price": round(avg_price, 4) if avg_price is not None else None,
        "ventas_gallons": ventas_gallons,
        "ventas_amount": ventas_amount,
        "final_teorico_gallons": final_teorico_gallons,
        "final_teorico_amount": final_teorico_amount,
        "real_gallons": real_gallons,
        "real_reading_date": real_reading_date,
        "diferencia_gallons": diferencia_gallons,
        "diferencia_amount": diferencia_amount,
        "diferencia_pct": diferencia_pct,
    }


def get_available_years():
    years = set(fisico_db.get_invoice_years()) | set(fisico_db.get_settings_years())
    return sorted(years)
