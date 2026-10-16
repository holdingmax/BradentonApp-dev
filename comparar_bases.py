"""
Compara la app sobre SQLite y sobre PostgreSQL con los mismos datos reales
(pase a Postgres, punto 6 de PLAN_POSTGRES.md). Antes de cada commit de un
módulo, sus páginas, descargas y escrituras tienen que dar idénticas.

    python comparar_bases.py jh_facturas [--datos CARPETA]

- Toma dos copias de la carpeta de datos (por defecto, la de la app; los
  originales no se tocan): una para correr la app sobre SQLite y otra para
  la corrida sobre Postgres.
- Copia el módulo a Postgres con migrar_sqlite_a_postgres.py --vaciar, así
  que solo corre contra un Postgres de esta PC (DATABASE_URL apuntando a
  localhost, nunca a testing ni producción).
- En cada corrida abre las páginas del módulo con el test client de Flask
  (logueado como admin, sin contraseña), hace sus pruebas de escritura y
  vuelve a abrir las páginas. Después compara todo: el HTML, el texto de
  los PDF y las celdas de los Excel.
"""

import argparse
import difflib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import date

APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_HOSTS = {"localhost", "127.0.0.1", "::1"}
CONFIG_FILES = ("users.json", "chase_master_rules.json", "chase_rules.json", "proveedores_dynamic_suppliers.json", "proveedores_pago_rules.json")
MONTHS = [(2026, 7), (2026, 8), (2026, 9), (2026, 10)]


def _months(path):
    return [f"{path}?year={y}&month={m}" for y, m in MONTHS]


# Páginas que leen cada módulo. Se agregan a medida que se pasa cada uno.
PAGES = {
    "jh_facturas": _months("/carga-datos/eft/varias") + ["/carga-datos/eft/proximos", "/controles/rapidos/eft", "/carga-datos/eft"],
    "ice_machine": _months("/carga-datos/depositos"),
    "reporte_mensual": [url for path in ("/carga-datos/reporte-mensual", "/carga-datos/reporte-diario", "/reporte/historial",
                                         "/reporte/store-info/historial", "/controles/cierre", "/controles/cierre/exportar/pdf",
                                         "/controles/cmv") for url in _months(path)],
    "cheques": ["/carga-datos/chase/cheques"],
    "control_cmv": _months("/controles/cmv"),
    "depositos": [url for path in ("/carga-datos/depositos", "/controles/depositos", "/carga-datos/chase") for url in _months(path)],
    "horas_trabajo": ["/carga-datos/horas-trabajo"] + _months("/carga-datos/horas-trabajo/historial"),
    "jh_mensual": [url for path in ("/carga-datos/eft/jh", "/carga-datos/eft/historial", "/carga-datos/eft/cupones/historial",
                                    "/carga-datos/reportes/eft/pdf", "/controles/tarjetas", "/controles/cierre",
                                    "/controles/rapidos/tarjetas") for url in _months(path)]
                  + ["/carga-datos/eft/proximos", "/controles/rapidos/eft"],
    "documents": ["/buscar/documentos?q=close"],  # /documentos/... redirige mientras el guardado esté apagado
    "fisico": [url for path in ("/fisico", "/fisico/exportar", "/fisico/exportar/pdf") for url in _months(path)]
              + ["/carga-datos/eft/proximos", "/controles/rapidos/eft"],
    "caja": [f"/carga-datos/caja?year=2026&month={m}" for m in range(1, 7)]
            + [url for path in ("/carga-datos/caja", "/carga-datos/caja/exportar", "/carga-datos/caja/exportar/pdf",
                                "/carga-datos/reportes/caja/pdf", "/controles/caja", "/controles/rapidos/caja") for url in _months(path)]
            + ["/carga-datos/reportes"],
    "gettel_toyota": [url for path in ("/carga-datos/gettel", "/carga-datos/gettel/historial", "/carga-datos/reportes/gettel/excel",
                                       "/carga-datos/reportes/gettel/pdf", "/controles/kia-toyota", "/controles/rapidos/kia")
                      for url in _months(path)],
    "lottery": ["/carga-datos/lottery", "/carga-datos/lottery/dia/2026-09-15", "/carga-datos/lottery/dia/2026-10-04",
                "/carga-datos/lottery/dia/2026-07-31"]
               + [url for path in ("/carga-datos/lottery/historial", "/carga-datos/lottery/exportar", "/carga-datos/lottery/exportar/pdf",
                                   "/carga-datos/lottery/mensual", "/carga-datos/lottery/resumen-mensual", "/carga-datos/reportes/lottery/pdf",
                                   "/controles/lottery", "/controles/caja", "/carga-datos/caja") for url in _months(path)],
}


def _escenario_jh_facturas():
    """Altas, repetidos, conflicto y baja de facturas varias (datos de prueba,
    solo en las copias)."""
    import jh_facturas_db

    invoice = {"invoice_number": "SI-999901", "invoice_date": date(2026, 10, 3), "due_date": date(2026, 10, 6),
               "amount": 131.004, "kind": "Mako", "description": "Prueba del comparador", "detail": "con 'comillas' y 100%"}
    results = [jh_facturas_db.save_invoice(invoice, "Prueba (Mako).pdf"), jh_facturas_db.save_invoice(invoice)]
    try:
        jh_facturas_db.save_invoice({**invoice, "amount": 140.0})
    except ValueError as exc:
        results.append(str(exc))
    month = jh_facturas_db.get_month_invoices(2026, 10)
    results.append([{k: v for k, v in inv.items() if k != "uploaded_at"} for inv in month])
    jh_facturas_db.delete_invoice("SI-999901")
    results.append(len(jh_facturas_db.list_all_invoices()))
    return results


def _escenario_ice_machine():
    """Alta, reemplazo (mismo número) y baja de un Payment Summary de prueba."""
    import ice_machine_db

    summary = {"summary_no": "999901", "reference": "IND-PRUEBA", "from_date": "2026-10-03", "to_date": "2026-10-09",
               "transactions": 12, "gross": 150.0, "process_fees": 4.5, "service_fees": 1.0, "net": 144.5,
               "notes": "con 'comillas' y 100%"}
    results = [ice_machine_db.save_summary(summary, "Payment Summary prueba.pdf"),
               ice_machine_db.save_summary({**summary, "net": 140.25, "notes": None})]
    results.append([{k: v for k, v in s.items() if k != "uploaded_at"} for s in ice_machine_db.list_month(2026, 10)])
    results += [ice_machine_db.delete_summary("999901"), ice_machine_db.delete_summary("999901")]
    results.append(len(ice_machine_db.list_month(2026, 10)))
    return results


def _escenario_reporte_mensual():
    """El reporte de septiembre guardado como diciembre, corregido a mano, listado y borrado."""
    import reporte_mensual_db

    report = reporte_mensual_db.get_report(2026, 9)
    copy = {**report, "month": 12, "from_date": "2026-12-01", "to_date": "2026-12-31"}
    reporte_mensual_db.save_report(copy, "Resumen Ventas prueba.pdf")
    store_info = {**report["store_info"], "cash": (report["store_info"]["cash"] or 0) + 1.5, "network_revenue": None}
    departments = report["departments"][:5] + [{"department": "PRUEBA 100%", "count": 3, "amount": 0.0}]
    reporte_mensual_db.update_report(2026, 12, store_info, departments, 123.45, 7, ["aviso de prueba"], from_sheet=True)
    saved = reporte_mensual_db.get_report(2026, 12)
    results = [{k: v for k, v in saved.items() if k not in ("updated_at", "edited_at", "sheet_added_at")}, reporte_mensual_db.list_months()]
    reporte_mensual_db.delete_report(2026, 12)
    results.append(reporte_mensual_db.get_report(2026, 12))
    return results


def _escenario_cheques():
    """Cheque cargado a mano, corregido, con datos automáticos, buscado y borrado."""
    import cheques_db

    fields = {"manual_date": "2026-10-05", "manual_amount": 250.5, "manual_supplier": "Prueba 100%", "manual_invoice": "F-1",
              "manual_debit_date": None, "voided": False}
    check_id = cheques_db.save_manual(None, 99901, fields)
    results = [check_id, cheques_db.save_manual(check_id, 99901, {**fields, "manual_amount": 260.0, "voided": True})]
    try:
        cheques_db.save_manual(None, 99901, fields)
    except ValueError as exc:
        results.append(str(exc))
    cheques_db.set_auto_fields(check_id, {"auto_supplier_key": "coca", "auto_supplier_label": "Coca-Cola", "auto_invoice_no": "123",
                                          "auto_invoice_date": "2026-10-01", "auto_amount": 260.0})
    results += [cheques_db.find_by_number(99901), cheques_db.get_check(check_id), len(cheques_db.list_checks()),
                cheques_db.delete_check(check_id), cheques_db.delete_check(check_id), cheques_db.find_by_number(99901)]
    return [_without(r, "uploaded_at") for r in results]


def _escenario_control_cmv():
    """Un reporte de Elistar de septiembre copiado a diciembre, reemplazado y borrado."""
    import control_cmv_db

    september = control_cmv_db.get_month(2026, 9)
    results = [sorted(september)]
    for kind, report in september.items():
        copy = {k: v for k, v in report.items() if k not in ("filename", "uploaded_at")}
        copy.update(month=12)
        control_cmv_db.save_report(copy, "Depts prueba.xls")
        control_cmv_db.save_report({**copy, "nota": "reemplazo con 100%"}, "Depts prueba 2.xls")
    results.append({k: _without(v, "uploaded_at") for k, v in control_cmv_db.get_month(2026, 12).items()})
    for kind in september:
        control_cmv_db.delete_report(2026, 12, kind)
    results.append(control_cmv_db.get_month(2026, 12))
    return results


def _escenario_depositos():
    """Recibos de prueba (uno sin N°, para el orden de los vacíos), duplicado, cambio de tipo, edición y baja."""
    import depositos_db

    a = depositos_db.add_deposit(2026, 10, "2026-10-08", 1234.5, "Recibo con 100%", "555001", "normal", "", "Dep (prueba).pdf", 0)
    b = depositos_db.add_deposit(2026, 10, "2026-10-08", 99.99, None, None, "normal", "", "Dep (prueba).pdf", 1)
    results = [depositos_db.find_duplicate("555001", "2026-10-08", 1234.501), depositos_db.find_duplicate("555001", "2026-10-08", 1234.6),
               depositos_db.find_by_source_page("Dep (prueba).pdf", 1), depositos_db.find_by_source_page("Dep (prueba).pdf", 0, "555001")]
    depositos_db.set_kind(a, "ice", "Ice Machine")
    depositos_db.update_deposit(b, "2026-10-09", 100.0, "editado", "food_truck")
    results += [[_without(d, "uploaded_at") for d in depositos_db.list_month(2026, 10)],
                [d["id"] for d in depositos_db.list_kind("ice")], _without(depositos_db.get_deposit(b), "uploaded_at")]
    for deposit_id in (a, b):
        depositos_db.delete_deposit(deposit_id)
    results.append(len(depositos_db.list_month(2026, 10)))
    return results


def _escenario_horas_trabajo():
    """Una semana de prueba cargada, vuelta a subir, con un empleado agregado, editado y borrado."""
    import horas_trabajo_db

    employees = [{"employee_name": "Prueba Uno", "hours_label": "26:48", "hours": 26.8}, {"employee_name": "Prueba Dos", "hours": 40}]
    week = horas_trabajo_db.upsert_week("2026-10-12", "2026-10-05", "2026-10-11", employees)
    again = horas_trabajo_db.upsert_week("2026-10-12", "2026-10-05", "2026-10-11", employees[:1], source="ocr")
    extra = horas_trabajo_db.add_employee(week, "Prueba Tres 100%", hours=5.5, rate=16.0, deduct=10.0, hours_label="05:30")
    horas_trabajo_db.update_employee(extra, hours=6.0, deduct=0.0)
    results = [week == again, horas_trabajo_db.get_employee_week(extra), horas_trabajo_db.known_employee_names(),
               [[_without(dict(w), "created_at", "employees"), [_without(dict(e), "id") for e in w.get("employees", [])]]
                for w in horas_trabajo_db.get_month_weeks(2026, 10)]]
    horas_trabajo_db.delete_employee(extra)
    horas_trabajo_db.delete_week(week)
    results += [horas_trabajo_db.get_week(week), len(horas_trabajo_db.get_month_weeks(2026, 10))]
    return results


def _escenario_jh_mensual():
    """Los reportes de J.H. de septiembre sumados a diciembre: primero un renglón, después todos, y la baja."""
    import jh_mensual_db

    september = jh_mensual_db.get_reports(2026, 9)
    results = [sorted(september)]
    for kind, report in sorted(september.items()):
        month_report = {**report, "month": 12, "from_date": "2026-12-01", "to_date": "2026-12-31"}
        results.append(jh_mensual_db.merge_report(kind, {**month_report, "rows": report["rows"][:1]}, "prueba.pdf"))
        results.append(jh_mensual_db.merge_report(kind, month_report, "prueba 2.pdf"))
        results.append(len(jh_mensual_db.get_rows(kind)))
    results.append({k: _without(v, "updated_at") for k, v in jh_mensual_db.get_reports(2026, 12).items()})
    for kind in september:
        jh_mensual_db.delete_report(kind, 2026, 12)
    results += [jh_mensual_db.get_reports(2026, 12), len(jh_mensual_db.get_all_invoice_numbers())]
    return results


def _escenario_documents():
    """Con el guardado prendido solo en la copia: listar, buscar, guardar uno de prueba, leerlo y borrarlo."""
    import tempfile as _tempfile

    import documents_db

    documents_db.GUARDAR_DOCUMENTOS = True
    try:
        clean = lambda docs: [_without(d, "uploaded_at", "stored_path") for d in docs]
        results = [clean(documents_db.list_all_documents("lottery")), clean(documents_db.search_documents("CLOSE", limit=5))]
        with _tempfile.NamedTemporaryFile("wb", suffix=".pdf", delete=False) as handle:
            handle.write(b"%PDF-1.4 prueba")
        doc_id = documents_db.store_document("lottery", handle.name, "Prueba 100%.pdf", 2026, 12, label="prueba")
        results += [doc_id, clean(documents_db.list_documents("lottery", 2026, 12)), _without(documents_db.get_document(doc_id), "uploaded_at", "stored_path")]
        documents_db.delete_document(doc_id)
        results.append(documents_db.get_document(doc_id))
        os.remove(handle.name)
    finally:
        documents_db.GUARDAR_DOCUMENTOS = False
    return results


def _escenario_fisico():
    """Factura de combustible de prueba con renglones, buscada sin distinguir mayúsculas, editada y borrada; ajustes del mes."""
    import fisico_db

    lines = [{"product_code": "087", "description": "Regular 100%", "qty_billing": 6000, "qty_freight": 6000,
              "unit_price_wo_tax": 2.1, "unit_price_with_tax": 2.75, "line_total": 16500.0},
             {"product_code": "093", "description": "Premium", "qty_billing": 2600, "qty_freight": 2600,
              "unit_price_wo_tax": 2.844, "unit_price_with_tax": 3.494, "line_total": 9084.4}]
    invoice_id = fisico_db.add_invoice("2026-12-03", "2026-12-08", "SI-PRUEBA1", 8600, 25584.4, source="pdf", bol_number="B1",
                                       lines=lines, source_filename="Fuel prueba.pdf")
    clean = lambda d: _without(d, "created_at", "updated_at") if d else d
    results = [invoice_id, clean(fisico_db.find_invoice_by_number("si-prueba1")), [_without(l, "id") for l in fisico_db.get_invoice_lines(invoice_id)]]
    fisico_db.update_invoice(invoice_id, "2026-12-04", "2026-12-09", "SI-PRUEBA1", 8601, 25600.0)
    fisico_db.set_month_initial_override(2026, 12, 2746.15, 7000.5)
    fisico_db.set_month_real_ending(2026, 12, 3000, "2026-12-31")
    results += [[clean(i) for i in fisico_db.get_month_invoices(2026, 12)], fisico_db.get_invoice_years(),
                _without(fisico_db.get_month_settings(2026, 12) or {}, "updated_at"), fisico_db.get_settings_years()]
    fisico_db.delete_invoice(invoice_id)
    results += [fisico_db.find_invoice_by_number("SI-PRUEBA1"), len(fisico_db.list_all_invoices())]
    return results


def _escenario_caja():
    """Gastos de prueba (con N° de ticket como número), repetidos, pendientes y saldos del mes."""
    import caja_db

    a = caja_db.add_expense_item("2026-12-05", 50.0, "Gasto con 100%", source="ocr", trans_no=77)
    b = caja_db.add_expense_item("2026-12-05", 50.0, None)
    results = [caja_db.find_same_expense("2026-12-05", 50.0, trans_no=77), caja_db.find_same_expense("2026-12-05", 50.004),
               caja_db.find_same_expense("2026-12-05", 50.0, trans_no="78")]
    caja_db.update_expense_detail(b, "detalle nuevo")
    pending = caja_db.add_pending_expense(2026, 12, "2026-12-06", 12.5, "ticket dudoso", "Gasto prueba.pdf")
    results += [caja_db.add_pending_expense(2026, 12, "2026-12-06", 12.5, "otra vez", "Gasto prueba.pdf"),
                caja_db.get_month_pending_expenses(2026, 12), caja_db.delete_pending_expense(pending), caja_db.delete_pending_expense(pending)]
    caja_db.set_month_opening_balance(2026, 12, 1234.56)
    caja_db.set_month_closing_override(2026, 12, 999.99)
    results += [caja_db.get_month_expense_items(2026, 12), caja_db.get_month_expenses(2026, 12), caja_db.get_expense_years(),
                _without(caja_db.get_month_settings(2026, 12) or {}, "updated_at")]
    for item in (a, b):
        caja_db.delete_expense_item(item)
    results.append(caja_db.get_month_expense_items(2026, 12))
    return [[_without(x, "created_at") for x in r] if isinstance(r, list) else _without(r, "created_at") for r in results]


def _escenario_gettel_toyota():
    """Días de Gettel y Toyota de prueba, pagos (con N° de transacción en otra caja), ajuste del mes y bajas."""
    from datetime import date as _date

    import gettel_db

    results = [gettel_db.upsert_vendor_totals("gettel", {_date(2026, 12, 1): {"amount": 812.34, "gallons": 250.5},
                                                         _date(2026, 12, 2): {"amount": 0.0, "gallons": 0.0}}),
               gettel_db.upsert_vendor_totals("toyota", {_date(2026, 12, 1): {"amount": 99.99, "gallons": 30.0}}, source="pdf")]
    pago = gettel_db.add_pago(_date(2026, 12, 5), 1, "TX-9001", 1234.5, source="pdf", empresa="Kia")
    gettel_db.update_pago(pago, _date(2026, 12, 6), 2, "TX-9001", 1300.0, empresa="Toyota")
    gettel_db.set_month_pendiente_anterior_override(2026, 12, 456.78)
    clean = lambda rows: [_without(r, "created_at", "updated_at") for r in rows]
    results += [_without(gettel_db.get_day("2026-12-01"), "updated_at"), clean(gettel_db.get_month_days(2026, 12)),
                gettel_db.get_month_gettel_amount(2026, 12), gettel_db.get_day_gettel_amount("2026-12-01"),
                _without(gettel_db.find_pago_by_transaction("tx-9001"), "created_at", "updated_at"), clean(gettel_db.get_month_pagos(2026, 12)),
                gettel_db.get_pago_years(), gettel_db.get_toyota_days_years(), gettel_db.get_earliest_month(),
                _without(gettel_db.get_pago_month_settings(2026, 12) or {}, "updated_at"), gettel_db.get_pago_settings_years()]
    gettel_db.delete_pago(pago)
    results.append(gettel_db.find_pago_by_transaction("TX-9001"))
    return results


def _escenario_lottery():
    """Un día de prueba por sus tres caminos, fechas de pago de los bloques (con el reajuste) y el reporte mensual."""
    import lottery_db

    real = lottery_db.get_day("2026-09-15")
    sales = {k: real.get(k) for k in lottery_db._SALES_REPORT_FIELDS}
    lottery_db.upsert_department_fields("2026-12-07", 3, 45.5, 2, 20.0, "ocr", "Close Store prueba.pdf")
    lottery_db.upsert_sales_report_fields("2026-12-07", sales, "portal", "Daily Sales prueba.pdf")
    lottery_db.upsert_manual_day("2026-12-08", {**sales, "online_count": 1, "online_net_sales": 10})
    lottery_db.blank_day("2026-12-09")
    results = [_without(lottery_db.get_day("2026-12-07"), "updated_at"), sorted(lottery_db.get_days_between("2026-12-01", "2026-12-31")),
               lottery_db.set_block_chase_date(2026, 49, "2026-12-11"), lottery_db.set_block_chase_date(2026, 48, "2026-12-03"),
               lottery_db.build_month_blocks(2026, 12), lottery_db.monthly_debit_total(2026, 12), lottery_db.get_available_years(),
               lottery_db.compute_month_closing(2026, 12)]
    report = lottery_db.get_monthly_report(2026, 9)
    if report:
        lottery_db.save_monthly_report({**report, "month": 12}, "Monthly prueba.pdf")
        lottery_db.save_monthly_report({**report, "month": 12, "warnings": ["reemplazo"]}, "Monthly prueba 2.pdf")
        results.append(_without(lottery_db.get_monthly_report(2026, 12), "updated_at"))
        lottery_db.delete_monthly_report(2026, 12)
    results += [lottery_db.get_monthly_report(2026, 12), lottery_db.set_block_chase_date(2026, 48, None)]
    return results


def _without(value, *keys):
    return {k: v for k, v in value.items() if k not in keys} if isinstance(value, dict) else value


SCENARIOS = {"jh_facturas": _escenario_jh_facturas, "ice_machine": _escenario_ice_machine,
             "reporte_mensual": _escenario_reporte_mensual, "cheques": _escenario_cheques,
             "control_cmv": _escenario_control_cmv, "depositos": _escenario_depositos,
             "horas_trabajo": _escenario_horas_trabajo, "jh_mensual": _escenario_jh_mensual,
             "documents": _escenario_documents, "fisico": _escenario_fisico,
             "caja": _escenario_caja, "gettel_toyota": _escenario_gettel_toyota,
             "lottery": _escenario_lottery}


# --- Una corrida (proceso aparte por motor) --------------------------------

def _normalize_html(text):
    text = re.sub(r"\b\d{1,2}:\d{2}(:\d{2})?\b", "HH:MM", text)  # horas de "generado a las ..."
    return text


def _body(response):
    kind = response.headers.get("Content-Type", "")
    data = response.get_data()
    if "pdf" in kind:
        from pypdf import PdfReader

        return "\n".join(page.extract_text() or "" for page in PdfReader(io.BytesIO(data)).pages)
    if "spreadsheet" in kind or "excel" in kind:
        import openpyxl

        book = openpyxl.load_workbook(io.BytesIO(data))
        return "\n".join(f"{sheet.title}!{cell.coordinate}={cell.value!r}" for sheet in book for row in sheet.iter_rows() for cell in row if cell.value is not None)
    if "json" in kind or "html" in kind or "text" in kind:
        return _normalize_html(data.decode("utf-8", "replace"))
    return f"<{len(data)} bytes {kind}>"


_TIME_KEYS = {"updated_at", "created_at", "uploaded_at", "changed_at", "edited_at", "sheet_added_at", "hidden_at", "registrado_en"}


def _drop_times(value):
    """Saca las horas de guardado (cambian entre corridas) de lo que devuelve un escenario."""
    if hasattr(value, "keys") and not isinstance(value, dict):
        value = {k: value[k] for k in value.keys()}
    if isinstance(value, dict):
        return {k: _drop_times(v) for k, v in value.items() if k not in _TIME_KEYS}
    if isinstance(value, (list, tuple)):
        return [_drop_times(v) for v in value]
    return value


def _jsonable(value):
    """Filas (sqlite3.Row o db.Row) como dict; lo demás como texto."""
    if hasattr(value, "keys"):
        return {k: value[k] for k in value.keys()}
    return str(value)


def _capture(modules, out_path):
    os.startfile = lambda *a, **k: None  # que ninguna ruta abra Excel en la PC
    code_dir = os.environ.get("BRADENTON_CODE_DIR") or APP_DIR
    sys.path.insert(0, code_dir)
    import flask_login.utils
    import webapp

    app = webapp.app
    environ = {"REMOTE_ADDR": "127.0.0.1", "HTTP_USER_AGENT": "comparador-bases"}
    with app.test_request_context(environ_base=environ):
        identifier = flask_login.utils._create_identifier()
    client = app.test_client()
    client.environ_base.update(environ)
    with client.session_transaction() as session:
        session.update({"_user_id": "admin", "_fresh": True, "_id": identifier, "_last_seen": time.time()})

    def pages(tag):
        found = {}
        for module in modules:
            for url in PAGES.get(module, []):
                response = client.get(url)
                found[f"{tag} {url}"] = {"status": response.status_code, "body": _body(response)}
        return found

    result = pages("antes")
    for module in modules:
        if module in SCENARIOS:
            result[f"escritura {module}"] = {"status": 0, "body": json.dumps(_drop_times(SCENARIOS[module]()), ensure_ascii=False, indent=1, default=_jsonable)}
    result.update(pages("después"))
    with open(out_path, "w", encoding="utf-8") as handle:
        json.dump(result, handle, ensure_ascii=False)


# --- Orquestación ----------------------------------------------------------

def _copy_data(source, target):
    os.makedirs(target)
    for name in os.listdir(source):
        if name.endswith(".db"):
            import sqlite3

            src = sqlite3.connect(f"file:{os.path.join(source, name)}?mode=ro", uri=True)
            dst = sqlite3.connect(os.path.join(target, name))
            src.backup(dst)
            src.close()
            dst.close()
        elif name.endswith(".json"):
            shutil.copy2(os.path.join(source, name), target)
    for name in CONFIG_FILES:
        if os.path.isfile(os.path.join(APP_DIR, name)):
            shutil.copy2(os.path.join(APP_DIR, name), target)


def _run_capture(modules, data_dir, out_path, database_url, code_dir=None):
    env = {k: v for k, v in os.environ.items() if k not in ("DATABASE_URL", "BRADENTON_CODE_DIR")}
    env.update({"BRADENTON_DATA_DIR": data_dir, "PYTHONIOENCODING": "utf-8"})
    if code_dir:
        env["BRADENTON_CODE_DIR"] = code_dir
    if database_url:
        env["DATABASE_URL"] = database_url
    done = subprocess.run([sys.executable, os.path.abspath(__file__), "--capturar", out_path, *modules], env=env, cwd=APP_DIR)
    if done.returncode:
        sys.exit(f"La corrida {'Postgres' if database_url else 'SQLite'} falló (ver arriba).")
    with open(out_path, encoding="utf-8") as handle:
        return json.load(handle)


def main():
    parser = argparse.ArgumentParser(description="Compara la app sobre SQLite y sobre PostgreSQL con los mismos datos.")
    parser.add_argument("modulos", nargs="+")
    parser.add_argument("--datos", help="carpeta de datos de origen (por defecto, la de la app)")
    parser.add_argument("--original", metavar="COMMIT",
                        help="además, corre el código de ese commit (antes del pase) sobre SQLite: lo nuevo no tiene que cambiar nada")
    parser.add_argument("--capturar", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.capturar:
        return _capture(args.modulos, args.capturar)

    sys.path.insert(0, APP_DIR)
    import app_paths
    import db

    if not db.uses_postgres():
        db.load_dotenv()
    url = db.database_url()
    import psycopg

    if not url or (psycopg.conninfo.conninfo_to_dict(url).get("host") or "localhost") not in LOCAL_HOSTS:
        sys.exit("El comparador vacía y recarga tablas: necesita DATABASE_URL apuntando a un Postgres de esta PC (127.0.0.1).")
    unknown = [m for m in args.modulos if m not in PAGES]
    if unknown:
        sys.exit(f"Sin páginas registradas para: {', '.join(unknown)}")

    with tempfile.TemporaryDirectory(prefix="bradenton_comparar_") as folder:
        lite_dir, pg_dir = os.path.join(folder, "sqlite"), os.path.join(folder, "postgres")
        source = args.datos or app_paths.DATA_DIR
        _copy_data(source, lite_dir)
        _copy_data(source, pg_dir)
        # Se recargan todos los módulos que ya están en Postgres, no solo los
        # que se comparan: las páginas leen también los otros.
        with psycopg.connect(url) as pg:
            migrated = [r[0] for r in pg.execute(
                "SELECT DISTINCT table_schema FROM information_schema.tables WHERE table_schema = ANY(%s)", (list(db.SCHEMAS),)
            ).fetchall()]
        missing = [m for m in args.modulos if m not in migrated]
        if missing:
            sys.exit(f"Sin tablas en Postgres para: {', '.join(missing)} (¿falta alembic upgrade head?)")
        copy = subprocess.run([sys.executable, os.path.join(APP_DIR, "migrar_sqlite_a_postgres.py"), *sorted(migrated), "--origen", pg_dir, "--vaciar"],
                              cwd=APP_DIR, env={**os.environ, "DATABASE_URL": url, "PYTHONIOENCODING": "utf-8"})
        if copy.returncode:
            sys.exit("La copia a Postgres no quedó idéntica: no se compara.")
        lite = _run_capture(args.modulos, lite_dir, os.path.join(folder, "sqlite.json"), None)
        pg = _run_capture(args.modulos, pg_dir, os.path.join(folder, "postgres.json"), url)
        original = None
        if args.original:
            code_dir, orig_dir = os.path.join(folder, "codigo_original"), os.path.join(folder, "sqlite_original")
            subprocess.run(["git", "worktree", "add", "--detach", "--quiet", code_dir, args.original], cwd=APP_DIR, check=True)
            try:
                _copy_data(source, orig_dir)
                original = _run_capture(args.modulos, orig_dir, os.path.join(folder, "original.json"), None, code_dir)
            finally:
                subprocess.run(["git", "worktree", "remove", "--force", code_dir], cwd=APP_DIR)

    differences, bad_status = 0, []
    print()
    for key in lite:
        a, b = lite[key], pg.get(key)
        same = b is not None and a == b
        differences += not same
        if a["status"] not in (0, 200) or (b and b["status"] not in (0, 200)) or 'name="password"' in a["body"]:
            bad_status.append(f"{key} ({a['status']}/{b['status'] if b else '-'})")
        if not same:
            print(f"DISTINTO  {key}  ({a['status']}/{b['status'] if b else '-'})")
            if b is not None:
                diff = difflib.unified_diff(a["body"].splitlines(), b["body"].splitlines(), "SQLite", "Postgres", lineterm="", n=1)
                for line in list(diff)[:40]:
                    print("    " + line[:200])
    print(f"{len(lite) - differences} de {len(lite)} comparaciones idénticas (SQLite contra Postgres).")
    if original is not None:
        changed = 0
        for key in lite:
            if original.get(key) != lite[key]:
                changed += 1
                print(f"CAMBIÓ CONTRA EL ORIGINAL  {key}")
                if key in original:
                    diff = difflib.unified_diff(original[key]["body"].splitlines(), lite[key]["body"].splitlines(), "original", "nuevo", lineterm="", n=1)
                    for line in list(diff)[:30]:
                        print("    " + line[:200])
        print(f"{len(lite) - changed} de {len(lite)} iguales al código original sobre SQLite.")
        differences += changed
    if bad_status:
        print("Páginas que no respondieron bien (o pidieron login):\n  " + "\n  ".join(bad_status))
    print("Todo idéntico." if not differences and not bad_status else "HAY QUE REVISAR antes del commit.")
    sys.exit(1 if differences or bad_status else 0)


if __name__ == "__main__":
    main()
