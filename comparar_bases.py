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


def _without(value, *keys):
    return {k: v for k, v in value.items() if k not in keys} if isinstance(value, dict) else value


SCENARIOS = {"jh_facturas": _escenario_jh_facturas, "ice_machine": _escenario_ice_machine,
             "reporte_mensual": _escenario_reporte_mensual, "cheques": _escenario_cheques,
             "control_cmv": _escenario_control_cmv, "depositos": _escenario_depositos,
             "horas_trabajo": _escenario_horas_trabajo}


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


def _jsonable(value):
    """Filas (sqlite3.Row o db.Row) como dict; lo demás como texto."""
    if hasattr(value, "keys"):
        return {k: value[k] for k in value.keys()}
    return str(value)


def _capture(modules, out_path):
    os.startfile = lambda *a, **k: None  # que ninguna ruta abra Excel en la PC
    sys.path.insert(0, APP_DIR)
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
            result[f"escritura {module}"] = {"status": 0, "body": json.dumps(SCENARIOS[module](), ensure_ascii=False, indent=1, default=_jsonable)}
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


def _run_capture(modules, data_dir, out_path, database_url):
    env = {k: v for k, v in os.environ.items() if k != "DATABASE_URL"}
    env.update({"BRADENTON_DATA_DIR": data_dir, "PYTHONIOENCODING": "utf-8"})
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
    print(f"{len(lite) - differences} de {len(lite)} comparaciones idénticas.")
    if bad_status:
        print("Páginas que no respondieron bien (o pidieron login):\n  " + "\n  ".join(bad_status))
    print("Todo idéntico." if not differences and not bad_status else "HAY QUE REVISAR antes del commit.")
    sys.exit(1 if differences or bad_status else 0)


if __name__ == "__main__":
    main()
