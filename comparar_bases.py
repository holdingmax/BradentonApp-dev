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


SCENARIOS = {"jh_facturas": _escenario_jh_facturas}


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
            result[f"escritura {module}"] = {"status": 0, "body": json.dumps(SCENARIOS[module](), ensure_ascii=False, indent=1, default=str)}
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
        copy = subprocess.run([sys.executable, os.path.join(APP_DIR, "migrar_sqlite_a_postgres.py"), *args.modulos, "--origen", pg_dir, "--vaciar"],
                              cwd=APP_DIR, env={**os.environ, "DATABASE_URL": url, "PYTHONIOENCODING": "utf-8"})
        if copy.returncode:
            sys.exit("La copia a Postgres no quedó idéntica: no se compara.")
        lite = _run_capture(args.modulos, lite_dir, os.path.join(folder, "sqlite.json"), None)
        pg = _run_capture(args.modulos, pg_dir, os.path.join(folder, "postgres.json"), url)

    differences = 0
    print(f"\n{'qué':70} {'estado':>11}  resultado")
    for key in lite:
        a, b = lite[key], pg.get(key)
        same = b is not None and a == b
        differences += not same
        status = f"{a['status']}/{b['status'] if b else '-'}"
        print(f"{key[:70]:70} {status:>11}  {'idéntico' if same else 'DISTINTO'}")
        if not same and b is not None:
            diff = difflib.unified_diff(a["body"].splitlines(), b["body"].splitlines(), "SQLite", "Postgres", lineterm="", n=1)
            for line in list(diff)[:40]:
                print("    " + line[:200])
    print("\nTodo idéntico." if not differences else f"\n{differences} diferencias: revisar antes del commit.")
    sys.exit(1 if differences else 0)


if __name__ == "__main__":
    main()
