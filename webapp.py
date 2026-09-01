"""
BradentonApp — versión web (Flask).

Primer paso de la migración de escritorio (Tkinter) a web, para poder
correr como servicio en Render y entrar en Toolbox. Reusa la lógica de
negocio ya extraída a módulos sin dependencia de Tkinter (chase_rules.py,
cmv_costo.py, etc.) — nunca reimplementa esa lógica acá.

Un módulo por vez: hoy solo está Chase Bank. El resto se va sumando
igual que el desktop, probando cada uno antes de seguir con el próximo.
"""

import os
import tempfile

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from flask_login import (
    LoginManager,
    UserMixin,
    current_user,
    login_required,
    login_user,
    logout_user,
)

import auth
from chase_rules import (
    add_dynamic_rule as add_chase_rule,
    delete_dynamic_rule_by_index as delete_chase_custom_rule,
    delete_master_rule_by_index as delete_chase_master_rule,
    edit_dynamic_rule_by_index as edit_chase_custom_rule,
    edit_master_rule_by_index as edit_chase_master_rule,
    list_display_rules as list_chase_display_rules,
    process_chase_categorization,
)
from caja import apply_chase_deposits, apply_lottery_cuenta_final
from cmv_costo import update_master_costo_todos_bulk
from cupones_append import (
    MonthlyReportFullyDuplicateError,
    NoPendingCouponsError,
    append_monthly_cupones,
    resync_cupones_only,
)
from eft_cta_cte import EFT_DUPLICATE_ALERT, eft_already_loaded_in_workbook, extract_eft_data, update_excel_workbook
from gettel_toyota_parser import (
    merge_gettel_toyota_into_master,
    merge_gettel_toyota_pdf_into_master,
    process_gettel_pagos,
)
from monthly_sales import process_monthly_sales
from proveedores import append_supplier_invoices, append_supplier_payments
from reporte_diario import process_lottery, process_reporte_diario, process_store_info

def _load_or_create_secret_key():
    """
    Persist the session secret key on disk (gitignored) instead of
    regenerating it on every restart — otherwise every reload of the dev
    server (Flask's debug reloader restarts often) would log everyone out.
    """
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".flask_secret_key")
    if os.path.isfile(path):
        with open(path, "rb") as handle:
            key = handle.read()
        if key:
            return key
    key = os.urandom(32)
    with open(path, "wb") as handle:
        handle.write(key)
    return key


app = Flask(__name__)
app.secret_key = _load_or_create_secret_key()

login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Iniciá sesión para continuar."
login_manager.login_message_category = "error"


class WebUser(UserMixin):
    def __init__(self, username, is_admin):
        self.id = username
        self.is_admin = is_admin


@login_manager.user_loader
def load_user(username):
    user = auth.get_user(username)
    if user is None:
        return None
    return WebUser(username, user.get("is_admin", False))


def _is_safe_next_url(target):
    """Only allow redirecting to an in-app relative path after login."""
    return bool(target) and target.startswith("/") and not target.startswith("//")


@app.before_request
def require_login():
    if request.endpoint in ("login", "static") or request.endpoint is None:
        return None
    if not current_user.is_authenticated:
        return redirect(url_for("login", next=request.path))
    return None


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("index"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = auth.verify_user(username, password)
        if user is None:
            flash("Usuario o contraseña incorrectos.", "error")
        else:
            login_user(WebUser(username, user.get("is_admin", False)))
            next_url = request.args.get("next")
            return redirect(next_url if _is_safe_next_url(next_url) else url_for("index"))

    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


@app.route("/admin/users", methods=["GET", "POST"])
@login_required
def admin_users():
    if not current_user.is_admin:
        flash("No tenés permiso para acceder a esta página.", "error")
        return redirect(url_for("index"))

    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "create":
                # Nuevas cuentas siempre no-admin — solo la cuenta admin
                # inicial tiene ese rol por ahora, sin UI para promover otras.
                auth.create_user(
                    request.form.get("username", ""),
                    request.form.get("password", ""),
                    is_admin=False,
                )
                flash("Usuario creado.", "success")
            elif action == "reset_password":
                auth.set_password(
                    request.form.get("username", "").strip(),
                    request.form.get("new_password", ""),
                )
                flash("Contraseña actualizada.", "success")
            elif action == "delete":
                auth.delete_user(
                    request.form.get("username", "").strip(),
                    current_username=current_user.id,
                )
                flash("Usuario eliminado.", "success")
        except ValueError as exc:
            flash(str(exc), "error")
        return redirect(url_for("admin_users"))

    return render_template("admin_users.html", users=auth.list_users())

# Same per-module accent colors the desktop app used (ui_theme.py SectionTheme,
# now retired) — kept here purely as brand identity/wayfinding across pages.
# "icon" is inline SVG markup (rendered with |safe in index.html) chosen to
# match each module's real-world subject, not just a generic placeholder —
# e.g. a bank for Chase (a bank statement), a calendar for Reporte Diario
# (a daily report). "code" is kept too: some templates/emails may still want
# a compact text badge, but the Home grid now shows the icon instead.
_ICON_BANK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 10 12 4l9 6"/><path d="M4 10v9M9 10v9M15 10v9M20 10v9"/><path d="M2 21h20"/></svg>'
_ICON_COINS = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="6" rx="7" ry="3"/><path d="M5 6v6c0 1.7 3.1 3 7 3s7-1.3 7-3V6"/><path d="M5 12v6c0 1.7 3.1 3 7 3s7-1.3 7-3v-6"/></svg>'
_ICON_CAR = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 13l1.5-4.5A2 2 0 0 1 6.4 7h11.2a2 2 0 0 1 1.9 1.5L21 13"/><path d="M3 13v4a1 1 0 0 0 1 1h1a1 1 0 0 0 1-1v-1h12v1a1 1 0 0 0 1 1h1a1 1 0 0 0 1-1v-4"/><circle cx="7.5" cy="17" r="1.6"/><circle cx="16.5" cy="17" r="1.6"/></svg>'
_ICON_CALENDAR = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/></svg>'
_ICON_TICKET = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v2a2 2 0 0 0 0 4v2a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-2a2 2 0 0 0 0-4z"/><path d="M13 7v10" stroke-dasharray="2 2"/></svg>'
_ICON_EXCHANGE = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h13l-3-3M20 17H7l3 3"/></svg>'
_ICON_TRUCK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="7" width="13" height="10" rx="1"/><path d="M14 10h4l3 3v4h-7z"/><circle cx="6" cy="18.5" r="1.6"/><circle cx="17.5" cy="18.5" r="1.6"/></svg>'
_ICON_REGISTER = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="10" width="18" height="10" rx="1"/><path d="M6 10V7a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v3"/><path d="M9 15h6"/></svg>'

TOOLS = [
    {
        "key": "chase",
        "code": "CH",
        "icon": _ICON_BANK,
        "label": "Chase Bank",
        "url": "/chase",
        "description": "Categoriza movimientos bancarios contra las reglas de Detalle.",
        "accent": "#16A34A",
        "accent_soft": "#DCF3E3",
    },
    {
        "key": "cmv",
        "code": "CMV",
        "icon": _ICON_COINS,
        "label": "CMV",
        "url": "/cmv",
        "description": "Costo por UPC (COSTO.TODOS) y ventas del POS por departamento.",
        "accent": "#7C3AED",
        "accent_soft": "#E9E0FC",
    },
    {
        "key": "gettel",
        "code": "GT",
        "icon": _ICON_CAR,
        "label": "Gettel / Toyota",
        "url": "/gettel",
        "description": "Cupones diarios (Excel o PDF) y pagos hacia el master Cierre.",
        "accent": "#0D9488",
        "accent_soft": "#D6F1EE",
    },
    {
        "key": "reporte",
        "code": "RD",
        "icon": _ICON_CALENDAR,
        "label": "Reporte Diario",
        "url": "/reporte",
        "description": "Ventas por Departamento y Store Info desde el PDF de cierre diario.",
        "accent": "#0284C7",
        "accent_soft": "#D7EFFB",
    },
    {
        "key": "lottery",
        "code": "LT",
        "icon": _ICON_TICKET,
        "label": "Lottery",
        "url": "/lottery",
        "description": "Daily Sales Report y PDF Diario hacia el Excel de Lottery.",
        "accent": "#0284C7",
        "accent_soft": "#D7EFFB",
    },
    {
        "key": "eft",
        "code": "EFT",
        "icon": _ICON_EXCHANGE,
        "label": "Cupones y EFT",
        "url": "/eft",
        "description": "PDF de EFT bancario a Cta Cte, y reporte mensual de Cupones.",
        "accent": "#3B5BDB",
        "accent_soft": "#DDE3FA",
    },
    {
        "key": "proveedores",
        "code": "PR",
        "icon": _ICON_TRUCK,
        "label": "Proveedores",
        "url": "/proveedores",
        "description": "Facturas de compra por proveedor y pagos vía Chase al Cta Cte.",
        "accent": "#DB2777",
        "accent_soft": "#FBD9EA",
    },
    {
        "key": "caja",
        "code": "CJ",
        "icon": _ICON_REGISTER,
        "label": "Caja",
        "url": "/caja",
        "description": "Depósitos Chase y Cuenta Final Lottery hacia las columnas K/N/S de CAJA.",
        "accent": "#EA580C",
        "accent_soft": "#FCE3D2",
    },
]

THEME_BY_KEY = {
    tool["key"]: {"accent": tool["accent"], "accent_soft": tool["accent_soft"]}
    for tool in TOOLS
}


def _new_workspace_dir():
    return tempfile.mkdtemp(prefix="bradenton_web_")


def _save_upload_to_workspace(upload, workdir=None):
    """
    Save an uploaded file to a fresh (or given) temp dir; return its local path.

    Keeps the original filename byte-for-byte (parens, spaces, accents) —
    several parsers (Gettel Pagos, most Proveedores extractors) read the
    payment number/vendor/invoice date straight off the filename, so
    anything that mangles it (werkzeug's secure_filename, a timestamp
    prefix) breaks them. Collisions are avoided with a per-file
    subdirectory instead of touching the filename itself; os.path.basename
    still strips any directory components a browser might send.
    """
    filename = os.path.basename(upload.filename.replace("\\", "/"))
    if not filename or filename in (".", ".."):
        raise ValueError("Nombre de archivo inválido.")
    workdir = workdir or _new_workspace_dir()
    file_dir = tempfile.mkdtemp(dir=workdir)
    temp_path = os.path.join(file_dir, filename)
    upload.save(temp_path)
    return temp_path, filename


def _save_uploads_to_workspace(uploads, workdir=None):
    """Save several uploaded files to the same temp dir; return their local paths."""
    workdir = workdir or _new_workspace_dir()
    paths = []
    for upload in uploads:
        if not upload or not upload.filename:
            continue
        temp_path, _filename = _save_upload_to_workspace(upload, workdir=workdir)
        paths.append(temp_path)
    return paths


@app.route("/")
def index():
    return render_template("index.html", tools=TOOLS)


@app.route("/chase", methods=["GET", "POST"])
def chase():
    if request.method == "GET":
        return render_template(
            "chase.html", chase_rules=list_chase_display_rules(), **THEME_BY_KEY["chase"]
        )

    upload = request.files.get("chase_file")
    if upload is None or not upload.filename:
        flash("Seleccioná un archivo CSV o Excel de Chase.", "error")
        return redirect(url_for("chase"))

    try:
        temp_path, filename = _save_upload_to_workspace(upload)
        updated_count, total_rows = process_chase_categorization(temp_path)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("chase"))

    return send_file(
        temp_path,
        as_attachment=True,
        download_name=filename,
    )


def _require_admin_for_chase_rules():
    if not current_user.is_admin:
        flash("Solo un administrador puede gestionar las reglas de Chase.", "error")
        return False
    return True


@app.route("/chase/rules/save", methods=["POST"])
def chase_rules_save():
    if not _require_admin_for_chase_rules():
        return redirect(url_for("chase"))

    keyword = request.form.get("keyword", "")
    detail = request.form.get("detail", "")
    rule_type = request.form.get("rule_type", "").strip()
    index = request.form.get("index", "").strip()

    try:
        if not rule_type or not index:
            add_chase_rule(keyword, detail)
            flash("Regla creada.", "success")
        elif rule_type == "master":
            edit_chase_master_rule(index, keyword, detail)
            flash("Regla Maestra actualizada.", "success")
        elif rule_type == "custom":
            edit_chase_custom_rule(index, keyword, detail)
            flash("Regla actualizada.", "success")
        else:
            flash("Tipo de regla inválido.", "error")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("chase"))


@app.route("/chase/rules/delete", methods=["POST"])
def chase_rules_delete():
    if not _require_admin_for_chase_rules():
        return redirect(url_for("chase"))

    rule_type = request.form.get("rule_type", "").strip()
    index = request.form.get("index", "").strip()

    try:
        if rule_type == "master":
            delete_chase_master_rule(index)
            flash("Regla Maestra eliminada.", "success")
        elif rule_type == "custom":
            delete_chase_custom_rule(index)
            flash("Regla eliminada.", "success")
        else:
            flash("Seleccioná una regla de la tabla antes de eliminar.", "error")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("chase"))


@app.route("/cmv")
def cmv():
    return render_template("cmv.html", **THEME_BY_KEY["cmv"])


@app.route("/cmv/costo", methods=["POST"])
def cmv_costo():
    master_upload = request.files.get("master_file")
    dept_uploads = request.files.getlist("dept_files")
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel maestro CMV.", "error")
        return redirect(url_for("cmv"))
    if not dept_uploads or not any(u.filename for u in dept_uploads):
        flash("Seleccioná uno o más archivos de departamento.", "error")
        return redirect(url_for("cmv"))

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        dept_paths = _save_uploads_to_workspace(dept_uploads, workdir=workdir)
        temp_xlsx_path, _file_stats, _total_parsed, rows_updated, _upcs, _count = (
            update_master_costo_todos_bulk(master_path, dept_paths)
        )
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("cmv"))

    return send_file(
        temp_xlsx_path,
        as_attachment=True,
        download_name=master_filename,
    )


@app.route("/cmv/ventas", methods=["POST"])
def cmv_ventas():
    master_upload = request.files.get("master_file")
    sales_uploads = request.files.getlist("sales_files")
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel maestro CMV.", "error")
        return redirect(url_for("cmv"))
    if not sales_uploads or not any(u.filename for u in sales_uploads):
        flash("Seleccioná uno o más reportes de ventas del POS.", "error")
        return redirect(url_for("cmv"))

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        sales_paths = _save_uploads_to_workspace(sales_uploads, workdir=workdir)
        _combined, temp_master_path = process_monthly_sales(sales_paths, master_path)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("cmv"))

    return send_file(
        temp_master_path,
        as_attachment=True,
        download_name=master_filename,
    )


@app.route("/gettel")
def gettel():
    return render_template("gettel.html", **THEME_BY_KEY["gettel"])


@app.route("/gettel/cupones", methods=["POST"])
def gettel_cupones():
    source_upload = request.files.get("source_file")
    master_upload = request.files.get("master_file")
    if source_upload is None or not source_upload.filename:
        flash("Seleccioná el Excel o PDF/Foto de origen (cupones diarios).", "error")
        return redirect(url_for("gettel"))
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel de destino (master Cierre).", "error")
        return redirect(url_for("gettel"))

    try:
        workdir = _new_workspace_dir()
        source_path, _source_filename = _save_upload_to_workspace(source_upload, workdir=workdir)
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)

        is_pdf = os.path.splitext(source_path)[1].lower() == ".pdf"
        if is_pdf:
            preview_path, rows_matched, vendor, days_found, _diagnostics = (
                merge_gettel_toyota_pdf_into_master(source_path, master_path, launch=False)
            )
        else:
            preview_path, rows_matched, gettel_days, toyota_days, _unmatched = (
                merge_gettel_toyota_into_master(source_path, master_path, launch=False)
            )
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("gettel"))

    return send_file(
        preview_path,
        as_attachment=True,
        download_name=master_filename,
    )


@app.route("/gettel/pagos", methods=["POST"])
def gettel_pagos():
    master_upload = request.files.get("master_file")
    pdf_uploads = request.files.getlist("pdf_files")
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel de destino (master Cierre).", "error")
        return redirect(url_for("gettel"))
    if not pdf_uploads or not any(u.filename for u in pdf_uploads):
        flash("Seleccioná uno o más PDF de pagos.", "error")
        return redirect(url_for("gettel"))

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        pdf_paths = _save_uploads_to_workspace(pdf_uploads, workdir=workdir)
        preview_path, _summary = process_gettel_pagos(master_path, pdf_paths)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("gettel"))

    return send_file(
        preview_path,
        as_attachment=True,
        download_name=master_filename,
    )


@app.route("/reporte")
def reporte():
    return render_template("reporte.html", **THEME_BY_KEY["reporte"])


def _reporte_pdf_upload():
    """Shared validation + upload-saving for the two Reporte Diario forms."""
    master_upload = request.files.get("master_file")
    pdf_uploads = request.files.getlist("pdf_files")
    if not pdf_uploads or not any(u.filename for u in pdf_uploads):
        flash("Seleccioná uno o más PDF diarios.", "error")
        return None
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel de destino.", "error")
        return None

    workdir = _new_workspace_dir()
    master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
    pdf_paths = _save_uploads_to_workspace(pdf_uploads, workdir=workdir)
    return master_path, master_filename, pdf_paths


@app.route("/reporte/ventas", methods=["POST"])
def reporte_ventas():
    saved = _reporte_pdf_upload()
    if saved is None:
        return redirect(url_for("reporte"))
    master_path, master_filename, pdf_paths = saved

    try:
        temp_path, _summary = process_reporte_diario(master_path, pdf_paths)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("reporte"))

    return send_file(temp_path, as_attachment=True, download_name=master_filename)


@app.route("/reporte/store-info", methods=["POST"])
def reporte_store_info():
    saved = _reporte_pdf_upload()
    if saved is None:
        return redirect(url_for("reporte"))
    master_path, master_filename, pdf_paths = saved

    try:
        temp_path, _summary = process_store_info(master_path, pdf_paths)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("reporte"))

    return send_file(temp_path, as_attachment=True, download_name=master_filename)


@app.route("/lottery")
def lottery():
    return render_template("lottery.html", **THEME_BY_KEY["lottery"])


def _lottery_pdf_upload():
    """Shared validation + upload-saving for the two Lottery forms."""
    master_upload = request.files.get("master_file")
    pdf_uploads = request.files.getlist("pdf_files")
    if not pdf_uploads or not any(u.filename for u in pdf_uploads):
        flash("Seleccioná uno o más PDF.", "error")
        return None
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel de Lottery.", "error")
        return None

    workdir = _new_workspace_dir()
    master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
    pdf_paths = _save_uploads_to_workspace(pdf_uploads, workdir=workdir)
    return master_path, master_filename, pdf_paths


@app.route("/lottery/sales-report", methods=["POST"])
def lottery_sales_report():
    saved = _lottery_pdf_upload()
    if saved is None:
        return redirect(url_for("lottery"))
    master_path, master_filename, pdf_paths = saved

    try:
        temp_path, _summary = process_lottery(master_path, [], pdf_paths)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("lottery"))

    return send_file(temp_path, as_attachment=True, download_name=master_filename)


@app.route("/lottery/department", methods=["POST"])
def lottery_department():
    saved = _lottery_pdf_upload()
    if saved is None:
        return redirect(url_for("lottery"))
    master_path, master_filename, pdf_paths = saved

    try:
        temp_path, _summary = process_lottery(master_path, pdf_paths, [])
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("lottery"))

    return send_file(temp_path, as_attachment=True, download_name=master_filename)


@app.route("/eft")
def eft():
    return render_template("eft.html", **THEME_BY_KEY["eft"])


@app.route("/eft/cta-cte", methods=["POST"])
def eft_cta_cte_route():
    pdf_upload = request.files.get("pdf_file")
    master_upload = request.files.get("master_file")
    if pdf_upload is None or not pdf_upload.filename:
        flash("Seleccioná un archivo PDF de EFT.", "error")
        return redirect(url_for("eft"))
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel Ledger.", "error")
        return redirect(url_for("eft"))

    try:
        workdir = _new_workspace_dir()
        pdf_path, _pdf_filename = _save_upload_to_workspace(pdf_upload, workdir=workdir)
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)

        header_data, paid_invoices, credit_coupons = extract_eft_data(pdf_path)
        if eft_already_loaded_in_workbook(master_path, header_data, credit_coupons):
            flash(EFT_DUPLICATE_ALERT, "error")
            return redirect(url_for("eft"))

        update_excel_workbook(master_path, header_data, paid_invoices, credit_coupons)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("eft"))

    return send_file(master_path, as_attachment=True, download_name=master_filename)


@app.route("/eft/cupones", methods=["POST"])
def eft_cupones():
    master_upload = request.files.get("master_file")
    monthly_upload = request.files.get("monthly_report_file")
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel Ledger.", "error")
        return redirect(url_for("eft"))

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)

        if monthly_upload is not None and monthly_upload.filename:
            monthly_path, _monthly_filename = _save_upload_to_workspace(monthly_upload, workdir=workdir)
            saved_path, _summary = append_monthly_cupones(master_path, monthly_path)
        else:
            saved_path, _summary = resync_cupones_only(master_path)
    except NoPendingCouponsError as exc:
        flash(str(exc), "error")
        return redirect(url_for("eft"))
    except MonthlyReportFullyDuplicateError as exc:
        flash(str(exc), "error")
        return redirect(url_for("eft"))
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("eft"))

    return send_file(saved_path, as_attachment=True, download_name=master_filename)


@app.route("/proveedores")
def proveedores():
    return render_template("proveedores.html", **THEME_BY_KEY["proveedores"])


@app.route("/proveedores/facturas", methods=["POST"])
def proveedores_facturas():
    master_upload = request.files.get("master_file")
    pdf_uploads = request.files.getlist("pdf_files")
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel Ledger.", "error")
        return redirect(url_for("proveedores"))
    if not pdf_uploads or not any(u.filename for u in pdf_uploads):
        flash("Seleccioná uno o más PDF de facturas.", "error")
        return redirect(url_for("proveedores"))

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        pdf_paths = _save_uploads_to_workspace(pdf_uploads, workdir=workdir)
        temp_path, summary = append_supplier_invoices(master_path, pdf_paths)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("proveedores"))

    if summary["failed"]:
        detail = "; ".join(f"{item['filename']}: {item['error']}" for item in summary["failed"])
        flash(f"{len(summary['failed'])} factura(s) no se pudieron cargar y se omitieron: {detail}", "error")

    return send_file(temp_path, as_attachment=True, download_name=master_filename)


@app.route("/proveedores/pagos", methods=["POST"])
def proveedores_pagos():
    master_upload = request.files.get("master_file")
    bank_upload = request.files.get("bank_file")
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel Ledger.", "error")
        return redirect(url_for("proveedores"))
    if bank_upload is None or not bank_upload.filename:
        flash("Seleccioná el extracto de Chase ya categorizado.", "error")
        return redirect(url_for("proveedores"))

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        bank_path, _bank_filename = _save_upload_to_workspace(bank_upload, workdir=workdir)
        temp_path, _summary = append_supplier_payments(master_path, bank_path)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("proveedores"))

    return send_file(temp_path, as_attachment=True, download_name=master_filename)


def _format_date_amounts(date_amounts):
    return ", ".join(
        f"{d.strftime('%d/%m')} (${amount:,.2f})" for d, amount in sorted(date_amounts.items())
    )


@app.route("/caja")
def caja():
    return render_template("caja.html", **THEME_BY_KEY["caja"])


@app.route("/caja/chase", methods=["POST"])
def caja_chase():
    master_upload = request.files.get("master_file")
    chase_upload = request.files.get("chase_file")
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel Cierre.", "error")
        return redirect(url_for("caja"))
    if chase_upload is None or not chase_upload.filename:
        flash("Seleccioná el Excel de Chase ya categorizado.", "error")
        return redirect(url_for("caja"))

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        chase_path, _chase_filename = _save_upload_to_workspace(chase_upload, workdir=workdir)
        temp_path, summary = apply_chase_deposits(master_path, chase_path)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("caja"))

    if summary["deposits_unmatched"]:
        flash(
            "Depósitos sin fecha en CAJA (no se cargaron): "
            + _format_date_amounts(summary["deposits_unmatched"]),
            "error",
        )
    if summary["food_ice_unmatched"]:
        flash(
            "Food Truck/Hielo sin fecha en CAJA (no se cargaron): "
            + _format_date_amounts(summary["food_ice_unmatched"]),
            "error",
        )
    return send_file(temp_path, as_attachment=True, download_name=master_filename)


@app.route("/caja/lottery", methods=["POST"])
def caja_lottery():
    master_upload = request.files.get("master_file")
    lottery_upload = request.files.get("lottery_file")
    if master_upload is None or not master_upload.filename:
        flash("Seleccioná el Excel Cierre.", "error")
        return redirect(url_for("caja"))
    if lottery_upload is None or not lottery_upload.filename:
        flash("Seleccioná el Excel de Lottery.", "error")
        return redirect(url_for("caja"))

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        lottery_path, _lottery_filename = _save_upload_to_workspace(lottery_upload, workdir=workdir)
        temp_path, summary = apply_lottery_cuenta_final(master_path, lottery_path)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("caja"))

    if summary["unmatched"]:
        flash(
            "Fechas de Lottery sin fila en CAJA (no se cargaron): "
            + _format_date_amounts(summary["unmatched"]),
            "error",
        )
    if summary["missing_cached_value"]:
        dates_str = ", ".join(d.strftime("%d/%m") for d in sorted(summary["missing_cached_value"]))
        flash(
            "Días sin CUENTA FINAL calculada en el Lottery (abrilo y guardalo en Excel para "
            f"que recalcule las fórmulas, después volvé a intentar): {dates_str}",
            "error",
        )

    return send_file(temp_path, as_attachment=True, download_name=master_filename)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
