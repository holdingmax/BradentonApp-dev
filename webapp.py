"""
BradentonApp — versión web (Flask).

Primer paso de la migración de escritorio (Tkinter) a web, para poder
correr como servicio en Render y entrar en Toolbox. Reusa la lógica de
negocio ya extraída a módulos sin dependencia de Tkinter (chase_rules.py,
cmv_costo.py, etc.) — nunca reimplementa esa lógica acá.

Un módulo por vez: hoy solo está Chase Bank. El resto se va sumando
igual que el desktop, probando cada uno antes de seguir con el próximo.
"""

import calendar
import json
import os
import re
import sqlite3
import shutil
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from urllib.parse import quote, urlsplit

from flask import Flask, abort, flash, jsonify, redirect, render_template, request, send_file, session, url_for
from flask_login import (
    LoginManager,
    UserMixin,
    current_user,
    login_required,
    login_user,
    logout_user,
)

import auth
import chase_db
import cheques_db
import depositos
import depositos_db
import control_cmv
import cuenta_kia_toyota
import cuenta_kia_toyota_db
import control_cmv_db
import control_depositos
import ice_machine
import ice_machine_db
import control_cierre
import reporte_mensual
import reporte_mensual_db
import control_tarjetas
import cupones_detalle
import jh_mensual
import jh_mensual_db
import lottery_mensual
import controles_rapidos
from cheques import check_number_from_chase_description, extract_checks_from_pdf
from chase_rules import (
    add_dynamic_rule as add_chase_rule,
    build_chase_export_workbook,
    build_chase_pdf_report,
    categorize_chase_description,
    delete_dynamic_rule_by_index as delete_chase_custom_rule,
    delete_master_rule_by_index as delete_chase_master_rule,
    edit_dynamic_rule_by_index as edit_chase_custom_rule,
    edit_master_rule_by_index as edit_chase_master_rule,
    extract_chase_transactions,
    list_display_rules as list_chase_display_rules,
)
import caja_db
import gastos_caja
from caja import (
    CHASE_DETALLE_FOOD_TRUCK,
    CHASE_DETALLE_ICE,
    build_caja_export_pdf,
    build_caja_export_workbook,
    build_month_report_from_db as build_caja_month_report,
    build_caja_pdf_resumen,
    get_available_years as get_caja_available_years,
)
CHASE_DETALLE_VACCUMMS = control_depositos.CHASE_DETALLE[control_depositos.VACCUMMS]
from cmv_costo import _consolidate_department_files, compare_cost_snapshots, update_master_costo_todos_bulk
import eft_db
from eft_cta_cte import EFT_DUPLICATE_ALERT, extract_eft_data
from cupones_append import expand_monthly_records_for_storage, read_monthly_coupon_rows
from gettel_toyota_parser import (
    detect_vendor_from_ocr_text,
    merge_gettel_toyota_into_master,
    merge_gettel_toyota_pdf_into_master,
    process_gettel_pagos,
    summarize_origin_workbook,
    summarize_pdf_report,
    VENDOR_GETTEL,
    VENDOR_TOYOTA,
)
from controles_caja import check_caja_mayores
from controles_cierre_mensual import check_department_sales_monthly, check_store_info_monthly
from controles_cupones import check_cupones_pending
from controles_lottery_mensual import check_lottery_monthly
from controles_mercaderia import check_mercaderia_invoices
from controles_valuacion import check_and_complete_valuation
from monthly_sales import _resolve_sheet_name, parse_monthly_sales_file, process_monthly_sales
from proveedores import append_supplier_invoices, append_supplier_payments, extract_invoices_from_pdf
from proveedores import _PDF_EXTRACTION_EXCEPTIONS, invoice_errors
from proveedores import list_supplier_registry_entries, match_supplier_for_chase_description
import proveedores_pago_rules
from proveedores_dynamic_extractors import (
    FIELD_LABELS as DYNAMIC_FIELD_LABELS,
    FIELDS as DYNAMIC_FIELDS,
    PDF_READ_EXCEPTIONS as DYNAMIC_PDF_READ_EXCEPTIONS,
    add_dynamic_supplier,
    analyze_sample as analyze_dynamic_sample,
    build_rule_fields as build_dynamic_rule_fields,
    delete_dynamic_supplier,
    extract_with_dynamic_rule,
    load_dynamic_suppliers,
)
from reporte_diario import (
    DEPARTMENT_GROUPS,
    DepartmentPagePrefetch,
    build_store_info_export_pdf,
    build_store_info_export_workbook,
    build_store_info_pdf_resumen,
    extract_department_sales_for_day,
    extract_lottery_department_fields_from_pdf,
    extract_lottery_receipt_fields_from_sales_report,
    extract_store_info_for_day,
    group_department_sales,
    process_lottery,
    process_reporte_diario,
    process_store_info,
    real_store_info_total_sales,
)
import proyecciones
import fisico
import fisico_db
import fisico_invoice_parser
import reportes_db
import lottery_db
import gettel_db
import gettel_pagos as gettel_pagos_logic
import gettel_pagos_parser
import gettel_reportes
import cmv_db
import documents_db
import proveedores_db
import proveedores_productos
import horas_trabajo_db
import horas_trabajo
from horas_trabajo import extract_hours_report
import jobs
from balance_mensual import replace_mayor_sheets

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


# Archivos temporales (auditoría 2026-09, webapp.py:808/1142/826): cada
# subida y cada export crean su carpeta temporal y ningún camino la borraba
# (en esta PC ya había ~430 carpetas, 265 MB de copias de PDFs y Excel
# contables). Todo lo temporal del proceso va a una carpeta propia
# (tempfile.tempdir) y un hilo borra lo que tenga más de 2 días -- ningún
# job ni descarga dura tanto. También limpia las carpetas viejas que quedaron
# sueltas en el temporal del sistema antes de este cambio.
_SYSTEM_TEMP_DIR = tempfile.gettempdir()
_APP_TEMP_DIR = os.path.join(_SYSTEM_TEMP_DIR, "bradenton_app")
os.makedirs(_APP_TEMP_DIR, exist_ok=True)
tempfile.tempdir = _APP_TEMP_DIR
_TEMP_MAX_AGE_SECONDS = 2 * 24 * 3600
_LEGACY_TEMP_PREFIXES = (
    "bradenton_web_", "caja_export_", "caja_reporte_", "chase_export_", "chase_reporte_",
    "eft_reporte_", "fisico_export_", "gettel_pagos_export_", "gettel_reporte_", "lottery_export_",
    "lottery_reporte_", "proveedores_reporte_", "reporte_diario_reporte_", "storeinfo_export_",
)


def _remove_old_entries(folder, max_age, prefixes=None):
    now = time.time()
    try:
        names = os.listdir(folder)
    except OSError:
        return
    for name in names:
        if prefixes is not None and not name.startswith(prefixes):
            continue
        path = os.path.join(folder, name)
        try:
            if now - os.path.getmtime(path) < max_age:
                continue
            if os.path.isdir(path):
                shutil.rmtree(path, ignore_errors=True)
            else:
                os.remove(path)
        except OSError:
            continue


def _temp_cleanup_loop():
    while True:
        _remove_old_entries(_APP_TEMP_DIR, _TEMP_MAX_AGE_SECONDS)
        _remove_old_entries(_SYSTEM_TEMP_DIR, _TEMP_MAX_AGE_SECONDS, _LEGACY_TEMP_PREFIXES)
        time.sleep(6 * 3600)


threading.Thread(target=_temp_cleanup_loop, daemon=True, name="temp-cleanup").start()

app = Flask(__name__)
app.secret_key = _load_or_create_secret_key()
# Protección CSRF (auditoría 2026-09, webapp.py:145/302): las cookies de
# sesión no viajan en un POST que arranca en otro sitio (SameSite=Lax), y
# además _reject_cross_site_posts (abajo) rechaza todo POST cuyo Origin o
# Referer sea de otro host. En Render (HTTPS) las cookies van además Secure.
app.config.update(
    SESSION_COOKIE_SAMESITE="Lax",
    REMEMBER_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_HTTPONLY=True,
    REMEMBER_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SECURE=bool(os.environ.get("RENDER")),
    REMEMBER_COOKIE_SECURE=bool(os.environ.get("RENDER")),
)

# Costos por unidad sin redondear (Productos, Precios de la botonera).
app.jinja_env.filters["exact_money"] = proveedores_productos.exact_money

login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Iniciá sesión para continuar."
login_manager.login_message_category = "error"


class WebUser(UserMixin):
    def __init__(self, username, is_admin, must_change_password=False):
        self.id = username
        self.is_admin = is_admin
        self.must_change_password = must_change_password


@login_manager.user_loader
def load_user(username):
    user = auth.get_user(username)
    if user is None:
        return None
    return WebUser(
        username,
        user.get("is_admin", False),
        user.get("must_change_password", False),
    )


@app.before_request
def _reject_cross_site_posts():
    """
    Defensa CSRF sin tokens: un POST/PUT/PATCH/DELETE solo se acepta si el
    navegador dice que viene de esta misma app (Origin, o Referer si no hay
    Origin). Los navegadores siempre mandan Origin en un POST; un pedido sin
    ninguno de los dos (un script, el test client) se deja pasar porque no
    es un ataque desde otra página abierta en el navegador del usuario.
    """
    if request.method in ("GET", "HEAD", "OPTIONS"):
        return None
    source = request.headers.get("Origin") or request.headers.get("Referer")
    if source == "null":
        return ("Pedido rechazado: origen no válido.", 403)
    if not source:
        return None
    if urlsplit(source).netloc.lower() != request.host.lower():
        return ("Pedido rechazado: viene de otro sitio.", 403)
    return None


@app.before_request
def require_login():
    if request.endpoint in ("login", "static") or request.endpoint is None:
        return None
    if not current_user.is_authenticated:
        return redirect(url_for("login"))
    return None


# Cambio de contraseña obligatorio de "primer login" -- endpoints que SÍ
# tienen que seguir funcionando aunque el usuario todavía lo tenga pendiente
# (la propia página de cambio, y salir).
_PASSWORD_CHANGE_EXEMPT_ENDPOINTS = {"perfil_password", "logout", "static"}


@app.before_request
def _require_password_change():
    if not current_user.is_authenticated:
        return None
    if request.endpoint in _PASSWORD_CHANGE_EXEMPT_ENDPOINTS or request.endpoint is None:
        return None
    if getattr(current_user, "must_change_password", False):
        return redirect(url_for("perfil_password"))
    return None


@app.after_request
def _download_error_as_message(response):
    """
    Descargas sin recargar (pedido del usuario, 2026-10-05): base.html pide
    cada PDF/Excel por detrás con el encabezado X-Download-Check. Si la ruta
    no puede armar el archivo (no hay datos ese mes, el PDF ya no está
    guardado...) hace lo de siempre, avisar con flash y redirigir; acá ese
    redirect se cambia por el mensaje solo, que el navegador muestra en el
    aviso flotante sin recargar la página. Sin el encabezado (un link abierto
    a mano, otra pestaña) todo sigue igual que antes.
    """
    if not request.headers.get("X-Download-Check") or not 300 <= response.status_code < 400:
        return response
    flashes = session.pop("_flashes", None) or []
    if flashes:
        category, message = next(((c, m) for c, m in flashes if c == "error"), flashes[0])
    elif urlsplit(response.location or "").path == url_for("login"):
        category, message = "error", "La sesión venció: volvé a entrar."
    else:
        category, message = "error", "No se pudo generar el archivo."
    message_response = jsonify({"error": str(message), "level": "error" if category == "error" else "warning"})
    message_response.status_code = 409
    return message_response


# Endpoints alcanzables desde LOS DOS lados de la bifurcación Carga de
# Datos/Excels (ej. /reporte/historial, linkeado tanto desde reporte.html
# como desde la barra lateral global) o transversales a los dos (cuenta,
# manual) -- no cambian session["app_side"], solo lo leen. Ver
# _track_app_side más abajo.
_NEUTRAL_SIDE_ENDPOINTS = {
    "reporte_historial",
    "reporte_store_info_historial",
    "reporte_dia",
    "reporte_dia_pdf",
    "reporte_dia_departamentos",
    "reporte_dia_store_info",
    "manual",
    "perfil_password",
    "admin_users",
}


@app.before_request
def _track_app_side():
    """
    Recuerda de qué lado de la bifurcación Carga de Datos/Excels está el
    usuario (session["app_side"]) para que las páginas COMPARTIDAS (ej.
    /reporte/historial, sin ningún prefijo de URL que las distinga) sepan
    qué barra de navegación mostrar arriba -- pedido explícito del usuario
    (2026-09-11): "ninguna redirección puede llevarte a la página de excel"
    estando del lado de Carga de Datos, y viceversa. El prefijo de la URL
    solo no alcanza para esas páginas neutras, hace falta memoria de sesión.
    Default "carga_datos" (no "excels") desde que se sacó el chooser inicial
    (ver "/") -- ese es el lado al que cae cualquiera apenas inicia sesión.
    """
    if not current_user.is_authenticated:
        return
    if request.endpoint is None or request.endpoint in ("login", "static", "logout"):
        return
    if request.endpoint in _NEUTRAL_SIDE_ENDPOINTS:
        session.setdefault("app_side", "carga_datos")
        return
    session["app_side"] = "carga_datos" if request.path.startswith("/carga-datos") else "excels"


# Rutas que solo sirven para ver, subir o listar archivos originales -- con
# el guardado de documentos apagado (documents_db.GUARDAR_DOCUMENTOS, pedido
# explícito del usuario 2026-09-28: "que la página solo sirva para extraer
# los datos... para que no haga falta el documento") quedan fuera de
# servicio y mandan al Menú. El código de cada ruta sigue intacto, para
# reactivarlo con solo volver el interruptor a True.
_DOCUMENT_ENDPOINTS = {
    "documentos_index",
    "carga_datos_documentos",
    "carga_datos_documentos_subir",
    "carga_datos_documento_descargar",
    "carga_datos_documento_eliminar",
    "reporte_documentos",
    "reporte_documentos_mensual_subir",
    "carga_datos_lottery_documentos",
    "carga_datos_lottery_documentos_mensual_subir",
    "reporte_dia_pdf",
    "carga_datos_lottery_dia_pdf",
    "chase_cheque_pdf",
    "controles_deposito_pdf",
}


@app.before_request
def _block_document_routes():
    if documents_db.GUARDAR_DOCUMENTOS or request.endpoint not in _DOCUMENT_ENDPOINTS:
        return None
    flash("Los documentos originales ya no se guardan: los datos quedan cargados en cada módulo.", "warning")
    return redirect(url_for("home"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("home"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = auth.verify_user(username, password)
        if user is None:
            flash("Usuario o contraseña incorrectos.", "error")
        else:
            remember = bool(request.form.get("remember"))
            login_user(
                WebUser(username, user.get("is_admin", False), user.get("must_change_password", False)),
                remember=remember,
            )
            if user.get("must_change_password", False):
                return redirect(url_for("perfil_password"))
            return redirect(url_for("home"))

    return render_template("login.html")


# POST: con GET cualquier página externa podía cerrar la sesión con un link
# o una imagen (auditoría 2026-09).
@app.route("/logout", methods=["POST"])
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


@app.route("/perfil/password", methods=["GET", "POST"])
@login_required
def perfil_password():
    forced = getattr(current_user, "must_change_password", False)
    if request.method == "POST":
        current_password = request.form.get("current_password", "")
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")
        if auth.verify_user(current_user.id, current_password) is None:
            flash("La contraseña actual no es correcta.", "error")
        elif new_password != confirm_password:
            flash("Las contraseñas nuevas no coinciden.", "error")
        else:
            try:
                auth.set_password(current_user.id, new_password)
                if forced:
                    flash("Contraseña actualizada. Ya podés usar la app normalmente.", "success")
                    return redirect(url_for("home"))
                flash("Contraseña actualizada correctamente.", "success")
            except ValueError as exc:
                flash(str(exc), "error")
        return redirect(url_for("perfil_password"))
    return render_template("perfil_password.html", forced=forced)


@app.route("/admin/users", methods=["GET", "POST"])
@login_required
def admin_users():
    if not current_user.is_admin:
        flash("No tenés permiso para acceder a esta página.", "error")
        return redirect(url_for("carga_datos_index"))

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
                new_password = request.form.get("new_password", "")
                confirm_password = request.form.get("confirm_password", "")
                if new_password != confirm_password:
                    flash("Las contraseñas no coinciden.", "error")
                else:
                    auth.set_password(
                        request.form.get("username", "").strip(),
                        new_password,
                        must_change_password=True,
                    )
                    flash("Contraseña actualizada. Esa persona va a tener que cambiarla al iniciar sesión.", "success")
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


@app.route("/manual")
@login_required
def manual():
    return render_template("manual.html")

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
_ICON_CHECKLIST = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 3h6a1 1 0 0 1 1 1v1H8V4a1 1 0 0 1 1-1z"/><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M8.5 12.5l2 2 4-4"/></svg>'
_ICON_SCALE = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v18"/><path d="M8 21h8"/><path d="M5 7h14"/><path d="M5 7l-3.5 6.5a3.2 3.2 0 0 0 6.4 0z"/><path d="M19 7l-3.5 6.5a3.2 3.2 0 0 0 6.4 0z"/></svg>'
# Ícono genérico para resultados de búsqueda de archivos (PDF/Excel ya
# guardados) -- pedido explícito del usuario (2026-09-16), distinto del
# ícono de cada módulo para que se note de un vistazo que es un archivo.
_ICON_FILE = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>'
_ICON_FUEL = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="3" y1="22" x2="15" y2="22"/><line x1="4" y1="9" x2="14" y2="9"/><path d="M14 22V4a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v18"/><path d="M14 13h2a2 2 0 0 1 2 2v2a2 2 0 0 0 2 2v0a2 2 0 0 0 2-2V9.83a2 2 0 0 0-.59-1.42L18 5"/></svg>'
_ICON_CLOCK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3.5 2"/></svg>'

TOOLS = [
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
        "key": "proveedores",
        "code": "PR",
        "icon": _ICON_TRUCK,
        "label": "Proveedores",
        "url": "/proveedores",
        "description": "Facturas de compra por proveedor y pagos vía Chase al Cta Cte.",
        "accent": "#DB2777",
        "accent_soft": "#FBD9EA",
    },
]

# Balance Mensual (`/balance-mensual`, balance_mensual.py) -- construido en
# una sesión anterior pero TODAVÍA NO CONFIRMADO por el usuario ("no es algo
# que este listo", pedido explícito 2026-09-14) -- sacado de TOOLS para que
# no aparezca ni en la grilla de Herramientas ni en la búsqueda del header
# como si fuera un módulo terminado. La ruta y el código siguen intactos,
# solo dejaron de anunciarse -- si el usuario confirma que está listo,
# devolver esta entrada a TOOLS de nuevo (no reescribir nada).

# Segunda sección de la app, hermana de Herramientas (TOOLS): cada entrada acá
# es un módulo que recibe un Excel ya cerrado de fin de mes y verifica que
# esté en orden, en vez de transformarlo. Ver CLAUDE.md, "Módulo Controles".
CONTROLS = [
    {
        "key": "cierre_mensual",
        "code": "CM",
        "icon": _ICON_CHECKLIST,
        "label": "Cierre Mensual",
        "url": "/controles/cierre-mensual",
        "description": "Cruza el Resumen de Ventas mensual del POS contra Store Info del Excel Cierre.",
        "accent": "#334155",
        "accent_soft": "#E2E8F0",
    },
    {
        "key": "lottery_mensual",
        "code": "LT",
        "icon": _ICON_CHECKLIST,
        "label": "Lottery Mensual",
        "url": "/controles/lottery-mensual",
        "description": "Cruza el Monthly Sales Report de Florida Lottery contra el Excel de Lottery del mes.",
        "accent": "#0284C7",
        "accent_soft": "#D7EFFB",
    },
    {
        "key": "cupones",
        "code": "CP",
        "icon": _ICON_CHECKLIST,
        "label": "Cupones",
        "url": "/controles/cupones",
        "description": "Cruza el saldo del Mayor de Recaudación a Liquidar contra los cupones sin aplicar a un EFT.",
        "accent": "#3B5BDB",
        "accent_soft": "#DDE3FA",
    },
    {
        "key": "mercaderia",
        "code": "MC",
        "icon": _ICON_CHECKLIST,
        "label": "Mercadería",
        "url": "/controles/mercaderia",
        "description": "Cruza las facturas del mes en Proveedores contra el Mayor de Mercadería en C-Store.",
        "accent": "#DB2777",
        "accent_soft": "#FBD9EA",
    },
    {
        "key": "valuacion",
        "code": "VL",
        "icon": _ICON_CHECKLIST,
        "label": "Valuación de Existencia Final",
        "url": "/controles/valuacion",
        "description": "Cruza la Existencia Final según contabilidad contra la valuación de Chevron Category Cost Report.",
        "accent": "#059669",
        "accent_soft": "#D1FAE5",
    },
    {
        "key": "caja",
        "code": "CJ",
        "icon": _ICON_CHECKLIST,
        "label": "Caja",
        "url": "/controles/caja",
        "description": "Cruza depósitos Ice Machine/Food Truck y columna K, y gastos en efectivo (columna M), contra los Mayores de Chase y de Caja.",
        "accent": "#EA580C",
        "accent_soft": "#FCE3D2",
    },
]

# Nueva sección "Controles" del Menú -- pedido explícito del usuario
# (2026-09-28), reemplaza a la tarjeta Documentos: controles del mes completo
# hechos con los datos YA cargados en Carga de Datos, sin Excel de por
# medio. Arranca vacía salvo Depósitos (que vivía en Documentos) y se va
# llenando de a uno. Los 6 controles viejos de arriba (CONTROLS, basados en
# Excel) quedan ocultos: siguen andando por URL directa, pero no se listan
# ni aparecen en la búsqueda.
_ICON_DEPOSIT = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="6" width="20" height="12" rx="2"/><circle cx="12" cy="12" r="2.5"/><path d="M6 12h.01M18 12h.01"/></svg>'

# "Cierre" reemplazó a "Caja" (Depósitos contra Chase) en la lista (pedido
# del usuario, 2026-10-06). Desde el 2026-10-07 /controles/depositos es
# "Control Depósitos": todos los depósitos y los pagos de la máquina de hielo
# contra Chase (control_depositos.py).
CONTROLES_SECTIONS = [
    {
        "key": "control_cierre",
        "code": "CI",
        "icon": _ICON_SCALE,
        "label": "Cierre",
        "url": "/controles/cierre",
        "description": "Los asientos de cierre del mes.",
        "accent": "#16A34A",
        "accent_soft": "#DCF3E3",
    },
    {
        "key": "control_lottery",
        "code": "LT",
        "icon": _ICON_TICKET,
        "label": "Lottery",
        "url": "/controles/lottery",
        "description": "El mes de Lottery contra los reportes diarios y los pagos en Chase.",
        "accent": "#EA580C",
        "accent_soft": "#FDE3D3",
    },
    {
        "key": "control_tarjetas",
        "code": "TC",
        "icon": _ICON_EXCHANGE,
        "label": "Tarjetas y Cupones",
        "url": "/controles/tarjetas",
        "description": "Las ventas con tarjeta contra lo que acredita J.H.",
        "accent": "#3B5BDB",
        "accent_soft": "#DDE3FA",
    },
    {
        "key": "control_depositos",
        "code": "DP",
        "icon": _ICON_DEPOSIT,
        "label": "Control Depósitos",
        "url": "/controles/depositos",
        "description": "Los depósitos del mes contra Chase.",
        "accent": "#0891B2",
        "accent_soft": "#D5F0F6",
    },
    {
        "key": "control_cmv",
        "code": "CV",
        "icon": _ICON_COINS,
        "label": "Control CMV",
        "url": "/controles/cmv",
        "description": "Las ventas de cada departamento: reportes diarios, reporte mensual, Elistar y CMV.",
        "accent": "#7C3AED",
        "accent_soft": "#E9E0FC",
    },
    {
        "key": "control_kia_toyota",
        "code": "KT",
        "icon": _ICON_CAR,
        "label": "Kia y Toyota",
        "url": "/controles/kia-toyota",
        "description": "Cuánto debe cada concesionaria: lo que cargan a cuenta menos lo que pagan.",
        "accent": "#0D9488",
        "accent_soft": "#D6F1EE",
    },
    {
        "key": "control_productos",
        "code": "PR",
        "icon": _ICON_TRUCK,
        "label": "Productos de proveedores",
        "url": "/controles/productos",
        "description": "Los cambios de precio de los productos de cada proveedor.",
        "accent": "#DB2777",
        "accent_soft": "#FBD9EA",
    },
]

# Tercera sección de la app, hermana de Herramientas/Controles pero del OTRO
# lado de la bifurcación "Carga de Datos vs. Excels" (ver CLAUDE.md) -- cada
# entrada acá es un módulo que YA NO escribe ningún Excel, solo lee/extrae y
# guarda en una base propia (mismo patrón, key/code/icon/label/url/
# description/accent, que TOOLS/CONTROLS). Se va sumando de a uno, en el
# orden en que cada módulo de Herramientas se convierte (empezando por
# Reporte Diario/Lottery, después Chase Bank) -- ver "Conversión módulo por
# módulo a Carga de Datos" en CLAUDE.md.
CARGA_DATOS_TOOLS = [
    {
        "key": "carga_reporte",
        "code": "RD",
        "icon": _ICON_CALENDAR,
        "label": "Reportes Diario/Mensual",
        "url": "/carga-datos/reporte-diario",
        "description": "Los reportes de cierre del día y del mes.",
        "accent": "#0284C7",
        "accent_soft": "#D7EFFB",
    },
    {
        "key": "carga_lottery",
        "code": "LT",
        "icon": _ICON_TICKET,
        "label": "Lottery",
        "url": "/carga-datos/lottery",
        "description": "Los reportes de ventas de Lottery.",
        "accent": "#0284C7",
        "accent_soft": "#D7EFFB",
    },
    {
        "key": "carga_chase",
        "code": "CH",
        "icon": _ICON_BANK,
        "label": "Chase Bank",
        "url": "/carga-datos/chase",
        "description": "El extracto del banco, categorizado solo.",
        "accent": "#16A34A",
        "accent_soft": "#DCF3E3",
    },
    {
        "key": "carga_eft",
        "code": "EFT",
        "icon": _ICON_EXCHANGE,
        "label": "EFT y Cupones",
        "url": "/carga-datos/eft",
        "description": "Los EFT y cupones de J.H. Williams.",
        "accent": "#3B5BDB",
        "accent_soft": "#DDE3FA",
    },
    {
        "key": "carga_caja",
        "code": "CJ",
        "icon": _ICON_REGISTER,
        "label": "Caja",
        "url": "/carga-datos/caja",
        "description": "El efectivo del mes, armado solo con lo ya cargado.",
        "accent": "#EA580C",
        "accent_soft": "#FCE3D2",
    },
    {
        "key": "carga_gettel",
        "code": "GT",
        "icon": _ICON_CAR,
        "label": "Gettel / Toyota",
        "url": "/carga-datos/gettel",
        "description": "Los cupones de combustible de Gettel y Toyota.",
        "accent": "#0D9488",
        "accent_soft": "#D6F1EE",
    },
    {
        "key": "carga_depositos",
        "code": "DP",
        "icon": _ICON_DEPOSIT,
        "label": "Depósitos",
        "url": "/carga-datos/depositos",
        "description": "Los comprobantes de depósito y los pagos de la máquina de hielo.",
        "accent": "#0891B2",
        "accent_soft": "#D5F0F6",
    },
    {
        # Nunca aparece en la grilla de Herramientas (ver el filtro de
        # carga_datos_index más abajo, mismo criterio que "carga_caja") --
        # el usuario probó tenerla como tarjeta propia acá y pidió sacarla
        # el mismo día (2026-09-21): "los pagos de los cupones van dentro
        # del modulo de gettel, no separados". Sigue existiendo solo para
        # que THEME_BY_KEY tenga su tema -- se llega siempre por la barra
        # lateral, panel "Gettel / Toyota" -> "Cargar Pagos"/"Cuadro de Pagos".
        "key": "carga_gettel_pagos",
        "code": "GP",
        "icon": _ICON_EXCHANGE,
        "label": "Gettel -- Pagos de Cupones",
        "url": "/carga-datos/gettel/pagos",
        "description": "Los recibos de pago de los cupones.",
        "accent": "#0D9488",
        "accent_soft": "#D6F1EE",
    },
    {
        "key": "carga_cmv",
        "code": "CMV",
        "icon": _ICON_COINS,
        "label": "CMV",
        "url": "/carga-datos/cmv",
        "description": "Costos y ventas por departamento.",
        "accent": "#7C3AED",
        "accent_soft": "#E9E0FC",
    },
    {
        "key": "carga_proveedores",
        "code": "PR",
        "icon": _ICON_TRUCK,
        "label": "Proveedores",
        "url": "/carga-datos/proveedores",
        "description": "Las facturas de compra.",
        "accent": "#DB2777",
        "accent_soft": "#FBD9EA",
    },
    {
        "key": "carga_horas",
        "code": "HT",
        "icon": _ICON_CLOCK,
        "label": "Horas de Trabajo",
        "url": "/carga-datos/horas-trabajo",
        "description": "Las horas de los empleados y su sueldo.",
        "accent": "#0891B2",
        "accent_soft": "#D3F0F4",
    },
    {
        "key": "carga_combustible",
        "code": "CB",
        "icon": _ICON_FUEL,
        "label": "Combustible",
        "url": "/carga-datos/combustible",
        "description": "Las compras de combustible.",
        "accent": "#92400E",
        "accent_soft": "#FDE8CE",
    },
    {
        # Nunca aparece en la grilla de Herramientas (ver el filtro de
        # carga_datos_index más abajo, mismo criterio que "carga_caja") --
        # existe solo para que THEME_BY_KEY tenga el tema de /fisico
        # (la página de "Cuadro del mes", que no es una carga en sí misma,
        # las facturas se cargan en /carga-datos/combustible de arriba).
        "key": "fisico",
        "code": "FI",
        "icon": _ICON_FUEL,
        "label": "Físico",
        "url": "/fisico",
        "description": "El combustible que debería haber contra el real.",
        "accent": "#92400E",
        "accent_soft": "#FDE8CE",
    },
]

# Módulo "Reportes" (pedido explícito del usuario, 2026-09-17) -- reemplaza
# al scaffold vacío que antes vivía en /carga-datos/controles ("vamos a
# cambiar el modulo de controles por ese nombre y vamos a empezar con el
# chase la opcion de exportar en un PDF"). Por ahora todos dan PDF nada más
# ("todos estos van a dar la opcion de PDF de momento luego vemos que
# incorporamos") -- `ready=False` en un tool de esta lista significa
# "Próximamente", mismo criterio visual que ya usaba /controles cuando
# arrancó sin ningún módulo (span.tool-card, sin link).
REPORTES_TOOLS = [
    {
        "key": "reportes_chase",
        "icon": _ICON_BANK,
        "label": "Chase Bank",
        "description": "El resumen del mes por categoría.",
        "accent": "#16A34A",
        "ready": True,
        "pdf_endpoint": "reportes_chase_pdf",
    },
    {
        "key": "reportes_eft",
        "icon": _ICON_EXCHANGE,
        "label": "EFT y Cupones",
        "description": "Los EFT y cupones del mes.",
        "accent": "#3B5BDB",
        "ready": True,
        "pdf_endpoint": "reportes_eft_pdf",
    },
    {
        "key": "reportes_diario",
        "icon": _ICON_CALENDAR,
        "label": "Reporte Diario",
        "description": "El resumen de ventas del mes.",
        "accent": "#0284C7",
        "ready": True,
        "pdf_endpoint": "reportes_diario_pdf",
    },
    {
        "key": "reportes_lottery",
        "icon": _ICON_TICKET,
        "label": "Lottery",
        "description": "El resumen de Lottery del mes.",
        "accent": "#0284C7",
        "ready": True,
        "pdf_endpoint": "reportes_lottery_pdf",
    },
    {
        "key": "reportes_caja",
        "icon": _ICON_REGISTER,
        "label": "Caja",
        "description": "El resumen del efectivo del mes.",
        "accent": "#EA580C",
        "ready": True,
        "pdf_endpoint": "reportes_caja_pdf",
    },
    {
        # Pedido explícito del usuario (2026-09-22): "resumen mensual en
        # reportes, donde se va a mostrar la cantidad de facturas que
        # llegaron de un proveedor y cual fue el total del mes" -- mismo
        # espíritu que la vieja hoja "RESUMEN COMPRAS" del Excel Ledger,
        # ver proveedores_db.build_proveedores_pdf_report.
        "key": "reportes_proveedores",
        "icon": _ICON_TRUCK,
        "label": "Proveedores",
        "description": "Las compras del mes por proveedor.",
        "accent": "#DB2777",
        "ready": True,
        "pdf_endpoint": "reportes_proveedores_pdf",
    },
    {
        # Pedido explícito del usuario (2026-09-21, corregido 2026-09-22):
        # "el reporte de gettel deberia ser un PDF, no un excel" -- mismas
        # 4 secciones que el Excel (Pendiente mes anterior / mes actual /
        # Pago Cupones / Pendiente mes actual resultante, ver
        # gettel_reportes.build_gettel_pdf_report), resumidas sin colores,
        # mismo criterio que el resto de los reportes de este módulo. El
        # Excel con el formato/colores real sigue existiendo -- se mudó al
        # propio módulo Gettel ("Cuadro del mes" en la barra lateral, ver
        # carga_datos_gettel_historial.html -> reportes_gettel_excel).
        "key": "reportes_gettel",
        "icon": _ICON_CAR,
        "label": "Gettel — Cupones",
        "description": "Los cupones del mes y lo que queda pendiente.",
        "accent": "#0D9488",
        "ready": True,
        "pdf_endpoint": "reportes_gettel_pdf",
    },
]

THEME_BY_KEY = {
    tool["key"]: {"accent": tool["accent"], "accent_soft": tool["accent_soft"]}
    for tool in TOOLS + CONTROLS + CARGA_DATOS_TOOLS
}

# Índice para la barra de búsqueda del header (base.html) -- pedido
# explícito del usuario 2026-09-06: buscar un módulo por nombre desde
# cualquier página, entendiendo palabras parecidas (typos), mostrando
# todos los módulos relacionados de las secciones a la vez. Se inyecta
# solo (sin que cada ruta tenga que pasarlo) porque base.html lo necesita
# en TODAS las páginas, no solo en los índices de sección.
@app.context_processor
def inject_search_index():
    # TOOLS (Herramientas/Excels: CMV, Gettel/Toyota, Reporte Diario, Lottery,
    # Proveedores) se sacó de la búsqueda -- pedido explícito del usuario
    # (2026-09-15): ya no usa ese lado para nada (solo Carga de Datos), y
    # buscar el nombre de un módulo y tocar el resultado equivocado lo
    # mandaba ahí sin querer, con un Excel pidiéndose de la nada. El módulo
    # en sí sigue existiendo (alcanzable por URL directa si hiciera falta),
    # solo se sacó de la búsqueda para no ofrecerlo por error.
    # `today` -- pedido explícito del usuario (2026-09-16): "si en algún
    # momento te encontrás parado en el mes actual, debajo de este diga
    # 'mes actual'" en vez del link "Ir al mes actual" -- se necesita en
    # las ~13 páginas con navegación de mes, así que se inyecta acá en vez
    # de agregarlo a mano en cada ruta.
    # `guardar_documentos` -- ver documents_db.GUARDAR_DOCUMENTOS: los
    # templates lo usan para ocultar los links "Ver PDF" de archivos que ya
    # estaban guardados de antes.
    return {
        "SEARCH_INDEX": CONTROLES_SECTIONS + CARGA_DATOS_TOOLS,
        "today": date.today(),
        "guardar_documentos": documents_db.GUARDAR_DOCUMENTOS,
    }


@app.route("/buscar/documentos")
def buscar_documentos():
    """
    Búsqueda por nombre de archivo (PDF/Excel ya guardados), aparte de la
    búsqueda de módulos -- pedido explícito del usuario (2026-09-16):
    "quiero que la barra de busqueda sirva para encontrar tanto como los
    modulos, como PDF por su nombre, y excels tambien por el nombre". La
    llama el JS de base.html vía fetch (debounced), nunca bloquea el
    renderizado de la página como el índice de módulos (que se manda
    siempre, entero, en cada página) -- acá se consulta bajo demanda, con
    un límite chico, porque con años de archivos guardados mandar TODO en
    cada página dejaría de ser viable.
    """
    query = (request.args.get("q") or "").strip()
    if len(query) < 2 or not documents_db.GUARDAR_DOCUMENTOS:
        return jsonify([])

    results = []
    for doc in documents_db.search_documents(query, limit=8):
        info = _DOCUMENTS_MODULES.get(doc["module"], {})
        parts = [info.get("title", doc["module"])]
        if doc.get("label"):
            parts.append(doc["label"])
        parts.append(f"{doc['month']:02d}/{doc['year']}")
        results.append({
            "label": doc["filename"],
            "description": " · ".join(parts),
            "url": url_for("carga_datos_documento_descargar", document_id=doc["id"]),
            "icon": _ICON_FILE,
        })

    for row in reportes_db.search_pdfs(query, limit=8):
        fname = os.path.basename(row["pdf_filename"])
        results.append({
            "label": fname,
            "description": f"Reporte Diario · {row['date']}",
            "url": url_for("reporte_dia_pdf", report_date=row["date"]),
            "icon": _ICON_FILE,
        })

    for row in lottery_db.search_pdfs(query, limit=8):
        fname = os.path.basename(row["pdf_filename"])
        kind_label = "PDF Diario" if row["kind"] == "department" else "Daily Sales Report"
        results.append({
            "label": fname,
            "description": f"Lottery ({kind_label}) · {row['date']}",
            "url": url_for("carga_datos_lottery_dia_pdf", report_date=row["date"], kind=row["kind"]),
            "icon": _ICON_FILE,
        })

    return jsonify(results[:20])


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


# ---------------------------------------------------------------------------
# Chequeo de fecha del nombre de archivo vs. la fecha leída del PDF --
# pedido explícito del usuario (2026-09-15): "por probar intente subir
# varios pdf con fecha de agosto y me dejo subirlos... tomar la fecha de
# algun lado, ya sea del titulo del pdf o de la fecha dentro de la
# factura, para evitar cargar pdf por error a un mes que no les
# corresponda". Aplica solo donde UN archivo = UNA fecha puntual (Reporte
# Diario, Lottery Daily Sales Report, EFT, facturas de Proveedores) -- los
# módulos donde un solo archivo cubre todo un mes (Gettel/Toyota, CMV
# Ventas) no tienen "la fecha equivocada" que chequear de esta forma.
# ---------------------------------------------------------------------------

# El lookbehind evita leer como DD-MM el final de una fecha completa
# ("Scan 2026-08-01" daba (8, 1) y rechazaba el PDF correcto).
_REPORTE_DIARIO_FILENAME_DATE_RE = re.compile(r"(?<![\d.\-])(\d{1,2})[.\-](\d{1,2})\s*$")
_FULL_DATE_DMY_RE = re.compile(r"(?<!\d)(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})(?!\d)")
_FULL_DATE_YMD_RE = re.compile(r"(?<!\d)(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})(?!\d)")


def _reporte_pdf_canonical_filename(business_date):
    """
    Nombre CLARO y fijo con el que se guarda la copia de cada PDF de
    Reporte Diario -- pedido explícito del usuario (2026-09-18): "que se
    guarden de forma clara... que cada día coincida con el día que se le
    manda". Antes se guardaba con el nombre tal cual lo subió el usuario
    (mismo criterio de "nunca sanitizar" que el resto del proyecto, pero
    acá terminaba guardando cualquier cosa -- un nombre de escaneo
    genérico, una copia "(1)", etc.) -- ahora, sin importar cómo se llamaba
    el archivo real, la copia guardada siempre usa este formato ("Close
    Store DD-MM.pdf", el mismo patrón que ya reconoce _reporte_filename_
    day_month) con la fecha real de negocio ya calculada (con el +1 día ya
    aplicado) -- así el nombre del archivo guardado siempre es fiel al día
    que representa, se pueda o no confiar en el nombre original.
    """
    return f"Close Store {business_date:%d-%m}.pdf"


def _reporte_filename_day_month(filename):
    """
    "Close Store DD-MM.pdf" (sin año, siempre al final del nombre) -- el
    formato ya conocido de este proyecto para Reporte Diario. Devuelve
    (día, mes) o None si el nombre no termina en ese patrón exacto -- nunca
    arriesga un falso positivo adivinando sobre un nombre distinto.
    """
    stem = os.path.splitext(filename)[0]
    match = _REPORTE_DIARIO_FILENAME_DATE_RE.search(stem)
    if not match:
        return None
    day, month = int(match.group(1)), int(match.group(2))
    if not (1 <= day <= 31 and 1 <= month <= 12):
        return None
    return (day, month)


def _filename_date_candidates(filename):
    """
    Cualquier fecha completa (día+mes+año) que aparezca en el nombre de
    archivo, en cualquier separador (./-) y en cualquiera de los dos
    órdenes (día-mes o mes-día, ya que este proyecto tiene archivos de
    ambos estilos -- "Invoice N DD.MM.YYYY.pdf" vs. reportes de EE.UU. tipo
    "..._8-1-2026.pdf") -- se prueban las dos lecturas y se descartan las
    que no sean una fecha válida. Un año de 2 dígitos se interpreta como
    20XX. Nunca produce falsos positivos sobre un N° de factura suelto
    porque exige un separador entre los 3 grupos, no solo dígitos pegados.
    """
    stem = os.path.splitext(filename)[0]
    candidates = set()
    for match in _FULL_DATE_DMY_RE.finditer(stem):
        a, b, y = match.groups()
        year = int(y) if len(y) == 4 else 2000 + int(y)
        for day, month in ((int(a), int(b)), (int(b), int(a))):
            try:
                candidates.add(date(year, month, day))
            except ValueError:
                pass
    for match in _FULL_DATE_YMD_RE.finditer(stem):
        y, a, b = match.groups()
        year = int(y)
        for day, month in ((int(a), int(b)), (int(b), int(a))):
            try:
                candidates.add(date(year, month, day))
            except ValueError:
                pass
    return candidates


def _filename_date_mismatch(filename, target_date):
    """
    True si el nombre de archivo trae una fecha completa (día+mes+año) y
    NINGUNA coincide con `target_date` -- False si coincide alguna, o si el
    nombre no trae ninguna fecha completa (no hay nada que chequear, se
    deja pasar como siempre). `target_date` acepta date o datetime.

    Bug real corregido (2026-09-17): `datetime` es subclase de `date`, así
    que `isinstance(target_date, date)` da True también para un
    `datetime` -- la condición vieja (`not isinstance(...)`) nunca
    disparaba la conversión a `.date()`, y comparar un `datetime` contra
    un `set` de `date` (los candidatos del nombre de archivo) siempre da
    "no coincide" en Python aunque el día/mes/año sean idénticos. Esto
    rechazaba SIEMPRE los EFT (`parsed` ahí es un `datetime.strptime`,
    nunca un `date` plano) con "la fecha leída no coincide con la del
    nombre de archivo" pese a coincidir de verdad -- confirmado contra 4
    EFT reales de julio-2026 que el usuario no podía cargar. Chequear
    `isinstance(target_date, datetime)` en vez de `date` es lo correcto:
    solo un `datetime` de verdad necesita bajar a `.date()` antes de
    comparar contra los candidatos (que siempre son `date` planos).
    """
    if isinstance(target_date, datetime):
        target_date = target_date.date()
    candidates = _filename_date_candidates(filename)
    if not candidates:
        return False
    return target_date not in candidates


def _is_ajax_request():
    """
    El JS genérico de base.html (form.ajax-process-form) manda este header
    cuando procesa el formulario por fetch en vez de dejar que el navegador
    navegue -- ahí un flash()+redirect no sirve (no hay recarga de página
    de por medio para mostrarlo, y si la respuesta es un archivo -- ver
    _success_response -- ni siquiera hay redirect), así que hace falta
    responder con JSON en error o con un header de aviso en éxito, en vez
    de la sesión de flash de Flask.
    """
    return request.headers.get("X-Ajax-Request") == "1"


def _error_response(message):
    """
    Reporta una falla dura (nada se pudo procesar). Un pedido por fetch
    recibe JSON así el popup de base.html lo muestra al instante; un submit
    de formulario común (sin JS) cae al flash()+redirect de siempre -- acá
    sí funciona porque todavía no hay ningún send_file de por medio.
    """
    if _is_ajax_request():
        return jsonify({"error": message}), 400
    flash(message, "error")
    return redirect(request.referrer or url_for("carga_datos_index"))


def _open_result_for_user(path):
    """
    Abre el resultado ya procesado en Excel, en esta misma PC -- pedido
    explícito del usuario (2026-09-02): quiere verlo al toque en vez de ir a
    buscarlo a la carpeta Descargas.

    Solo se llama acá, desde el handler de una request HTTP real -- nunca
    desde los módulos de negocio (cmv_costo.py, proveedores.py, etc.), que
    Claude también llama directo (sin pasar por Flask) para verificar un fix
    antes de pedirle al usuario que lo pruebe él mismo. Mantener esta
    llamada fuera de esos módulos es lo que garantiza que esas verificaciones
    nunca abran Excel solas en la PC del usuario (ver "Nunca abrir Excel en
    la PC del usuario" en CLAUDE.md) mientras que un click real en
    "Procesar" sí lo abre.
    """
    if os.name != "nt":
        # os.startfile no existe fuera de Windows (Render corre Linux) --
        # ahí no hay "abrir en Excel" posible, el navegador ya sirve la
        # descarga igual, así que no hacemos nada.
        return
    try:
        os.startfile(path)
    except OSError:
        pass


def _success_response(temp_path, download_name, notice=None, notice_level="warning"):
    """
    Abre el archivo procesado en Excel (ver _open_result_for_user) y lo sirve
    en la respuesta -- ya no fuerza la descarga a la carpeta Descargas del
    navegador (el JS de base.html dejó de disparar esa descarga, pedido
    explícito del usuario), pero el archivo real sigue viajando en el blob
    de la respuesta porque Proveedores lo reusa para encadenar Facturas →
    Pagos sin que el usuario tenga que volver a seleccionarlo (ver
    chain-master-result en proveedores.html).

    El aviso opcional (éxito parcial: algo no se pudo cargar solo) tiene que
    mostrarse en el momento aunque la respuesta sea una descarga de archivo,
    nunca una página HTML -- un flash() acá quedaría en cola de sesión y
    aparecería recién en la próxima página que el usuario visite, fuera de
    contexto (bug real, ya documentado en CLAUDE.md). Un pedido por fetch
    recibe el aviso en un header que el JS de base.html muestra al toque,
    arriba de la página; un submit común (sin JS) no tiene forma de mostrar
    nada en el momento junto a una descarga, así que cae a flash() como
    mejor esfuerzo -- caso raro, todos los formularios de módulo ya mandan
    el pedido por fetch.
    """
    _open_result_for_user(temp_path)
    response = send_file(temp_path, as_attachment=True, download_name=download_name)
    if notice:
        if _is_ajax_request():
            response.headers["X-App-Notice"] = quote(notice)
            response.headers["X-App-Notice-Level"] = notice_level
        else:
            flash(notice, notice_level)
    return response


@app.route("/")
def home():
    """
    Menú real de entrada (2026-09-22, pedido explícito del usuario --
    revierte el redirect directo de 2026-09-12, ver el docstring viejo en
    el historial de git): "no deberia tener que elegir solo desde la barra
    de arriba... la idea es que cuando un usuario abra la pagina tenga que
    elegir entre esos dos modulos" -- Carga de Datos y Reportes, cada uno
    con su propia tarjeta grande (Proyecciones sigue sin mostrarse, mismo
    criterio que el section-switch de `base.html`). `login()` redirige acá
    en vez de directo a `carga_datos_index`.
    """
    return render_template("home_menu.html")


@app.route("/excels")
def excels_index():
    # Ya no muestra la grilla de Herramientas -- pedido explícito del
    # usuario (2026-09-15): "el único modulo que le cargamos un excel es a
    # chase y a CMV [y eso ya lo hace por Carga de Datos]... no hace falta"
    # -- confirmó que no usa nada de este lado. Mismo criterio que "/" (ver
    # home() arriba): redirige en vez de 404, para que un link/favorito
    # viejo no se encuentre con un error -- solo deja de mostrar contenido
    # de Excels. Los módulos en sí (TOOLS, sus rutas propias como /cmv,
    # /reporte, etc.) no se tocaron -- si el usuario pide sacarlos también,
    # es un pedido aparte, módulo por módulo (mismo criterio de siempre).
    return redirect(url_for("carga_datos_index"))


@app.route("/carga-datos")
def carga_datos_index():
    # Caja no tiene ningún upload propio (cruza en el momento lo que ya
    # guardaron Chase/Lottery/Reporte Diario) -- pedido explícito del
    # usuario (2026-09-12, cuarta tanda): no le corresponde una tarjeta acá
    # ("no se le tiene que cargar ningun PDF o excel para completar"), solo
    # queda accesible desde la barra lateral (grupo Book Keeping).
    # Gettel -- Pagos de Cupones (2026-09-21, corrección de la misma tarde):
    # el usuario pidió sacarla de acá -- "los pagos de los cupones van
    # dentro del modulo de gettel, no separados" -- vuelve a ser accesible
    # SOLO desde la barra lateral (panel "Gettel / Toyota" -> "Cargar
    # Pagos"), mismo criterio que carga_caja/fisico.
    upload_tools = [
        tool for tool in CARGA_DATOS_TOOLS
        if tool["key"] not in ("carga_caja", "fisico", "carga_gettel_pagos")
    ]
    return render_template("carga_datos_index.html", tools=upload_tools)


@app.route("/carga-datos/reportes")
def carga_datos_reportes():
    """
    Módulo "Reportes" (pedido explícito del usuario, 2026-09-17) --
    reemplaza al scaffold vacío que antes vivía acá mismo bajo el nombre
    "Controles" ("vamos a cambiar el modulo de controles por ese nombre y
    vamos a empezar con el chase la opcion de exportar en un PDF"). Por
    ahora solo Chase Bank tiene su reporte armado (PDF agrupado por
    Detalle, ver chase_rules.build_chase_pdf_report) -- el resto (EFT/
    Cupones, Reporte Diario, Lottery) queda "Próximamente", mismo criterio
    que usó /controles cuando arrancó sin ningún módulo.
    """
    today = date.today()
    return render_template(
        "carga_datos_reportes.html",
        tools=REPORTES_TOOLS,
        month_names=_MONTH_NAMES_ES,
        years_by_key={
            "reportes_chase": chase_db.get_available_years() or [today.year],
            "reportes_eft": eft_db.get_deposit_years() or [today.year],
            "reportes_diario": reportes_db.get_store_info_years() or [today.year],
            "reportes_lottery": lottery_db.get_available_years() or [today.year],
            "reportes_caja": get_caja_available_years() or [today.year],
            "reportes_proveedores": proveedores_db.get_available_years() or [today.year],
            "reportes_gettel": gettel_reportes.get_available_years() or [today.year],
        },
        current_year=today.year,
        current_month=today.month,
    )


def _report_without_data(message):
    """
    Un reporte mensual sin ningún dato ese mes no se arma vacío (pedido del
    usuario, 2026-10-05): se avisa, y con la descarga sin recargar de
    base.html el aviso sale sin recargar la página.
    """
    flash(message, "error")
    return redirect(url_for("carga_datos_reportes"))


@app.route("/carga-datos/reportes/chase/pdf")
def reportes_chase_pdf():
    """
    Descarga el PDF de Reportes -> Chase Bank -- resumen por Detalle con el
    total de cada categoría (pedido explícito del usuario, 2026-09-17: "que
    en el PDF no se muestren todos los movimientos, sino que esten los
    detalles y al lado el total de montos a ese detalle"), a diferencia del
    Excel de /carga-datos/chase/exportar (ese sí lista cada movimiento).
    Ver chase_rules.build_chase_pdf_report.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    transactions = chase_db.get_month_transactions(year, month)
    if not transactions:
        flash("No hay ningún movimiento guardado ese mes para el reporte.", "error")
        return redirect(url_for("carga_datos_reportes"))

    workspace_dir = tempfile.mkdtemp(prefix="chase_reporte_")
    dest_path = os.path.join(workspace_dir, f"Chase Reporte {month:02d}-{year}.pdf")
    build_chase_pdf_report(transactions, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/reportes/eft/pdf")
def reportes_eft_pdf():
    """
    Descarga el PDF de Reportes -> EFT y Cupones -- pedido explícito del
    usuario, 2026-09-17: "Lo mismo quiero que hagas con el Reporte de los
    EFT incluyendo todos los datos que se tengan del mes que se
    selecciono, y lo mismo estar incluido en ese PDF los cupones cargados
    hasta ese momento y cuanto acumulan". Ver eft_db.build_eft_pdf_report.
    Un mes sin ningún EFT ni cupón no se descarga (2026-10-05): antes salía
    igual, con la tabla de EFT vacía y solo el acumulado de Cupones.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    prefix = f"{year:04d}-{month:02d}-"
    if not eft_db.get_month_deposits(year, month) and not any(
        day.startswith(prefix) for day in eft_db.get_coupon_gross_by_date()
    ):
        return _report_without_data("No hay ningún EFT ni cupón cargado ese mes para el reporte.")

    workspace_dir = tempfile.mkdtemp(prefix="eft_reporte_")
    dest_path = os.path.join(workspace_dir, f"EFT Reporte {month:02d}-{year}.pdf")
    eft_db.build_eft_pdf_report(year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/reportes/reporte-diario/pdf")
def reportes_diario_pdf():
    """
    Descarga el PDF de Reportes -> Reporte Diario -- pedido explícito del
    usuario (2026-09-17, sesión siguiente: "quiero que pongas los pdf de
    reportes en reporte diario y lottery como los otros dos, mas
    resumido"). Ya no reusa el PDF día-por-día de Store Info (ese sigue
    intacto en /reporte/store-info/historial) -- ver
    reporte_diario.build_store_info_pdf_resumen, resumen de una sola
    tabla Detalle/Total, mismo criterio que Chase/EFT.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    store_info_rows = _build_store_info_rows(year, month)
    if not any(
        row.get(field) is not None
        for row in store_info_rows
        for field in ("from_time", "volume", "total_sales", "cash", "non_fuel_total", "total_revenue")
    ):
        return _report_without_data("No hay ningún Reporte Diario cargado ese mes para el reporte.")
    workspace_dir = tempfile.mkdtemp(prefix="reporte_diario_reporte_")
    dest_path = os.path.join(workspace_dir, f"Reporte Diario {month:02d}-{year}.pdf")
    build_store_info_pdf_resumen(store_info_rows, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/reportes/lottery/pdf")
def reportes_lottery_pdf():
    """
    Descarga el PDF de Reportes -> Lottery -- pedido explícito del usuario
    (2026-09-17, sesión siguiente: "quiero que pongas los pdf de reportes
    en reporte diario y lottery como los otros dos, mas resumido y con
    los totales bien hecho"). Ya no reusa el PDF día-por-día/bloques de 7
    días (ese sigue intacto en /carga-datos/lottery/historial) -- ver
    lottery_db.build_lottery_pdf_resumen, que suma el MES CALENDARIO
    completo (no por bloque ISO) y excluye a propósito las columnas que
    no son sumables (Cash Balance, Skoff Count/Sales).
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    days = lottery_db.get_days_between(date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1]))
    if not any(
        day.get(field) is not None
        for day in days.values()
        for field in ("online_count", "online_net_sales", "sales", "pagos", "skoff_count", "skoff_net_sales", "pays_amount")
    ):
        return _report_without_data("No hay ningún día de Lottery cargado ese mes para el reporte.")

    workspace_dir = tempfile.mkdtemp(prefix="lottery_reporte_")
    dest_path = os.path.join(workspace_dir, f"Lottery Reporte {month:02d}-{year}.pdf")
    lottery_db.build_lottery_pdf_resumen(year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/reportes/caja/pdf")
def reportes_caja_pdf():
    """
    Descarga el PDF de Reportes -> Caja -- pedido explícito del usuario
    (2026-09-17, sesión siguiente: "agrega tambien el de caja"), mismo
    criterio Detalle/Total que los otros 3 reportes de este módulo. Ver
    caja.build_caja_pdf_resumen -- reusa build_caja_month_report (mismo
    cálculo, ya validado, que usa /carga-datos/caja y su export
    día-por-día en /carga-datos/caja/exportar/pdf, sin tocar ninguno).
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    report = build_caja_month_report(year, month)
    # Un mes futuro viene con todo en 0, por eso se mira que haya algún monto y no solo que no sea None.
    if not any(
        row.get(field)
        for row in report["rows"]
        for field in ("total_sales", "cash", "deposit", "expenses_cash", "cuenta_final", "food_ice")
    ):
        return _report_without_data("No hay datos de Caja ese mes para el reporte (ni Reporte Diario, ni Chase, ni Lottery, ni gastos).")
    workspace_dir = tempfile.mkdtemp(prefix="caja_reporte_")
    dest_path = os.path.join(workspace_dir, f"Caja Reporte {month:02d}-{year}.pdf")
    build_caja_pdf_resumen(report, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/reportes/proveedores/pdf")
def reportes_proveedores_pdf():
    """
    Descarga el PDF de Reportes -> Proveedores -- pedido explícito del
    usuario (2026-09-22): "un resumen mensual en reportes, donde se va a
    mostrar la cantidad de facturas que llegaron de un proveedor y cual
    fue el total del mes" (mismo espíritu que RESUMEN COMPRAS del Excel
    real). Ver proveedores_db.build_proveedores_pdf_report.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    if not proveedores_db.get_month_invoices(year, month):
        return _report_without_data("No hay ninguna factura de proveedores cargada ese mes para el reporte.")

    workspace_dir = tempfile.mkdtemp(prefix="proveedores_reporte_")
    dest_path = os.path.join(workspace_dir, f"Proveedores Reporte {month:02d}-{year}.pdf")
    proveedores_db.build_proveedores_pdf_report(year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/reportes/gettel/excel")
def reportes_gettel_excel():
    """
    Descarga el Excel de Gettel con el formato/colores del Excel real --
    pedido explícito del usuario (2026-09-21): "ese excel que te pase es
    el que quiero que uses para crear el reporte de gettel, usando los
    formatos y colores que tiene" -- 4 hojas (Pendiente mes anterior / mes
    actual / Pago Cupones / Pendiente mes actual resultante), ver
    gettel_reportes.py. Se mudó de "Reportes" a este mismo módulo
    (2026-09-22, pedido explícito: "el excel de gettel deberia ir en su
    apartado de la barra lateral dentro de cuadro del mes") -- linkeado
    desde carga_datos_gettel_historial.html, ya no desde Reportes (que
    ahora usa reportes_gettel_pdf, un PDF resumido).
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    report = gettel_reportes.resolve_month(year, month)
    # Un mes que todavía no llegó arrastra igual el pendiente del anterior: tampoco se arma.
    if (year, month) > (today.year, today.month) or not report["has_any_data"]:
        return _report_without_data("No hay datos de Gettel ese mes para el reporte.")
    workspace_dir = tempfile.mkdtemp(prefix="gettel_reporte_")
    dest_path = os.path.join(workspace_dir, f"Gettel Reporte {month:02d}-{year}.xlsx")
    gettel_reportes.build_gettel_reportes_workbook(report, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/reportes/gettel/pdf")
def reportes_gettel_pdf():
    """
    Descarga el PDF de Reportes -> Gettel -- pedido explícito del usuario
    (2026-09-22): "el reporte de gettel deberia ser un PDF, no un excel".
    Mismas 4 secciones que el Excel (ver reportes_gettel_excel, que sigue
    existiendo del lado de Gettel/Cuadro del mes), resumidas sin colores --
    ver gettel_reportes.build_gettel_pdf_report.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    report = gettel_reportes.resolve_month(year, month)
    # Un mes que todavía no llegó arrastra igual el pendiente del anterior: tampoco se arma.
    if (year, month) > (today.year, today.month) or not report["has_any_data"]:
        return _report_without_data("No hay datos de Gettel ese mes para el reporte.")
    workspace_dir = tempfile.mkdtemp(prefix="gettel_reporte_")
    dest_path = os.path.join(workspace_dir, f"Gettel Reporte {month:02d}-{year}.pdf")
    gettel_reportes.build_gettel_pdf_report(report, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


# Módulo "Proyecciones" -- pedido explícito del usuario (2026-09-18): un
# TERCER apartado, hermano de Herramientas y Reportes (ver el switch de 3
# vías en base.html), con Ventas por Departamento y Store Info
# proyectados a fin de mes -- ver proyecciones.py para la fórmula real
# decodificada de la planilla de cierre (filas 37/38/39, columnas hasta
# la O) y todo lo que queda deliberadamente afuera de esta v1.
@app.route("/carga-datos/proyecciones")
def carga_datos_proyecciones():
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    projection = proyecciones.build_projection(year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "carga_datos_proyecciones.html",
        projection=projection,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        accent="#B45309",
        accent_soft="#FFFBEB",
    )


@app.route("/carga-datos/proyecciones/margenes", methods=["POST"])
def carga_datos_proyecciones_margenes():
    """
    Guarda los márgenes editados a mano desde la propia página de
    Proyecciones (uno por categoría de Ventas por Departamento, ver
    proyecciones.DEFAULT_MARGINS) -- pedido explícito del usuario
    (2026-09-18). Se tipean como porcentaje (ej. "15" para 15%) y se
    guardan como fracción (0.15), mismo criterio que ya usaba la
    planilla real. Un campo vacío o inválido deja ese margen sin tocar
    (nunca lo borra ni lo pone en 0 por accidente).
    """
    margins = {}
    for label, _members in DEPARTMENT_GROUPS:
        raw = (request.form.get(f"margin_{label}") or "").strip()
        if not raw:
            continue
        try:
            margins[label] = float(raw) / 100.0
        except ValueError:
            continue
    if margins:
        reportes_db.save_department_margins(margins)

    year = request.form.get("year", type=int) or date.today().year
    month = request.form.get("month", type=int) or date.today().month
    return redirect(url_for("carga_datos_proyecciones", year=year, month=month))


# Módulo "Combustible" + apartado "Fisico" -- pedido explícito del usuario
# (2026-09-18, misma sesión que Proyecciones): cargar a mano las facturas
# de compra de combustible ("Combustible", herramienta de Carga de Datos)
# para armar el cuadro de reconciliación mensual Inventario Teórico vs.
# lectura física real ("Fisico", apartado propio de la barra lateral) --
# ver fisico.py/fisico_db.py para la fórmula real decodificada de un
# ejemplo (`hoja_fisico.xlsx`). El usuario confirmó (2026-09-18)
# preferir carga manual mientras no hubiera un ejemplo real de factura;
# el 2026-09-21 subió 5 facturas reales y pidió que el PDF se lea solo
# -- ver fisico_invoice_parser.py y carga_datos_combustible_subir_pdf()
# más abajo. La carga manual queda como respaldo para cuando no hay un
# PDF limpio (o es de otro formato/proveedor).
@app.route("/carga-datos/combustible")
def carga_datos_combustible():
    """
    Pedido explícito del usuario (2026-09-21): se saca la carga a mano --
    esta página queda solo para subir el PDF de la factura. El detalle del
    mes (facturas ya cargadas, editar, eliminar) vive en /fisico -- mismo
    criterio que ya usa Gettel/Toyota ("Cargar" separado de "Cuadro del
    mes"). No hace falta año/mes acá: cada factura se archiva sola, en el
    mes de su propia Fecha de Factura leída del PDF (ver
    _run_carga_datos_combustible_job).
    """
    _active_job = jobs.get_active_job("combustible")
    return render_template(
        "carga_datos_combustible.html",
        resume_job_id=(_active_job["id"] if _active_job else None),
        **THEME_BY_KEY["carga_combustible"],
    )


@app.route("/carga-datos/combustible/<int:invoice_id>/editar", methods=["POST"])
def carga_datos_combustible_editar(invoice_id):
    """
    Corrige a mano un campo mal leído de una factura ya cargada por PDF --
    la única edición que queda disponible ahora que se sacó la carga
    manual (pedido explícito del usuario, 2026-09-21). El detalle por
    grado (`lines`) no se edita acá, solo el agregado que usa
    fisico.build_month_report.
    """
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    invoice_date = (request.form.get("invoice_date") or "").strip()
    due_date = (request.form.get("due_date") or "").strip() or None
    invoice_number = request.form.get("invoice_number")
    gallons_raw = (request.form.get("gallons") or "").strip()
    amount_raw = (request.form.get("amount") or "").strip()

    try:
        if not invoice_date:
            raise ValueError("Falta la Fecha de Factura.")
        gallons = float(gallons_raw)
        amount = float(amount_raw)
        fisico_db.update_invoice(invoice_id, invoice_date, due_date, invoice_number, gallons, amount)
        flash("Factura corregida.", "success")
        redirect_year, redirect_month = (int(part) for part in invoice_date.split("-")[:2])
    except ValueError:
        flash("Revisá la Fecha, los Galones y el Monto -- tienen que ser válidos.", "error")
        redirect_year, redirect_month = year, month

    return redirect(url_for("fisico_view", year=redirect_year, month=redirect_month))


@app.route("/carga-datos/combustible/subir-pdf", methods=["POST"])
def carga_datos_combustible_subir_pdf():
    """
    Carga por PDF -- pedido explícito de Alfonso (2026-09-21), con 5
    facturas reales del proveedor de combustible como ejemplo (ver
    fisico_invoice_parser.py para el detalle de qué se lee y por qué).
    Mismo patrón de siempre para lotes de PDF (jobs.py + threading, ver
    _run_carga_datos_proveedores_job): cada archivo se procesa aislado --
    un PDF roto o de otro formato no tira abajo el resto del lote, y no
    se guarda nada de esa factura puntual (fisico_invoice_parser nunca
    adivina con baja confianza). Duplicado = mismo N° de factura ya
    guardado, en cualquier mes -- mismo criterio que
    proveedores_db.save_invoice.
    """
    uploads = [f for f in request.files.getlist("pdf_files") if f and f.filename]
    if not uploads:
        return _error_response("Seleccioná uno o más PDF de factura de combustible.")

    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths), kind="combustible")
    threading.Thread(target=_run_carga_datos_combustible_job, args=(job_id, paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


def _run_carga_datos_combustible_job(job_id, paths):
    """Corre en su propio hilo -- mismo patrón que _run_carga_datos_proveedores_job, ver ese docstring."""
    try:
        saved = []
        duplicates = []
        failed = []

        for index, path in enumerate(paths, start=1):
            filename = os.path.basename(path)
            try:
                result = fisico_invoice_parser.extract_fuel_invoice(path)
            except fisico_invoice_parser.PDF_READ_EXCEPTIONS as exc:
                failed.append({"filename": filename, "error": str(exc)})
                jobs.update_job(job_id, done=index, total=len(paths))
                continue

            existing = fisico_db.find_invoice_by_number(result["invoice_number"])
            if existing:
                duplicates.append({"filename": filename, "invoice_number": result["invoice_number"]})
                jobs.update_job(job_id, done=index, total=len(paths))
                continue

            try:
                documents_db.store_document(
                    "combustible", path, filename,
                    result["invoice_date"].year, result["invoice_date"].month,
                    label=result["invoice_number"],
                )
            except Exception:
                pass

            fisico_db.add_invoice(
                result["invoice_date"],
                result["due_date"],
                result["invoice_number"],
                result["total_gallons"],
                result["total_amount_due"],
                source="pdf",
                bol_number=result["bol_number"],
                lines=result["lines"],
                source_filename=filename,
            )
            saved.append({
                "filename": filename,
                "invoice_number": result["invoice_number"],
                "date": result["invoice_date"],
            })
            jobs.update_job(job_id, done=index, total=len(paths))

        parts = []
        if saved:
            parts.append(f"{len(saved)} factura(s) guardada(s).")
        if duplicates:
            nums = ", ".join(d["invoice_number"] for d in duplicates)
            parts.append(f"{len(duplicates)} factura(s) ya estaban cargadas y se omitieron ({nums}).")
        if failed:
            for item in failed:
                parts.append(f"{item['filename']}: {item['error']}")
        if not parts:
            parts.append("No se guardó ninguna factura de este lote.")
        level = (
            "success" if (saved and not duplicates and not failed)
            else ("error" if not saved else "warning")
        )

        if saved:
            # Pedido explícito del usuario (2026-09-21): el resultado ya no
            # se mira en esta misma página (que ahora es solo el upload) --
            # va directo al Cuadro del mes de Físico, en el mes de la
            # PRIMER factura guardada del lote (auto-detectado de su propia
            # Fecha de Factura, no de ningún selector).
            d = saved[0]["date"]
            redirect_url = f"/fisico?year={d.year}&month={d.month}"
        else:
            # Nada se guardó (todo duplicado/fallido) -- se queda en la
            # página de carga para que se vea el aviso de qué pasó.
            redirect_url = "/carga-datos/combustible"

        jobs.update_job(
            job_id, status="done", done=len(paths), total=len(paths),
            notice=" ".join(parts), notice_level=level, redirect_url=redirect_url,
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


@app.route("/carga-datos/combustible/<int:invoice_id>/eliminar", methods=["POST"])
def carga_datos_combustible_eliminar(invoice_id):
    """El botón de eliminar vive en /fisico ahora (la tabla de facturas se movió ahí), no en esta página de carga."""
    fisico_db.delete_invoice(invoice_id)
    flash("Factura eliminada.", "success")
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    return redirect(url_for("fisico_view", year=year, month=month))


@app.route("/fisico")
def fisico_view():
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    report = fisico.build_month_report(year, month)
    docs_by_filename = {}
    for doc in documents_db.list_all_documents("combustible"):
        docs_by_filename.setdefault(doc["filename"], doc)
    for inv in report["invoices"]:
        inv["document"] = docs_by_filename.get(inv.get("source_filename"))
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "fisico.html",
        report=report,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["fisico"],
    )


@app.route("/fisico/exportar")
def fisico_exportar():
    """Excel NUEVO (nunca toca ningún archivo real) con el Inventario Teórico + facturas del mes -- ver fisico.build_fisico_export_workbook."""
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    report = fisico.build_month_report(year, month)
    workspace_dir = tempfile.mkdtemp(prefix="fisico_export_")
    dest_path = os.path.join(workspace_dir, f"Fisico {month:02d}-{year}.xlsx")
    fisico.build_fisico_export_workbook(report, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/fisico/exportar/pdf")
def fisico_exportar_pdf():
    """Versión PDF del export de arriba -- ver fisico.build_fisico_pdf_report."""
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    report = fisico.build_month_report(year, month)
    workspace_dir = tempfile.mkdtemp(prefix="fisico_export_pdf_")
    dest_path = os.path.join(workspace_dir, f"Fisico {month:02d}-{year}.pdf")
    fisico.build_fisico_pdf_report(report, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/fisico/inicial", methods=["POST"])
def fisico_ajustar_inicial():
    """
    Override manual del Inventario Inicial Teórico -- hace falta para el
    primer mes que se usa este módulo (no hay mes anterior del que
    encadenar) y queda disponible siempre por si hace falta corregirlo,
    mismo criterio que carga_datos_caja_saldo. Vacío borra el override
    (vuelve a encadenarse del mes anterior).
    """
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    gallons_raw = (request.form.get("initial_gallons") or "").strip()
    amount_raw = (request.form.get("initial_amount") or "").strip()

    # Galones y monto van juntos (los dos para guardar, los dos vacíos para
    # borrar): con uno solo el Inicial quedaba a medias sin aviso.
    if bool(gallons_raw) != bool(amount_raw):
        flash("Cargá los galones y el monto juntos (o dejá los dos vacíos para borrar el ajuste).", "error")
        return redirect(url_for("fisico_view", year=year, month=month))

    try:
        gallons_value = float(gallons_raw) if gallons_raw else None
        amount_value = float(amount_raw) if amount_raw else None
        fisico_db.set_month_initial_override(year, month, gallons_value, amount_value)
        flash("Inventario Inicial guardado.", "success")
    except ValueError:
        flash("No se pudo guardar: revisá que los galones/monto sean números válidos.", "error")

    return redirect(url_for("fisico_view", year=year, month=month))


@app.route("/fisico/lectura-real", methods=["POST"])
def fisico_lectura_real():
    """Lectura física real de los tanques -- siempre un dato externo, nunca se calcula. Vacío la borra (mes vuelve a "pendiente")."""
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    gallons_raw = (request.form.get("real_gallons") or "").strip()
    reading_date = (request.form.get("real_reading_date") or "").strip() or None

    try:
        gallons_value = float(gallons_raw) if gallons_raw else None
        fisico_db.set_month_real_ending(year, month, gallons_value, reading_date)
        flash("Lectura física guardada.", "success")
    except ValueError:
        flash("No se pudo guardar: los galones tienen que ser un número válido.", "error")

    return redirect(url_for("fisico_view", year=year, month=month))


@app.route("/carga-datos/reporte-mensual")
def carga_datos_reporte_mensual():
    """
    Reporte Mensual (pedido del usuario, 2026-10-06): se sube el resumen de
    ventas del mes del POS y se cruza con la suma de los reportes diarios del
    mes (pedido del usuario, chat 21: no llamarlo asiento). Lectura y cruce
    en reporte_mensual.py.
    """
    year, month = _cierre_month()
    store_info_rows = _build_store_info_rows(year, month)
    entries = control_cierre.build_month_entries(store_info_rows, year, month)
    report, cross = _monthly_cross(year, month, entries)
    # Lo que no coincide fuera de los importes del cruce (hay que corregirlo igual).
    other_store_info, other_departments = [], []
    if cross:
        days_loaded = sum(1 for r in store_info_rows if r.get("store_info_source"))
        comparison = reporte_mensual.store_info_comparison(_store_info_totals(store_info_rows), report, days_loaded)
        other_store_info = [c for c in comparison.values() if not c["ok"]]
        departments = reporte_mensual.department_comparison(reportes_db.get_month_department_totals(year, month), report)
        other_departments = [d for d in departments or [] if not d["ok"]]
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    return render_template(
        "carga_datos_reporte_mensual.html",
        report=report,
        editable_fields=reporte_mensual.EDITABLE_STORE_INFO,
        entries=entries,
        cross=cross,
        other_store_info=other_store_info,
        other_departments=other_departments,
        notice=control_cierre.missing_notice(entries) if entries else None,
        pending_items=reporte_mensual.pending_items(report) if report else [],
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["reporte"],
    )


@app.route("/carga-datos/reporte-mensual/subir", methods=["POST"])
def carga_datos_reporte_mensual_subir():
    uploads = request.files.getlist("reporte_files")
    if not uploads or not any(u.filename for u in uploads):
        return _error_response("Seleccioná el PDF del reporte mensual.")
    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths), kind="reporte_mensual")
    threading.Thread(target=_run_reporte_mensual_job, args=(job_id, paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


def _run_reporte_mensual_job(job_id, paths):
    """Aislado por archivo: cada PDF es el reporte de un mes y reemplaza el de ese mes."""
    try:
        loaded, problems, sheets, last = [], [], [], None
        for index, path in enumerate(paths, start=1):
            try:
                try:
                    report = reporte_mensual.extract_monthly_report(path)
                except ValueError as exc:
                    # Una hoja suelta (sin "PERIOD FROM") es la que se pidió de
                    # nuevo: si hay un solo mes con hoja pendiente, va a ese mes.
                    pending = _pending_monthly_reports()
                    if "PERIOD FROM" not in str(exc) or not pending:
                        raise
                    if len(pending) > 1:
                        raise ValueError("Ese archivo no es un reporte mensual completo. Si es la hoja que faltaba, "
                                         "subila en \"Hoja pendiente\", eligiendo el mes.") from None
                    notice, left = _apply_monthly_sheet(pending[0]["year"], pending[0]["month"], path)
                    (problems if left else sheets).append(notice)
                    last = (pending[0]["year"], pending[0]["month"])
                    jobs.update_job(job_id, done=index, total=len(paths))
                    continue
                reporte_mensual_db.save_report(report, os.path.basename(path))
                label = f"{_MONTH_NAMES_ES[report['month'] - 1]} {report['year']}"
                loaded.append(label)
                last = (report["year"], report["month"])
                problems.extend(f"{label}: {warning}" for warning in report["warnings"])
            except ValueError as exc:
                problems.append(str(exc))
            except Exception as exc:
                print(f"[carga-datos/reporte-mensual] {path}: {exc}")
                problems.append("Un archivo no se pudo leer.")
            jobs.update_job(job_id, done=index, total=len(paths))

        parts = []
        if loaded:
            parts.append("Reporte mensual cargado: " + ", ".join(loaded) + ".")
        parts.extend(sheets + problems)
        done = loaded or sheets or last
        level = "success" if done and not problems else ("warning" if done else "error")
        jobs.update_job(
            job_id, status="done", done=len(paths), total=len(paths),
            notice=" ".join(parts), notice_level=level,
            redirect_url=(f"/carga-datos/reporte-mensual?year={last[0]}&month={last[1]}" if last
                          else "/carga-datos/reporte-mensual"),
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


def _pending_monthly_reports():
    """Meses cuyo reporte mensual tiene algo sin leer: [{year, month, label, items}]."""
    pending = []
    for year, month in reporte_mensual_db.list_months():
        items = reporte_mensual.pending_items(reporte_mensual_db.get_report(year, month))
        if items:
            pending.append({"year": year, "month": month, "items": items,
                            "label": f"{_MONTH_NAMES_ES[month - 1]} {year}"})
    return pending


def _apply_monthly_sheet(year, month, path):
    """Completa el reporte del mes con la hoja pedida de nuevo; devuelve el aviso. ValueError si no sirve."""
    report = reporte_mensual_db.get_report(year, month)
    if report is None:
        raise ValueError("Ese mes no tiene reporte mensual cargado.")
    sheet = reporte_mensual.extract_replacement_sheet(path)
    store_info, departments, printed_amount, printed_count, warnings, filled, conflicts = (
        reporte_mensual.apply_replacement_sheet(report, sheet)
    )
    reporte_mensual_db.update_report(year, month, store_info, departments, printed_amount, printed_count,
                                     warnings, from_sheet=True)
    label = f"{_MONTH_NAMES_ES[month - 1]} {year}"
    parts = [f"{label}: se completó con la hoja nueva " + ", ".join(filled) + "."]
    if conflicts:
        parts.append("No se pisó lo ya guardado que no coincide con la hoja: " + ", ".join(conflicts) + ".")
    left = reporte_mensual.pending_items(reporte_mensual_db.get_report(year, month))
    parts.append("Sigue faltando: " + ", ".join(left) + "." if left else "Ya no le falta nada.")
    return " ".join(parts), bool(conflicts or left)


@app.route("/carga-datos/reporte-mensual/hoja", methods=["POST"])
def carga_datos_reporte_mensual_hoja():
    """
    Hoja pedida de nuevo (pedido del usuario, 2026-10-07): la hoja que salió
    borrosa, mandada otra vez por el manager, completa lo que le faltaba al
    reporte del mes (solo lo vacío; ver reporte_mensual.apply_replacement_sheet).
    """
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    uploads = request.files.getlist("hoja_files")
    if not (year and month):
        return _error_response("Elegí el mes al que le falta la hoja.")
    if not uploads or not any(u.filename for u in uploads):
        return _error_response("Seleccioná la hoja (PDF o foto).")
    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths), kind="reporte_mensual")
    threading.Thread(target=_run_reporte_mensual_hoja_job, args=(job_id, year, month, paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


def _run_reporte_mensual_hoja_job(job_id, year, month, paths):
    """Cada archivo es una hoja (o varias) del mismo mes; una que no sirve no frena las demás."""
    try:
        notices, problems, still_pending = [], [], False
        for index, path in enumerate(paths, start=1):
            try:
                notice, left = _apply_monthly_sheet(year, month, path)
                notices.append(notice)
                still_pending = left
            except ValueError as exc:
                problems.append(str(exc))
            except Exception as exc:
                print(f"[carga-datos/reporte-mensual/hoja] {path}: {exc}")
                problems.append("Un archivo no se pudo leer.")
            jobs.update_job(job_id, done=index, total=len(paths))
        level = "error" if not notices else ("warning" if problems or still_pending else "success")
        jobs.update_job(
            job_id, status="done", done=len(paths), total=len(paths),
            notice=" ".join(notices + problems), notice_level=level,
            redirect_url=f"/carga-datos/reporte-mensual?year={year}&month={month}",
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


@app.route("/carga-datos/reporte-mensual/editar", methods=["POST"])
def carga_datos_reporte_mensual_editar():
    """Corrección a mano del reporte mensual guardado (pedido del usuario, 2026-10-07)."""
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    report = reporte_mensual_db.get_report(year, month) if year and month else None
    if report is None:
        flash("Ese mes no tiene reporte mensual cargado.", "error")
    else:
        try:
            store_info, departments, printed_amount, printed_count, warnings = reporte_mensual.edited_report(
                report, request.form
            )
        except ValueError as exc:
            flash(f"No se guardó: {exc}", "error")
        else:
            reporte_mensual_db.update_report(
                year, month, store_info, departments, printed_amount, printed_count, warnings
            )
    return redirect(url_for("carga_datos_reporte_mensual", year=year, month=month))


@app.route("/carga-datos/reporte-mensual/eliminar", methods=["POST"])
def carga_datos_reporte_mensual_eliminar():
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    if year and month:
        reporte_mensual_db.delete_report(year, month)
    return redirect(url_for("carga_datos_reporte_mensual", year=year, month=month))


@app.route("/carga-datos/reporte-diario")
def carga_datos_reporte_diario():
    """
    Carga directa de Reporte Diario del lado Carga de Datos -- solo PDF, sin
    ningún campo de Excel (pedido explícito del usuario 2026-09-11, parte del
    plan de dejar este lado autosuficiente para subir datos sin depender de
    Herramientas). Ver carga_datos_reporte_diario_subir más abajo.

    "Reportes Diario/Mensual" (pedido del usuario, 2026-10-06): abajo, en la
    misma página, se carga el reporte mensual (carga_datos_reporte_mensual_
    subir); su página propia queda solo para verlo y eliminarlo.
    """
    _active_job = jobs.get_active_job("reporte_diario")
    _monthly_job = jobs.get_active_job("reporte_mensual")
    return render_template(
        "carga_datos_reporte_diario.html",
        resume_job_id=(_active_job["id"] if _active_job else None),
        resume_monthly_job_id=(_monthly_job["id"] if _monthly_job else None),
        pending_reports=_pending_monthly_reports(),
        **THEME_BY_KEY["reporte"],
    )


def _run_carga_datos_reporte_diario_job(job_id, pdf_paths):
    """
    Corre en su propio hilo (ver carga_datos_reporte_diario_subir) -- mismo
    patrón que _run_reporte_ventas_job (jobs.py): el cuerpo entero va
    envuelto en un único try/except para que cualquier excepción, en
    cualquier paso, termine el job con status="error" en vez de dejarlo
    pegado en "running" para siempre. Pedido explícito del usuario
    (2026-09-16, segunda tanda): "una barra de progreso real... no una
    animación estática y repetitiva" -- antes esta carga corría de forma
    síncrona dentro del propio POST, mostrando el mismo rayado indeterminado
    que cualquier otro form mientras el servidor trabajaba; ahora reporta
    el avance real (PDF ya procesados/total) por polling, igual que ya hacía
    el lado Herramientas de este mismo módulo.

    Los PDF siguientes se leen por adelantado en otros hilos mientras se
    guarda el actual (reporte_diario.DepartmentPagePrefetch, 2026-10-06): la
    misma lectura, más rápida en un lote.
    """
    prefetch = DepartmentPagePrefetch(pdf_paths)
    try:
        days_complete = set()
        days_partial = set()
        days_subtotal_mismatch = set()
        days_doubtful = {}
        days_unverified = set()
        days_total_sales_mismatch = {}
        days_store_info_missing = {}
        days_lottery_missing = {}
        lottery_unreadable = 0
        files_unreadable = 0
        date_mismatches = 0
        first_date = None

        for index, pdf_path in enumerate(pdf_paths, start=1):
            prefetch.wait(index - 1)
            filename = os.path.basename(pdf_path)
            filename_day_month = _reporte_filename_day_month(filename)
            day_date = None
            got_departments = False
            got_store_info = False
            got_store_info_partial = False
            file_had_mismatch = False

            def _check_filename_date(candidate_date):
                if filename_day_month and (candidate_date.day, candidate_date.month) != filename_day_month:
                    raise ValueError(
                        f"la fecha leída ({candidate_date.isoformat()}) no coincide con la fecha "
                        f"del nombre de archivo ({filename_day_month[0]:02d}-{filename_day_month[1]:02d})"
                    )

            try:
                result = extract_department_sales_for_day(pdf_path)
                candidate_date = result["date"]
                _check_filename_date(candidate_date)
                pdf_relpath = reportes_db.store_pdf_copy(
                    candidate_date, pdf_path, _reporte_pdf_canonical_filename(candidate_date)
                )
                reportes_db.replace_department_sales(
                    candidate_date, result["records"], pdf_filename=pdf_relpath,
                )
                day_date = candidate_date
                got_departments = True
                if result.get("subtotal_mismatch"):
                    days_subtotal_mismatch.add(candidate_date)
                if result.get("doubtful_departments"):
                    days_doubtful[candidate_date] = result["doubtful_departments"]
                if result.get("total_unverified"):
                    days_unverified.add(candidate_date)
                # El total impreso al pie de la tabla queda guardado para el
                # chequeo "Total impreso" del día (antes había que tipearlo).
                printed = result.get("printed_totals")
                if printed:
                    reportes_db.upsert_printed_totals(candidate_date, printed["amount"], printed["count"])
            except Exception as exc:
                if "no coincide con la fecha del nombre" in str(exc):
                    file_had_mismatch = True
                print(f"[carga-datos/reporte-diario] departamentos de {pdf_path}: {exc}")

            if not file_had_mismatch:
                try:
                    result = extract_store_info_for_day(pdf_path)
                    candidate_date = result["date"]
                    _check_filename_date(candidate_date)
                    pdf_relpath = reportes_db.store_pdf_copy(
                        candidate_date, pdf_path, _reporte_pdf_canonical_filename(candidate_date)
                    )
                    reportes_db.upsert_store_info(
                        candidate_date, result["fields"], source="ocr", pdf_filename=pdf_relpath,
                        keep_existing_for_none=True,
                    )
                    if result["fields"].get("total_sales_mismatch"):
                        days_total_sales_mismatch[candidate_date] = result["fields"]["total_sales_mismatch"]
                    # Lo que el OCR no pudo leer quedó vacío (no se adivina):
                    # se guarda el resto y se avisa qué falta completar.
                    if result["fields"].get("missing_fields"):
                        days_store_info_missing[candidate_date] = result["fields"]["missing_fields"]
                    day_date = candidate_date
                    got_store_info = not result["fields"].get("missing_fields")
                    got_store_info_partial = bool(result["fields"].get("missing_fields"))
                except Exception as exc:
                    if "no coincide con la fecha del nombre" in str(exc):
                        file_had_mismatch = True
                    print(f"[carga-datos/reporte-diario] store info de {pdf_path}: {exc}")

            if not file_had_mismatch:
                try:
                    fields = extract_lottery_department_fields_from_pdf(pdf_path)
                    _check_filename_date(fields["report_date"])
                    lottery_relpath = lottery_db.store_pdf_copy(fields["report_date"], pdf_path, filename)
                    # Lo dudoso viene en None y deja lo que ya había (ver
                    # upsert_department_fields); se avisa qué falta.
                    lottery_db.upsert_department_fields(
                        fields["report_date"], fields["online_count"], fields["online_net_sales"],
                        fields["skoff_count"], fields["skoff_net_sales"], source="ocr", pdf_filename=lottery_relpath,
                    )
                    if fields["missing"]:
                        days_lottery_missing[fields["report_date"]] = fields["missing"]
                except Exception as exc:
                    if "no coincide con la fecha del nombre" not in str(exc):
                        lottery_unreadable += 1
                    print(f"[carga-datos/reporte-diario] ONLINE/SKOFF de {pdf_path}: {exc}")

            if file_had_mismatch:
                date_mismatches += 1
                files_unreadable += 1
            elif day_date is None:
                files_unreadable += 1
            else:
                if first_date is None or day_date < first_date:
                    first_date = day_date
                if got_departments and got_store_info:
                    days_complete.add(day_date)
                elif got_departments or got_store_info or got_store_info_partial:
                    days_partial.add(day_date)
                else:
                    files_unreadable += 1

            jobs.update_job(job_id, done=index, total=len(pdf_paths))

        parts = []
        if days_complete:
            parts.append(f"{len(days_complete)} día(s) guardado(s) completos.")
        if days_partial:
            dates_txt = ", ".join(sorted(d.isoformat() for d in days_partial))
            parts.append(f"{len(days_partial)} día(s) quedaron incompletos ({dates_txt}) — completalos a mano.")
        if date_mismatches:
            parts.append(
                f"{date_mismatches} archivo(s) rechazados: la fecha real del PDF no coincide con la del "
                "nombre de archivo — revisá que no sea de otro mes."
            )
        if files_unreadable - date_mismatches:
            parts.append(f"{files_unreadable - date_mismatches} archivo(s) no se pudieron leer en absoluto.")
        if days_subtotal_mismatch:
            dates_txt = ", ".join(sorted(d.isoformat() for d in days_subtotal_mismatch))
            parts.append(
                f"{len(days_subtotal_mismatch)} día(s) con la suma de departamentos distinta del total "
                f"impreso en el PDF ({dates_txt}) — es señal de que el OCR se salteó alguna fila (ej. un "
                "departamento esporádico como GIFT CARD), revisá Ventas por Departamento y completalo a mano si falta algo."
            )
        if days_store_info_missing:
            dates_txt = "; ".join(
                f"{d.strftime('%d/%m')}: {', '.join(fields)}" for d, fields in sorted(days_store_info_missing.items())
            )
            parts.append(
                f"Store Info: lo demás se guardó, pero no se pudo leer con seguridad ({dates_txt}) — "
                "quedó vacío, completalo a mano."
            )
        if days_doubtful:
            dates_txt = "; ".join(
                f"{d.strftime('%d/%m')}: {', '.join(items)}" for d, items in sorted(days_doubtful.items())
            )
            parts.append(
                f"Ventas por Departamento: valores que el OCR no pudo leer con seguridad ({dates_txt}) — "
                "quedaron vacíos, completalos a mano en el día."
            )
        if days_lottery_missing:
            dates_txt = "; ".join(
                f"{d.strftime('%d/%m')}: {', '.join(items)}" for d, items in sorted(days_lottery_missing.items())
            )
            parts.append(
                f"Lottery: ventas que el OCR no pudo leer con seguridad ({dates_txt}) — no se cargaron, "
                "completalas a mano en el día de Lottery."
            )
        if lottery_unreadable:
            parts.append(
                f"Lottery: {lottery_unreadable} archivo(s) sin las ventas ONLINE/SKOFF "
                "(no se pudo leer la fecha o la página) — completalas a mano en Lottery."
            )
        if days_unverified:
            dates_txt = ", ".join(sorted(d.isoformat() for d in days_unverified))
            parts.append(
                f"{len(days_unverified)} día(s) sin el total impreso legible al pie de Ventas por Departamento "
                f"({dates_txt}) — no se pudo verificar la suma, revisalos contra el PDF."
            )
        if days_total_sales_mismatch:
            dates_txt = ", ".join(
                f"{d.isoformat()} (${v:+,.2f})" for d, v in sorted(days_total_sales_mismatch.items())
            )
            parts.append(
                f"{len(days_total_sales_mismatch)} día(s) donde Store Info no cierra contra el Total Sales "
                f"impreso ({dates_txt}) — algún monto se leyó mal, revisalo a mano en el día."
            )

        if not parts:
            notice, level = "No se pudo guardar nada de este lote.", "error"
        else:
            notice = " ".join(parts)
            level = "warning" if (days_partial or files_unreadable or days_subtotal_mismatch or days_doubtful or days_unverified or days_total_sales_mismatch or days_lottery_missing or lottery_unreadable) else "success"

        redirect_url = (
            f"/reporte/historial?year={first_date.year}&month={first_date.month}"
            if first_date else "/carga-datos/reporte-diario"
        )
        jobs.update_job(
            job_id, status="done", done=len(pdf_paths), total=len(pdf_paths),
            notice=notice, notice_level=level, redirect_url=redirect_url,
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")
    finally:
        prefetch.close()


@app.route("/carga-datos/reporte-diario/subir", methods=["POST"])
def carga_datos_reporte_diario_subir():
    """
    Guarda directo en reportes_db/lottery_db -- nunca genera ni toca ningún
    Excel (a diferencia de /reporte/ventas y /reporte/store-info, que
    siguen viviendo del lado Excels). Usa las mismas funciones de extracción
    "puras" que ya alimentan el guardado-espejo de ese otro lado
    (extract_department_sales_for_day / extract_store_info_for_day /
    extract_lottery_department_fields_from_pdf) -- acá son el ÚNICO camino
    de escritura, no un paso adicional después de un Excel. Cada PDF se
    procesa aislado (puede acertar Departamentos, Store Info, ninguno, o los
    dos) para que un archivo con problema no tumbe el resto del lote.

    Corre en segundo plano (jobs.py) con progreso real PDF a PDF -- pedido
    explícito del usuario (2026-09-16, segunda tanda), ver el docstring de
    _run_carga_datos_reporte_diario_job.
    """
    pdf_uploads = request.files.getlist("pdf_files")
    if not pdf_uploads or not any(u.filename for u in pdf_uploads):
        return _error_response("Seleccioná uno o más PDF de cierre diario.")

    pdf_paths = _save_uploads_to_workspace(pdf_uploads)
    job_id = jobs.create_job(len(pdf_paths), kind="reporte_diario")
    threading.Thread(target=_run_carga_datos_reporte_diario_job, args=(job_id, pdf_paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(pdf_paths)})


@app.route("/carga-datos/lottery")
def carga_datos_lottery():
    """Carga directa de Lottery (Daily Sales Report) del lado Carga de Datos -- solo PDF."""
    _active_job = jobs.get_active_job("lottery")
    _monthly_job = jobs.get_active_job("lottery_mensual")
    return render_template(
        "carga_datos_lottery.html",
        resume_job_id=(_active_job["id"] if _active_job else None),
        resume_monthly_job_id=(_monthly_job["id"] if _monthly_job else None),
        **THEME_BY_KEY["lottery"],
    )


@app.route("/carga-datos/lottery/subir", methods=["POST"])
def carga_datos_lottery_subir():
    """
    Análogo a carga_datos_reporte_diario_subir, para el Daily Sales Report
    de Florida Lottery -- único camino de escritura, sin ningún Excel de
    por medio. ONLINE/SKOFF ya se completan solos al subir Reporte Diario
    (ver carga_datos_reporte_diario_subir) -- acá solo hace falta el resto
    de las columnas de este documento.
    """
    pdf_uploads = request.files.getlist("pdf_files")
    if not pdf_uploads or not any(u.filename for u in pdf_uploads):
        return _error_response("Seleccioná uno o más PDF de Daily Sales Report.")

    pdf_paths = _save_uploads_to_workspace(pdf_uploads)
    job_id = jobs.create_job(len(pdf_paths), kind="lottery")
    threading.Thread(target=_run_carga_datos_lottery_job, args=(job_id, pdf_paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(pdf_paths)})


def _run_carga_datos_lottery_job(job_id, pdf_paths):
    """Corre en su propio hilo -- mismo patrón que _run_carga_datos_reporte_diario_job, ver ese docstring."""
    try:
        saved_dates = []
        failed = 0
        date_mismatches = 0
        incomplete = 0
        for index, pdf_path in enumerate(pdf_paths, start=1):
            filename = os.path.basename(pdf_path)
            try:
                fields = extract_lottery_receipt_fields_from_sales_report(pdf_path)
                # Un campo que no se pudo leer queda en None con un warning:
                # se guarda lo demás (el None no pisa lo ya guardado, ver
                # upsert_sales_report_fields) y se avisa para completarlo.
                if fields.get("warning"):
                    incomplete += 1
                if _filename_date_mismatch(filename, fields["report_date"]):
                    date_mismatches += 1
                    raise ValueError("la fecha leída no coincide con la del nombre de archivo")
                pdf_relpath = lottery_db.store_pdf_copy(fields["report_date"], pdf_path, filename)
                lottery_db.upsert_sales_report_fields(fields["report_date"], fields, source="ocr", pdf_filename=pdf_relpath)
                saved_dates.append(fields["report_date"])
            except Exception as exc:
                print(f"[carga-datos/lottery] {pdf_path}: {exc}")
                failed += 1
            jobs.update_job(job_id, done=index, total=len(pdf_paths))

        parts = []
        if saved_dates:
            parts.append(f"{len(saved_dates)} día(s) guardado(s).")
        if date_mismatches:
            parts.append(
                f"{date_mismatches} archivo(s) rechazados: la fecha real no coincide con la del "
                "nombre de archivo — revisá que no sea de otro mes."
            )
        if failed - date_mismatches:
            parts.append(f"{failed - date_mismatches} archivo(s) no se pudieron leer.")
        if incomplete:
            parts.append(f"{incomplete} día(s) con algún campo que no se pudo leer: completalo a mano.")
        if not parts:
            notice, level = "No se pudo guardar nada de este lote.", "error"
        else:
            notice, level = " ".join(parts), ("warning" if (failed or incomplete) else "success")

        if saved_dates:
            first = min(saved_dates)
            redirect_url = f"/carga-datos/lottery/historial?year={first.year}&month={first.month}"
        else:
            redirect_url = "/carga-datos/lottery"

        jobs.update_job(
            job_id, status="done", done=len(pdf_paths), total=len(pdf_paths),
            notice=notice, notice_level=level, redirect_url=redirect_url,
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


def _run_monthly_upload_job(job_id, paths, read_and_save, done_prefix, redirect_for, fallback_url):
    """
    Carga de reportes mensuales (Lottery, J.H.), aislada por archivo: cada PDF
    reemplaza el reporte de su mes. `read_and_save(path)` lo lee, lo guarda y
    devuelve (etiqueta, (año, mes), avisos ya completos); un ValueError es un
    aviso corto.
    """
    try:
        loaded, problems, last = [], [], None
        for index, path in enumerate(paths, start=1):
            try:
                label, last, warnings = read_and_save(path)
                loaded.append(label)
                problems.extend(warnings)
            except ValueError as exc:
                problems.append(str(exc))
            except Exception as exc:
                print(f"[{done_prefix}] {path}: {exc}")
                problems.append("Un archivo no se pudo leer.")
            jobs.update_job(job_id, done=index, total=len(paths))
        parts = [f"{done_prefix}: " + ", ".join(loaded) + "."] if loaded else []
        parts.extend(problems)
        level = "success" if loaded and not problems else ("warning" if loaded else "error")
        jobs.update_job(
            job_id, status="done", done=len(paths), total=len(paths),
            notice=" ".join(parts), notice_level=level,
            redirect_url=redirect_for(*last) if last else fallback_url,
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


def _save_lottery_monthly_report(path):
    report = lottery_mensual.extract_monthly_report(path)
    lottery_db.save_monthly_report(report, os.path.basename(path))
    label = f"{_MONTH_NAMES_ES[report['month'] - 1]} {report['year']}"
    return label, (report["year"], report["month"]), [f"{label}: {warning}" for warning in report["warnings"]]


@app.route("/carga-datos/lottery/mensual/subir", methods=["POST"])
def carga_datos_lottery_mensual_subir():
    """Monthly Sales Report de Lottery (pedido del usuario, 2026-10-06): se cruza en Controles → Lottery."""
    uploads = request.files.getlist("monthly_files")
    if not uploads or not any(u.filename for u in uploads):
        return _error_response("Seleccioná el PDF o el Excel del Monthly Sales Report.")
    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths), kind="lottery_mensual")
    threading.Thread(
        target=_run_monthly_upload_job,
        args=(job_id, paths, _save_lottery_monthly_report, "Reporte mensual de Lottery cargado",
              lambda y, m: f"/carga-datos/lottery/mensual?year={y}&month={m}", "/carga-datos/lottery/mensual"),
        daemon=True,
    ).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


@app.route("/carga-datos/lottery/mensual")
def carga_datos_lottery_mensual():
    """El Monthly Sales Report guardado de un mes, tal cual lo imprime el portal, para verlo y eliminarlo."""
    year, month = _cierre_month()
    report = lottery_db.get_monthly_report(year, month)
    cross = lottery_mensual.cross_check(report, year, month) if report else None
    return render_template(
        "carga_datos_lottery_mensual.html",
        report=report,
        cross=cross,
        **_month_nav(year, month),
        **THEME_BY_KEY["carga_lottery"],
    )


@app.route("/carga-datos/lottery/mensual/eliminar", methods=["POST"])
def carga_datos_lottery_mensual_eliminar():
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    if year and month:
        lottery_db.delete_monthly_report(year, month)
    return redirect(url_for("carga_datos_lottery_mensual", year=year, month=month))


def _month_nav(year, month):
    """year/month, el nombre del mes y los vecinos, para las páginas con navegación de mes."""
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    return {
        "year": year, "month": month, "month_name": _MONTH_NAMES_ES[month - 1],
        "prev_year": prev_year, "prev_month": prev_month, "next_year": next_year, "next_month": next_month,
    }


# Etiquetas de las 11 columnas que suma la fila Subtotal -- pedido
# explícito del usuario (2026-09-16, segunda tanda): "los totales al final
# del bloque... que te muestren la sumatoria de valores que hacen para
# llegar a ese número". Mismo orden que lottery_db._SUBTOTAL_SUM_FIELDS.
_LOTTERY_SUBTOTAL_LABELS = {
    "online_count": "Recuento (ONLINE)",
    "online_net_sales": "Ventas $ (ONLINE)",
    "sales": "Ventas (Terminal)",
    "pagos": "Pagos (Terminal)",
    "comis": "Comis (Terminal)",
    "prize_free_plays": "Premios FP (Terminal)",
    "total_comm": "Total Comm (Terminal)",
    "pays_units": "Pagos U (SKOFF)",
    "pays_amount": "Pagos $ (SKOFF)",
    "skoff_sales_amount": "Monto Ventas (SKOFF)",
    "sales_comm": "Comm Ventas (SKOFF)",
}


def _decorate_lottery_blocks_with_breakdowns(blocks):
    """
    Agrega, a cada bloque ya armado por lottery_db.build_month_blocks, los
    desgloses que necesitan los cuadros flotantes de verificación de
    lottery_historial.html -- Subtotal (suma simple de los 7 días) y
    Debito/Cuenta Final (fórmula sobre el propio Subtotal, ver
    lottery_db._build_block) -- sin tocar lottery_db.py, que no sabe nada
    de cómo se presenta esto en pantalla.
    """
    for block in blocks:
        block["day_breakdown_rows"] = {
            field: [(f"{d['date'][8:10]}/{d['date'][5:7]}", d.get(field)) for d in block["days"]]
            for field in _LOTTERY_SUBTOTAL_LABELS
        }
        subtotal = block["subtotal"]
        debito = block["debito"]
        block["debito_breakdowns"] = {
            "online_net_sales": [  # E =+E-F
                ("Ventas $ (ONLINE, Subtotal)", subtotal.get("online_net_sales")),
                ("Ventas (Terminal, Subtotal) — resta", -subtotal["sales"] if subtotal.get("sales") is not None else None),
            ],
            "sales": [  # F =+F+G+I+K+10
                ("Ventas (Terminal, Subtotal)", subtotal.get("sales")),
                ("Pagos (Terminal, Subtotal)", subtotal.get("pagos")),
                ("Comis (Terminal, Subtotal)", subtotal.get("comis")),
                ("Premios FP (Terminal, Subtotal)", subtotal.get("prize_free_plays")),
                ("Cargo fijo", 10),
            ],
            "pays_amount": [  # Q
                ("Pagos $ (SKOFF, Subtotal)", subtotal.get("pays_amount")),
                ("Monto Ventas (SKOFF, Subtotal)", subtotal.get("skoff_sales_amount")),
                ("Comm Ventas (SKOFF, Subtotal)", subtotal.get("sales_comm")),
            ],
            "net_debit": [  # V =+F+Q (de la fila Debito)
                ("Ventas (F, Debito)", debito.get("sales")),
                ("Pagos $ (Q, Debito)", debito.get("pays_amount")),
            ],
        }
        for d in block["days"]:
            d["cuenta_final_breakdown"] = [  # X =-G-Q
                ("Pagos (G), resta", -d["pagos"] if d.get("pagos") is not None else None),
                ("Pagos $ (Q), resta", -d["pays_amount"] if d.get("pays_amount") is not None else None),
            ]
    return blocks


@app.route("/carga-datos/lottery/historial")
def carga_datos_lottery_historial():
    """Reporte mensual en bloques de 7 días -- ver lottery_db.build_month_blocks / CLAUDE.md."""
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    blocks = _decorate_lottery_blocks_with_breakdowns(lottery_db.build_month_blocks(year, month))
    # Aviso de pago faltante -- pedido explícito del usuario (2026-09-14):
    # cruzar la fecha de Chase Bank YA CONFIRMADA de cada bloque contra los
    # movimientos de Chase ya guardados (Detalle "LOTTERY", ver
    # chase_rules.py) -- si ese día no tiene ningún pago real de Lottery,
    # es señal de que la fecha está mal o de que el pago todavía no se
    # cargó/no llegó. Solo se chequean fechas CONFIRMADAS -- la sugerencia
    # automática (sin confirmar todavía) no se compara contra nada.
    for block in blocks:
        if block["chase_bank_date"]:
            block["chase_payment_missing"] = not chase_db.has_detalle_on_date(
                datetime.strptime(block["chase_bank_date"], "%Y-%m-%d").date(), "LOTTERY"
            )
    debit_total = lottery_db.monthly_debit_total(year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "lottery_historial.html",
        blocks=blocks,
        debit_total=debit_total,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        today_iso=today.isoformat(),
        **THEME_BY_KEY["lottery"],
    )


@app.route("/carga-datos/lottery/exportar")
def carga_datos_lottery_exportar():
    """
    Excel NUEVO (nunca toca el archivo real) con los bloques de Lottery del
    mes -- pedido explícito del usuario (2026-09-19), mismo criterio que
    Store Info/Caja. Ver lottery_db.build_lottery_export_workbook.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    workspace_dir = tempfile.mkdtemp(prefix="lottery_export_")
    dest_path = os.path.join(workspace_dir, f"Lottery {month:02d}-{year}.xlsx")
    lottery_db.build_lottery_export_workbook(year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/lottery/exportar/pdf")
def carga_datos_lottery_exportar_pdf():
    """Versión PDF del export de arriba -- ver lottery_db.build_lottery_export_pdf."""
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    workspace_dir = tempfile.mkdtemp(prefix="lottery_export_pdf_")
    dest_path = os.path.join(workspace_dir, f"Lottery {month:02d}-{year}.pdf")
    lottery_db.build_lottery_export_pdf(year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/lottery/bloque/<int:iso_year>/<int:iso_week>/chase", methods=["POST"])
def carga_datos_lottery_bloque_chase(iso_year, iso_week):
    chase_date_raw = (request.form.get("chase_bank_date") or "").strip()
    try:
        chase_date = datetime.strptime(chase_date_raw, "%Y-%m-%d").date() if chase_date_raw else None
    except ValueError:
        flash("Fecha inválida.", "error")
    else:
        corrected = lottery_db.set_block_chase_date(iso_year, iso_week, chase_date)
        message = "Fecha de Chase Bank guardada." if chase_date else "Fecha de Chase Bank borrada."
        if corrected:
            # Auto-corrección de la cadencia (pedido explícito del usuario
            # 2026-09-12) -- avisar cuáles otros bloques se realinearon solos
            # para que no parezca magia si el usuario nota el cambio después.
            message += f" {len(corrected)} otro(s) bloque(s) se realinearon solos a la cadencia de 7 días."
        # "info" (no "success"): los avisos de éxito ya no se muestran, pero
        # este cuenta algo que cambió además de lo que tocó el usuario.
        flash(message, "info" if corrected else "success")
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    return redirect(url_for("carga_datos_lottery_historial", year=year, month=month))


@app.route("/carga-datos/lottery/dia/<report_date>")
def carga_datos_lottery_dia(report_date):
    try:
        parsed_date = _parse_report_date(report_date)
    except ValueError:
        flash("Fecha inválida.", "error")
        return redirect(url_for("carga_datos_lottery_historial"))

    day = lottery_db.get_day(parsed_date) or lottery_db.blank_day(parsed_date)
    day = lottery_db.decorate_day(day)
    return render_template(
        "lottery_dia.html",
        report_date=parsed_date,
        day=day,
        **THEME_BY_KEY["lottery"],
    )


@app.route("/carga-datos/lottery/dia/<report_date>", methods=["POST"])
def carga_datos_lottery_dia_guardar(report_date):
    fields = {name: request.form.get(name, "").strip() for name in lottery_db.ALL_DAY_FIELDS}
    fields = {name: value for name, value in fields.items() if value != ""}
    try:
        lottery_db.upsert_manual_day(report_date, fields)
    except ValueError:
        flash("No se pudo guardar: revisá que los montos sean números válidos.", "error")
        return redirect(url_for("carga_datos_lottery_dia", report_date=report_date))

    flash("Día guardado.", "success")
    return redirect(url_for("carga_datos_lottery_dia", report_date=report_date))


@app.route("/carga-datos/lottery/dia/<report_date>/pdf/<kind>")
def carga_datos_lottery_dia_pdf(report_date, kind):
    day = lottery_db.get_day(report_date)
    column = {"department": "department_pdf_filename", "sales_report": "sales_report_pdf_filename"}.get(kind)
    relpath = day.get(column) if day and column else None
    path = lottery_db.absolute_pdf_path(relpath) if relpath else None
    if not path or not os.path.isfile(path):
        flash("No hay ningún PDF guardado para este día.", "error")
        return redirect(url_for("carga_datos_lottery_dia", report_date=report_date))
    # Vista previa por default (inline), descarga forzada solo con
    # ?mode=download -- pedido explícito del usuario (2026-09-19), ver
    # templates/_pdf_links.html.
    force_download = request.args.get("mode") == "download"
    return send_file(path, as_attachment=force_download, download_name=os.path.basename(path))


@app.route("/documentos/lottery")
def carga_datos_lottery_documentos():
    """
    PDFs diarios ya guardados este mes -- ver lottery_db.get_month_pdf_list.
    Suma un cuadrito aparte (pedido explícito del usuario, 2026-09-19) para
    el PDF mensual real de Florida Lottery -- solo se guarda, nunca se lee
    ni se procesa (a diferencia del Excel de "Resumen mensual", que si se
    lee -- ver carga_datos_lottery_resumen_mensual). Reusa documents_db.py,
    módulo "lottery_resumen_mensual" -- mismo nombre de siempre, solo
    cambió DÓNDE vive el formulario de carga.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    pdfs = lottery_db.get_month_pdf_list(year, month)
    monthly_pdfs = documents_db.list_documents("lottery_resumen_mensual", year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "lottery_documentos.html",
        pdfs=pdfs,
        monthly_pdfs=monthly_pdfs,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["lottery"],
    )


@app.route("/documentos/lottery/mensual/subir", methods=["POST"])
def carga_datos_lottery_documentos_mensual_subir():
    """
    Sube el único PDF mensual de Florida Lottery -- solo se guarda, ver
    arriba. Un solo PDF por mes -- pedido explícito del usuario (2026-09-16):
    "que no te deje cargar más PDF y que desaparezca el cuadro para cargar
    hasta que se elimine el PDF que se cargó anteriormente". El formulario
    ya se oculta solo en el template cuando ya hay uno -- este chequeo es
    el respaldo del lado del servidor (nunca confiar solo en ocultar un
    botón).
    """
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    upload = request.files.get("pdf_file")
    if not year or not month or not (1 <= month <= 12):
        flash("Elegí a qué mes corresponde este PDF.", "error")
        return redirect(url_for("carga_datos_lottery_documentos"))
    if documents_db.list_documents("lottery_resumen_mensual", year, month):
        flash("Ya hay un PDF mensual cargado este mes -- eliminalo antes de subir otro.", "error")
        return redirect(url_for("carga_datos_lottery_documentos", year=year, month=month))
    if upload is None or not upload.filename:
        flash("Seleccioná el PDF mensual de Lottery.", "error")
        return redirect(url_for("carga_datos_lottery_documentos", year=year, month=month))

    path, filename = _save_upload_to_workspace(upload)
    documents_db.store_document("lottery_resumen_mensual", path, filename, year, month)
    flash("PDF mensual guardado.", "success")
    return redirect(url_for("carga_datos_lottery_documentos", year=year, month=month))


@app.route("/carga-datos/lottery/resumen-mensual")
def carga_datos_lottery_resumen_mensual():
    """
    Cierre mensual de Lottery -- pedido explícito del usuario (2026-09-16):
    "lo ideal es no tener que subir ningún excel, ya que los datos debería
    poder extraerlos de la misma lottery que se carga en la página". Las
    dos tablas que este cuadro tiene en el Excel real ("LIQUIDACION CIERRE
    LOTTERY" / "LIQUIDACION CIERRE RECAUDACION COMISIONES") resultaron ser
    puras sumas de columnas que ya viven en lottery_days -- se recalculan
    en el momento (ver lottery_db.compute_month_closing), nada se sube ni
    se lee de ningún archivo. El único valor manual real es "Gastos
    Adminits-Loteria" ($150 por default, editable) -- ver la ruta de abajo.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    closing = lottery_db.compute_month_closing(year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "lottery_resumen_mensual.html",
        closing=closing,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["lottery"],
    )




@app.route("/controles")
def controles():
    # La tarjeta de un control muestra su alerta sin tener que entrar.
    alerts = {}
    try:
        latest = control_tarjetas.latest_status(reportes_db.get_card_sales_by_date(), eft_db.get_coupon_gross_by_date())
        if latest and latest["status"] == "alert":
            alerts["control_tarjetas"] = f"Alerta: pendiente ${latest['pending']:,.2f} al {_fmt_ddmmyyyy(latest['date'])}"
    except Exception as exc:
        print(f"[controles] estado de Tarjetas y Cupones: {exc}")
    try:
        today = date.today()
        deposits_control = _depositos_month_control(today.year, today.month)
        if deposits_control["issues"]:
            alerts["control_depositos"] = f"{len(deposits_control['issues'])} cosa(s) para revisar este mes"
    except Exception as exc:
        print(f"[controles] estado de Control Depósitos: {exc}")
    return render_template("controles_index.html", controls=CONTROLES_SECTIONS, alerts=alerts)


@app.route("/controles/cierre-mensual", methods=["GET", "POST"])
def control_cierre_mensual():
    # A diferencia de Herramientas, este control nunca descarga un archivo
    # -- solo lee los dos que suben y muestra el resultado en la misma
    # página (ver CLAUDE.md, "Módulo Controles": reporte en pantalla,
    # verde/rojo por chequeo) -- por eso no usa ajax-process-form ni
    # _success_response, un submit normal alcanza.
    result = None
    department_result = None
    if request.method == "POST":
        cierre_upload = request.files.get("cierre_file")
        ventas_upload = request.files.get("ventas_file")
        pdf_upload = request.files.get("monthly_pdf")
        if cierre_upload is None or not cierre_upload.filename:
            flash("Seleccioná el Excel Cierre del mes.", "error")
        elif ventas_upload is None or not ventas_upload.filename:
            flash("Seleccioná el Excel de Ventas del mes.", "error")
        elif pdf_upload is None or not pdf_upload.filename:
            flash("Seleccioná el PDF de Resumen de Ventas del mes.", "error")
        else:
            try:
                workdir = _new_workspace_dir()
                cierre_path, _cierre_filename = _save_upload_to_workspace(cierre_upload, workdir=workdir)
                ventas_path, _ventas_filename = _save_upload_to_workspace(ventas_upload, workdir=workdir)
                pdf_path, _pdf_filename = _save_upload_to_workspace(pdf_upload, workdir=workdir)
            except Exception as exc:
                flash(f"Error: {exc}", "error")
            else:
                # Los dos chequeos se aíslan por separado -- un problema con
                # el Excel de Ventas (o con la sección "Department Sales
                # Report" del PDF) no debe tirar abajo el resultado de Store
                # Info si ese sí se pudo calcular bien, y viceversa (bug real
                # corregido 2026-09-17: antes un solo try/except compartido
                # descartaba ambos resultados apenas uno de los dos fallaba).
                try:
                    result = check_store_info_monthly(cierre_path, pdf_path)
                except Exception as exc:
                    flash(f"Error en Store Info: {exc}", "error")
                try:
                    department_result = check_department_sales_monthly(ventas_path, pdf_path)
                except Exception as exc:
                    flash(f"Error en Ventas por Departamento: {exc}", "error")
    return render_template(
        "control_cierre_mensual.html",
        result=result,
        department_result=department_result,
        **THEME_BY_KEY["cierre_mensual"],
    )


@app.route("/controles/lottery-mensual", methods=["GET", "POST"])
def control_lottery_mensual():
    # Mismo criterio que control_cierre_mensual: solo lectura, reporte en
    # pantalla, sin ajax-process-form ni _success_response.
    result = None
    if request.method == "POST":
        lottery_upload = request.files.get("lottery_file")
        pdf_upload = request.files.get("monthly_pdf")
        if lottery_upload is None or not lottery_upload.filename:
            flash("Seleccioná el Excel de Lottery del mes.", "error")
        elif pdf_upload is None or not pdf_upload.filename:
            flash("Seleccioná el Monthly Sales Report (PDF) del mes.", "error")
        else:
            try:
                workdir = _new_workspace_dir()
                lottery_path, _lottery_filename = _save_upload_to_workspace(lottery_upload, workdir=workdir)
                pdf_path, _pdf_filename = _save_upload_to_workspace(pdf_upload, workdir=workdir)
                result = check_lottery_monthly(lottery_path, pdf_path)
            except Exception as exc:
                flash(f"Error: {exc}", "error")
    return render_template(
        "control_lottery_mensual.html", result=result, **THEME_BY_KEY["lottery_mensual"]
    )


@app.route("/controles/cupones", methods=["GET", "POST"])
def control_cupones():
    # Mismo criterio que los otros controles: solo lectura, reporte en
    # pantalla, sin ajax-process-form ni _success_response.
    result = None
    if request.method == "POST":
        mayor_upload = request.files.get("mayor_file")
        eft_upload = request.files.get("eft_excel_file")
        if mayor_upload is None or not mayor_upload.filename:
            flash("Seleccioná el Mayor de Recaudación a Liquidar.", "error")
        elif eft_upload is None or not eft_upload.filename:
            flash("Seleccioná el Excel de Aplicacion TC y EFT.", "error")
        else:
            try:
                workdir = _new_workspace_dir()
                mayor_path, _mayor_filename = _save_upload_to_workspace(mayor_upload, workdir=workdir)
                eft_path, _eft_filename = _save_upload_to_workspace(eft_upload, workdir=workdir)
                result = check_cupones_pending(mayor_path, eft_path)
            except Exception as exc:
                flash(f"Error: {exc}", "error")
    return render_template("control_cupones.html", result=result, **THEME_BY_KEY["cupones"])


@app.route("/controles/mercaderia", methods=["GET", "POST"])
def control_mercaderia():
    # Mismo criterio que los otros controles: solo lectura, reporte en
    # pantalla, sin ajax-process-form ni _success_response.
    result = None
    if request.method == "POST":
        proveedores_upload = request.files.get("proveedores_file")
        mayor_upload = request.files.get("mayor_file")
        if proveedores_upload is None or not proveedores_upload.filename:
            flash("Seleccioná el Excel de Proveedores (Cta Cte).", "error")
        elif mayor_upload is None or not mayor_upload.filename:
            flash("Seleccioná el Mayor de Mercadería en C-Store.", "error")
        else:
            try:
                workdir = _new_workspace_dir()
                proveedores_path, _proveedores_filename = _save_upload_to_workspace(proveedores_upload, workdir=workdir)
                mayor_path, _mayor_filename = _save_upload_to_workspace(mayor_upload, workdir=workdir)
                result = check_mercaderia_invoices(proveedores_path, mayor_path)
            except Exception as exc:
                flash(f"Error: {exc}", "error")
    return render_template("control_mercaderia.html", result=result, **THEME_BY_KEY["mercaderia"])


@app.route("/controles/valuacion", methods=["GET", "POST"])
def control_valuacion():
    # A diferencia de los otros controles, este SÍ completa una copia del
    # Excel BGS (nunca el original) -- pedido explícito del usuario
    # (2026-09-08): "quiero que si complete el excel, pero que el archivo
    # aparezca abajo del reporte para que el usuario lo pueda descargar
    # por su cuenta". Se guarda en un temporal y se ofrece aparte con un
    # link (ver control_valuacion_descargar) en vez de forzar la descarga
    # como hace un módulo de Herramientas.
    result = None
    if request.method == "POST":
        mayor_upload = request.files.get("mayor_file")
        bgs_upload = request.files.get("bgs_file")
        chevron_upload = request.files.get("chevron_file")
        if mayor_upload is None or not mayor_upload.filename:
            flash("Seleccioná el Mayor de Mercadería en C-Store.", "error")
        elif bgs_upload is None or not bgs_upload.filename:
            flash("Seleccioná el Excel BGS.", "error")
        elif chevron_upload is None or not chevron_upload.filename:
            flash("Seleccioná el Chevron Category Cost Report.", "error")
        else:
            try:
                workdir = _new_workspace_dir()
                mayor_path, _mayor_filename = _save_upload_to_workspace(mayor_upload, workdir=workdir)
                bgs_path, _bgs_filename = _save_upload_to_workspace(bgs_upload, workdir=workdir)
                chevron_path, _chevron_filename = _save_upload_to_workspace(chevron_upload, workdir=workdir)
                result = check_and_complete_valuation(mayor_path, bgs_path, chevron_path)
                # El temporal de una corrida anterior (si había uno, ej. el
                # usuario corrigió algo y volvió a subir los 3 archivos) se
                # borra antes de reemplazar el puntero en la sesión -- si no,
                # queda huérfano para siempre en la carpeta temporal del
                # sistema (bug real corregido 2026-09-17: nada lo borraba
                # nunca, ni siquiera después de descargarlo).
                previous_path = session.get("valuacion_download_path")
                if previous_path and previous_path != result["download_path"] and os.path.isfile(previous_path):
                    try:
                        os.remove(previous_path)
                    except OSError:
                        pass
                session["valuacion_download_path"] = result["download_path"]
                session["valuacion_download_filename"] = result["download_filename"]
            except Exception as exc:
                flash(f"Error: {exc}", "error")
    return render_template("control_valuacion.html", result=result, **THEME_BY_KEY["valuacion"])


@app.route("/controles/valuacion/descargar")
def control_valuacion_descargar():
    path = session.get("valuacion_download_path")
    filename = session.get("valuacion_download_filename")
    if not path or not filename or not os.path.isfile(path):
        flash("El archivo ya no está disponible -- volvé a correr el control de nuevo.", "error")
        return redirect(url_for("control_valuacion"))
    return send_file(path, as_attachment=True, download_name=filename)


@app.route("/controles/caja", methods=["GET", "POST"])
def control_caja():
    # Mismo criterio que los otros controles: solo lectura, reporte en
    # pantalla, sin ajax-process-form ni _success_response. A diferencia de
    # los otros 4, esta ruta sí hace POST/Redirect/GET -- pedido explícito
    # del usuario (2026-09-10): que al recargar la página (F5) el cuadro de
    # comparación anterior desaparezca y quede limpia. El resultado viaja
    # en `session` (formateado a tipos JSON simples desde controles_caja.py)
    # y se lee con `.pop()`, así que se muestra una sola vez -- justo
    # después de enviar el formulario -- y cualquier recarga posterior de
    # esa misma página ya no lo encuentra.
    if request.method == "POST":
        cierre_upload = request.files.get("cierre_file")
        mayor_chase_upload = request.files.get("mayor_chase_file")
        mayor_caja_upload = request.files.get("mayor_caja_file")
        if cierre_upload is None or not cierre_upload.filename:
            flash("Seleccioná el Excel Cierre.", "error")
        elif mayor_chase_upload is None or not mayor_chase_upload.filename:
            flash("Seleccioná el Mayor de la cuenta Chase Bank.", "error")
        elif mayor_caja_upload is None or not mayor_caja_upload.filename:
            flash("Seleccioná el Mayor de la cuenta Caja.", "error")
        else:
            try:
                workdir = _new_workspace_dir()
                cierre_path, _cierre_filename = _save_upload_to_workspace(cierre_upload, workdir=workdir)
                mayor_chase_path, _mayor_chase_filename = _save_upload_to_workspace(mayor_chase_upload, workdir=workdir)
                mayor_caja_path, _mayor_caja_filename = _save_upload_to_workspace(mayor_caja_upload, workdir=workdir)
                session["caja_control_result"] = check_caja_mayores(cierre_path, mayor_chase_path, mayor_caja_path)
            except Exception as exc:
                flash(f"Error: {exc}", "error")
        return redirect(url_for("control_caja"))
    result = session.pop("caja_control_result", None)
    return render_template("control_caja.html", result=result, **THEME_BY_KEY["caja"])


@app.route("/carga-datos/chase", methods=["GET", "POST"])
def chase():
    """
    Categoriza el extracto de Chase y lo guarda en chase_db.py -- ya NO
    genera ni descarga ningún Excel (2026-09-11, ver CLAUDE.md "Conversión
    de Chase Bank a Carga de Datos"). Cada carga recategoriza fresca contra
    las reglas vigentes -- subir el mismo extracto (o uno solapado) de
    nuevo actualiza el Detalle guardado en vez de duplicar el movimiento.
    """
    if request.method == "GET":
        return render_template(
            "chase.html", chase_rules=list_chase_display_rules(), **THEME_BY_KEY["carga_chase"]
        )

    upload = request.files.get("chase_file")
    if upload is None or not upload.filename:
        return _error_response("Seleccioná un archivo CSV o Excel de Chase.")

    temp_path = None
    filename = None
    try:
        temp_path, filename = _save_upload_to_workspace(upload)
        rows, total_rows = extract_chase_transactions(temp_path)
    except Exception as exc:
        # pandas a veces incrusta la ruta completa (que contiene el nombre
        # real del archivo) en el texto de su propia excepción -- a
        # diferencia de otros módulos, acá no hay una capa propia que arme
        # un mensaje corto antes, así que se scrubea el texto crudo.
        message = str(exc)
        if temp_path:
            message = message.replace(temp_path, "el archivo")
        if filename:
            message = message.replace(filename, "el archivo")
        return _error_response(f"Error: {message}")

    if not rows:
        return _error_response("No se encontró ningún movimiento con fecha válida en el archivo.")

    # Vínculo directo con un proveedor puntual (2026-09-21, pedido explícito
    # del usuario) -- independiente del Detalle "PROVEEDORES" ya calculado
    # arriba por extract_chase_transactions, ver
    # proveedores.match_supplier_for_chase_description.
    for row in rows:
        row["supplier_key"] = _resolve_supplier_key_for_chase(row["description"])

    inserted, updated = chase_db.upsert_transactions(rows, source_filename=filename)
    skipped = total_rows - len(rows)
    # Los recibos de Ice Machine ya cargados categorizan sus depósitos.
    linked = _link_ice_deposits_to_chase()
    uncategorized = max(0, sum(1 for row in rows if not row["detalle"]) - linked)

    parts = [f"{len(rows)} movimiento(s) guardado(s) ({inserted} nuevo(s), {updated} actualizado(s))."]
    if linked:
        parts.append(f"{linked} depósito(s) categorizados por su recibo (Ice Machine o Vaccumms).")
    if uncategorized:
        parts.append(f"{uncategorized} sin ninguna regla que matcheara.")
    if skipped:
        parts.append(f"{skipped} fila(s) sin fecha válida, no se guardaron.")
    flash(" ".join(parts), "warning" if (uncategorized or skipped) else "success")

    first_date = min(row["posting_date"] for row in rows)
    return redirect(url_for("chase_historial", year=first_date.year, month=first_date.month))


def _require_admin(message):
    """
    Gate genérico para acciones admin-only (reglas de Chase, agregar
    proveedores nuevos sin código, etc.) -- flashea `message` y devuelve
    False si el usuario logueado no es admin.
    """
    if not current_user.is_admin:
        flash(message, "error")
        return False
    return True


def _resolve_supplier_key_for_chase(description):
    """Solo la supplier_key (sin el label) -- forma que pide chase_db.recategorize_all."""
    supplier_key, _supplier_label = match_supplier_for_chase_description(description)
    return supplier_key


def _recategorize_all_chase_and_flash():
    """
    Recalcula Detalle y proveedor vinculado de TODOS los movimientos de
    Chase ya guardados contra las reglas vigentes -- pedido explícito del
    usuario (2026-09-21): crear/editar/eliminar una regla (de Chase o de
    pago a proveedores) tiene que aplicarse solo a lo ya cargado, sin
    obligar a resubir el extracto del banco de nuevo. Se llama después de
    cualquier cambio a chase_rules.json/chase_master_rules.json/
    proveedores_pago_rules.json. Nunca toca un valor ya corregido a mano
    (ver chase_db.recategorize_all).
    """
    detalle_changed, supplier_changed = chase_db.recategorize_all(
        categorize_chase_description, _resolve_supplier_key_for_chase
    )
    if detalle_changed or supplier_changed:
        parts = []
        if detalle_changed:
            parts.append(f"{detalle_changed} movimiento(s) recategorizado(s)")
        if supplier_changed:
            parts.append(f"{supplier_changed} movimiento(s) vinculado(s)/desvinculado(s) de un proveedor")
        flash("Aplicado a lo ya guardado: " + ", ".join(parts) + ".", "info")


@app.route("/carga-datos/chase/rules/save", methods=["POST"])
def chase_rules_save():
    if not _require_admin("Solo un administrador puede gestionar las reglas de Chase."):
        return redirect(url_for("chase"))

    keyword = request.form.get("keyword", "")
    detail = request.form.get("detail", "")
    rule_type = request.form.get("rule_type", "").strip()
    index = request.form.get("index", "").strip()
    # Si vienen (edición desde la tabla), protegen contra la carrera de
    # índice desactualizado: otra pestaña/sesión pudo haber editado o
    # borrado una regla en el medio, corriendo los índices de todo lo que
    # está después -- ver _check_expected_rule en chase_rules.py.
    expected_keyword = request.form.get("expected_keyword") or None
    expected_detail = request.form.get("expected_detail") or None

    try:
        if not rule_type or not index:
            add_chase_rule(keyword, detail)
            flash("Regla creada.", "success")
        elif rule_type == "master":
            edit_chase_master_rule(index, keyword, detail, expected_keyword, expected_detail)
            flash("Regla Maestra actualizada.", "success")
        elif rule_type == "custom":
            edit_chase_custom_rule(index, keyword, detail, expected_keyword, expected_detail)
            flash("Regla actualizada.", "success")
        else:
            flash("Tipo de regla inválido.", "error")
            return redirect(url_for("chase"))
        _recategorize_all_chase_and_flash()
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("chase"))


@app.route("/carga-datos/chase/rules/delete", methods=["POST"])
def chase_rules_delete():
    if not _require_admin("Solo un administrador puede gestionar las reglas de Chase."):
        return redirect(url_for("chase"))

    rule_type = request.form.get("rule_type", "").strip()
    index = request.form.get("index", "").strip()
    expected_keyword = request.form.get("expected_keyword") or None
    expected_detail = request.form.get("expected_detail") or None

    try:
        if rule_type == "master":
            delete_chase_master_rule(index, expected_keyword, expected_detail)
            flash("Regla Maestra eliminada.", "success")
        elif rule_type == "custom":
            delete_chase_custom_rule(index, expected_keyword, expected_detail)
            flash("Regla eliminada.", "success")
        else:
            flash("Seleccioná una regla de la tabla antes de eliminar.", "error")
            return redirect(url_for("chase"))
        _recategorize_all_chase_and_flash()
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("chase"))


@app.route("/carga-datos/chase/historial")
def chase_historial():
    """Movimientos de Chase guardados de un mes, ordenados por fecha -- ver chase_db.py."""
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    transactions = chase_db.get_month_transactions(year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    # Vínculo a proveedor (2026-09-21) -- supplier_options alimenta el
    # <select> del popover de categorización manual; supplier_labels
    # resuelve el nombre a mostrar de la supplier_key ya guardada en cada
    # movimiento (chase_db solo guarda la clave, nunca el label -- así un
    # proveedor renombrado no queda desactualizado acá).
    supplier_options = list_supplier_registry_entries()
    supplier_labels = {entry["key"]: entry["label"] for entry in supplier_options}

    return render_template(
        "chase_historial.html",
        transactions=transactions,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        # "DEPOSITO VENTA ICE" siempre disponible para elegir a mano -- esos
        # depósitos chicos ya no se categorizan solos (ver
        # chase_rules._split_small_deposit), así que puede no estar usado todavía.
        known_details=sorted(set(chase_db.list_known_details()) | {"DEPOSITO VENTA ICE", CHASE_DETALLE_VACCUMMS}),
        supplier_options=supplier_options,
        supplier_labels=supplier_labels,
        **THEME_BY_KEY["carga_chase"],
    )


@app.route("/carga-datos/chase/exportar")
def chase_exportar():
    """
    Descarga un Excel NUEVO (nunca toca ningún archivo del banco) con los
    movimientos ya categorizados de un mes -- pedido explícito del usuario
    (2026-09-14), ver chase_rules.build_chase_export_workbook.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    transactions = chase_db.get_month_transactions(year, month)
    if not transactions:
        flash("No hay ningún movimiento guardado ese mes para exportar.", "error")
        return redirect(url_for("chase_historial", year=year, month=month))

    workspace_dir = tempfile.mkdtemp(prefix="chase_export_")
    dest_path = os.path.join(workspace_dir, f"Chase {month:02d}-{year}.xlsx")
    build_chase_export_workbook(transactions, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/chase/categorizar", methods=["POST"])
def chase_categorizar():
    """
    Categorización manual de un movimiento puntual -- pedido explícito del
    usuario (2026-09-12): "los datos sin categorizar del chase se puedan
    categorizar". Queda pegado (detalle_source="manual") aunque se vuelva a
    subir el mismo extracto después, ver chase_db.set_manual_detalle.

    2026-09-21, ampliado: el mismo form también deja elegir a mano a qué
    proveedor pertenece el movimiento (`supplier_key`, vía
    chase_db.set_manual_supplier) -- necesario para un pago sin ningún
    texto reusable en la Descripción (ej. un cheque, "CHECK 1770"), donde
    ninguna regla de palabra clave puede resolverlo sola.

    2026-09-22: vincular un proveedor a un movimiento que todavía no
    tenía Detalle lo marca solo como "PROVEEDORES" -- pedido explícito
    del usuario ("cuando lo haga deberia ponerse automaticamente
    PROVEEDORES al lado del asiento... asi va a quedar mas prolijo"),
    para que un asiento vinculado nunca se quede mostrando "Sin
    categorizar". Nunca pisa un Detalle que el movimiento ya tenía (de
    una regla o cargado a mano antes) ni lo que el usuario haya tipeado
    a mano en el mismo form -- solo completa el hueco cuando estaba
    vacío de los dos lados.

    2026-09-22, más tarde: al revés -- si el movimiento SÍ tenía un
    proveedor vinculado y se lo saca (seleccionando "ninguno"), y el
    Detalle sigue siendo el "PROVEEDORES" que quedó puesto por ese mismo
    vínculo (sin que el usuario lo haya tipeado distinto en el mismo
    form), vuelve a quedar "Sin categorizar" -- pedido explícito del
    usuario: "si me confundi y no era de eso... deberia quedar en sin
    categorizar" en vez de seguir mostrando PROVEEDORES sin nadie
    vinculado.
    """
    posting_date = request.form.get("posting_date", "").strip()
    description = request.form.get("description", "")
    amount_raw = request.form.get("amount", "").strip()
    detalle = request.form.get("detalle", "").strip()
    current_detalle = request.form.get("current_detalle", "").strip()
    current_supplier_key = request.form.get("current_supplier_key", "").strip()
    supplier_key = request.form.get("supplier_key", "").strip()
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)

    try:
        amount = float(amount_raw)
    except ValueError:
        flash("No se pudo identificar el movimiento (monto inválido).", "error")
        return redirect(url_for("chase_historial", year=year, month=month))

    if supplier_key and not current_detalle and not detalle:
        detalle = "PROVEEDORES"
    elif (
        not supplier_key
        and current_supplier_key
        and detalle == "PROVEEDORES"
        and current_detalle == "PROVEEDORES"
    ):
        detalle = ""

    # Cada campo se marca "manual" solo si de verdad cambió: guardar el
    # popover tal cual (por ejemplo para tocar solo el Detalle) no debe
    # congelar contra futuras reglas un Detalle ni un proveedor que venían
    # de una regla (auditoría 2026-09, webapp.py:2751). El rowcount de lo
    # que sí se guarda es la señal de "el movimiento existe".
    found = True
    if supplier_key != current_supplier_key:
        found = chase_db.set_manual_supplier(posting_date, description, amount, supplier_key)
    if detalle != current_detalle:
        found = chase_db.set_manual_detalle(posting_date, description, amount, detalle) and found

    if not found:
        # Sin aviso de éxito -- pedido explícito del usuario (2026-09-22):
        # "esta notificacion quiero que la quites, no hace falta" -- el
        # badge de Detalle (y el popover del proveedor vinculado, ver
        # chase_historial.html) ya muestran el resultado al instante, sin
        # hacer falta un flash aparte. El único aviso que queda es el de
        # error real (el movimiento ya no está guardado).
        flash("No se encontró ese movimiento -- puede que ya no esté guardado.", "error")
    return redirect(url_for("chase_historial", year=year, month=month))


# ---------------------------------------------------------------------------
# Chase Bank -> Cheques (2026-09-23) -- pedido explícito del usuario: guardar
# los PDF de los cheques propios (sueltos o adjuntos a la factura del
# proveedor) y llevar el control en UN solo cuadro, en orden de N° de cheque,
# igual que su Excel "Cheques for pays suppliers. Control.xls" (N° / Fecha /
# Monto / Proveedor / Invoice N° / Debitado en Chase).
#
# De dónde sale cada dato (lo tipeado a mano siempre gana):
# - N°: OCR de la esquina del cheque (cheques.py).
# - Proveedor / Invoice / Fecha / Monto: la factura del proveedor que viene en
#   el mismo PDF (se lee al subir, mismo motor que Carga de Datos -> Proveedores);
#   si el cheque vino suelto, el proveedor vinculado al "CHECK n" de Chase o el
#   nombre del archivo, y la factura se busca entre las ya cargadas de ese
#   proveedor por el mismo monto.
# - Debitado en Chase: el movimiento "CHECK n" de chase_db, cruzado en el
#   momento (una carga nueva del extracto lo actualiza sola).
# ---------------------------------------------------------------------------


def _supplier_from_filename(filename, registry_entries):
    """'Cheque 1773. Midtown Wholesale.pdf' -> ('midtown', 'Midtown Wholesale LLC'), o (None, None)."""
    # Antes alcanzaba con la primera palabra del proveedor: "Florida Audit
    # Service" se leía como FPL ("Florida Power & Light"). Ahora tiene que
    # aparecer la clave del proveedor, o TODAS las palabras significativas de
    # su nombre (auditoría 2026-09, cheques 1749/1761).
    generic = {"inc", "llc", "co", "corp", "company", "the", "of", "and", "usa", "dist", "fl"}
    name = re.sub(r"[^a-z0-9]+", " ", (filename or "").lower())
    words = set(name.split())
    for entry in registry_entries:
        label_words = [w for w in re.sub(r"[^a-z0-9]+", " ", entry["label"].lower()).split() if w not in generic]
        key_words = entry["key"].split("_")
        joined_key = entry["key"].replace("_", "")
        if (all(w in words for w in key_words)
                or (len(joined_key) >= 6 and joined_key in name.replace(" ", ""))  # "Sky Harvest" -> skyharvest
                or (label_words and all(w in words for w in label_words))):
            return entry["key"], entry["label"]
    return None, None


def _pdf_without_pages(pdf_path, skip_pages):
    """Copia temporal del PDF sin las páginas del cheque (algunos extractores de proveedor se confunden con él)."""
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(pdf_path)
    keep = [i for i in range(len(reader.pages)) if i not in skip_pages]
    if not keep or len(keep) == len(reader.pages):
        return None
    writer = PdfWriter()
    for i in keep:
        writer.add_page(reader.pages[i])
    out_path = os.path.join(_new_workspace_dir(), os.path.basename(pdf_path))
    with open(out_path, "wb") as handle:
        writer.write(handle)
    return out_path


def _read_check_invoice(pdf_path, filename, check_pages=()):
    """
    Datos automáticos del cheque a partir de la factura del proveedor que
    viene en el mismo PDF (se lee sin las páginas del cheque, que confunden a
    algunos extractores). Nunca lanza: si se reconoce el proveedor pero no la
    factura, queda al menos el proveedor; si nada, el proveedor según el
    nombre del archivo; y si tampoco, vacío.
    """
    import proveedores as proveedores_module
    auto = {}
    candidates = []
    trimmed = None
    try:
        trimmed = _pdf_without_pages(pdf_path, set(check_pages)) if check_pages else None
    except Exception as exc:
        print(f"[cheques] no se pudo separar la factura de {filename}: {exc}")
    candidates = [path for path in (trimmed, pdf_path) if path]
    registry = proveedores_module._effective_supplier_registry()
    for path in candidates:
        try:
            supplier_key = proveedores_module._detect_supplier(path)
        except Exception:
            continue
        auto.update(auto_supplier_key=supplier_key, auto_supplier_label=registry[supplier_key]["label"])
        try:
            result = registry[supplier_key]["extract"](path)
        except Exception as exc:
            print(f"[cheques] proveedor {supplier_key} reconocido pero sin factura legible en {filename}: {exc}")
            continue
        invoices, _ = proveedores_module.invoice_errors(result if isinstance(result, list) else [result])
        if invoices:
            auto["auto_invoice_no"] = " ".join(str(inv["invoice_no"]) for inv in invoices if inv.get("invoice_no")) or None
            dates = [inv["date"] for inv in invoices if inv.get("date")]
            if dates:
                auto["auto_invoice_date"] = max(dates).strftime("%Y-%m-%d")
            amounts = [inv["amount"] for inv in invoices if inv.get("amount") is not None]
            if amounts:
                auto["auto_amount"] = round(sum(amounts), 2)
        break
    if not auto.get("auto_supplier_key"):
        key, label = _supplier_from_filename(filename, list_supplier_registry_entries())
        if key:
            auto.update(auto_supplier_key=key, auto_supplier_label=label)
    return auto


def _fmt_ddmmyyyy(iso_text):
    return f"{iso_text[8:10]}/{iso_text[5:7]}/{iso_text[:4]}" if iso_text and len(iso_text) >= 10 else None


def _match_invoice_by_amount(invoices, amount, before_iso):
    """
    La factura del proveedor con ese monto exacto (anterior al débito), solo
    si es única: con dos o más candidatas no se adivina (auditoría 2026-09,
    elegir "la más reciente" asignaba la misma factura a varios cheques).
    """
    if amount is None:
        return None
    candidates = [
        inv for inv in invoices
        if abs(inv["amount"] - amount) < 0.01 and (not before_iso or inv["invoice_date"] <= before_iso)
    ]
    return candidates[0] if len(candidates) == 1 else None


_MAX_CHECK_GAP = 50


def _build_cheques_rows():
    """Una fila por N° de cheque (cargados, cobrados en Chase, y los huecos entre medio), en orden."""
    records = cheques_db.list_checks()
    registry = list_supplier_registry_entries()
    supplier_labels = {entry["key"]: entry["label"] for entry in registry}

    chase_by_number = {}
    for tx in chase_db.get_check_transactions():
        number = check_number_from_chase_description(tx["description"])
        if number is not None:
            chase_by_number.setdefault(number, tx)

    by_number = {r["check_number"]: r for r in records if r["check_number"] is not None}
    numbers = set(by_number) | set(chase_by_number)
    # Huecos: se ven igual que en el Excel. Solo entre números cercanos -- un
    # N° mal leído por el OCR (ej. 775 en vez de 1775) no tiene que llenar el
    # cuadro con cientos de filas vacías.
    known = sorted(numbers)
    for low, high in zip(known, known[1:]):
        if high - low <= _MAX_CHECK_GAP:
            numbers |= set(range(low + 1, high))

    invoice_cache = {}

    def supplier_invoices(key):
        if key not in invoice_cache:
            invoice_cache[key] = proveedores_db.get_supplier_invoices(key) if key else []
        return invoice_cache[key]

    def build(record, number):
        record = record or {}
        tx = chase_by_number.get(number) if number is not None else None
        chase_amount = round(-tx["amount"], 2) if tx else None
        supplier_key = record.get("auto_supplier_key") or (tx.get("supplier_key") if tx else None)
        if not supplier_key and not record.get("manual_supplier"):
            supplier_key, _ = _supplier_from_filename(record.get("source_filename"), registry)
        supplier = (
            record.get("manual_supplier")
            or record.get("auto_supplier_label")
            or supplier_labels.get(supplier_key)
        )
        if not supplier and tx:
            supplier = re.sub(r"^\s*CHE(?:CK|QUE)\s*#?\s*\d+\s*(\d{2}/\d{2}\s*)?", "", tx["description"], flags=re.I).strip() or None

        amount = record.get("manual_amount")
        if amount is None:
            amount = chase_amount if chase_amount is not None else record.get("auto_amount")

        invoice_no = record.get("manual_invoice") or record.get("auto_invoice_no")
        invoice_date = record.get("auto_invoice_date")
        if not invoice_no and supplier_key and amount is not None:
            match = _match_invoice_by_amount(supplier_invoices(supplier_key), amount, tx["posting_date"] if tx else None)
            if match:
                invoice_no, invoice_date = match["invoice_no"], invoice_date or match["invoice_date"]

        debit_iso = record.get("manual_debit_date") or (tx["posting_date"] if tx else None)
        date_iso = record.get("manual_date") or invoice_date
        if record.get("voided"):
            status = "Anulado"
        elif number is None:
            status = "Sin N°"
        elif debit_iso:
            status = "Cobrado"
        elif record:
            status = "Pendiente"
        else:
            status = "Sin registro"
        return {
            "id": record.get("id"),
            "number": number,
            "has_pdf": documents_db.GUARDAR_DOCUMENTOS and bool(record.get("check_pdf")),
            "source_filename": record.get("source_filename"),
            "date": _fmt_ddmmyyyy(date_iso),
            "date_iso": date_iso,
            "amount": amount,
            "supplier": supplier,
            "invoice_no": invoice_no,
            "debit": _fmt_ddmmyyyy(debit_iso),
            "debit_iso": debit_iso,
            "status": status,
            "voided": bool(record.get("voided")),
            "manual": {k: record.get(k) for k in (
                "manual_date", "manual_amount", "manual_supplier", "manual_invoice", "manual_debit_date")},
        }

    rows = [build(by_number.get(n), n) for n in sorted(numbers)]
    rows += [build(r, None) for r in records if r["check_number"] is None]
    return rows


@app.route("/carga-datos/chase/cheques")
def chase_cheques():
    rows = _build_cheques_rows()
    stats = {
        "total": sum(1 for r in rows if r["has_pdf"]),
        "cobrados": sum(1 for r in rows if r["status"] == "Cobrado"),
        "pendientes": sum(1 for r in rows if r["status"] == "Pendiente"),
        "sin_pdf": sum(1 for r in rows if r["status"] == "Cobrado" and not r["has_pdf"]),
    }
    _active_job = jobs.get_active_job("cheques")
    return render_template(
        "chase_cheques.html",
        rows=rows,
        stats=stats,
        supplier_names=sorted({e["label"] for e in list_supplier_registry_entries()}
                              | {r["supplier"] for r in rows if r["supplier"]}),
        resume_job_id=(_active_job["id"] if _active_job else None),
        **THEME_BY_KEY["carga_chase"],
    )


@app.route("/carga-datos/chase/cheques/subir", methods=["POST"])
def chase_cheques_subir():
    pdf_uploads = request.files.getlist("pdf_files")
    if not pdf_uploads or not any(u.filename for u in pdf_uploads):
        return _error_response("Seleccioná uno o más PDF con cheques.")
    pdf_paths = _save_uploads_to_workspace(pdf_uploads)
    job_id = jobs.create_job(len(pdf_paths), kind="cheques")
    threading.Thread(target=_run_chase_cheques_job, args=(job_id, pdf_paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(pdf_paths)})


def _run_chase_cheques_job(job_id, pdf_paths):
    """Mismo patrón que _run_carga_datos_lottery_job: aislado por PDF, avisos cortos sin nombres de archivo."""
    try:
        saved = unreadable_number = duplicates = no_check = failed = 0
        for index, pdf_path in enumerate(pdf_paths, start=1):
            filename = os.path.basename(pdf_path)
            try:
                found = extract_checks_from_pdf(pdf_path)
                if not found:
                    no_check += 1
                original_rel = None
                auto = None
                for item in found:
                    number = item["number"]
                    if number is not None:
                        existing = cheques_db.find_by_number(number)
                    else:
                        # Sin N° legible el duplicado se reconoce por archivo y
                        # página: si no, cada resubida crea otra fila "Sin N°".
                        existing = cheques_db.find_by_source(filename, item["page"])
                        if existing:
                            duplicates += 1
                            continue
                    # Sin guardado de documentos un cheque leído por OCR
                    # queda con check_pdf vacío, igual que uno cargado a
                    # mano -- number_source es lo que los distingue.
                    if existing and (existing["check_pdf"] or existing["number_source"] == "ocr"):
                        duplicates += 1
                        continue
                    if original_rel is None:
                        original_rel = cheques_db.store_original(pdf_path, filename)
                        # Un PDF con varios cheques comparte la misma factura -- se lee una vez.
                        auto = _read_check_invoice(pdf_path, filename, [c["page"] for c in found]) if len(found) == 1 else {}
                    cheques_db.add_check(number, item["image"], original_rel, filename, item["page"], auto=auto)
                    saved += 1
                    if number is None:
                        unreadable_number += 1
            except Exception as exc:
                print(f"[carga-datos/chase/cheques] {pdf_path}: {exc}")
                failed += 1
            jobs.update_job(job_id, done=index, total=len(pdf_paths))

        parts = []
        if saved:
            parts.append(f"{saved} cheque(s) guardado(s).")
        if unreadable_number:
            parts.append(f"{unreadable_number} sin N° legible — completalo a mano en el cuadro.")
        if duplicates:
            parts.append(f"{duplicates} ya estaban cargados y se omitieron.")
        if no_check:
            parts.append(f"{no_check} archivo(s) sin ningún cheque adentro.")
        if failed:
            parts.append(f"{failed} archivo(s) no se pudieron leer.")
        if saved:
            level = "warning" if (unreadable_number or no_check or failed) else "success"
        else:
            level = "warning" if duplicates and not (no_check or failed) else "error"
        jobs.update_job(
            job_id, status="done", done=len(pdf_paths), total=len(pdf_paths),
            notice=" ".join(parts) or "No se encontró ningún cheque.", notice_level=level,
            redirect_url="/carga-datos/chase/cheques",
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


@app.route("/carga-datos/chase/cheques/<int:check_id>/pdf")
def chase_cheque_pdf(check_id):
    """El cheque recortado y derecho; ?original=1 sirve el PDF completo subido, ?descargar=1 lo baja."""
    check = cheques_db.get_check(check_id)
    original = request.args.get("original") == "1"
    relpath = (check or {}).get("original_pdf" if original else "check_pdf")
    path = cheques_db.absolute_path(relpath) if relpath else None
    if not path or not os.path.isfile(path):
        flash("No se encontró el archivo guardado de ese cheque.", "error")
        return redirect(url_for("chase_cheques"))
    if original:
        name = check["source_filename"] or "original.pdf"
    elif check["check_number"]:
        name = f"Cheque {check['check_number']}.pdf"
    else:
        name = f"Cheque sin numero {check_id}.pdf"
    return send_file(path, mimetype="application/pdf",
                     as_attachment=request.args.get("descargar") == "1", download_name=name)


def _parse_money_field(raw):
    raw = (raw or "").replace("$", "").replace(",", "").strip()
    if not raw:
        return None
    return round(float(raw), 2)


@app.route("/carga-datos/chase/cheques/guardar", methods=["POST"])
def chase_cheque_guardar():
    """Alta o corrección a mano de una fila del cuadro (un campo vacío vuelve a lo automático)."""
    form = request.form
    check_id = form.get("check_id", type=int)
    raw_number = (form.get("check_number") or "").strip()
    try:
        if raw_number and not raw_number.isdigit():
            raise ValueError("El N° de cheque tiene que ser un número.")
        number = int(raw_number) if raw_number else None
        if check_id is None and number is None:
            raise ValueError("Poné el N° de cheque.")
        try:
            amount = _parse_money_field(form.get("manual_amount"))
        except ValueError:
            raise ValueError("El monto no es un número válido.")
        fields = {
            "manual_date": form.get("manual_date") or None,
            "manual_amount": amount,
            "manual_supplier": (form.get("manual_supplier") or "").strip() or None,
            "manual_invoice": (form.get("manual_invoice") or "").strip() or None,
            "manual_debit_date": form.get("manual_debit_date") or None,
            "voided": 1 if form.get("voided") else 0,
        }
        cheques_db.save_manual(check_id, number, fields)
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("chase_cheques") + (f"#cheque-{raw_number}" if raw_number else ""))


@app.route("/carga-datos/chase/cheques/<int:check_id>/eliminar", methods=["POST"])
def chase_cheque_eliminar(check_id):
    if not cheques_db.delete_check(check_id):
        flash("Ese cheque ya no existe.", "error")
    return redirect(url_for("chase_cheques"))


@app.route("/cmv")
def cmv():
    return render_template("cmv.html", **THEME_BY_KEY["cmv"])


@app.route("/cmv/costo", methods=["POST"])
def cmv_costo():
    master_upload = request.files.get("master_file")
    dept_uploads = request.files.getlist("dept_files")
    if master_upload is None or not master_upload.filename:
        return _error_response("Seleccioná el Excel maestro CMV.")
    if not dept_uploads or not any(u.filename for u in dept_uploads):
        return _error_response("Seleccioná uno o más archivos de departamento.")

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        dept_paths = _save_uploads_to_workspace(dept_uploads, workdir=workdir)
        temp_xlsx_path, _file_stats, _total_parsed, rows_updated, _upcs, _count, failed_files = (
            update_master_costo_todos_bulk(master_path, dept_paths)
        )
    except Exception as exc:
        return _error_response(f"Error: {exc}")

    notice = f"{failed_files} archivo(s) de departamento no se pudieron leer." if failed_files else None
    return _success_response(temp_xlsx_path, master_filename, notice=notice)


@app.route("/cmv/ventas", methods=["POST"])
def cmv_ventas():
    master_upload = request.files.get("master_file")
    sales_uploads = request.files.getlist("sales_files")
    if master_upload is None or not master_upload.filename:
        return _error_response("Seleccioná el Excel maestro CMV.")
    if not sales_uploads or not any(u.filename for u in sales_uploads):
        return _error_response("Seleccioná uno o más reportes de ventas del POS.")

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        sales_paths = _save_uploads_to_workspace(sales_uploads, workdir=workdir)
        _combined, temp_master_path, summary = process_monthly_sales(sales_paths, master_path)
    except Exception as exc:
        return _error_response(f"Error: {exc}")

    # Un archivo o departamento problemático no aborta la carga (ver
    # monthly_sales.py) -- el resto se guarda igual, y el aviso se muestra
    # al toque junto con la descarga (nunca con flash(): la respuesta es
    # una descarga de archivo, no una página, así que quedaría en cola y
    # aparecería fuera de contexto -- mismo bug ya documentado para
    # Proveedores/Caja).
    notice_parts = []
    if summary["failed_files"]:
        notice_parts.append(f"{len(summary['failed_files'])} archivo(s) de ventas no se pudieron leer.")
    if summary["unmapped_departments"]:
        notice_parts.append(
            f"{len(summary['unmapped_departments'])} departamento(s) no se pudieron ubicar en el maestro."
        )
    if summary["sheets_failed"]:
        notice_parts.append(
            f"{len(summary['sheets_failed'])} hoja(s) de departamento no se pudieron actualizar."
        )
    if summary.get("resumen_failed"):
        notice_parts.append(
            "No se pudo actualizar la hoja RESUMEN (los departamentos sí se guardaron bien)."
        )

    return _success_response(
        temp_master_path,
        master_filename,
        notice=" ".join(notice_parts) or None,
        notice_level="error",
    )


@app.route("/gettel")
def gettel():
    return render_template("gettel.html", **THEME_BY_KEY["gettel"])


@app.route("/gettel/cupones", methods=["POST"])
def gettel_cupones():
    source_upload = request.files.get("source_file")
    master_upload = request.files.get("master_file")
    if source_upload is None or not source_upload.filename:
        return _error_response("Seleccioná el Excel o PDF/Foto de origen (cupones diarios).")
    if master_upload is None or not master_upload.filename:
        return _error_response("Seleccioná el Excel de destino (master Cierre).")

    try:
        workdir = _new_workspace_dir()
        source_path, _source_filename = _save_upload_to_workspace(source_upload, workdir=workdir)
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)

        is_pdf = os.path.splitext(source_path)[1].lower() == ".pdf"
        notice_parts = []
        if is_pdf:
            preview_path, rows_matched, vendor, days_found, diagnostics = (
                merge_gettel_toyota_pdf_into_master(source_path, master_path)
            )
            unmatched = diagnostics.get("unmatched_days") or []
            if unmatched:
                notice_parts.append(f"{len(unmatched)} día(s) del PDF no matchearon ninguna fila en el destino.")
            if diagnostics.get("printed_subtotal_found") and not (
                diagnostics.get("amount_matches_subtotal") and diagnostics.get("gallons_matches_subtotal")
            ):
                notice_parts.append(
                    "El total impreso en el PDF no coincide con lo leído — revise el OCR."
                )
        else:
            preview_path, rows_matched, gettel_days, toyota_days, unmatched = (
                merge_gettel_toyota_into_master(source_path, master_path)
            )
            if unmatched:
                notice_parts.append(f"{len(unmatched)} día(s) del origen no matchearon ninguna fila en el destino.")
    except Exception as exc:
        return _error_response(f"Error: {exc}")

    return _success_response(preview_path, master_filename, notice=" ".join(notice_parts) or None)


@app.route("/gettel/pagos", methods=["POST"])
def gettel_pagos():
    master_upload = request.files.get("master_file")
    pdf_uploads = request.files.getlist("pdf_files")
    if master_upload is None or not master_upload.filename:
        return _error_response("Seleccioná el Excel de destino (master Cierre).")
    if not pdf_uploads or not any(u.filename for u in pdf_uploads):
        return _error_response("Seleccioná uno o más PDF de pagos.")

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        pdf_paths = _save_uploads_to_workspace(pdf_uploads, workdir=workdir)
        preview_path, summary = process_gettel_pagos(master_path, pdf_paths)
    except Exception as exc:
        return _error_response(f"Error: {exc}")

    notice_parts = []
    if summary.get("files_failed_to_parse"):
        notice_parts.append(f"{summary['files_failed_to_parse']} archivo(s) no se pudieron leer.")
    if summary.get("batches_failed_to_write"):
        notice_parts.append(f"{summary['batches_failed_to_write']} pago(s) no se pudieron escribir.")
    receipt_warnings = [r for r in summary.get("batch_results", []) if r.get("warning")]
    if receipt_warnings:
        notice_parts.append(
            f"{len(receipt_warnings)} pago(s) con algo para revisar (OCR ilegible en algún campo)."
        )
    return _success_response(preview_path, master_filename, notice=" ".join(notice_parts) or None)


@app.route("/reporte")
def reporte():
    return render_template("reporte.html", **THEME_BY_KEY["reporte"])


def _reporte_pdf_upload():
    """Shared validation + upload-saving for the two Reporte Diario forms."""
    master_upload = request.files.get("master_file")
    pdf_uploads = request.files.getlist("pdf_files")
    if not pdf_uploads or not any(u.filename for u in pdf_uploads):
        return None, _error_response("Seleccioná uno o más PDF diarios.")
    if master_upload is None or not master_upload.filename:
        return None, _error_response("Seleccioná el Excel de destino.")

    workdir = _new_workspace_dir()
    master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
    pdf_paths = _save_uploads_to_workspace(pdf_uploads, workdir=workdir)
    return (master_path, master_filename, pdf_paths), None


def _reporte_batch_notice(summary, day_key="calendar_day"):
    """
    Arma un aviso corto (sin nombres de archivo) a partir del summary de
    process_reporte_diario/process_store_info/process_lottery: cuántos
    archivos/días quedaron aislados por un problema, más cualquier
    diagnóstico propio ya calculado por día (BS vs. cargado, subtotal OCR
    vs. impreso, fuente faltante) que antes se calculaba y nunca llegaba a
    mostrarse -- ver "Reporte Diario + Lottery" en la auditoría de
    2026-09-03.
    """
    parts = []
    if summary.get("files_with_bad_filename"):
        parts.append(f"{summary['files_with_bad_filename']} archivo(s) con nombre no reconocido.")
    if summary.get("files_failed_to_parse"):
        parts.append(f"{summary['files_failed_to_parse']} archivo(s) no se pudieron leer.")
    if summary.get("days_failed_to_write"):
        parts.append(f"{summary['days_failed_to_write']} día(s) no matchearon ninguna fila.")
    if summary.get("dates_failed_to_write"):
        parts.append(f"{summary['dates_failed_to_write']} fecha(s) no matchearon ninguna fila.")

    day_warnings = [
        (result.get(day_key) or result.get("report_date"), result["warning"])
        for result in summary.get("batch_results", [])
        if result.get("warning")
    ]
    if day_warnings:
        detail = "; ".join(f"día {day}: {warning}" for day, warning in day_warnings)
        parts.append(f"{len(day_warnings)} día(s) con algo para revisar — {detail}")

    return " ".join(parts) or None


def _parse_paths_concurrently(pdf_paths, parse_fn, progress_callback=None):
    """
    Corre parse_fn(pdf_path) para cada PDF, en paralelo cuando hay más de
    uno -- mismo patrón (y mismo tope de workers) que reporte_diario.py:
    _parse_pdfs_concurrently, reusado acá para los guardados-espejo de
    Reporte Diario.

    Bug real encontrado reproduciendo el reporte del usuario ("se tardó
    mucho y después no cargó nada", 2026-09-14): los guardados-espejo
    releían cada PDF SECUENCIAL, uno por uno -- con un lote de 5 PDFs
    reales esto agregó ~8 minutos de trabajo después de que la lectura
    principal (que sí corre en paralelo) ya había terminado. Acá se lee en
    paralelo igual que la lectura principal -- la escritura en la base
    (rápida) sigue siendo secuencial después, sobre los resultados ya
    parseados.

    Devuelve {pdf_path: (parsed_value, error)} -- exactamente uno de los
    dos es None. `progress_callback(pdf_path)`, si viene, se llama apenas
    termina de parsearse cada PDF (éxito o error).
    """
    def _parse_one(pdf_path):
        try:
            return pdf_path, parse_fn(pdf_path), None
        except Exception as exc:
            return pdf_path, None, exc
        finally:
            if progress_callback:
                progress_callback(pdf_path)

    results = {}
    if len(pdf_paths) <= 1:
        for pdf_path in pdf_paths:
            path, value, error = _parse_one(pdf_path)
            results[path] = (value, error)
    else:
        max_workers = min(len(pdf_paths), os.cpu_count() or 4, 6)
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(_parse_one, pdf_path) for pdf_path in pdf_paths]
            for future in as_completed(futures):
                path, value, error = future.result()
                results[path] = (value, error)
    return results


def _persist_reporte_diario_departments(pdf_paths, progress_callback=None):
    """
    Guarda en reportes_db (además de escribir el Excel de siempre, que ya
    pasó antes de llamar a esto) los departamentos de cada PDF del lote --
    "memoria" nueva de la página, ver CLAUDE.md. Nunca puede romper la
    respuesta principal: el Excel ya se generó bien antes de llegar acá,
    así que cualquier problema en este paso se ignora en silencio (queda
    solo en la consola del servidor) en vez de tapar la descarga real.

    `progress_callback(pdf_path)`, si viene, se llama después de leer cada
    PDF (ver _parse_paths_concurrently) -- para que un lote grande siga
    mostrando avance real en vez de parecer trabado.
    """
    outcomes = _parse_paths_concurrently(pdf_paths, extract_department_sales_for_day, progress_callback)
    for pdf_path, (result, error) in outcomes.items():
        if error is not None:
            print(f"[reportes_db] no se pudo guardar departamentos de {pdf_path}: {error}")
            continue
        try:
            pdf_relpath = reportes_db.store_pdf_copy(
                result["date"], pdf_path, _reporte_pdf_canonical_filename(result["date"])
            )
            reportes_db.replace_department_sales(
                result["date"], result["records"], pdf_filename=pdf_relpath,
            )
        except Exception as exc:
            print(f"[reportes_db] no se pudo guardar departamentos de {pdf_path}: {exc}")


def _persist_lottery_department_fields(pdf_paths, progress_callback=None):
    """
    Primer paso del "interconectado" entre módulos de Carga de Datos (pedido
    explícito del usuario 2026-09-11): el PDF de cierre diario de Reporte
    Diario ya trae, en su Department Sales Report, las mismas filas
    ONLINE/SKOFF que Lottery necesita para D/E/N/O -- exactamente lo mismo
    que ya hacía el Excel (un solo PDF alimentando dos libros distintos).
    Se llama junto con _persist_reporte_diario_departments (tanto desde el
    lado Excels como desde Carga de Datos) para que subir un PDF de Reporte
    Diario UNA sola vez ya deje esa parte de Lottery lista, sin otro upload
    aparte -- solo falta el "Daily Sales Report" del portal de Lottery
    (F/G/H/I/K/P/Q/R/S), que es un documento realmente distinto. Aislado por
    PDF y nunca puede romper la respuesta principal, mismo criterio que el
    resto de estos guardados-espejo. `progress_callback` -- ver
    _persist_reporte_diario_departments.
    """
    outcomes = _parse_paths_concurrently(pdf_paths, extract_lottery_department_fields_from_pdf, progress_callback)
    for pdf_path, (fields, error) in outcomes.items():
        if error is not None:
            print(f"[lottery_db] no se pudo guardar ONLINE/SKOFF de {pdf_path}: {error}")
            continue
        try:
            filename = os.path.basename(pdf_path)
            pdf_relpath = lottery_db.store_pdf_copy(fields["report_date"], pdf_path, filename)
            lottery_db.upsert_department_fields(
                fields["report_date"], fields["online_count"], fields["online_net_sales"],
                fields["skoff_count"], fields["skoff_net_sales"], source="ocr", pdf_filename=pdf_relpath,
            )
        except Exception as exc:
            print(f"[lottery_db] no se pudo guardar ONLINE/SKOFF de {pdf_path}: {exc}")


def _persist_lottery_sales_report_fields(pdf_paths):
    """
    Análogo a _persist_lottery_department_fields, para el otro documento que
    alimenta Lottery -- el "Daily Sales Report" del portal de Florida
    Lottery (F/G/H/I/K/P/Q/R/S). Se llama desde el flujo de Excels
    (/lottery/sales-report) además de Carga de Datos, para que subir ese
    PDF por CUALQUIERA de los dos lados deje lottery_db al día -- mismo
    criterio de "un solo upload, las dos memorias" que ya rige Reporte
    Diario. Aislado por PDF, nunca puede romper la respuesta principal.
    """
    for pdf_path in pdf_paths:
        try:
            fields = extract_lottery_receipt_fields_from_sales_report(pdf_path)
            filename = os.path.basename(pdf_path)
            pdf_relpath = lottery_db.store_pdf_copy(fields["report_date"], pdf_path, filename)
            lottery_db.upsert_sales_report_fields(fields["report_date"], fields, source="ocr", pdf_filename=pdf_relpath)
        except Exception as exc:
            print(f"[lottery_db] no se pudo guardar Daily Sales Report de {pdf_path}: {exc}")


def _persist_reporte_diario_store_info(pdf_paths, progress_callback=None):
    """Análogo a _persist_reporte_diario_departments, para Store Info. `progress_callback` -- ver ese docstring."""
    outcomes = _parse_paths_concurrently(pdf_paths, extract_store_info_for_day, progress_callback)
    for pdf_path, (result, error) in outcomes.items():
        if error is not None:
            print(f"[reportes_db] no se pudo guardar Store Info de {pdf_path}: {error}")
            continue
        try:
            pdf_relpath = reportes_db.store_pdf_copy(
                result["date"], pdf_path, _reporte_pdf_canonical_filename(result["date"])
            )
            reportes_db.upsert_store_info(
                result["date"], result["fields"], source="ocr", pdf_filename=pdf_relpath, keep_existing_for_none=True,
            )
        except Exception as exc:
            print(f"[reportes_db] no se pudo guardar Store Info de {pdf_path}: {exc}")


@app.route("/reporte/ventas", methods=["POST"])
def reporte_ventas():
    saved, error = _reporte_pdf_upload()
    if error is not None:
        return error
    master_path, master_filename, pdf_paths = saved

    try:
        temp_path, summary = process_reporte_diario(master_path, pdf_paths)
    except Exception as exc:
        return _error_response(f"Error: {exc}")

    _persist_reporte_diario_departments(pdf_paths)
    _persist_lottery_department_fields(pdf_paths)

    return _success_response(temp_path, master_filename, notice=_reporte_batch_notice(summary))


@app.route("/reporte/store-info", methods=["POST"])
def reporte_store_info():
    saved, error = _reporte_pdf_upload()
    if error is not None:
        return error
    master_path, master_filename, pdf_paths = saved

    try:
        temp_path, summary = process_store_info(master_path, pdf_paths)
    except Exception as exc:
        return _error_response(f"Error: {exc}")

    _persist_reporte_diario_store_info(pdf_paths)

    return _success_response(temp_path, master_filename, notice=_reporte_batch_notice(summary))


# ---------------------------------------------------------------------------
# Carga en segundo plano -- pedido explícito del usuario (2026-09-14): "la
# pagina tarda mucho en cargar los reportes diarios varios a la vez... ver
# si ese proceso podria hacerse por aparte mientras se trabaja en otra
# cosa" + "queria que muestre el progreso real... que vaya moviendo el
# porcentaje". Reusa las mismas process_reporte_diario/process_store_info
# de siempre (con un progress_callback opcional nuevo, ver reporte_diario.py)
# -- nada de la lógica de negocio se duplica, esto solo mueve el trabajo
# pesado a un hilo aparte y expone su avance real por polling (ver jobs.py).
# Rutas nuevas y separadas de /reporte/ventas y /reporte/store-info de
# arriba -- esas dos siguen intactas, sin ningún riesgo de regresión.
# ---------------------------------------------------------------------------

def _run_reporte_ventas_job(job_id, master_path, pdf_paths, master_filename):
    """
    Corre en su propio hilo (ver reporte_ventas_async) -- TODO el cuerpo va
    envuelto en un try/except, no solo la llamada a process_reporte_diario.
    Antes, una excepción en cualquiera de los pasos de después (guardado-
    espejo, abrir el resultado, armar el aviso) quedaba sin atrapar -- un
    hilo de Python que revienta así no tira abajo el servidor, pero tampoco
    avisa a nadie: el job se quedaba pegado en "running" para siempre y el
    navegador seguía sondeando sin que nada cambie nunca (bug real,
    encontrado reproduciendo el reporte del usuario de "se tardó mucho y
    después no cargó nada" -- ver CLAUDE.md).
    """
    def on_progress(done, total):
        jobs.update_job(job_id, done=done, total=total)

    try:
        temp_path, summary = process_reporte_diario(master_path, pdf_paths, progress_callback=on_progress)

        # A partir de acá no hay más progreso por-PDF de la lectura -- el
        # Excel ya se escribió y falta el guardado-espejo (que vuelve a leer
        # cada PDF, más lento que la lectura de arriba porque no corre en
        # paralelo) -- "phase" le avisa al navegador que siga esperando en
        # vez de parecer trabado en 100%, y "done"/"total" se reusan para
        # el progreso REAL de este segundo paso (2 guardados por PDF acá:
        # departamentos + Lottery), no una cuenta regresiva inventada.
        save_total = len(pdf_paths) * 2
        save_done = [0]

        def on_save_progress(_pdf_path):
            save_done[0] += 1
            jobs.update_job(job_id, done=save_done[0], total=save_total)

        jobs.update_job(job_id, phase="saving", done=0, total=save_total)
        _persist_reporte_diario_departments(pdf_paths, progress_callback=on_save_progress)
        _persist_lottery_department_fields(pdf_paths, progress_callback=on_save_progress)
        _open_result_for_user(temp_path)
        jobs.update_job(
            job_id, status="done", done=save_total, total=save_total,
            result_path=temp_path, result_filename=master_filename,
            notice=_reporte_batch_notice(summary),
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


def _run_reporte_store_info_job(job_id, master_path, pdf_paths, master_filename):
    """Análoga a _run_reporte_ventas_job -- ver ese docstring."""
    def on_progress(done, total):
        jobs.update_job(job_id, done=done, total=total)

    try:
        temp_path, summary = process_store_info(master_path, pdf_paths, progress_callback=on_progress)

        save_total = len(pdf_paths)
        save_done = [0]

        def on_save_progress(_pdf_path):
            save_done[0] += 1
            jobs.update_job(job_id, done=save_done[0], total=save_total)

        jobs.update_job(job_id, phase="saving", done=0, total=save_total)
        _persist_reporte_diario_store_info(pdf_paths, progress_callback=on_save_progress)
        _open_result_for_user(temp_path)
        jobs.update_job(
            job_id, status="done", done=save_total, total=save_total,
            result_path=temp_path, result_filename=master_filename,
            notice=_reporte_batch_notice(summary),
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


@app.route("/reporte/ventas/async", methods=["POST"])
def reporte_ventas_async():
    saved, error = _reporte_pdf_upload()
    if error is not None:
        return error
    master_path, master_filename, pdf_paths = saved

    job_id = jobs.create_job(len(pdf_paths))
    threading.Thread(
        target=_run_reporte_ventas_job, args=(job_id, master_path, pdf_paths, master_filename), daemon=True
    ).start()
    return jsonify({"job_id": job_id, "total": len(pdf_paths)})


@app.route("/reporte/store-info/async", methods=["POST"])
def reporte_store_info_async():
    saved, error = _reporte_pdf_upload()
    if error is not None:
        return error
    master_path, master_filename, pdf_paths = saved

    job_id = jobs.create_job(len(pdf_paths))
    threading.Thread(
        target=_run_reporte_store_info_job, args=(job_id, master_path, pdf_paths, master_filename), daemon=True
    ).start()
    return jsonify({"job_id": job_id, "total": len(pdf_paths)})


@app.route("/jobs/<job_id>/status")
def job_status(job_id):
    job = jobs.get_job(job_id)
    if job is None:
        # 200, no 404 -- un job que no existe más (ej. el servidor se
        # reinició a mitad de una carga) es un resultado válido a mostrarle
        # al usuario, no una falla de red -- si fuera 404 el polling de
        # base.html lo trataría como error de conexión y reintentaría para
        # siempre en vez de avisar y frenar.
        return jsonify({"status": "not_found"})
    return jsonify(
        {
            "status": job["status"],
            "phase": job.get("phase", "parsing"),
            "done": job["done"],
            "total": job["total"],
            "notice": job.get("notice"),
            "notice_level": job.get("notice_level"),
            "error": job.get("error"),
            # `redirect_url` -- pedido explícito del usuario (2026-09-16,
            # segunda tanda): las cargas de Carga de Datos (a diferencia de
            # las de Herramientas, que se quedan en la misma página) tienen
            # que terminar en la página de "ya guardado" correspondiente --
            # solo lo llevan los jobs que lo necesitan (ver jobs.update_job
            # en cada _run_*_job de Carga de Datos), None para el resto.
            "redirect_url": job.get("redirect_url"),
        }
    )


@app.route("/jobs/<job_id>/cancel", methods=["POST"])
def job_cancel(job_id):
    """Botón "Cancelar carga" de la barra de progreso (ver jobs.cancel_job)."""
    job = jobs.cancel_job(job_id)
    if job is None:
        return jsonify({"status": "not_found"})
    return jsonify({"status": job["status"], "done": job["done"], "total": job["total"]})


@app.route("/jobs/<job_id>/ack", methods=["POST"])
def job_ack(job_id):
    """
    Pedido explícito del usuario (2026-09-21, ver el docstring de
    jobs.get_active_job): el JS llama esto apenas terminó de mostrarle al
    usuario el resultado final (notice o redirect) de un job -- para que
    ese job deje de aparecer como "activo" la próxima vez que se entra a
    esa página de carga. No hace falta esperar la respuesta ni reintentar
    si falla (el peor caso es que el aviso se repita una vez más).
    """
    jobs.acknowledge_job(job_id)
    return ("", 204)


@app.route("/jobs/<job_id>/download")
def job_download(job_id):
    job = jobs.get_job(job_id)
    if job is None or job.get("status") != "done" or not job.get("result_path"):
        flash("Ese resultado ya no está disponible.", "error")
        return redirect(url_for("reporte"))
    return send_file(job["result_path"], as_attachment=True, download_name=job["result_filename"])


_MONTH_NAMES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

# Mismos nombres "amigables" que ya usa reporte_historial.html para las
# columnas fijas del mes (Tabacco/SODA/BEER-WINE/LOTTO/KIA-TOY/Resto) --
# acá se reusan para el cuadro flotante de "Non Fuel" de Store Info (ver
# reporte_store_info_historial más abajo), que muestra el mismo desglose
# por categoría como respaldo del monto.
_CATEGORY_DISPLAY_LABELS = {
    "TABACCO": "Tabacco",
    "SODA": "SODA",
    "BEER/WINE": "BEER/WINE",
    "LOTERY/LOTTO": "LOTTO",
    "Gettel": "KIA/TOY",
    "RESTO": "Resto",
}


@app.route("/reporte/historial")
def reporte_historial():
    """
    Ventas por Departamento de un mes completo -- solo esa parte, ver
    CLAUDE.md. Store Info vive en su propia página (reporte_store_info_
    historial) -- pedido explícito del usuario (2026-09-12): separarlas en
    la barra lateral "como si fueran los Excels de Ventas y de Cierre",
    que en la vida real siempre fueron dos archivos distintos.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    overview = reportes_db.get_month_overview(year, month)
    # Vista previa por día -- pedido explícito del usuario (2026-09-12,
    # cuarta tanda): el espacio de la fila alcanzaba para más que la lista
    # de departamentos crudos ("datos inecesarios") -- ahora muestra las
    # mismas 6 categorías (TABACCO/SODA/...) del resumen del mes, pero
    # calculadas para ESE día puntual, para poder comparar días sin entrar
    # a "Ver/editar". Corrección 2026-09-16: pasaron de pastillas (solo las
    # categorías con algo cargado) a 6 columnas fijas de la propia tabla --
    # las 6 SIEMPRE se devuelven, aunque den $0, para que la columna
    # correspondiente muestre "0" en vez de desaparecer.
    for day in overview:
        day_groups, _unmatched = group_department_sales(day["department_detail"])
        day["category_groups"] = day_groups
    department_totals = reportes_db.get_month_department_totals(year, month)
    department_groups, department_unmatched = group_department_sales(department_totals)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    # Reporte mensual del POS, si está cargado: columnas para comparar.
    monthly = reporte_mensual_db.get_report(year, month)

    return render_template(
        "reporte_historial.html",
        overview=overview,
        department_totals=department_totals,
        department_groups=department_groups,
        department_unmatched=department_unmatched,
        monthly_categories=reporte_mensual.category_comparison(department_groups, monthly) if monthly else None,
        monthly_departments=reporte_mensual.department_comparison(department_totals, monthly) if monthly else None,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        today_iso=today.isoformat(),
        **THEME_BY_KEY["reporte"],
    )


@app.route("/reporte/historial/eliminar", methods=["POST"])
def reporte_historial_eliminar():
    """
    Elimina Ventas por Departamento de uno o más días seleccionados con
    checkbox -- pedido explícito del usuario (2026-09-15), para poder
    borrar días que quedaron mal (ej. con el nombre viejo "GETTEL/TOYOTA")
    y volver a cargarlos limpio. Nunca toca Store Info de esos días -- son
    dos extracciones independientes, mismo criterio que el resto del
    módulo.
    """
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    selected_dates = request.form.getlist("dates")
    deleted = 0
    for date_str in selected_dates:
        rows = reportes_db.delete_departments_for_date(date_str)
        if rows:
            deleted += 1
    if deleted:
        flash(f"{deleted} día(s) de Ventas por Departamento eliminados.", "success")
    else:
        flash("No se seleccionó ningún día con departamentos cargados.", "error")
    return redirect(url_for("reporte_historial", year=year, month=month))


def _build_store_info_rows(year, month):
    """
    Store Info de un mes, con los campos calculados (Total Fuel/Total
    Sales/Total Revenue/gettel_amount y sus desgloses) ya resueltos --
    extraído de reporte_store_info_historial para poder reusarlo también
    en la exportación a Excel (ver reporte_store_info_exportar), sin
    duplicar esta lógica.
    """
    store_info_rows = reportes_db.get_month_store_info(year, month)
    # Aviso de horario -- pedido explícito del usuario (2026-09-14): si un
    # día termina a una hora y el siguiente no arranca exactamente ahí, es
    # señal de que algo quedó mal cargado (un día sin reporte en el medio,
    # un horario mal leído por OCR, etc.) -- se marcan las dos horas
    # (la de "hasta" del día que corta y la de "desde" del que sigue) para
    # que salten a la vista. Un día sin datos (from_time/to_time en None)
    # no tiene nada que comparar, se saltea sin marcar nada.
    for cur, nxt in zip(store_info_rows, store_info_rows[1:]):
        if cur.get("to_time") and nxt.get("from_time") and cur["to_time"] != nxt["from_time"]:
            cur["time_warn_to"] = True
            nxt["time_warn_from"] = True

    # "Total Fuel" y el "Total Sales" real -- pedido explícito del usuario
    # (2026-09-16). Las columnas de categoría (Tabacco/Soda/Beer-Wine/Lotto/
    # VS/Resto) que se mostraban acá se sacaron el mismo día ("eso ya
    # igual se puede ver el día a día en las ventas por departamento") --
    # pero el monto de la categoría "Gettel" (columna VS del Excel real)
    # sigue haciendo falta puertas adentro, sin mostrarse como columna,
    # para la fórmula de Total Sales de abajo.
    department_detail_by_date = {d["date"]: d["department_detail"] for d in reportes_db.get_month_overview(year, month)}
    for row in store_info_rows:
        detail = department_detail_by_date.get(row["date"], [])
        groups, _unmatched = group_department_sales(detail)
        gettel_amount = next((g["amount"] for g in groups if g["label"] == "Gettel"), 0.0)
        row["gettel_amount"] = gettel_amount
        # Sin departamentos ese día, LOTTO y VS quedan en 0 (el control Cierre lo avisa).
        row["has_departments"] = bool(detail)
        # Categorías crudas (TABACCO/SODA/BEER-WINE/LOTERY-LOTTO/Gettel/
        # RESTO) -- hace falta puertas adentro para la exportación a Excel
        # (columnas I-N de la hoja real, ver build_store_info_export_workbook),
        # nunca se muestra como columna en esta página.
        row["category_amounts"] = {g["label"]: g["amount"] for g in groups}

        # "Total Fuel" -- pedido explícito del usuario (2026-09-16): la
        # columna real de Store Info (H) nunca se había mostrado -- es
        # Sales Fuel (bruto) + Desc. Comb (el descuento, ya guardado en
        # negativo), el neto de combustible después del descuento. Se
        # calcula acá, nunca se guarda -- mismo criterio que el resto de
        # esta página, que solo muestra lo que ya está en la base.
        if row.get("sales_fuel") is not None and row.get("desc_comb") is not None:
            row["total_fuel"] = round(row["sales_fuel"] + row["desc_comb"], 2)
        else:
            row["total_fuel"] = None

        # "Total Sales" real -- bug real encontrado por el usuario
        # (2026-09-16): el valor guardado se leía tal cual lo imprime el
        # propio PDF ("Total Sales $X"), sin restar VS -- eso significa que
        # nunca reaccionaba a un cambio en la categoría "Gettel" de Ventas
        # por Departamento, aunque el usuario borrara y volviera a cargar
        # el día. La fórmula real de Store Info!R es Total Fuel + Non Fuel
        # + Desc Otros + Tax Collect - VS (ver reporte_diario._extract_
        # store_info_fields) -- no se implementaba antes porque VS no se
        # podía calcular; ahora sí (es la categoría "Gettel" de arriba), así
        # que se recalcula acá para mostrar el valor REAL en vez del
        # impreso -- reacciona solo a cualquier cambio en Ventas por
        # Departamento de ese día, sin tener que volver a cargar Store Info.
        # El valor impreso en el PDF (lo que de verdad valida el OCR al
        # extraer, y lo que se puede corregir a mano) sigue guardado tal
        # cual en la base -- ver reporte_dia_store_info, no se tocó.
        real_total_sales, _gettel = real_store_info_total_sales(row, detail)
        if real_total_sales is not None:
            row["total_sales"] = real_total_sales

        # Desglose para los cuadros flotantes de "qué valores usaron para
        # llegar a ese resultado" (Total Fuel/Total Sales/Total Rev) --
        # pedido explícito del usuario (2026-09-16).
        row["total_fuel_breakdown"] = [
            ("Sales Fuel", row.get("sales_fuel")),
            ("Desc. Comb", row.get("desc_comb")),
        ]
        row["total_sales_breakdown"] = [
            ("Total Fuel", row.get("total_fuel")),
            ("Non Fuel", row.get("non_fuel_total")),
            ("Desc. Otros", row.get("desc_otros")),
            ("Tax Collect", row.get("tax_collect")),
            ("VS (Gettel, resta)", -gettel_amount if gettel_amount else 0.0),
        ]
        row["total_revenue_breakdown"] = [
            ("Cash", row.get("cash")),
            ("Tarjeta/Crédito", round(sum(row["credit_terms"]), 2) if row.get("credit_terms") else 0.0),
            ("Other", row.get("other_amount")),
            ("Local Acc.", row.get("local_accounts")),
        ]

        # "Non Fuel" -- pedido explícito del usuario (2026-09-16): a
        # diferencia de Total Fuel/Total Sales/Total Revenue (fórmulas
        # calculadas a partir de otros campos de Store Info), Non Fuel es un
        # valor crudo leído del PDF -- lo que respalda ese monto son las
        # ventas reales del día en Ventas por Departamento (todo lo que no
        # es combustible). Se muestran las mismas 6 categorías, con los
        # mismos nombres, que ya usa esa página -- más un total para
        # comparar a simple vista contra el Non Fuel de esta fila.
        row["non_fuel_breakdown"] = [
            (_CATEGORY_DISPLAY_LABELS.get(g["label"], g["label"]), g["amount"]) for g in groups
        ]
        row["non_fuel_categories_total"] = round(sum(g["amount"] for g in groups), 2)
    return store_info_rows


def _store_info_totals(rows):
    """
    Fila de totales del mes para /reporte/store-info/historial -- pedido
    explícito del usuario (2026-09-19): "no tenemos una fila extra despues
    del ultimo dia en el que muestra los totales sumados". Un campo con
    NINGÚN día cargado queda en None ("—" en el template) en vez de 0, para
    no dar a entender que se sabe que ese total es cero.
    """
    fields = (
        "volume", "sales_fuel", "desc_comb", "total_fuel", "non_fuel_total",
        "desc_otros", "tax_collect", "total_sales", "cash", "local_accounts",
        "other_amount", "network_revenue", "total_revenue",
    )
    totals = {field: 0.0 for field in fields}
    any_value = {field: False for field in fields}
    tc_total, tc_any = 0.0, False
    for row in rows:
        for field in fields:
            value = row.get(field)
            if value is not None:
                totals[field] += value
                any_value[field] = True
        credit_terms = row.get("credit_terms") or []
        if credit_terms:
            tc_total += sum(credit_terms)
            tc_any = True
    result = {field: (round(totals[field], 2) if any_value[field] else None) for field in fields}
    result["tc"] = round(tc_total, 2) if tc_any else None
    return result


@app.route("/reporte/store-info/historial")
def reporte_store_info_historial():
    """Store Info de un mes completo -- página propia, ver reporte_historial de arriba."""
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    store_info_rows = _build_store_info_rows(year, month)
    store_info_totals = _store_info_totals(store_info_rows)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    # Reporte mensual del POS, si está cargado: filas para comparar con el total.
    monthly = reporte_mensual_db.get_report(year, month)
    days_loaded = sum(1 for r in store_info_rows if r.get("store_info_source"))

    return render_template(
        "reporte_store_info_historial.html",
        store_info_rows=store_info_rows,
        store_info_totals=store_info_totals,
        monthly_store_info=(
            reporte_mensual.store_info_comparison(store_info_totals, monthly, days_loaded) if monthly else None
        ),
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        today_iso=today.isoformat(),
        **THEME_BY_KEY["reporte"],
    )


@app.route("/reporte/store-info/exportar")
def reporte_store_info_exportar():
    """
    Descarga un Excel NUEVO (nunca toca ningún archivo real) con Store Info
    de un mes ya guardado -- pedido explícito del usuario (2026-09-18), ver
    reporte_diario.build_store_info_export_workbook.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    store_info_rows = _build_store_info_rows(year, month)
    if not any(r.get("store_info_source") for r in store_info_rows):
        flash("No hay ningún Store Info guardado ese mes para exportar.", "error")
        return redirect(url_for("reporte_store_info_historial", year=year, month=month))

    workspace_dir = tempfile.mkdtemp(prefix="storeinfo_export_")
    dest_path = os.path.join(workspace_dir, f"Store Info {month:02d}-{year}.xlsx")
    build_store_info_export_workbook(store_info_rows, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/reporte/store-info/exportar/pdf")
def reporte_store_info_exportar_pdf():
    """
    Versión PDF (básica, sin colores, solo líneas y bordes) del export de
    arriba -- pedido explícito del usuario (2026-09-19): "solo necesitaria
    que salieran los datos limpios en un PDF basico". Ver
    reporte_diario.build_store_info_export_pdf.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    store_info_rows = _build_store_info_rows(year, month)
    if not any(r.get("store_info_source") for r in store_info_rows):
        flash("No hay ningún Store Info guardado ese mes para exportar.", "error")
        return redirect(url_for("reporte_store_info_historial", year=year, month=month))

    workspace_dir = tempfile.mkdtemp(prefix="storeinfo_export_pdf_")
    dest_path = os.path.join(workspace_dir, f"Store Info {month:02d}-{year}.pdf")
    build_store_info_export_pdf(store_info_rows, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


def _parse_report_date(report_date):
    if isinstance(report_date, date):
        return report_date
    return datetime.strptime(report_date, "%Y-%m-%d").date()


@app.route("/documentos/reporte-diario")
def reporte_documentos():
    """
    PDFs de cierre diario ya guardados este mes -- ver CLAUDE.md, barra
    lateral por módulo. Suma un cuadrito aparte para el PDF de resumen
    mensual (pedido explícito del usuario, 2026-09-16: "igual que en la
    Lottery" -- vive en Documentos, un solo PDF por mes, reemplaza la
    página propia que tenía antes). Reusa documents_db.py, módulo
    "reporte_diario_resumen_mensual".
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    pdfs = reportes_db.get_month_pdf_list(year, month)
    for pdf in pdfs:
        pdf["filename"] = os.path.basename(pdf["pdf_filename"]) if pdf["pdf_filename"] else None
    monthly_pdfs = documents_db.list_documents("reporte_diario_resumen_mensual", year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "reporte_documentos.html",
        pdfs=pdfs,
        monthly_pdfs=monthly_pdfs,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        today_iso=today.isoformat(),
        **THEME_BY_KEY["reporte"],
    )


@app.route("/documentos/reporte-diario/mensual/subir", methods=["POST"])
def reporte_documentos_mensual_subir():
    """
    Sube el único PDF de resumen mensual -- solo se guarda, ver arriba. Un
    solo PDF por mes, con guarda del lado del servidor (mismo criterio que
    el equivalente de Lottery, ver carga_datos_lottery_documentos_mensual_subir).
    """
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    upload = request.files.get("pdf_file")
    if not year or not month or not (1 <= month <= 12):
        flash("Elegí a qué mes corresponde este PDF.", "error")
        return redirect(url_for("reporte_documentos"))
    if documents_db.list_documents("reporte_diario_resumen_mensual", year, month):
        flash("Ya hay un PDF mensual cargado este mes -- eliminalo antes de subir otro.", "error")
        return redirect(url_for("reporte_documentos", year=year, month=month))
    if upload is None or not upload.filename:
        flash("Seleccioná el PDF de resumen mensual.", "error")
        return redirect(url_for("reporte_documentos", year=year, month=month))

    path, filename = _save_upload_to_workspace(upload)
    documents_db.store_document("reporte_diario_resumen_mensual", path, filename, year, month)
    flash("PDF mensual guardado.", "success")
    return redirect(url_for("reporte_documentos", year=year, month=month))


@app.route("/reporte/dia/<report_date>")
def reporte_dia(report_date):
    try:
        parsed_date = _parse_report_date(report_date)
    except ValueError:
        flash("Fecha inválida.", "error")
        return redirect(url_for("reporte_historial"))

    day = reportes_db.get_day(parsed_date)
    department_groups, department_unmatched = group_department_sales(day["departments"])

    # "Total Sales" real de la tarjeta resumen -- mismo fix que reporte_
    # store_info_historial (ver ahí el porqué): se recalcula con Total Fuel
    # + Non Fuel + Desc Otros + Tax Collect - VS (categoría "Gettel" de
    # arriba) en vez de mostrar el valor impreso en el PDF tal cual, así
    # reacciona solo a cualquier cambio en los departamentos de este mismo
    # día. `None` si todavía falta algún componente -- la tarjeta cae al
    # valor guardado en ese caso.
    store_info = day["store_info"]
    gettel_amount = next((g["amount"] for g in department_groups if g["label"] == "Gettel"), 0.0)
    computed_total_sales = None
    if store_info and None not in (
        store_info.get("sales_fuel"), store_info.get("desc_comb"),
        store_info.get("non_fuel_total"), store_info.get("desc_otros"), store_info.get("tax_collect"),
    ):
        total_fuel = store_info["sales_fuel"] + store_info["desc_comb"]
        computed_total_sales = round(
            total_fuel + store_info["non_fuel_total"] + store_info["desc_otros"] + store_info["tax_collect"] - gettel_amount, 2
        )

    return render_template(
        "reporte_dia.html",
        report_date=parsed_date,
        departments=day["departments"],
        department_groups=department_groups,
        department_unmatched=department_unmatched,
        store_info=store_info,
        computed_total_sales=computed_total_sales,
        pdf_filename=day["pdf_filename"],
        printed_total_sales=day["printed_total_sales"],
        printed_total_units=day["printed_total_units"],
        known_departments=reportes_db.list_known_departments(),
        **THEME_BY_KEY["reporte"],
    )


@app.route("/reporte/dia/<report_date>/pdf")
def reporte_dia_pdf(report_date):
    day = reportes_db.get_day(report_date)
    path = reportes_db.absolute_pdf_path(day["pdf_filename"]) if day["pdf_filename"] else None
    if not path or not os.path.isfile(path):
        flash("No hay ningún PDF guardado para este día.", "error")
        return redirect(url_for("reporte_dia", report_date=report_date))
    # Vista previa por default (inline), descarga forzada solo con
    # ?mode=download -- pedido explícito del usuario (2026-09-19), ver
    # templates/_pdf_links.html.
    force_download = request.args.get("mode") == "download"
    return send_file(path, as_attachment=force_download, download_name=os.path.basename(path))


@app.route("/reporte/dia/<report_date>/departamentos", methods=["POST"])
def reporte_dia_departamentos(report_date):
    names = request.form.getlist("dept_name")
    counts = request.form.getlist("dept_count")
    amounts = request.form.getlist("dept_amount")
    to_delete = set(request.form.getlist("dept_delete"))

    saved = 0
    skipped = 0
    for name, count_raw, amount_raw in zip(names, counts, amounts):
        name = name.strip()
        if not name:
            continue
        if name in to_delete:
            reportes_db.delete_department_row(report_date, name)
            continue
        count_raw = (count_raw or "").strip()
        amount_raw = (amount_raw or "").strip()
        if not count_raw and not amount_raw:
            continue
        try:
            count = int(float(count_raw)) if count_raw else None
            amount = float(amount_raw) if amount_raw else None
        except ValueError:
            skipped += 1
            continue
        reportes_db.upsert_department_row(report_date, name, count, amount, source="manual")
        saved += 1

    printed_sales_raw = (request.form.get("printed_total_sales") or "").strip()
    printed_units_raw = (request.form.get("printed_total_units") or "").strip()
    try:
        printed_sales = float(printed_sales_raw) if printed_sales_raw else None
        printed_units = float(printed_units_raw) if printed_units_raw else None
        reportes_db.upsert_printed_totals(report_date, printed_sales, printed_units)
    except ValueError:
        skipped += 1

    if skipped:
        flash(f"{saved} departamento(s) guardado(s). {skipped} con un número/monto inválido, no se guardaron.", "warning")
    else:
        flash(f"{saved} departamento(s) guardado(s).", "success")
    return redirect(url_for("reporte_dia", report_date=report_date))


_THOUSANDS_AMOUNT_RE = re.compile(r"-?\d{1,3}(,\d{3})+(\.\d+)?")


def _parse_amount_list(raw):
    """
    Lista de montos tipeada a mano ("500.00, 300.00"). Los montos se separan
    con coma y espacio, punto y coma o salto de línea; una coma de miles
    ("2,001.68", el mismo formato que muestra la app) es parte del monto y
    no un separador: antes se guardaba como 2.00 y 1.68 (auditoría 2026-09).
    """
    amounts = []
    for piece in re.split(r"[;\n]|,\s+", raw or ""):
        piece = piece.strip().replace("$", "")
        if not piece:
            continue
        if _THOUSANDS_AMOUNT_RE.fullmatch(piece):
            amounts.append(float(piece.replace(",", "")))
            continue
        amounts.extend(float(part.strip()) for part in piece.split(",") if part.strip())
    return amounts


@app.route("/reporte/dia/<report_date>/store-info", methods=["POST"])
def reporte_dia_store_info(report_date):
    credit_terms_raw = request.form.get("credit_terms", "")
    try:
        credit_terms = _parse_amount_list(credit_terms_raw)
        fields = {
            "from_time": request.form.get("from_time") or None,
            "to_time": request.form.get("to_time") or None,
            "volume": request.form.get("volume", ""),
            "sales_fuel": request.form.get("sales_fuel", ""),
            "desc_comb": request.form.get("desc_comb", ""),
            "non_fuel_total": request.form.get("non_fuel_total", ""),
            "desc_otros": request.form.get("desc_otros", ""),
            "tax_collect": request.form.get("tax_collect", ""),
            "total_sales": request.form.get("total_sales", ""),
            "cash": request.form.get("cash", ""),
            "local_accounts": request.form.get("local_accounts", ""),
            "other_amount": request.form.get("other_amount", ""),
            "network_revenue": request.form.get("network_revenue", ""),
            "total_revenue": request.form.get("total_revenue", ""),
            "credit_terms": credit_terms,
        }
        reportes_db.upsert_store_info(report_date, fields, source="manual")
    except ValueError:
        flash("No se pudo guardar: revisá que los montos sean números válidos.", "error")
        return redirect(url_for("reporte_dia", report_date=report_date))

    flash("Store Info guardado.", "success")
    return redirect(url_for("reporte_dia", report_date=report_date))


@app.route("/lottery")
def lottery():
    return render_template("lottery.html", **THEME_BY_KEY["lottery"])


def _lottery_pdf_upload():
    """Shared validation + upload-saving for the two Lottery forms."""
    master_upload = request.files.get("master_file")
    pdf_uploads = request.files.getlist("pdf_files")
    if not pdf_uploads or not any(u.filename for u in pdf_uploads):
        return None, _error_response("Seleccioná uno o más PDF.")
    if master_upload is None or not master_upload.filename:
        return None, _error_response("Seleccioná el Excel de Lottery.")

    workdir = _new_workspace_dir()
    master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
    pdf_paths = _save_uploads_to_workspace(pdf_uploads, workdir=workdir)
    return (master_path, master_filename, pdf_paths), None


@app.route("/lottery/sales-report", methods=["POST"])
def lottery_sales_report():
    saved, error = _lottery_pdf_upload()
    if error is not None:
        return error
    master_path, master_filename, pdf_paths = saved

    try:
        temp_path, summary = process_lottery(master_path, [], pdf_paths)
    except Exception as exc:
        return _error_response(f"Error: {exc}")

    _persist_lottery_sales_report_fields(pdf_paths)

    return _success_response(temp_path, master_filename, notice=_reporte_batch_notice(summary))


@app.route("/lottery/department", methods=["POST"])
def lottery_department():
    saved, error = _lottery_pdf_upload()
    if error is not None:
        return error
    master_path, master_filename, pdf_paths = saved

    try:
        temp_path, summary = process_lottery(master_path, pdf_paths, [])
    except Exception as exc:
        return _error_response(f"Error: {exc}")

    _persist_lottery_department_fields(pdf_paths)

    return _success_response(temp_path, master_filename, notice=_reporte_batch_notice(summary))


@app.route("/carga-datos/eft")
def carga_datos_eft():
    """
    Carga directa de EFT y Cupones del lado Carga de Datos -- solo PDF/
    reporte mensual, sin ningún Excel. Ver eft_db.py.
    """
    _active_job = jobs.get_active_job("eft")
    _jh_job = jobs.get_active_job("jh_mensual")
    _batch_job = jobs.get_active_job("cupones_detalle")
    return render_template(
        "carga_datos_eft.html",
        resume_job_id=(_active_job["id"] if _active_job else None),
        resume_jh_job_id=(_jh_job["id"] if _jh_job else None),
        resume_batch_job_id=(_batch_job["id"] if _batch_job else None),
        **THEME_BY_KEY["carga_eft"],
    )


@app.route("/carga-datos/eft/cupones-detalle/subir", methods=["POST"])
def carga_datos_eft_cupones_detalle_subir():
    """
    Detalle de cupones (pedido del usuario, 2026-10-06): el "Credit Card
    Batch Detail" que se imprime de cada cupón en el portal de J.H. Cada
    cupón queda con su monto exacto y las fechas reales de sus batches.
    Lectura en cupones_detalle.py, guardado en eft_db.save_coupon_batches.
    """
    uploads = request.files.getlist("batch_files")
    if not uploads or not any(u.filename for u in uploads):
        return _error_response("Seleccioná los PDF del detalle de cupones (Credit Card Batch Detail).")
    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths), kind="cupones_detalle")
    threading.Thread(target=_run_cupones_detalle_job, args=(job_id, paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


def _run_cupones_detalle_job(job_id, paths):
    """Aislado por archivo: un PDF roto no frena a los demás."""
    try:
        results, problems = [], []
        batches = 0
        for index, path in enumerate(paths, start=1):
            try:
                groups = cupones_detalle.extract_detail_groups(path)
                results.extend(eft_db.save_coupon_detail(groups, os.path.basename(path)))
                batches += sum(len(g["batches"]) for g in groups)
                eft_db.identify_detail_groups(jh_mensual_db.get_coupon_rows())
            except ValueError as exc:
                problems.append(str(exc))
            except Exception as exc:
                print(f"[carga-datos/eft/cupones-detalle] {path}: {exc}")
                problems.append("Un archivo no se pudo leer.")
            jobs.update_job(job_id, done=index, total=len(paths))
        parts = []
        if results:
            new = [r for r in results if r["status"] == "new"]
            first = min(r["first_date"] for r in results)
            last = max(r["last_date"] for r in results)
            line = (f"Detalle de cupones cargado: {len(new)} depósito{'s' if len(new) != 1 else ''} nuevo{'s' if len(new) != 1 else ''}"
                    f" ({batches} batches), ventas del {_fmt_ddmmyyyy(first)[:5]} al {_fmt_ddmmyyyy(last)[:5]}")
            if len(results) > len(new):
                line += f"; {len(results) - len(new)} ya estaba{'n' if len(results) - len(new) != 1 else ''} cargado{'s' if len(results) - len(new) != 1 else ''}"
            parts.append(line + ".")
            ids = {r["id"] for r in results}
            unidentified = sum(1 for g in eft_db.get_detail_groups() if g["id"] in ids and not g["coupons"])
            if unidentified:
                parts.append(
                    f"{unidentified} todavía sin DDC: se identifican solos cuando se cargue el reporte mensual de cupones "
                    "(Credit Card Daily Summary) que los trae."
                )
        parts.extend(problems)
        level = "success" if results and not problems else ("warning" if results else "error")
        jobs.update_job(
            job_id, status="done", done=len(paths), total=len(paths),
            notice=" ".join(parts), notice_level=level, redirect_url="/controles/tarjetas",
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


def _save_jh_monthly_report(path):
    """
    Un reporte de J.H. de cualquier extensión, mes por mes: lo ya cargado se
    deja como está y se suma lo nuevo (jh_mensual_db.merge_report).
    """
    report = jh_mensual.extract_report(path, jh_mensual_db.get_rows)
    kind_label = jh_mensual.KIND_LABELS[report["kind"]]
    parts, warnings, last = [], [], None
    for month_report in report["months"]:
        result = jh_mensual_db.merge_report(report["kind"], month_report, os.path.basename(path))
        label = f"{_MONTH_NAMES_ES[month_report['month'] - 1].lower()} {month_report['year']}"
        news = []
        if result["added"]:
            news.append(f"{result['added']} nuevo{'s' if result['added'] != 1 else ''}")
        if result["updated"]:
            news.append(f"{result['updated']} actualizado{'s' if result['updated'] != 1 else ''}")
        parts.append(f"{label} ({', '.join(news) if news else 'ya estaba'})")
        if result["conflicts"]:
            keys = [_fmt_ddmmyyyy(k)[:5] if report["kind"] == "coupons" else k for k in result["conflicts"]]
            warnings.append(
                f"{kind_label} de {label}: {', '.join(keys)} {'vienen' if len(keys) != 1 else 'viene'} con otro importe "
                "que lo ya cargado (se dejó lo cargado)."
            )
        last = (month_report["year"], month_report["month"])
    if report["kind"] == "coupons":
        # El mismo reporte que se sube en Excel en "Cupones": carga también
        # Cupones, solo los DDC que faltan (pedido del usuario, 2026-10-06).
        raw = [
            {"coupon": ",".join(r["coupons"]), "gross": r["gross"], "fees": r["fees"], "net": r["net"],
             "date": "{d.month}/{d.day}/{d.year}".format(d=date.fromisoformat(r["date"]))}
            for month_report in report["months"] for r in month_report["rows"]
        ]
        inserted, _ = eft_db.insert_new_cupones(expand_monthly_records_for_storage(raw), os.path.basename(path))
        if inserted:
            parts.append(f"{inserted} {'cupones nuevos' if inserted != 1 else 'cupón nuevo'} en Cupones")
        eft_db.identify_detail_groups(jh_mensual_db.get_coupon_rows())
    parts.extend(report["notes"])
    return f"{kind_label}: {', '.join(parts)}", last, warnings


@app.route("/carga-datos/eft/jh/subir", methods=["POST"])
def carga_datos_eft_jh_subir():
    """
    Reportes mensuales de J.H. (pedido del usuario, 2026-10-06): EFT History,
    Invoice History y Credit Card Daily Summary, en PDF; cada uno se reconoce
    solo y se cruza en Controles → Tarjetas y Cupones.
    """
    uploads = request.files.getlist("jh_files")
    if not uploads or not any(u.filename for u in uploads):
        return _error_response("Seleccioná los PDF de los reportes mensuales de J.H.")
    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths), kind="jh_mensual")
    threading.Thread(
        target=_run_monthly_upload_job,
        args=(job_id, paths, _save_jh_monthly_report, "Reportes de J.H. cargados",
              lambda y, m: f"/carga-datos/eft/jh?year={y}&month={m}", "/carga-datos/eft/jh"),
        daemon=True,
    ).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


@app.route("/carga-datos/eft/jh")
def carga_datos_eft_jh():
    """Los reportes mensuales de J.H. guardados de un mes, para verlos y eliminarlos."""
    year, month = _cierre_month()
    reports = jh_mensual_db.get_reports(year, month)
    return render_template(
        "carga_datos_eft_jh.html",
        reports=reports,
        kinds=jh_mensual.KINDS,
        kind_labels=jh_mensual.KIND_LABELS,
        **_month_nav(year, month),
        **THEME_BY_KEY["carga_eft"],
    )


@app.route("/carga-datos/eft/jh/eliminar", methods=["POST"])
def carga_datos_eft_jh_eliminar():
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    kind = request.form.get("kind")
    if year and month and kind in jh_mensual.KINDS:
        jh_mensual_db.delete_report(kind, year, month)
    return redirect(url_for("carga_datos_eft_jh", year=year, month=month))


def _jh_month_checks(year, month, detail_groups=()):
    """Los cruces de los reportes mensuales de J.H. del mes (None el que no está cargado)."""
    reports = jh_mensual_db.get_reports(year, month)
    checks = {"reports": reports}
    checks["eft"] = jh_mensual.eft_check(reports["eft"], year, month) if "eft" in reports else None
    checks["invoices"] = (
        jh_mensual.invoice_check(reports["invoices"], year, month, jh_mensual_db.get_all_invoice_numbers())
        if "invoices" in reports else None
    )
    checks["coupons"] = (
        jh_mensual.coupon_check(reports["coupons"], year, month, detail_groups) if "coupons" in reports else None
    )
    return checks


@app.route("/carga-datos/eft/subir", methods=["POST"])
def carga_datos_eft_subir():
    """
    Extrae uno o más PDF de EFT (pedido explícito del usuario 2026-09-12:
    "quiero que se puedan cargar varios pdf de eft de una sola vez") y los
    guarda en eft_db -- nunca genera ni toca ningún Excel. Cada archivo se
    aísla en su propio try/except (mismo criterio de todo el proyecto: un
    PDF roto o duplicado no debe tirar abajo el resto del lote). Mismo
    chequeo de duplicado que el lado Excels (RCV + fecha + neto), pero
    contra la base en vez de un Ledger.
    """
    pdf_uploads = [f for f in request.files.getlist("pdf_file") if f and f.filename]
    if not pdf_uploads:
        return _error_response("Seleccioná uno o más PDF de EFT.")

    pdf_paths = _save_uploads_to_workspace(pdf_uploads)
    job_id = jobs.create_job(len(pdf_paths), kind="eft")
    threading.Thread(target=_run_carga_datos_eft_job, args=(job_id, pdf_paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(pdf_paths)})


def _run_carga_datos_eft_job(job_id, pdf_paths):
    """Corre en su propio hilo -- mismo patrón que _run_carga_datos_reporte_diario_job, ver ese docstring."""
    try:
        saved = 0
        duplicates = 0
        missing_ddc_total = 0
        skipped_total = 0
        failed = 0
        date_mismatches = 0
        last_saved_date = None

        for index, pdf_path in enumerate(pdf_paths, start=1):
            filename = os.path.basename(pdf_path)
            try:
                header_data, paid_invoices, credit_coupons, skipped_coupon_rows = extract_eft_data(pdf_path)
                if not credit_coupons:
                    raise ValueError("no se extrajeron cupones de tarjeta de crédito")

                eft_date = header_data.get("eft_date")
                try:
                    parsed = datetime.strptime(eft_date, "%m/%d/%Y") if eft_date else None
                except ValueError:
                    parsed = None
                # Sin fecha válida el EFT no aparecería en ningún mes (no se
                # vería ni se podría borrar) y cada resubida lo duplicaría:
                # se rechaza (auditoría 2026-09, webapp.py:4312).
                if parsed is None:
                    raise ValueError("no se pudo leer la fecha del EFT (MM/DD/YYYY)")
                # Mismo chequeo que Reporte Diario/Lottery (2026-09-15): estos
                # PDF suelen traer su propia fecha en el nombre (ej. "EFT
                # Nº21062 03.08.2026.pdf") -- si no coincide con la fecha real
                # leída del PDF, se rechaza en vez de guardarlo bajo una fecha
                # dudosa.
                if parsed and _filename_date_mismatch(filename, parsed):
                    date_mismatches += 1
                    raise ValueError("la fecha leída no coincide con la del nombre de archivo")

                net_total = sum(float(c.get("paid_amount") or 0.0) for c in credit_coupons)
                existing = eft_db.find_existing_deposit(
                    header_data.get("draft_no"), header_data.get("eft_date"), net_total
                )
                if existing:
                    duplicates += 1
                else:
                    eft_db.save_eft(header_data, paid_invoices, credit_coupons, source_filename=filename)
                    saved += 1
                    missing_ddc_total += sum(1 for c in credit_coupons if not c.get("coupon"))
                    skipped_total += skipped_coupon_rows or 0

                    if parsed and (last_saved_date is None or parsed > last_saved_date):
                        last_saved_date = parsed
                    try:
                        doc_when = parsed or datetime.now()
                        documents_db.store_document(
                            "eft", pdf_path, filename, doc_when.year, doc_when.month,
                            label=header_data.get("draft_no"),
                        )
                    except Exception as exc:
                        print(f"[documents_db] no se pudo guardar el documento de EFT {filename}: {exc}")
            except Exception as exc:
                print(f"[carga-datos/eft/subir] {filename}: {exc}")
                failed += 1
            jobs.update_job(job_id, done=index, total=len(pdf_paths))

        parts = []
        if saved:
            parts.append(f"{saved} EFT guardado(s).")
        if duplicates:
            parts.append(f"{duplicates} ya estaban cargado(s) y se omitieron.")
        # Líneas sin DDC que coinciden al centavo con un único cupón pendiente
        # se completan solas (eft_db.autolink_missing_ddc_by_amount).
        autolinked = 0
        try:
            autolinked = eft_db.autolink_missing_ddc_by_amount()
        except Exception as exc:
            print(f"[eft] no se pudieron completar DDC por monto: {exc}")
        if autolinked:
            parts.append(f"{autolinked} DDC faltante(s) se completaron solos: el monto coincide exacto con un cupón pendiente.")
            missing_ddc_total = max(0, missing_ddc_total - autolinked)
        if missing_ddc_total:
            parts.append(f"{missing_ddc_total} cupón(es) sin número DDC -- se puede agregar a mano abajo.")
        if skipped_total:
            parts.append(f"{skipped_total} fila(s) de cupón no se pudieron leer.")
        if date_mismatches:
            parts.append(
                f"{date_mismatches} archivo(s) rechazados: la fecha real no coincide con la del "
                "nombre de archivo — revisá que no sea de otro mes."
            )
        if failed - date_mismatches:
            parts.append(f"{failed - date_mismatches} archivo(s) no se pudieron procesar.")
        if not parts:
            parts.append("No se guardó ningún EFT de este lote.")
        level = (
            "success" if (saved and not (duplicates or missing_ddc_total or skipped_total or failed)) else
            ("error" if not saved else "warning")
        )

        if last_saved_date:
            redirect_url = f"/carga-datos/eft/historial?year={last_saved_date.year}&month={last_saved_date.month}"
        else:
            redirect_url = "/carga-datos/eft/historial"

        jobs.update_job(
            job_id, status="done", done=len(pdf_paths), total=len(pdf_paths),
            notice=" ".join(parts), notice_level=level, redirect_url=redirect_url,
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


@app.route("/carga-datos/eft/cupones/subir", methods=["POST"])
def carga_datos_eft_cupones_subir():
    """
    Lee el reporte mensual de Cupones y lo guarda en eft_db -- nunca genera
    ni toca ningún Excel. Cada DDC queda cruzado en el momento contra los
    EFT ya guardados (join en la lectura, ver eft_db.get_cupones_with_status),
    no hace falta ningún paso de "resincronizar" aparte como en el Excel.
    """
    monthly_upload = request.files.get("monthly_report_file")
    if monthly_upload is None or not monthly_upload.filename:
        flash("Seleccioná el reporte mensual de Cupones.", "error")
        return redirect(url_for("carga_datos_eft"))

    try:
        monthly_path, filename = _save_upload_to_workspace(monthly_upload)
        raw_rows = read_monthly_coupon_rows(monthly_path)
        if not raw_rows:
            raise ValueError("No se encontraron filas de cupones en el reporte mensual.")
        records = expand_monthly_records_for_storage(raw_rows)
    except Exception as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("carga_datos_eft"))

    inserted, updated, repeated = eft_db.upsert_cupones(records, source_filename=filename)
    try:
        today = date.today()
        documents_db.store_document("eft", monthly_path, filename, today.year, today.month, label="Reporte mensual de Cupones")
    except Exception as exc:
        print(f"[documents_db] no se pudo guardar el reporte mensual de Cupones {filename}: {exc}")
    flash(f"{inserted + updated} cupón(es) guardado(s) ({inserted} nuevo(s), {updated} actualizado(s)).", "success")
    if repeated:
        shown = ", ".join(repeated[:5]) + ("…" if len(repeated) > 5 else "")
        flash(
            f"El reporte trae {len(repeated)} DDC repetido(s) ({shown}): quedó guardado uno solo de cada uno, revisalos.",
            "warning",
        )
    return redirect(url_for("carga_datos_eft_historial"))


# Enlazar cada "factura que paga el EFT" con el documento real ya subido a
# Documentos (módulo "eft") -- pedido explícito del usuario (2026-09-14):
# "que pueda tocar con un link la invoice y eso lo va a reedirigir a la
# factura que se pago, linkeado de los mismos documentos por el nombre de
# invoice". No hay ninguna referencia guardada entre EFT y documento -- el
# cruce se recalcula en el momento (nunca un valor fijo que pueda quedar
# desactualizado) buscando el número de la factura (ej. "SI-212530" ->
# "212530") como subcadena del nombre de archivo o la etiqueta del
# documento -- así, subir un documento nuevo actualiza el link solo, sin
# tocar nada del lado del EFT.
_INVOICE_DIGITS_RE = re.compile(r"\d{4,}")


def _invoice_number_digits(text):
    if not text:
        return None
    match = _INVOICE_DIGITS_RE.search(text)
    return match.group(0) if match else None


def _match_invoice_document(invoice_text, docs):
    digits = _invoice_number_digits(invoice_text)
    if not digits:
        return None
    for doc in docs:
        haystack = f"{doc.get('filename') or ''} {doc.get('label') or ''}"
        if digits in (_invoice_number_digits(haystack) or ""):
            return doc
        if digits in haystack:
            return doc
    return None


@app.route("/carga-datos/eft/historial")
def carga_datos_eft_historial():
    """
    EFT del mes (con sus cupones/facturas, ordenados por la fecha del EFT)
    -- solo EFT, ver carga_datos_eft_cupones_historial para el historial
    completo de Cupones. Pedido explícito del usuario (2026-09-18): antes
    el historial ENTERO de Cupones (histórico, desde inicios de 2026) salía
    pegado debajo de los EFT de un solo mes en la misma página -- separados
    en dos apartados propios, cada uno con su propia navegación.
    """
    try:
        eft_db.autolink_missing_ddc_by_amount()
    except Exception as exc:
        print(f"[eft] no se pudieron completar DDC por monto: {exc}")
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    deposits = eft_db.get_month_deposits(year, month)
    eft_docs = documents_db.list_all_documents("eft")
    for entry in deposits:
        # Autosuma por columna (pedido explícito del usuario 2026-09-12) --
        # se calcula acá y no con el filtro |sum(attribute=...) de Jinja
        # porque un monto en None (columna nullable) rompería ese filtro.
        entry["totals"] = {
            "gross": sum(float(c.get("gross_amount") or 0.0) for c in entry["coupons"]),
            "fees": sum(float(c.get("fees_amount") or 0.0) for c in entry["coupons"]),
            "paid": sum(float(c.get("paid_amount") or 0.0) for c in entry["coupons"]),
        }
        for inv in entry.get("paid_invoices") or []:
            inv["document"] = _match_invoice_document(inv.get("invoice"), eft_docs)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "carga_datos_eft_historial.html",
        deposits=deposits,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["carga_eft"],
    )


@app.route("/carga-datos/eft/cupones/historial")
def carga_datos_eft_cupones_historial():
    """
    Cupones por mes (pedido del usuario, 2026-10-06; antes era el historial
    completo agrupado por año): los que aplicaron los EFT del mes, juntos
    por EFT con su total, los que quedaron pendientes al cierre del mes
    elegido y las ventas del mes que entraron en cupones del mes siguiente
    (eft_db.cupones_month_view). Columnas Fecha/DDC/Gross/Fee/Net Amount/
    Diferencia/EFT/Fecha EFT como antes (pedido del usuario, 2026-09-16).
    """
    try:
        backfilled = eft_db.backfill_grouped_cupones_from_eft()
        if backfilled:
            print(f"[cupones] {backfilled} cupón(es) de un batch completado(s) con datos de EFT ya cargados.")
        autolinked = eft_db.autolink_missing_ddc_by_amount()
        if autolinked:
            print(f"[cupones] {autolinked} DDC de EFT completado(s) por monto exacto.")
    except Exception as exc:
        print(f"[cupones] no se pudo completar cupones agrupados desde EFT: {exc}")

    year, month = _cierre_month()
    view = eft_db.cupones_month_view(year, month)
    # Días de venta de cada DDC, del detalle de cupones (el del depósito que lo trae).
    sale_days = {}
    for group in eft_db.get_detail_groups():
        for coupon in group["coupons"] or []:
            sale_days[coupon] = (group["first_date"], group["last_date"], len(group["coupons"]))
    for cp in view["applied"] + view["pending"]:
        days = sale_days.get(cp["coupon_id"] if "coupon_id" in cp else cp["coupons"][0])
        cp["sale_days"] = (
            {"from": _fmt_ddmmyyyy(days[0])[:5], "to": _fmt_ddmmyyyy(days[1])[:5], "group": days[2]} if days else None
        )
    for cp in view["applied"]:
        if cp["prev_month"]:
            parsed = eft_db._parse_cupon_date(cp.get("date"))
            cp["prev_month_label"] = _MONTH_NAMES_ES[parsed.month - 1].lower()
    for row in view["pending"]:
        for later in row["laters"]:
            later["label"] = f"{later['eft_date'].strftime('%d/%m/%Y')} ({_MONTH_NAMES_ES[later['eft_date'].month - 1].lower()})"
    for sub in view["pending_by_month"]:
        sub["label"] = _MONTH_NAMES_ES[sub["month"] - 1].lower()
    for group in view["next_month"]:
        group["deposit_label"] = group["deposit"].strftime("%d/%m/%Y") if group["deposit"] else None
        group["from_label"] = _fmt_ddmmyyyy(group["from"])[:5]
        group["to_label"] = _fmt_ddmmyyyy(group["to"])[:5]

    nav = _month_nav(year, month)
    following = _MONTH_NAMES_ES[nav["next_month"] - 1].lower()
    years = eft_db.get_cupones_years()
    return render_template(
        "carga_datos_eft_cupones_historial.html",
        view=view,
        following_name=following,
        today=date.today(),
        delete_month_years=years if year in years else sorted(set(years) | {year}),
        month_names=_MONTH_NAMES_ES,
        **nav,
        **THEME_BY_KEY["carga_eft"],
    )


@app.route("/carga-datos/eft/cupones/borrar-mes", methods=["POST"])
def carga_datos_eft_cupones_borrar_mes():
    """
    Borra todos los cupones guardados de un mes/año puntual -- pedido
    explícito del usuario (2026-09-17: "no hay forma de borrar los
    cupones cargados... debería haber una forma de borrar todos los
    cupones del mes si se quisiera"). Acción irreversible -- el template
    ya pide confirmación antes de mandar el POST.
    """
    try:
        year = int(request.form.get("year"))
        month = int(request.form.get("month"))
        if not (1 <= month <= 12):
            raise ValueError
    except (TypeError, ValueError):
        flash("Mes/año inválido.", "error")
        return redirect(url_for("carga_datos_eft_cupones_historial"))

    deleted = eft_db.delete_cupones_month(year, month)
    if deleted:
        flash(f"{deleted} cupón(es) de {_MONTH_NAMES_ES[month - 1]} {year} borrado(s).", "success")
    else:
        flash(f"No había ningún cupón guardado en {_MONTH_NAMES_ES[month - 1]} {year}.", "warning")
    return redirect(url_for("carga_datos_eft_cupones_historial", year=year, month=month))


@app.route("/carga-datos/eft/coupon/<int:eft_coupon_id>/editar", methods=["POST"])
def carga_datos_eft_coupon_editar(eft_coupon_id):
    """
    Corrige a mano la fila completa de un cupón de EFT (fecha/factura/DDC/
    Gross/Fees/Net Amount) -- pedido explícito del usuario (2026-09-16):
    "corregir" solo dejaba editar el DDC, "cuando se ponga corregir que te
    deje editar la fila". Reemplaza a la vieja carga_datos_eft_coupon_ddc
    (solo DDC).
    """
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    coupon_date = (request.form.get("date") or "").strip() or None
    invoice = (request.form.get("invoice") or "").strip() or None
    coupon_id = (request.form.get("coupon_id") or "").strip()
    gross_amount = request.form.get("gross_amount", type=float)
    fees_amount = request.form.get("fees_amount", type=float)
    paid_amount = request.form.get("paid_amount", type=float)
    ok = eft_db.update_coupon_row(
        eft_coupon_id,
        date=coupon_date,
        invoice=invoice,
        coupon=coupon_id,
        gross_amount=gross_amount,
        fees_amount=fees_amount,
        paid_amount=paid_amount,
    )
    flash("Cupón corregido." if ok else "No se encontró esa línea de EFT.", "success" if ok else "error")
    return redirect(url_for("carga_datos_eft_historial", year=year, month=month))


@app.route("/carga-datos/eft/<int:deposit_id>/eliminar", methods=["POST"])
def carga_datos_eft_eliminar(deposit_id):
    """
    Borra un EFT ya cargado -- pedido explícito del usuario (2026-09-14):
    "los EFT subidos no tienen forma de ser eliminados". Los documentos
    guardados en Documentos (módulo "eft") NO se tocan -- son archivos
    originales aparte, no dependen del registro del EFT.
    """
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    ok = eft_db.delete_deposit(deposit_id)
    flash("EFT eliminado." if ok else "Ese EFT ya no existe.", "success" if ok else "error")
    return redirect(url_for("carga_datos_eft_historial", year=year, month=month))


@app.route("/carga-datos/caja")
def carga_datos_caja():
    """
    Caja del lado Carga de Datos -- pedido explícito del usuario
    (2026-09-12): "que esta vez se completaria automaticamente con los
    datos que haya guardado en el chase de tal mes... tendria que verse
    como el cuadro del excel exactamente igual". Chase y Lottery se cargan
    por sus propios módulos, este reporte solo cruza lo que ya está
    guardado (ver caja.build_month_report_from_db). Lo único que se sube
    acá son los comprobantes de los gastos en efectivo (2026-10-04, ver
    carga_datos_caja_gastos_subir).
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    report = build_caja_month_report(year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    _active_job = jobs.get_active_job("caja_gastos")
    return render_template(
        "carga_datos_caja_historial.html",
        report=report,
        expense_items=caja_db.get_month_expense_items(year, month),
        pending_expenses=caja_db.get_month_pending_expenses(year, month),
        resume_job_id=(_active_job["id"] if _active_job else None),
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        today_iso=today.isoformat(),
        month_last_day_iso=date(year, month, calendar.monthrange(year, month)[1]).isoformat(),
        **THEME_BY_KEY["carga_caja"],
    )


@app.route("/carga-datos/caja/exportar")
def carga_datos_caja_exportar():
    """
    Descarga un Excel NUEVO (nunca toca ningún archivo real) con la Caja
    del mes -- pedido explícito del usuario (2026-09-19), mismo criterio
    que reporte_store_info_exportar: "hagamos algo igual del exportar la
    caja y que quede con el formato que tenia en el cierre".
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    report = build_caja_month_report(year, month)
    workspace_dir = tempfile.mkdtemp(prefix="caja_export_")
    dest_path = os.path.join(workspace_dir, f"Caja {month:02d}-{year}.xlsx")
    build_caja_export_workbook(report, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/caja/exportar/pdf")
def carga_datos_caja_exportar_pdf():
    """
    Versión PDF (básica, sin colores, solo líneas y bordes) del export de
    arriba -- pedido explícito del usuario (2026-09-19). Ver
    caja.build_caja_export_pdf.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    report = build_caja_month_report(year, month)
    workspace_dir = tempfile.mkdtemp(prefix="caja_export_pdf_")
    dest_path = os.path.join(workspace_dir, f"Caja {month:02d}-{year}.pdf")
    build_caja_export_pdf(report, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/caja/gastos/agregar", methods=["POST"])
def carga_datos_caja_gastos_agregar():
    """
    Agrega UN gasto en efectivo con su detalle (a quién se le pagó) --
    pedido explícito del usuario (2026-09-12, cuarta tanda): "poder
    escribirlos a mano en el sistema con un detalle de a quien
    pertenecen". Reemplaza el form viejo de un solo monto por día (sin
    detalle) -- un día puede tener varios gastos, cada uno con el suyo.
    """
    report_date = request.form.get("date")
    amount_raw = (request.form.get("amount") or "").strip()
    detail = request.form.get("detail")
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)

    try:
        amount = float(amount_raw)
    except ValueError:
        flash("El monto tiene que ser un número válido.", "error")
        return redirect(url_for("carga_datos_caja", year=year, month=month))

    # La fecha tiene que caer en el mes que se está viendo: si no, el gasto
    # se iba a otro mes sin aviso y parecía no haberse guardado.
    try:
        parsed_date = datetime.strptime(report_date or "", "%Y-%m-%d").date()
    except ValueError:
        flash("La fecha del gasto no es válida.", "error")
        return redirect(url_for("carga_datos_caja", year=year, month=month))
    if year and month and (parsed_date.year, parsed_date.month) != (year, month):
        flash(f"La fecha del gasto ({parsed_date:%d/%m/%Y}) no es de {month:02d}/{year}.", "error")
        return redirect(url_for("carga_datos_caja", year=year, month=month))

    caja_db.add_expense_item(report_date, amount, detail)
    flash("Gasto agregado.", "success")
    return redirect(url_for("carga_datos_caja", year=year, month=month))


@app.route("/carga-datos/caja/gastos/subir", methods=["POST"])
def carga_datos_caja_gastos_subir():
    """
    Gastos de Caja leídos de los comprobantes escaneados -- pedido del
    usuario (2026-10-04): "implementemos un sistema de OCR para subir los
    gastos del mes hechos con caja, así aunque no sean proveedores ya
    escaneados, que sea un intento de cargar de forma automática gastos del
    mes". Lo que se lee es el ticket "Paid Out" del POS que va abrochado a
    cada comprobante (ver gastos_caja.py).
    """
    uploads = request.files.getlist("gasto_files")
    if not uploads or not any(u.filename for u in uploads):
        return _error_response("Seleccioná uno o más comprobantes (PDF o foto).")
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths), kind="caja_gastos")
    threading.Thread(target=_run_caja_gastos_job, args=(job_id, paths, year, month), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


def _run_caja_gastos_job(job_id, paths, year, month):
    """
    Aislado por archivo, como los demás jobs. Nunca guarda un ticket con la
    fecha o el monto dudoso; un comprobante sin ticket de caja queda
    pendiente de confirmar (caja_db.add_pending_expense), no se carga solo.
    """
    try:
        saved = duplicates = failed = without_detail = 0
        other_months, no_slip, doubtful = [], [], []
        for index, path in enumerate(paths, start=1):
            try:
                filename = os.path.basename(path)
                result = gastos_caja.extract_cash_expenses(path, filename, default_year=year)
                name = result["detail"] or "un comprobante sin nombre"
                proposal = result["proposal"]
                if proposal is not None:
                    already = (proposal["date"] and proposal["amount"] is not None
                               and caja_db.find_same_expense(proposal["date"], proposal["amount"]))
                    if already:
                        duplicates += 1
                    elif caja_db.add_pending_expense(year, month, proposal["date"], proposal["amount"],
                                                     result["detail"], filename):
                        no_slip.append(name)
                    else:  # ese archivo ya estaba en "Para confirmar"
                        duplicates += 1
                # Ticket encontrado pero con la fecha o el monto dudoso: también
                # va a "Para confirmar", con lo que sí se pudo leer.
                for number, slip in enumerate(result["doubtful"], start=1):
                    pending_date = slip["date"] or gastos_caja.date_from_filename(filename, year)
                    if caja_db.add_pending_expense(year, month, pending_date, slip["amount"], result["detail"],
                                                   filename if number == 1 else f"{filename} (ticket {number})"):
                        doubtful.append(name)
                    else:
                        duplicates += 1
                for slip in result["slips"]:
                    if caja_db.find_same_expense(slip["date"], slip["amount"], slip["trans"]):
                        duplicates += 1
                        continue
                    caja_db.add_expense_item(slip["date"], slip["amount"], result["detail"],
                                             source="ocr", trans_no=slip["trans"])
                    saved += 1
                    without_detail += not result["detail"]
                    if year and month and (slip["date"].year, slip["date"].month) != (year, month):
                        other_months.append(slip["date"].strftime("%d/%m/%Y"))
            except Exception as exc:
                print(f"[carga-datos/caja/gastos] {path}: {exc}")
                failed += 1
            jobs.update_job(job_id, done=index, total=len(paths))

        parts = []
        if saved:
            parts.append(f"{saved} gasto(s) cargado(s) desde el ticket de la caja.")
        if other_months:
            parts.append(f"{len(other_months)} con fecha de otro mes ({', '.join(other_months)}): quedaron en su mes.")
        if duplicates:
            parts.append(f"{duplicates} ya estaban cargados y se omitieron.")
        if without_detail:
            parts.append(f"{without_detail} sin detalle (el nombre del archivo no dice a quién se le pagó): completalo en la tabla.")
        if doubtful:
            parts.append("Ticket de caja con la fecha o el monto dudoso (" + "; ".join(doubtful) + "): quedó abajo, "
                         "en \"Para confirmar\" -- revisalo contra el papel antes de agregarlo.")
        if no_slip:
            parts.append("Sin el ticket \"Paid Out\" de la caja (" + "; ".join(no_slip) + "): quedaron abajo, "
                         "en \"Para confirmar\", con la fecha y el total que se pudieron leer -- agregalos si "
                         "fueron en efectivo o descartalos.")
        if failed:
            parts.append(f"{failed} archivo(s) no se pudieron leer.")
        if saved:
            level = "warning" if (doubtful or no_slip or failed or without_detail or other_months) else "success"
        else:
            level = "warning" if (duplicates or no_slip or doubtful) and not failed else "error"
        jobs.update_job(
            job_id, status="done", done=len(paths), total=len(paths),
            notice=" ".join(parts) or "No se encontró ningún gasto.", notice_level=level,
            redirect_url=f"/carga-datos/caja?year={year}&month={month}" if year and month else "/carga-datos/caja",
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


@app.route("/carga-datos/caja/gastos/pendiente/<int:pending_id>/agregar", methods=["POST"])
def carga_datos_caja_gastos_pendiente_agregar(pending_id):
    """Confirma un comprobante sin ticket de caja: pasa a ser un gasto con lo que el usuario dejó en el form."""
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    try:
        amount = float((request.form.get("amount") or "").strip())
        parsed_date = datetime.strptime(request.form.get("date") or "", "%Y-%m-%d").date()
    except ValueError:
        flash("Completá la fecha y el monto del gasto antes de agregarlo.", "error")
        return redirect(url_for("carga_datos_caja", year=year, month=month))
    caja_db.add_expense_item(parsed_date, amount, request.form.get("detail"), source="comprobante")
    caja_db.delete_pending_expense(pending_id)
    if year and month and (parsed_date.year, parsed_date.month) != (year, month):
        flash(f"Gasto agregado en {parsed_date:%d/%m/%Y}, que es de otro mes.", "info")
    return redirect(url_for("carga_datos_caja", year=year, month=month))


@app.route("/carga-datos/caja/gastos/pendiente/<int:pending_id>/descartar", methods=["POST"])
def carga_datos_caja_gastos_pendiente_descartar(pending_id):
    caja_db.delete_pending_expense(pending_id)
    return redirect(url_for("carga_datos_caja", year=request.form.get("year", type=int),
                            month=request.form.get("month", type=int)))


@app.route("/carga-datos/caja/gastos/<int:item_id>/detalle", methods=["POST"])
def carga_datos_caja_gastos_detalle(item_id):
    caja_db.update_expense_detail(item_id, request.form.get("detail"))
    return redirect(url_for("carga_datos_caja", year=request.form.get("year", type=int),
                            month=request.form.get("month", type=int)))


@app.route("/carga-datos/caja/gastos/<int:item_id>/eliminar", methods=["POST"])
def carga_datos_caja_gastos_eliminar(item_id):
    caja_db.delete_expense_item(item_id)
    linked = proveedores_db.delete_manual_payment_by_caja_item(item_id)
    if linked:
        flash(
            f"Gasto eliminado. Era un pago a mano a {linked['supplier_label']}: también se borró de su cuenta corriente.",
            "info",
        )
    else:
        flash("Gasto eliminado.", "success")
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    return redirect(url_for("carga_datos_caja", year=year, month=month))


@app.route("/carga-datos/caja/saldo", methods=["POST"])
def carga_datos_caja_saldo():
    """
    Saldo Inicial (editable, encadenado del mes anterior por default) y un
    ajuste manual opcional de Saldo Final -- pedido explícito del usuario
    (2026-09-12): "el saldo inicial de la caja de un mes deberia ser el
    saldo final de la caja del mes anterior, este saldo se deberia poder
    editar junto con el saldo final de ser necesario". Un campo vacío borra
    el override (vuelve al valor encadenado/calculado).
    """
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    opening_raw = (request.form.get("opening_balance") or "").strip()
    closing_raw = (request.form.get("closing_balance_override") or "").strip()

    try:
        opening_value = float(opening_raw) if opening_raw else None
        closing_value = float(closing_raw) if closing_raw else None
        caja_db.set_month_opening_balance(year, month, opening_value)
        caja_db.set_month_closing_override(year, month, closing_value)
        flash("Saldo guardado.", "success")
    except ValueError:
        flash("No se pudo guardar: revisá que los montos sean números válidos.", "error")

    return redirect(url_for("carga_datos_caja", year=year, month=month))


@app.route("/carga-datos/gettel")
def carga_datos_gettel():
    """
    Gettel / Toyota -- Carga de Datos (2026-09-12, cuarta tanda): mismo
    origen que ya lee el módulo de Herramientas (Excel con hojas Gettel/
    Toyota, o PDF/foto por separado de cada uno) pero guardado directo en
    gettel_db, sin ningún Excel Cierre de destino -- pedido explícito del
    usuario: "quiero que empieces a crear los modulos de... el excel ese
    donde contengo los datos de gettel y toyota junto con sus gallons".

    Un solo cuadro con dos formularios (2026-09-22, pedido explícito del
    usuario): "la carga de pagos todavia no se puede hacer desde el modulo
    de gettel en cargar datos, ahi solo se pueden subir los dias, se
    deberia poder subir eso y tambien abajo los pagos, asi en la barra
    lateral no tengamos el Cargar y Cargar pagos" -- esta misma página
    ahora también sube los recibos de Pago (mismo form que antes vivía
    solo, en `carga_datos_gettel_pagos.html` / `/carga-datos/gettel/pagos`
    -- esa ruta sigue existiendo por compatibilidad de link viejo, ver su
    propio docstring, pero redirige acá).
    """
    _active_job = jobs.get_active_job("gettel")
    _active_pagos_job = jobs.get_active_job("gettel_pagos")
    return render_template(
        "carga_datos_gettel.html",
        resume_job_id=(_active_job["id"] if _active_job else None),
        resume_pagos_job_id=(_active_pagos_job["id"] if _active_pagos_job else None),
        **THEME_BY_KEY["carga_gettel"],
    )


@app.route("/carga-datos/gettel/subir", methods=["POST"])
def carga_datos_gettel_subir():
    uploads = request.files.getlist("source_files")
    if not uploads or not any(u.filename for u in uploads):
        return _error_response("Seleccioná uno o más Excel/PDF de cupones de Gettel y/o Toyota.")

    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths), kind="gettel")
    threading.Thread(target=_run_carga_datos_gettel_job, args=(job_id, paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


def _run_carga_datos_gettel_job(job_id, paths):
    """Corre en su propio hilo -- mismo patrón que _run_carga_datos_reporte_diario_job, ver ese docstring."""
    try:
        days_gettel = set()
        days_toyota = set()
        files_failed = 0
        files_to_review = 0
        first_date = None

        for index, path in enumerate(paths, start=1):
            ext = os.path.splitext(path)[1].lower()
            batch_days = set()
            try:
                if ext in (".xlsx", ".xlsm"):
                    gettel_totals, toyota_totals = summarize_origin_workbook(path)
                    if gettel_totals:
                        gettel_db.upsert_vendor_totals("gettel", gettel_totals, source="excel")
                        days_gettel.update(gettel_totals)
                    if toyota_totals:
                        gettel_db.upsert_vendor_totals("toyota", toyota_totals, source="excel")
                        days_toyota.update(toyota_totals)
                    batch_days = set(gettel_totals) | set(toyota_totals)
                elif ext == ".pdf":
                    vendor = detect_vendor_from_ocr_text(path)
                    if vendor is None:
                        raise ValueError("No se pudo determinar si el PDF es de Gettel o Toyota.")
                    totals_by_date, diagnostics = summarize_pdf_report(path)
                    if not totals_by_date:
                        raise ValueError("No se pudo leer ninguna fila del reporte.")
                    # Mismo control que /gettel/cupones: lo leído contra el
                    # subtotal impreso, más las filas sin Monto legible (que
                    # se suman como $0). Se guarda igual, pero con aviso.
                    subtotal_mismatch = diagnostics.get("printed_subtotal_found") and not (
                        diagnostics.get("amount_matches_subtotal") and diagnostics.get("gallons_matches_subtotal")
                    )
                    if subtotal_mismatch or diagnostics.get("rows_without_amount"):
                        files_to_review += 1
                    key = "gettel" if vendor == VENDOR_GETTEL[0] else "toyota"
                    gettel_db.upsert_vendor_totals(key, totals_by_date, source="pdf")
                    (days_gettel if key == "gettel" else days_toyota).update(totals_by_date)
                    batch_days = set(totals_by_date)
                else:
                    raise ValueError("Formato no reconocido (subí un .xlsx o un .pdf).")

                if batch_days:
                    batch_first = min(batch_days)
                    if first_date is None or batch_first < first_date:
                        first_date = batch_first

                try:
                    doc_when = (batch_days and min(batch_days)) or date.today()
                    documents_db.store_document(
                        "gettel_toyota", path, os.path.basename(path), doc_when.year, doc_when.month
                    )
                except Exception as exc:
                    print(f"[documents_db] no se pudo guardar el documento de Gettel/Toyota {path}: {exc}")
            except Exception as exc:
                print(f"[carga-datos/gettel] {path}: {exc}")
                files_failed += 1
            jobs.update_job(job_id, done=index, total=len(paths))

        parts = []
        if days_gettel:
            parts.append(f"Gettel: {len(days_gettel)} día(s) guardado(s).")
        if days_toyota:
            parts.append(f"Toyota: {len(days_toyota)} día(s) guardado(s).")
        if files_failed:
            parts.append(f"{files_failed} archivo(s) no se pudieron leer.")
        if files_to_review:
            parts.append(
                f"{files_to_review} reporte(s) no cierran contra el subtotal impreso o tienen filas sin monto legible: revisalos a mano."
            )

        if not parts:
            notice, level = "No se pudo guardar nada de este lote.", "error"
        else:
            notice, level = " ".join(parts), ("warning" if (files_failed or files_to_review) else "success")

        if first_date:
            redirect_url = f"/carga-datos/gettel/historial?year={first_date.year}&month={first_date.month}"
        else:
            redirect_url = "/carga-datos/gettel"

        jobs.update_job(
            job_id, status="done", done=len(paths), total=len(paths),
            notice=notice, notice_level=level, redirect_url=redirect_url,
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


# ---------------------------------------------------------------------------
# Gettel -- Pagos de Cupones (2026-09-19, pedido explícito del usuario):
# pestaña separada de carga (a mano, un cupón/transacción por vez, ver
# gettel_pagos.py) más un cuadro que replica la hoja real "Pago Cupones"
# (tabla + totales del mes conectados al módulo de días + Pendiente Mes
# Anterior encadenado + export Excel/PDF). Reemplaza la lectura automática
# de PDF de pagos de /gettel/pagos ("no esta leyendo bien los pagos").
# ---------------------------------------------------------------------------
@app.route("/carga-datos/gettel/pagos")
def carga_datos_gettel_pagos():
    """
    Ruta vieja -- redirige a /carga-datos/gettel (2026-09-22, pedido
    explícito del usuario: subir los días y subir los pagos quedan juntos
    en la misma página, para no tener "Cargar" y "Cargar Pagos" como dos
    entradas separadas en la barra lateral -- ver carga_datos_gettel()).
    Se conserva el endpoint (no se borra) porque templates/enlaces viejos
    (ej. gettel_pagos_cuadro.html) todavía apuntan acá con url_for --
    redirige en vez de 404, mismo criterio que /excels -> /carga-datos.
    """
    return redirect(url_for("carga_datos_gettel"))


@app.route("/carga-datos/gettel/pagos/subir-pdf", methods=["POST"])
def carga_datos_gettel_pagos_subir_pdf():
    """
    Carga por PDF -- reemplaza la carga a mano (pedido explícito del
    usuario, 2026-09-21) usando los 4 PDFs reales que subió como ejemplo
    (ver gettel_pagos_parser.py para el detalle de qué se lee y por qué,
    incluido el bug real que tenía la herramienta vieja /gettel/pagos).
    Mismo patrón de siempre para lotes de PDF (jobs.py + threading, ver
    _run_carga_datos_combustible_job): cada archivo se procesa aislado, y
    DENTRO de cada archivo cada página/recibo también se aisla -- un recibo
    ilegible no tira abajo el resto del mismo PDF. Duplicado = mismo N° de
    Transacción ya guardado, en cualquier mes (ese número es un contador
    corrido de la caja registradora, nunca se repite).
    """
    uploads = [f for f in request.files.getlist("pdf_files") if f and f.filename]
    if not uploads:
        return _error_response("Seleccioná uno o más PDF de pagos de Gettel.")

    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths), kind="gettel_pagos")
    threading.Thread(target=_run_carga_datos_gettel_pagos_job, args=(job_id, paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


def _run_carga_datos_gettel_pagos_job(job_id, paths):
    """Corre en su propio hilo -- mismo patrón que _run_carga_datos_combustible_job, ver ese docstring."""
    try:
        saved_receipts = []
        duplicate_count = 0
        failed = []
        file_notes = []

        for index, path in enumerate(paths, start=1):
            filename = os.path.basename(path)
            try:
                result = gettel_pagos_parser.extract_pagos_from_pdf(path)
            except gettel_pagos_parser.PDF_READ_EXCEPTIONS as exc:
                failed.append({"filename": filename, "error": str(exc)})
                jobs.update_job(job_id, done=index, total=len(paths))
                continue

            file_saved = 0
            file_dupes = 0
            for receipt in result["receipts"]:
                existing = gettel_db.find_pago_by_transaction(receipt["transc_n"])
                if existing:
                    file_dupes += 1
                    continue
                gettel_db.add_pago(
                    receipt["fecha"], result["pago_n"], receipt["transc_n"], receipt["total_cupon"],
                    source="pdf", empresa=result["empresa"],
                )
                saved_receipts.append(receipt)
                file_saved += 1

            duplicate_count += file_dupes
            note = f"{filename}: {file_saved} cupón(es) guardado(s) (Pago N° {result['pago_n']}, {result['empresa']})"
            if file_dupes:
                note += f", {file_dupes} ya estaba(n) cargado(s)"
            if result["page_warnings"]:
                note += f" -- sin leer con confianza: {' '.join(result['page_warnings'])}"
            file_notes.append(note)
            jobs.update_job(job_id, done=index, total=len(paths))

        parts = list(file_notes)
        if failed:
            for item in failed:
                parts.append(f"{item['filename']}: {item['error']}")
        if not parts:
            parts.append("No se guardó ningún cupón de este lote.")
        level = (
            "success" if (saved_receipts and not failed and not duplicate_count)
            else ("error" if not saved_receipts else "warning")
        )

        if saved_receipts:
            d = saved_receipts[0]["fecha"]
            redirect_url = f"/carga-datos/gettel/pagos/cuadro?year={d.year}&month={d.month}"
        else:
            redirect_url = "/carga-datos/gettel/pagos"

        jobs.update_job(
            job_id, status="done", done=len(paths), total=len(paths),
            notice=" | ".join(parts), notice_level=level, redirect_url=redirect_url,
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


@app.route("/carga-datos/gettel/pagos/<int:pago_id>/editar", methods=["POST"])
def carga_datos_gettel_pagos_editar(pago_id):
    """
    Corrige a mano un campo mal leído de un cupón ya cargado por PDF -- la
    única edición que queda disponible ahora que se sacó la carga manual
    (pedido explícito del usuario, 2026-09-21).
    """
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    fecha = (request.form.get("fecha") or "").strip()
    pago_n_raw = (request.form.get("pago_n") or "").strip()
    transc_n = request.form.get("transc_n")
    empresa = request.form.get("empresa")
    total_cupon_raw = (request.form.get("total_cupon") or "").strip()
    try:
        if not fecha:
            raise ValueError("Falta la Fecha.")
        pago_n = int(pago_n_raw) if pago_n_raw else None
        total_cupon = float(total_cupon_raw)
        gettel_db.update_pago(pago_id, fecha, pago_n, transc_n, total_cupon, empresa=empresa)
        flash("Cupón corregido.", "success")
        redirect_year, redirect_month = (int(part) for part in fecha.split("-")[:2])
    except ValueError:
        flash("Revisá la Fecha, el N° de Pago y el Total del Cupón -- tienen que ser válidos.", "error")
        redirect_year, redirect_month = year, month
    return redirect(url_for("carga_datos_gettel_pagos_cuadro", year=redirect_year, month=redirect_month))


@app.route("/carga-datos/gettel/pagos/<int:pago_id>/eliminar", methods=["POST"])
def carga_datos_gettel_pagos_eliminar(pago_id):
    """El botón de eliminar vive en el Cuadro de Pagos ahora (la tabla se movió ahí), no en la página de carga."""
    gettel_db.delete_pago(pago_id)
    flash("Pago eliminado.", "success")
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    return redirect(url_for("carga_datos_gettel_pagos_cuadro", year=year, month=month))


@app.route("/carga-datos/gettel/pagos/cuadro")
def carga_datos_gettel_pagos_cuadro():
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month
    report = gettel_pagos_logic.build_month_report(year, month)
    groups = gettel_pagos_logic.grouped_pagos(report["pagos"])
    period_labels = gettel_pagos_logic.get_period_labels(year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    return render_template(
        "gettel_pagos_cuadro.html", report=report, groups=groups, year=year, month=month,
        period_labels=period_labels,
        month_name=_MONTH_NAMES_ES[month - 1], prev_year=prev_year, prev_month=prev_month,
        next_year=next_year, next_month=next_month, **THEME_BY_KEY["carga_gettel_pagos"],
    )


@app.route("/carga-datos/gettel/pagos/pendiente", methods=["POST"])
def carga_datos_gettel_pagos_pendiente():
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    pendiente_raw = (request.form.get("pendiente_anterior") or "").strip()
    try:
        pendiente_value = float(pendiente_raw) if pendiente_raw else None
        gettel_db.set_month_pendiente_anterior_override(year, month, pendiente_value)
        flash("Pendiente Mes Anterior guardado.", "success")
    except ValueError:
        flash("No se pudo guardar: el monto tiene que ser un número válido.", "error")
    return redirect(url_for("carga_datos_gettel_pagos_cuadro", year=year, month=month))


@app.route("/carga-datos/gettel/pagos/exportar/excel")
def carga_datos_gettel_pagos_exportar_excel():
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month
    report = gettel_pagos_logic.build_month_report(year, month)
    workspace_dir = tempfile.mkdtemp(prefix="gettel_pagos_export_")
    dest_path = os.path.join(workspace_dir, f"Pago Cupones {month:02d}-{year}.xlsx")
    gettel_pagos_logic.build_pagos_export_workbook(report, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/gettel/pagos/exportar/pdf")
def carga_datos_gettel_pagos_exportar_pdf():
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month
    report = gettel_pagos_logic.build_month_report(year, month)
    workspace_dir = tempfile.mkdtemp(prefix="gettel_pagos_export_pdf_")
    dest_path = os.path.join(workspace_dir, f"Pago Cupones {month:02d}-{year}.pdf")
    gettel_pagos_logic.build_pagos_export_pdf(report, year, month, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/carga-datos/gettel/historial")
def carga_datos_gettel_historial():
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    days_by_date = {d["date"]: d for d in gettel_db.get_month_days(year, month)}
    # "Local Account" -- la fila que el Excel real cruza contra Gettel+Toyota
    # (DIF = Local Account - Gettel - Toyota). CORREGIDO 2026-09-15: se
    # confirmó contra el Excel real (Gettel-Toyota 08.2026, columna "Local
    # Account") que esta cifra es el campo CHICO de Store Info (Method of
    # Payment Totals, reportes_db.get_month_local_accounts) -- coincide
    # EXACTO día por día (ej. 523.63 el 07/08, 812.29 el 03/08). El
    # departamento "LOCAL ACCT" del Department Sales Report (usado antes)
    # era la fuente equivocada -- solo aparece esporádicamente (cargos
    # puntuales grandes, no una cifra diaria) y nunca coincide con la
    # columna real -- por eso la mayoría de los días quedaban sin Local
    # Account y el DIF de los días que sí tenían ese departamento daba un
    # número gigante y sin sentido (pedido explícito del usuario, aclaró
    # que las ventas de Gettel que "aparecían mal" eran ese DIF fantasma).
    local_account_by_date = reportes_db.get_month_local_accounts(year, month)
    _blank_gettel_day = {
        "date": None, "gettel_amount": None, "gettel_gallons": None,
        "toyota_amount": None, "toyota_gallons": None,
    }
    for day_key in local_account_by_date:
        if day_key not in days_by_date:
            days_by_date[day_key] = {**_blank_gettel_day, "date": day_key}

    days = []
    for day_key in sorted(days_by_date):
        day = dict(days_by_date[day_key])
        local_account = local_account_by_date.get(day_key)
        day["local_account"] = local_account
        if local_account is not None:
            day["dif"] = round(local_account - (day.get("gettel_amount") or 0.0) - (day.get("toyota_amount") or 0.0), 2)
        else:
            day["dif"] = None
        days.append(day)

    totals = {
        "gettel_amount": sum(d.get("gettel_amount") or 0.0 for d in days),
        "gettel_gallons": sum(d.get("gettel_gallons") or 0.0 for d in days),
        "toyota_amount": sum(d.get("toyota_amount") or 0.0 for d in days),
        "toyota_gallons": sum(d.get("toyota_gallons") or 0.0 for d in days),
        "local_account": sum(v or 0.0 for v in local_account_by_date.values()),
        # Diferencia total que quedó en el mes -- pedido explícito del
        # usuario (2026-09-19), suma de los DIF diarios (días sin Local
        # Account, dif=None, no aportan nada -- no hay nada que sumar ahí).
        "dif": round(sum(d["dif"] for d in days if d.get("dif") is not None), 2),
    }
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "carga_datos_gettel_historial.html",
        days=days,
        totals=totals,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["carga_gettel"],
    )


# ---------------------------------------------------------------------------
# Horas de Trabajo -- Carga de Datos (2026-09-18, pedido explícito del
# usuario): el reporte semanal "Clock In/Out Detail Report" (Chevron POS,
# uno por lunes) que antes se transcribía a mano al Excel "BDT. HOURS..."
# ahora se lee solo -- una semana = un bloque, igual que EFT. Horas/tarifa
# ($15/hora por default)/descuento quedan editables por empleado; el monto a
# pagar se calcula siempre en el momento (horas*tarifa−descuento), nunca se
# guarda ya calculado -- mismo criterio que DIF EFECT/Saldo en Caja.
# ---------------------------------------------------------------------------

@app.route("/carga-datos/horas-trabajo")
def carga_datos_horas_trabajo():
    _active_job = jobs.get_active_job("horas_trabajo")
    return render_template(
        "carga_datos_horas_trabajo.html",
        resume_job_id=(_active_job["id"] if _active_job else None),
        **THEME_BY_KEY["carga_horas"],
    )


@app.route("/carga-datos/horas-trabajo/subir", methods=["POST"])
def carga_datos_horas_trabajo_subir():
    uploads = request.files.getlist("report_files")
    if not uploads or not any(u.filename for u in uploads):
        return _error_response("Seleccioná uno o más PDF de Clock In/Out.")

    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths), kind="horas_trabajo")
    threading.Thread(target=_run_carga_datos_horas_trabajo_job, args=(job_id, paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


def _run_carga_datos_horas_trabajo_job(job_id, paths):
    """Corre en su propio hilo -- mismo patrón que _run_carga_datos_gettel_job."""
    try:
        weeks_saved = 0
        files_failed = 0
        unresolved_total = []
        first_date = None

        for index, path in enumerate(paths, start=1):
            try:
                data = extract_hours_report(path, known_names=horas_trabajo_db.known_employee_names())
                if data["report_date"] is None:
                    raise ValueError('No se pudo leer la fecha de "REPORT PRINTED".')
                if not data["employees"] and not data["unresolved_employees"]:
                    raise ValueError("No se pudo leer ningún empleado de este reporte.")

                document_id = None
                try:
                    document_id = documents_db.store_document(
                        "horas_trabajo", path, os.path.basename(path),
                        data["report_date"].year, data["report_date"].month,
                        label="Reporte semanal",
                    )
                except Exception as exc:
                    print(f"[documents_db] no se pudo guardar el reporte de Horas de Trabajo {path}: {exc}")

                horas_trabajo_db.upsert_week(
                    data["report_date"], data["period_from"], data["period_to"],
                    data["employees"], source="ocr", document_id=document_id,
                )
                weeks_saved += 1
                unresolved_total.extend((data["report_date"], name) for name in data["unresolved_employees"])
                if first_date is None or data["report_date"] < first_date:
                    first_date = data["report_date"]
            except Exception as exc:
                print(f"[carga-datos/horas-trabajo] {path}: {exc}")
                files_failed += 1
            jobs.update_job(job_id, done=index, total=len(paths))

        parts = []
        if weeks_saved:
            parts.append(f"{weeks_saved} semana(s) guardada(s).")
        # Un empleado por renglón (como el aviso de Proveedores): quién quedó
        # sin horas y por qué -- se agrega a mano desde el cuadro de esa semana.
        parts.extend(f"Reporte del {when:%d/%m/%Y}: cargá a mano las horas de {name}" for when, name in unresolved_total)
        if files_failed:
            parts.append(f"{files_failed} archivo(s) no se pudieron leer.")

        if not parts:
            notice, level = "No se pudo guardar nada de este lote.", "error"
        else:
            notice, level = "\n".join(parts), ("warning" if (files_failed or unresolved_total) else "success")

        if first_date:
            redirect_url = f"/carga-datos/horas-trabajo/historial?year={first_date.year}&month={first_date.month}"
        else:
            redirect_url = "/carga-datos/horas-trabajo"

        jobs.update_job(
            job_id, status="done", done=len(paths), total=len(paths),
            notice=notice, notice_level=level, redirect_url=redirect_url,
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


@app.route("/carga-datos/horas-trabajo/historial")
def carga_datos_horas_trabajo_historial():
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    entries = []
    for item in horas_trabajo_db.get_month_weeks(year, month):
        week = item["week"]
        employees = []
        week_total = 0.0
        for emp in item["employees"]:
            total_pay = round((emp["hours"] or 0.0) * (emp["rate"] or 0.0) - (emp["deduct"] or 0.0), 2)
            # HH:MM que se muestra al lado de las horas, y aviso si no coincide
            # con las horas que se pagan (filas viejas corregidas a mano en
            # formato H.MM, ej. 39:12 guardado como 39.12).
            label = emp.get("hours_label")
            label_mismatch = False
            if label:
                try:
                    label_hours, _ = horas_trabajo.parse_hours_input(label)
                    label_mismatch = abs(label_hours - (emp["hours"] or 0.0)) > 0.005
                except ValueError:
                    label = None
            employees.append({
                **emp,
                "total_pay": total_pay,
                "hours_display": horas_trabajo.hours_to_label(emp["hours"]),
                "report_label": label,
                "label_mismatch": label_mismatch,
            })
            week_total += total_pay
        document = documents_db.get_document(week["document_id"]) if week.get("document_id") else None
        entries.append({"week": week, "employees": employees, "week_total": round(week_total, 2), "document": document})

    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "carga_datos_horas_trabajo_historial.html",
        entries=entries,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["carga_horas"],
    )


@app.route("/carga-datos/horas-trabajo/empleado/<int:employee_id>/editar", methods=["POST"])
def carga_datos_horas_trabajo_empleado_editar(employee_id):
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    name = (request.form.get("employee_name") or "").strip()
    try:
        hours, hours_label = horas_trabajo.parse_hours_input(request.form.get("hours") or "0")
        rate = float((request.form.get("rate") or "0").strip())
        deduct = float((request.form.get("deduct") or "0").strip())
    except ValueError:
        flash("No se pudo guardar: las horas van como 26:48 (o 26.8) y tarifa/descuento como números.", "error")
        return redirect(url_for("carga_datos_horas_trabajo_historial", year=year, month=month))

    horas_trabajo_db.update_employee(
        employee_id, employee_name=name or None, hours=hours, rate=rate, deduct=deduct, hours_label=hours_label
    )
    flash("Empleado actualizado.", "success")
    return redirect(url_for("carga_datos_horas_trabajo_historial", year=year, month=month))


@app.route("/carga-datos/horas-trabajo/semana/<int:week_id>/empleado", methods=["POST"])
def carga_datos_horas_trabajo_empleado_agregar(week_id):
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    name = (request.form.get("employee_name") or "").strip()
    if not name:
        flash("Ingresá el nombre del empleado.", "error")
        return redirect(url_for("carga_datos_horas_trabajo_historial", year=year, month=month))

    try:
        hours, hours_label = horas_trabajo.parse_hours_input(request.form.get("hours") or "0")
        rate_raw = (request.form.get("rate") or "").strip()
        rate = float(rate_raw) if rate_raw else horas_trabajo_db.DEFAULT_HOURLY_RATE
        deduct = float((request.form.get("deduct") or "0").strip())
    except ValueError:
        flash("No se pudo agregar: las horas van como 26:48 (o 26.8) y tarifa/descuento como números.", "error")
        return redirect(url_for("carga_datos_horas_trabajo_historial", year=year, month=month))

    horas_trabajo_db.add_employee(week_id, name, hours=hours, rate=rate, deduct=deduct, hours_label=hours_label)
    flash("Empleado agregado.", "success")
    return redirect(url_for("carga_datos_horas_trabajo_historial", year=year, month=month))


@app.route("/carga-datos/horas-trabajo/empleado/<int:employee_id>/eliminar", methods=["POST"])
def carga_datos_horas_trabajo_empleado_eliminar(employee_id):
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    horas_trabajo_db.delete_employee(employee_id)
    flash("Empleado eliminado.", "success")
    return redirect(url_for("carga_datos_horas_trabajo_historial", year=year, month=month))


@app.route("/carga-datos/horas-trabajo/semana/<int:week_id>/eliminar", methods=["POST"])
def carga_datos_horas_trabajo_eliminar(week_id):
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    horas_trabajo_db.delete_week(week_id)
    flash("Semana eliminada. Los documentos guardados no se borran.", "success")
    return redirect(url_for("carga_datos_horas_trabajo_historial", year=year, month=month))


@app.route("/carga-datos/cmv")
def carga_datos_cmv():
    """
    CMV -- Carga de Datos (2026-09-12, cuarta tanda): Costo por UPC y
    Ventas mensuales por departamento, guardados directo en cmv_db, sin
    generar ningún Excel -- pedido explícito del usuario: "quiero que
    empieces a crear los modulos de lo que seria CMV donde cargariamos el
    costo de los productos como en el excel, y tambien las ventas
    mensuales que iran a cada departamento".
    """
    today = date.today()
    return render_template(
        "carga_datos_cmv.html",
        current_year=today.year,
        current_month=today.month,
        month_names=_MONTH_NAMES_ES,
        **THEME_BY_KEY["carga_cmv"],
    )


@app.route("/carga-datos/cmv/costo/subir", methods=["POST"])
def carga_datos_cmv_costo_subir():
    uploads = request.files.getlist("costo_files")
    if not uploads or not any(u.filename for u in uploads):
        flash("Seleccioná uno o más archivos de costo por departamento.", "error")
        return redirect(url_for("carga_datos_cmv"))
    # Fecha del CMV (default hoy): cada carga queda como una foto con esa
    # fecha, para compararla contra la anterior -- ver cmv_db.save_snapshot.
    try:
        cmv_date = datetime.strptime(request.form.get("cmv_date") or "", "%Y-%m-%d").date()
    except ValueError:
        cmv_date = date.today()

    paths = _save_uploads_to_workspace(uploads)
    try:
        combined, file_stats, failed_files = _consolidate_department_files(paths)
    except ValueError as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("carga_datos_cmv"))

    summary = cmv_db.replace_costs_for_departments(combined.to_dict("records"))
    cmv_db.save_snapshot(cmv_date)
    previous = next((s for s in cmv_db.list_snapshots() if s["loaded_on"] < cmv_date.isoformat()), None)

    today = date.today()
    for path in paths:
        try:
            documents_db.store_document("cmv_costo", path, os.path.basename(path), today.year, today.month)
        except Exception as exc:
            print(f"[documents_db] no se pudo guardar el documento de CMV Costo {path}: {exc}")

    parts = [f"{summary['departments']} departamento(s), {summary['rows']} producto(s) guardados."]
    if summary["price_changes"]:
        parts.append(f"{summary['price_changes']} cambio(s) de precio detectado(s).")
    if summary["partial"]:
        parts.append(
            f"Parece una carga parcial (trae menos del 80% de los {summary['stored_before']} productos ya "
            "guardados): se actualizaron estos productos y no se borró ninguno. Para el CMV completo, subí todas las páginas juntas."
        )
    if failed_files:
        parts.append(f"{failed_files} archivo(s) no se pudieron leer.")
    if previous:
        diff = compare_cost_snapshots(
            cmv_db.get_snapshot_items(previous["loaded_on"]), cmv_db.get_snapshot_items(cmv_date.isoformat())
        )
        parts.append(
            f"Contra el CMV del {_fmt_ddmmyyyy(previous['loaded_on'])}: {diff['cost_changes']} cambio(s) de costo, "
            f"{diff['price_changes']} de precio, {len(diff['added'])} producto(s) nuevo(s) y "
            f"{len(diff['removed'])} que ya no están."
        )
    flash(" ".join(parts), "warning" if (failed_files or summary["partial"]) else "success")
    if previous:
        return redirect(url_for("carga_datos_cmv_costo_comparar", fecha=cmv_date.isoformat(), contra=previous["loaded_on"]))
    return redirect(url_for("carga_datos_cmv_costo_historial"))


@app.route("/carga-datos/cmv/costo/comparar")
def carga_datos_cmv_costo_comparar():
    """
    Compara dos fotos del CMV por fecha (default: la última contra la
    anterior) -- pedido explícito del usuario (2026-09-28): "ver qué tanto
    cambió con los días cada vez que se quiera subir un CMV".
    """
    snapshots = cmv_db.list_snapshots()
    dates = [s["loaded_on"] for s in snapshots]
    fecha = request.args.get("fecha") if request.args.get("fecha") in dates else (dates[0] if dates else None)
    older = [d for d in dates if fecha and d < fecha]
    contra = request.args.get("contra") if request.args.get("contra") in older else (older[0] if older else None)
    diff = None
    if fecha and contra:
        diff = compare_cost_snapshots(cmv_db.get_snapshot_items(contra), cmv_db.get_snapshot_items(fecha))
    return render_template(
        "carga_datos_cmv_costo_comparar.html",
        snapshots=snapshots,
        fecha=fecha,
        contra=contra,
        older=older,
        diff=diff,
        fmt_date=_fmt_ddmmyyyy,
        **THEME_BY_KEY["carga_cmv"],
    )


_CMV_ALL_DEPARTMENTS = "__todos__"


@app.route("/carga-datos/cmv/costo/historial")
def carga_datos_cmv_costo_historial():
    departments = cmv_db.list_departments()
    selected = request.args.get("dept") or (departments[0]["dept_name"] if departments else None)
    # "Todos los productos" (pedido del usuario, 2026-10-02): el catálogo
    # entero en una tabla, ordenado por departamento y nombre.
    show_all = selected == _CMV_ALL_DEPARTMENTS
    if show_all:
        rows = sorted(cmv_db.get_all_costs(), key=lambda r: ((r.get("dept_name") or "").lower(), (r.get("name") or "").lower()))
    else:
        rows = cmv_db.get_costs_by_department(selected) if selected else []
    return render_template(
        "carga_datos_cmv_costo_historial.html",
        departments=departments,
        selected=selected,
        show_all=show_all,
        all_value=_CMV_ALL_DEPARTMENTS,
        total_products=sum(d["count"] for d in departments),
        rows=rows,
        price_changes=cmv_db.get_recent_price_changes(),
        **THEME_BY_KEY["carga_cmv"],
    )


@app.route("/carga-datos/cmv/ventas/subir", methods=["POST"])
def carga_datos_cmv_ventas_subir():
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    uploads = request.files.getlist("ventas_files")
    if not year or not month or not (1 <= month <= 12):
        flash("Elegí a qué mes corresponden estas ventas.", "error")
        return redirect(url_for("carga_datos_cmv"))
    if not uploads or not any(u.filename for u in uploads):
        flash("Seleccioná uno o más archivos de ventas mensuales.", "error")
        return redirect(url_for("carga_datos_cmv"))

    paths = _save_uploads_to_workspace(uploads)

    # Un mismo departamento crudo del POS que llega en dos archivos del lote
    # (el mismo export subido dos veces, "... (2).csv") se toma una sola
    # vez: queda el último archivo y se avisa (auditoría 2026-09).
    rows_by_raw_dept = {}
    repeated_departments = set()
    unmapped_departments = set()
    files_failed = 0
    for file_index, path in enumerate(paths):
        try:
            frame = parse_monthly_sales_file(path)
        except Exception as exc:
            print(f"[carga-datos/cmv/ventas] {path}: {exc}")
            files_failed += 1
            continue
        try:
            documents_db.store_document("cmv_ventas", path, os.path.basename(path), year, month)
        except Exception as exc:
            print(f"[documents_db] no se pudo guardar el documento de CMV Ventas {path}: {exc}")
        file_rows = {}
        for record in frame.to_dict("records"):
            raw_dept = (record.get("Dept Name") or "").strip()
            dept_name = _resolve_sheet_name(record.get("Dept Name"))
            if dept_name is None:
                unmapped_departments.add(raw_dept or "(sin nombre)")
                continue
            file_rows.setdefault(raw_dept.upper(), (dept_name, []))[1].append(
                {
                    "upc": record.get("UPC"),
                    "name": record.get("Name"),
                    "count": record.get("Count"),
                    "amount": record.get("Retail/Amount"),
                }
            )
        for raw_key, value in file_rows.items():
            if raw_key in rows_by_raw_dept:
                repeated_departments.add(raw_key)
            rows_by_raw_dept[raw_key] = value

    rows_by_dept = {}
    for dept_name, rows in rows_by_raw_dept.values():
        rows_by_dept.setdefault(dept_name, []).extend(rows)

    for dept_name, rows in rows_by_dept.items():
        try:
            cmv_db.replace_month_department_sales(year, month, dept_name, rows)
        except sqlite3.Error as exc:
            print(f"[carga-datos/cmv/ventas] {dept_name}: {exc}")
            files_failed += 1

    parts = []
    if rows_by_dept:
        parts.append(
            f"{len(rows_by_dept)} departamento(s), "
            f"{sum(len(r) for r in rows_by_dept.values())} producto(s) guardados."
        )
    if unmapped_departments:
        parts.append(f"{len(unmapped_departments)} departamento(s) sin hoja conocida, no se guardaron.")
    if repeated_departments:
        parts.append(
            f"{len(repeated_departments)} departamento(s) venían en más de un archivo: se tomó solo el último."
        )
    if files_failed:
        parts.append(f"{files_failed} archivo(s) no se pudieron leer.")
    if not parts:
        flash("No se pudo guardar nada de este lote.", "error")
    else:
        flash(" ".join(parts), "warning" if (unmapped_departments or files_failed or repeated_departments) else "success")

    return redirect(url_for("carga_datos_cmv_ventas_historial", year=year, month=month))


@app.route("/carga-datos/cmv/ventas/historial")
def carga_datos_cmv_ventas_historial():
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    totals = cmv_db.get_month_department_totals(year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "carga_datos_cmv_ventas_historial.html",
        totals=totals,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["carga_cmv"],
    )


# Sección "Documentos" -- tercera opción del Menú (pedido explícito del
# usuario, 2026-09-23): "dentro de el varios modulos en los que solo se
# contengan los documentos fisicos tipo Pdf o excel". Cada tarjeta abre la
# lista de archivos originales de un módulo -- los mismos que ya se guardaban
# al cargar (documents_db, y los PDF diarios de Reporte Diario/Lottery en sus
# propias bases), ahora juntos en un solo lugar en vez de repartidos por la
# barra lateral. Orden = el de la barra lateral.
DOCUMENTOS_SECTIONS = [
    {"label": "Reporte Diario", "endpoint": "reporte_documentos", "theme": "reporte", "icon": _ICON_CALENDAR,
     "description": "PDF de cierre diario, uno por día, y el resumen mensual."},
    {"label": "Chase Bank", "module": "chase", "theme": "carga_chase", "icon": _ICON_BANK,
     "description": "Extractos y comprobantes del banco."},
    {"label": "EFT y Cupones", "module": "eft", "theme": "carga_eft", "icon": _ICON_EXCHANGE,
     "description": "PDF de cada EFT, reportes de Cupones y facturas de J.H. Williams."},
    {"label": "Caja", "module": "caja", "theme": "carga_caja", "icon": _ICON_REGISTER,
     "description": "Comprobantes de depósito y planillas de gastos en efectivo."},
    {"label": "Gettel / Toyota", "module": "gettel_toyota", "theme": "carga_gettel", "icon": _ICON_CAR,
     "description": "Excel y PDF de cupones de combustible de los vendedores."},
    {"label": "CMV — Costo", "module": "cmv_costo", "theme": "carga_cmv", "icon": _ICON_COINS,
     "description": "Archivos de costo por departamento."},
    {"label": "CMV — Ventas", "module": "cmv_ventas", "theme": "carga_cmv", "icon": _ICON_COINS,
     "description": "Archivos de ventas mensuales por departamento."},
    {"label": "Lottery", "endpoint": "carga_datos_lottery_documentos", "theme": "lottery", "icon": _ICON_TICKET,
     "description": "PDF diarios (Department y Daily Sales Report) y el resumen mensual."},
    {"label": "Proveedores", "module": "proveedores", "theme": "carga_proveedores", "icon": _ICON_TRUCK,
     "description": "Facturas de proveedores, en una carpeta por proveedor."},
    {"label": "Horas de Trabajo", "module": "horas_trabajo", "theme": "carga_horas", "icon": _ICON_CLOCK,
     "description": "Reportes semanales de horas y comprobantes de pago."},
    {"label": "Combustible", "module": "combustible", "theme": "fisico", "icon": _ICON_FUEL,
     "description": "Facturas de combustible."},
]


@app.route("/documentos")
def documentos_index():
    sections = []
    for item in DOCUMENTOS_SECTIONS:
        if "module" in item:
            url = url_for("carga_datos_documentos", module_key=item["module"])
        else:
            url = url_for(item["endpoint"])
        sections.append({**item, **THEME_BY_KEY[item["theme"]], "url": url})
    return render_template("documentos_index.html", sections=sections)


# ---------------------------------------------------------------------------
# Documentos -> Depósitos (pedido explícito del usuario, 2026-09-23): cada
# recibo de depósito es una fila propia (un PDF con varios recibos se divide
# en una fila por página, cada una con su PDF), con fecha, monto y
# descripción leídos por OCR y editables. Ver depositos.py/depositos_db.py.
# ---------------------------------------------------------------------------

_CHASE_DEPOSIT_KINDS = {"FOOD TRUCK": depositos.FOOD_TRUCK, "DEPOSITO VENTA ICE": depositos.ICE_MACHINE,
                        CHASE_DETALLE_VACCUMMS: control_depositos.VACCUMMS}


def _chase_deposits(year, month):
    """Depósitos (monto positivo, descripción DEPOSIT) de Chase en ese mes."""
    return [
        r for r in chase_db.get_month_transactions(year, month)
        if (r.get("amount") or 0) > 0 and "DEPOSIT" in (r.get("description") or "").upper()
    ]


def _chase_kind_for(deposit_date, amount):
    """Food Truck / Ice Machine según cómo quedó categorizado ese depósito en Chase (si ya está cargado)."""
    if deposit_date is None or amount is None:
        return None
    for r in _chase_deposits(deposit_date.year, deposit_date.month):
        if r["posting_date"] == deposit_date.isoformat() and abs(r["amount"] - amount) < 0.005:
            kind = _CHASE_DEPOSIT_KINDS.get((r.get("detalle") or "").upper())
            if kind:
                return kind
    return None


def _card_detail_by_day(year, month, detail_groups, today=None):
    """Lo vendido con tarjeta por día contra los batches del POS del detalle de cupones de J.H. (o None)."""
    covered = sorted(g["last_date"] for g in detail_groups)
    month_first = date(year, month, 1)
    return control_tarjetas.build_detail_by_day(
        year, month, reportes_db.get_card_sales_by_date(),
        eft_db.get_detail_batches_between(
            (month_first - timedelta(days=10)).isoformat(),
            (month_first + timedelta(days=45)).isoformat(),
        ),
        covered[0] if covered else None, covered[-1] if covered else None, today=today or date.today(),
    )


@app.route("/controles/tarjetas")
def controles_tarjetas():
    """
    Control Tarjetas y Cupones (pedido del usuario, 2026-10-02): lo vendido
    con tarjeta en el C-store contra los cupones que acredita JH. Lo que
    todavía no se acreditó (unos 3 días por las 72 hs) no debería pasar de
    $15,000. Cálculo en control_tarjetas.py.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month
    control = control_tarjetas.build_month_control(
        year, month, reportes_db.get_card_sales_by_date(), eft_db.get_coupon_gross_by_date(), today=today,
    )
    for row in control["rows"]:
        row["date_display"] = _fmt_ddmmyyyy(row["date"])
    detail_groups = eft_db.get_detail_groups()
    detail = _card_detail_by_day(year, month, detail_groups, today)
    if detail:
        for row in detail["rows"]:
            row["date_display"] = _fmt_ddmmyyyy(row["date"])
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    return render_template(
        "controles_tarjetas.html",
        control=control,
        missing_days_display=[_fmt_ddmmyyyy(d) for d in control["missing_days"]],
        detail=detail,
        jh=_jh_month_checks(year, month, detail_groups),
        jh_labels=jh_mensual.KIND_LABELS,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["carga_eft"],
    )


@app.route("/controles/lottery")
def controles_lottery():
    """
    Control Lottery (pedido del usuario, 2026-10-06, chat 21): el Monthly
    Sales Report del portal contra la suma de los reportes diarios del mes, y
    el Debito de cada bloque semanal que se paga en el mes contra el pago en
    Chase. El asiento del mes queda pendiente (el usuario lo pasa después).
    Cálculo en lottery_mensual.py.
    """
    year, month = _cierre_month()
    report = lottery_db.get_monthly_report(year, month)
    section = next(s for s in CONTROLES_SECTIONS if s["key"] == "control_lottery")
    return render_template(
        "controles_lottery.html",
        report=report,
        cross=lottery_mensual.cross_check(report, year, month) if report else None,
        payments=lottery_mensual.payments_check(year, month),
        **_month_nav(year, month),
        accent=section["accent"],
        accent_soft=section["accent_soft"],
    )


def _cierre_month():
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month
    return year, month


def _cierre_period_label(entries):
    first, last = entries["first_day"], entries["last_day"]
    return f"Período: {_fmt_ddmmyyyy(first)} al {_fmt_ddmmyyyy(last)}"


def _monthly_cross(year, month, entries):
    """(reporte mensual guardado, cruce contra el asiento), o None donde falte."""
    report = reporte_mensual_db.get_report(year, month)
    if not report or not entries:
        return report, None
    return report, control_cierre.cross_check(entries, reporte_mensual.report_totals(report))


@app.route("/controles/cierre")
def controles_cierre():
    """
    Control Cierre (pedido del usuario, 2026-10-06): los dos asientos de
    cierre del mes de la hoja Store info del Excel Cierre, con los totales
    ya guardados. Cálculo en control_cierre.py.
    """
    year, month = _cierre_month()
    entries = control_cierre.build_month_entries(_build_store_info_rows(year, month), year, month)
    monthly_report, cross = _monthly_cross(year, month, entries)
    jh_cards = None
    if entries:
        # Las tarjetas del asiento contra J.H. (pedido del usuario, 2026-10-06).
        detail_groups = eft_db.get_detail_groups()
        jh_cards = control_cierre.jh_cards_check(
            entries, _card_detail_by_day(year, month, detail_groups),
            _jh_month_checks(year, month, detail_groups)["coupons"],
        )
    section = next(s for s in CONTROLES_SECTIONS if s["key"] == "control_cierre")
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    return render_template(
        "controles_cierre.html",
        entries=entries,
        monthly_report=monthly_report,
        cross=cross,
        jh_cards=jh_cards,
        fmt_ddmm=lambda iso: f"{iso[8:10]}/{iso[5:7]}",
        fmt_days=lambda days: ", ".join(f"{iso[8:10]}/{iso[5:7]}" for iso in days),
        notice=control_cierre.missing_notice(entries) if entries else None,
        period_label=_cierre_period_label(entries) if entries else None,
        entry_title=control_cierre.ENTRY_TITLE,
        account_jh=control_cierre.ACCOUNT_JH,
        account_gettel=control_cierre.ACCOUNT_GETTEL,
        lottery_note=control_cierre.LOTTERY_NOTE,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        accent=section["accent"],
        accent_soft=section["accent_soft"],
    )


@app.route("/controles/cierre/exportar/pdf")
def controles_cierre_pdf():
    """PDF de los dos asientos del mes (control_cierre.build_entries_pdf)."""
    year, month = _cierre_month()
    entries = control_cierre.build_month_entries(_build_store_info_rows(year, month), year, month)
    if not entries:
        flash("No hay ningún Store Info guardado ese mes: no hay asientos para exportar.", "error")
        return redirect(url_for("controles_cierre", year=year, month=month))
    dest_path = os.path.join(tempfile.mkdtemp(prefix="cierre_pdf_"), f"Asientos de cierre {month:02d}-{year}.pdf")
    control_cierre.build_entries_pdf(
        entries, f"Asientos de cierre — {_MONTH_NAMES_ES[month - 1]} {year}", _cierre_period_label(entries), dest_path,
        cross=_monthly_cross(year, month, entries)[1],
    )
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/controles/cierre/exportar")
def controles_cierre_excel():
    """
    El mismo Excel de Store Info de Reportes (build_store_info_export_
    workbook), más la fila de totales del mes y los dos asientos debajo.
    """
    year, month = _cierre_month()
    store_info_rows = _build_store_info_rows(year, month)
    entries = control_cierre.build_month_entries(store_info_rows, year, month)
    if not entries:
        flash("No hay ningún Store Info guardado ese mes para exportar.", "error")
        return redirect(url_for("controles_cierre", year=year, month=month))
    dest_path = os.path.join(tempfile.mkdtemp(prefix="cierre_excel_"), f"Store Info {month:02d}-{year} con asientos.xlsx")
    build_store_info_export_workbook(store_info_rows, year, month, dest_path)
    control_cierre.add_entries_to_store_info_workbook(dest_path, entries)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/controles/rapidos/<kind>")
def controles_rapidos_panel(kind):
    """
    Cuadro de la botonera de controles rápidos de abajo a la derecha
    (base.html, pedido del jefe, 2026-10-05): un resumen chico de Tarjetas o
    Caja que el navegador pide al tocar el botón y mete en el cuadro. Arranca
    en el mes actual; ?year=&month= pasa a otro mes (pedido del usuario,
    2026-10-07). Cálculo en controles_rapidos.py.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not 1 <= month <= 12:
        abort(404)
    day = controles_rapidos.reference_day(year, month, today)
    if day is None:  # mes futuro: el actual
        year, month, day = today.year, today.month, today
    if kind == "tarjetas":
        data = controles_rapidos.tarjetas_status(
            reportes_db.get_card_sales_by_date(), eft_db.get_coupon_gross_by_date(), today=day,
        )
    elif kind == "caja":
        data = controles_rapidos.caja_status(today=day)
    elif kind == "precios":
        pos_costs = cmv_db.get_all_costs()
        snapshots = cmv_db.list_snapshots()
        data = {
            "pos_loaded": bool(pos_costs),
            "cmv_date": snapshots[0]["loaded_on"] if snapshots else None,
            "suppliers": controles_rapidos.price_review(
                proveedores_db.get_all_invoice_lines(), pos_costs, _supplier_labels(), today=day,
            ),
        }
    else:
        abort(404)
    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    next_year, next_month = (year + 1, 1) if month == 12 else (year, month + 1)
    return render_template(
        "_controles_rapidos.html", kind=kind, data=data,
        month_label=f"{_MONTH_NAMES_ES[month - 1]} {year}",
        is_current=(year, month) == (today.year, today.month),
        this_month=f"{year}-{month}",
        prev_month=f"{prev_year}-{prev_month}",
        next_month=None if (year, month) == (today.year, today.month) else f"{next_year}-{next_month}",
        caja_limit=controles_rapidos.caja_limit(),
    )


@app.route("/controles/rapidos/caja/limite", methods=["POST"])
def controles_rapidos_caja_limite():
    """Límite de la alerta de Caja de la botonera (pedido del usuario, 2026-10-07)."""
    try:
        controles_rapidos.set_caja_limit(reporte_mensual.parse_amount(request.form.get("limit")))
    except ValueError as exc:
        return jsonify({"error": str(exc) if "límite" in str(exc) else "El límite no es un monto."}), 400
    return jsonify({"ok": True})


@app.route("/controles/rapidos/alertas")
def controles_rapidos_alertas():
    """Qué botón de la botonera marcar en rojo (mes actual), sin abrir el cuadro."""
    today = date.today()
    tarjetas = controles_rapidos.tarjetas_status(
        reportes_db.get_card_sales_by_date(), eft_db.get_coupon_gross_by_date(), today=today,
    )
    caja_data = controles_rapidos.caja_status(today=today)
    return jsonify({
        "tarjetas": bool(tarjetas and tarjetas["last"]["status"] == "alert"),
        "caja": bool(caja_data and caja_data["alert"]),
    })


# ---------------------------------------------------------------------------
# Depósitos y Control Depósitos (pedido del usuario, 2026-10-07): en Carga de
# Datos -> Depósitos se suben todos los PDF de depósitos del cajero (normales,
# Food Truck, Ice Machine, Vaccumms; cada recibo es una fila de depositos_db,
# corregible a mano) y los Payment Summary de la máquina de hielo
# (ice_machine.py). Controles -> Control Depósitos lo cruza todo con Chase
# (control_depositos.py).
# ---------------------------------------------------------------------------

@app.route("/carga-datos/depositos")
def carga_datos_depositos():
    year, month = _cierre_month()
    rows = depositos_db.list_month(year, month)
    for row in rows:
        row["group"] = control_depositos.receipt_kind(row.get("kind"))
        row["date_display"] = (
            f"{row['deposit_date'][8:10]}/{row['deposit_date'][5:7]}/{row['deposit_date'][:4]}"
            if row["deposit_date"] else None
        )
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    _active_job = jobs.get_active_job("depositos")
    return render_template(
        "carga_datos_depositos.html",
        rows=rows,
        total=round(sum(r["amount"] or 0 for r in rows), 2),
        summaries=ice_machine_db.list_month(year, month),
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        resume_job_id=(_active_job["id"] if _active_job else None),
        **THEME_BY_KEY["carga_depositos"],
    )


@app.route("/carga-datos/ice-food-truck")
def carga_datos_ice():
    return redirect(url_for("carga_datos_depositos", **request.args.to_dict()))


def _depositos_month_control(year, month):
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    # Un resumen que termina a fin de mes (o en un feriado) se acredita los
    # primeros días del mes siguiente: el cruce mira también esa semana.
    chase_rows = chase_db.get_month_transactions(year, month) + [
        r for r in chase_db.get_month_transactions(next_year, next_month) if r["posting_date"][8:10] <= "07"
    ]
    return control_depositos.build_month_control(
        year, month, ice_machine_db.list_month(year, month), depositos_db.list_month(year, month),
        chase_rows, chase_db.get_last_posting_date(),
        known_references=[s["reference"] for s in ice_machine_db.list_month(prev_year, prev_month)],
    )


@app.route("/controles/depositos")
def controles_depositos():
    year, month = _cierre_month()
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    return render_template(
        "controles_depositos.html",
        control=_depositos_month_control(year, month),
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["carga_depositos"],
    )


# ---------------------------------------------------------------------------
# Control CMV (pedido del usuario, 2026-10-07): las ventas de cada
# departamento del mes en los reportes diarios, el reporte mensual del POS,
# la página de Elistar (Depts Report o P & L, se suben acá) y las ventas
# cargadas en CMV. Cálculo en control_cmv.py.
# ---------------------------------------------------------------------------

def _cmv_month_control(year, month):
    reports = control_cmv_db.get_month(year, month)
    elistar = reports.get("depts") or reports.get("pl")
    monthly = reporte_mensual_db.get_report(year, month)
    fuel_by_date = {
        r["date"]: round((r["sales_fuel"] or 0.0) + (r["desc_comb"] or 0.0), 2)
        for r in reportes_db.get_month_store_info(year, month) if r["sales_fuel"] is not None
    }
    fuel_monthly = None
    if monthly and monthly["store_info"].get("sales_fuel") is not None:
        fuel_monthly = round(monthly["store_info"]["sales_fuel"] + (monthly["store_info"].get("desc_comb") or 0.0), 2)
    control = control_cmv.build_control(
        reportes_db.get_month_department_totals(year, month),
        monthly["departments"] if monthly and monthly.get("departments") else None,
        cmv_db.get_month_department_totals(year, month),
        elistar,
        reportes_db.get_month_departments_by_date(year, month),
        fuel_by_date,
        fuel_monthly,
    )
    return control, reports, elistar


@app.route("/controles/cmv")
def controles_cmv():
    year, month = _cierre_month()
    control, reports, elistar = _cmv_month_control(year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    return render_template(
        "controles_cmv.html",
        control=control,
        reports=reports,
        elistar=elistar,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["carga_cmv"],
    )


@app.route("/controles/cmv/subir", methods=["POST"])
def controles_cmv_subir():
    uploads = [u for u in request.files.getlist("files") if u and u.filename]
    if not uploads:
        return _error_response("Seleccioná el Depts Report o el P & L de Elistar.")
    saved, errors, last = [], [], None
    for upload in uploads:
        filename = os.path.basename(upload.filename)
        tmp_dir = tempfile.mkdtemp(prefix="elistar_")
        path = os.path.join(tmp_dir, filename)
        upload.save(path)
        try:
            report = control_cmv.read_elistar_report(path)
        except ValueError as exc:
            errors.append(str(exc))
            continue
        except Exception as exc:
            print(f"[controles/cmv] {filename}: {exc}")
            errors.append("Un archivo no se pudo leer: tiene que ser el .xls que baja Elistar (Depts Report o P & L).")
            continue
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        control_cmv_db.save_report(report, filename)
        last = (report["year"], report["month"])
        saved.append(f"{'Depts Report' if report['kind'] == 'depts' else 'P & L'} de "
                     f"{_MONTH_NAMES_ES[report['month'] - 1]} {report['year']}")
    if saved:
        flash("Guardado: " + ", ".join(saved) + ".", "success")
    for message in errors:
        flash(message, "error")
    if last:
        return redirect(url_for("controles_cmv", year=last[0], month=last[1]))
    return redirect(request.referrer or url_for("controles_cmv"))


@app.route("/controles/cmv/<int:year>/<int:month>/<kind>/eliminar", methods=["POST"])
def controles_cmv_eliminar(year, month, kind):
    control_cmv_db.delete_report(year, month, kind)
    return redirect(url_for("controles_cmv", year=year, month=month))


# ---------------------------------------------------------------------------
# Cuenta corriente de Kia y Toyota (pedido del usuario, 2026-10-07): lo que
# cargan a cuenta (vales) menos lo que pagan con la Amex, por empresa. El
# historial sale de las hojas de Gettel-Toyota de los Excel de Cierre; lo que
# la app ya tiene (Store Info, LOCAL ACCT, cupones y pagos de Gettel/Toyota)
# completa y manda. Cálculo en cuenta_kia_toyota.py.
# ---------------------------------------------------------------------------

def _kia_toyota_ledger():
    charges = cuenta_kia_toyota_db.get_charges()
    payments = cuenta_kia_toyota_db.get_payments()
    pos = {d: {"la": r["la"], "vs": r["vs"]} for d, r in cuenta_kia_toyota_db.get_pos_days().items()}
    first = min(list(charges) + [p["date"] for p in payments] or [date.today().isoformat()])
    year, month = int(first[:4]), int(first[5:7])
    today = date.today()
    seen = {(p["date"], p["transc"], round(p["amount"], 2)) for p in payments}
    while (year, month) <= (today.year, today.month):
        # Cupones del módulo Gettel/Toyota para los días que el Excel no trae.
        for d in gettel_db.get_month_days(year, month):
            if d["date"] not in charges and (d.get("gettel_amount") or d.get("toyota_amount")):
                charges[d["date"]] = {"la": None, "kia": d.get("gettel_amount"), "kia_gal": d.get("gettel_gallons"),
                                      "toyota": d.get("toyota_amount"), "toyota_gal": d.get("toyota_gallons")}
        for p in gettel_db.get_month_pagos(year, month):
            key = (p["fecha"], str(p["transc_n"]), round(p["total_cupon"] or 0, 2))
            if key not in seen and p["total_cupon"]:
                seen.add(key)
                company = {"kia": cuenta_kia_toyota.KIA, "gettel": cuenta_kia_toyota.KIA,
                           "toyota": cuenta_kia_toyota.TOYOTA}.get((p.get("empresa") or "").strip().lower())
                payments.append({"date": p["fecha"], "transc": str(p["transc_n"]), "amount": p["total_cupon"],
                                 "company": company})
        # Store Info y LOCAL ACCT de los reportes diarios mandan sobre el Excel.
        local_acct = {}
        for day, rows in reportes_db.get_month_departments_by_date(year, month).items():
            local_acct[day] = round(sum(r["amount"] or 0 for r in rows if r["department"] == "LOCAL ACCT"), 2)
        for r in reportes_db.get_month_store_info(year, month):
            if r.get("local_accounts") is not None:
                pos[r["date"]] = {"la": r["local_accounts"], "vs": local_acct.get(r["date"], 0.0)}
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    # Lo que cobró el POS (LOCAL ACCT) sin recibo anotado también es un pago.
    extra, over = cuenta_kia_toyota.pos_only_payments(
        payments, pos, cuenta_kia_toyota_db.get_pos_only_assignments())
    ledger = cuenta_kia_toyota.build_ledger(charges, payments + extra, pos)
    for o in over:
        ledger["issues"].append(f"El {o['date'][8:10]}/{o['date'][5:7]}/{o['date'][:4]} hay ${o['amount']:,.2f} "
                                f"de recibos anotados de más contra lo que cobró el POS.")
    return ledger


@app.route("/controles/kia-toyota")
def controles_kia_toyota():
    year, month = _cierre_month()
    ledger = _kia_toyota_ledger()
    month_key = f"{year:04d}-{month:02d}"
    month_summary = next((m for m in ledger["monthly"] if m["month"] == month_key), None)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)
    return render_template(
        "controles_kia_toyota.html",
        ledger=ledger,
        companies=cuenta_kia_toyota.COMPANIES,
        voided=cuenta_kia_toyota.VOIDED,
        month_summary=month_summary,
        days=cuenta_kia_toyota.month_days(ledger, year, month),
        rebate=cuenta_kia_toyota.REBATE_PER_GALLON,
        card_charge=cuenta_kia_toyota.CARD_CHARGE,
        late_days=cuenta_kia_toyota.LATE_PAYMENT_DAYS,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        month_names=_MONTH_NAMES_ES,
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["carga_gettel"],
    )


@app.route("/controles/kia-toyota/subir", methods=["POST"])
def controles_kia_toyota_subir():
    uploads = [u for u in request.files.getlist("files") if u and u.filename]
    if not uploads:
        return _error_response("Seleccioná los Excel de Cierre.")
    read, errors = [], []
    for upload in uploads:
        filename = os.path.basename(upload.filename)
        tmp_dir = tempfile.mkdtemp(prefix="kia_toyota_")
        path = os.path.join(tmp_dir, filename)
        upload.save(path)
        try:
            data = cuenta_kia_toyota.read_control_workbook(path)
            if not data["days"] and not data["payments"]:
                errors.append(f"{filename}: no tiene hojas de Gettel-Toyota.")
                continue
            new_days, new_payments = cuenta_kia_toyota_db.save_import(data, filename)
            read.append((len(data["days"]), new_payments))
        except Exception as exc:
            print(f"[controles/kia-toyota] {filename}: {exc}")
            errors.append(f"{filename}: no se pudo leer (tiene que ser un Excel de Cierre).")
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
    if read:
        flash(f"Leídos {len(read)} Excel: {sum(n for _, n in read)} pago(s) nuevo(s).", "success")
    for message in errors:
        flash(message, "error")
    return redirect(request.referrer or url_for("controles_kia_toyota"))


@app.route("/controles/kia-toyota/asignar", methods=["POST"])
def controles_kia_toyota_asignar():
    company = request.form.get("company")
    ids = [int(x) for x in request.form.get("ids", "").split(",") if x.strip().isdigit()]
    pos_date = request.form.get("pos_date") or ""
    if company not in cuenta_kia_toyota.COMPANIES + (cuenta_kia_toyota.VOIDED,) or not (ids or pos_date):
        return _error_response("Elegí Kia o Toyota.")
    if pos_date:
        cuenta_kia_toyota_db.assign_pos_only(pos_date, company)
    else:
        cuenta_kia_toyota_db.assign_company(ids, company)
    return redirect(request.referrer or url_for("controles_kia_toyota"))


@app.route("/controles/ice-food-truck")
def controles_ice():
    return redirect(url_for("controles_depositos", **request.args.to_dict()))


@app.route("/carga-datos/depositos/subir", methods=["POST"])
def carga_datos_depositos_subir():
    uploads = [u for u in request.files.getlist("pdf_files") if u and u.filename]
    if not uploads:
        return _error_response("Seleccioná los PDF de depósitos o los Payment Summary.")
    pdf_paths = _save_uploads_to_workspace(uploads)
    fallback = (request.form.get("year", type=int) or date.today().year,
                request.form.get("month", type=int) or date.today().month)
    job_id = jobs.create_job(len(pdf_paths), kind="depositos")
    threading.Thread(target=_run_depositos_job, args=(job_id, pdf_paths, fallback), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(pdf_paths)})


def _pdf_text(pdf_path):
    import pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join((page.extract_text() or "") for page in pdf.pages)


def _link_ice_deposits_to_chase():
    """
    Cada recibo de Ice Machine o de Vaccumms categoriza su depósito en Chase
    (misma fecha e importe): DEPOSITO VENTA ICE, la categoría que la Caja
    muestra como ICE MACHINE (pedido del usuario, 2026-10-07), y DEPOSITO
    VACCUMMS (2026-10-08). Lo corregido a mano en Chase no se toca; un
    depósito ya categorizado como otro de estos tipos (o Food Truck) tampoco.
    Devuelve cuántos movimientos de Chase se categorizaron.
    """
    from collections import Counter
    by_kind = {depositos.ICE_MACHINE: CHASE_DETALLE_ICE, control_depositos.VACCUMMS: CHASE_DETALLE_VACCUMMS}
    special = {CHASE_DETALLE_FOOD_TRUCK, *by_kind.values()}
    changed = 0
    for kind, detalle in by_kind.items():
        wanted = Counter((d["deposit_date"], round(d["amount"], 2)) for d in depositos_db.list_kind(kind)
                         if d["deposit_date"] and d["amount"] is not None)
        for (day, amount), count in wanted.items():
            rows = [r for r in chase_db.deposits_on(day) if abs(r["amount"] - amount) <= 0.005]
            need = count - sum(1 for r in rows if (r["detalle"] or "").upper() == detalle)
            candidates = [r for r in rows if r["detalle_source"] != "manual"
                          and (r["detalle"] or "").upper() not in special]
            for r in candidates[:max(need, 0)]:
                chase_db.set_deposit_detalle(r["rowid"], detalle)
                changed += 1
    return changed


def _run_depositos_job(job_id, pdf_paths, fallback_period):
    """Cada PDF es un Payment Summary (con texto) o un PDF de recibos de depósito (fotos); aislado por archivo."""
    try:
        summaries = replaced = deposits = incomplete = duplicates = failed = 0
        problems, last_period = [], None
        for index, pdf_path in enumerate(pdf_paths, start=1):
            filename = os.path.basename(pdf_path)
            try:
                text = _pdf_text(pdf_path)
                if ice_machine.is_payment_summary_text(text):
                    summary = ice_machine.read_payment_summary_text(text)
                    replaced += ice_machine_db.save_summary(summary, filename)
                    summaries += 1
                    last_period = ice_machine.month_of(summary["to_date"])
                else:
                    s, i, d, period = import_deposit_pdf(pdf_path, filename, fallback_period)
                    deposits, incomplete, duplicates = deposits + s, incomplete + i, duplicates + d
                    last_period = period or last_period
            except ValueError as exc:
                problems.append(str(exc))
            except Exception as exc:
                print(f"[carga-datos/depositos] {pdf_path}: {exc}")
                failed += 1
            jobs.update_job(job_id, done=index, total=len(pdf_paths))

        parts = []
        if summaries:
            parts.append(f"{summaries} Payment Summary guardado(s)" + (f" ({replaced} ya estaban y se actualizaron)" if replaced else "") + ".")
        if deposits:
            parts.append(f"{deposits} recibo(s) de depósito guardado(s).")
        if incomplete:
            parts.append(f"{incomplete} recibo(s) con algún dato sin leer: completalo en Depósitos.")
        if duplicates:
            parts.append(f"{duplicates} recibo(s) ya estaban cargados.")
        linked = _link_ice_deposits_to_chase() if deposits else 0
        if linked:
            parts.append(f"{linked} depósito(s) de Chase quedaron categorizados por su recibo (Ice Machine o Vaccumms).")
        parts.extend(problems)
        if failed:
            parts.append(f"{failed} archivo(s) no se pudieron leer.")
        done = summaries or deposits
        level = "success" if done and not (incomplete or problems or failed) else ("warning" if done or duplicates else "error")
        year, month = last_period or fallback_period
        jobs.update_job(
            job_id, status="done", done=len(pdf_paths), total=len(pdf_paths),
            notice=" ".join(parts) or "No se encontró nada para cargar.", notice_level=level,
            redirect_url=f"/carga-datos/depositos?year={year}&month={month}",
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


@app.route("/carga-datos/depositos/resumen/<summary_no>/eliminar", methods=["POST"])
def carga_datos_resumen_eliminar(summary_no):
    to_date = ice_machine_db.delete_summary(summary_no)
    if to_date:
        year, month = ice_machine.month_of(to_date)
        return redirect(url_for("carga_datos_depositos", year=year, month=month))
    return redirect(url_for("carga_datos_depositos"))


def import_deposit_pdf(pdf_path, filename, fallback_period):
    """Divide un PDF en depósitos y los guarda. Devuelve (guardados, incompletos, duplicados, primer_período)."""
    saved = incomplete = duplicates = 0
    first_period = None
    for item in depositos.extract_deposits_from_pdf(pdf_path, filename):
        deposit_date = item["date"]
        iso = deposit_date.isoformat() if deposit_date else None
        complete = item["amount"] is not None and deposit_date is not None and item["tx_number"] is not None
        # También por archivo y página cuando el recibo está completo: si el
        # usuario ya corrigió a mano el monto o la fecha, find_duplicate no
        # lo encuentra con lo que vuelve a leer el OCR y se duplicaba.
        if (
            (complete and depositos_db.find_duplicate(item["tx_number"], iso, item["amount"]))
            or depositos_db.find_by_source_page(filename, item["page"], item["tx_number"])
        ):
            duplicates += 1
            continue
        kind = item["kind"] or _chase_kind_for(deposit_date, item["amount"])
        year, month = (deposit_date.year, deposit_date.month) if deposit_date else fallback_period
        # pdf_path es NOT NULL en la base: sin guardado de documentos queda ''.
        rel = ""
        if documents_db.GUARDAR_DOCUMENTOS:
            rel = depositos_db.new_pdf_relpath(year, month)
            depositos.write_single_page_pdf(pdf_path, item["page"], depositos_db.absolute_path(rel))
        depositos_db.add_deposit(
            year, month, iso, item["amount"],
            depositos.default_description(item["tx_number"], kind),
            item["tx_number"], kind, rel, filename, item["page"],
        )
        saved += 1
        first_period = first_period or (year, month)
        if not complete:
            incomplete += 1
    return saved, incomplete, duplicates, first_period


@app.route("/controles/depositos/<int:deposit_id>/guardar", methods=["POST"])
def controles_deposito_guardar(deposit_id):
    deposit = depositos_db.get_deposit(deposit_id)
    if deposit is None:
        flash("Ese depósito ya no existe.", "error")
        return redirect(url_for("carga_datos_depositos"))
    raw_date = (request.form.get("deposit_date") or "").strip() or None
    try:
        if raw_date:
            datetime.strptime(raw_date, "%Y-%m-%d")
        amount = _parse_money_field(request.form.get("amount"))
    except ValueError:
        flash("La fecha o el monto no son válidos.", "error")
        return redirect(url_for("carga_datos_depositos", year=deposit["year"], month=deposit["month"]))
    description = (request.form.get("description") or "").strip() or None
    # La aclaración se edita en la propia descripción: "(Food Truck)" al final
    # lo saca del control contra Caja, borrarla lo vuelve un depósito normal.
    depositos_db.update_deposit(deposit_id, raw_date, amount, description,
                                depositos.kind_from_description(description))
    _link_ice_deposits_to_chase()
    updated = depositos_db.get_deposit(deposit_id)
    return redirect(url_for("carga_datos_depositos", year=updated["year"], month=updated["month"]) + f"#deposito-{deposit_id}")


@app.route("/controles/depositos/<int:deposit_id>/eliminar", methods=["POST"])
def controles_deposito_eliminar(deposit_id):
    deposit = depositos_db.delete_deposit(deposit_id)
    if deposit is None:
        flash("Ese depósito ya no existe.", "error")
        return redirect(url_for("carga_datos_depositos"))
    flash("Depósito eliminado.", "success")
    return redirect(url_for("carga_datos_depositos", year=deposit["year"], month=deposit["month"]))


@app.route("/controles/depositos/<int:deposit_id>/pdf")
def controles_deposito_pdf(deposit_id):
    deposit = depositos_db.get_deposit(deposit_id)
    if deposit is None:
        flash("Ese depósito ya no existe.", "error")
        return redirect(url_for("carga_datos_depositos"))
    name = (deposit.get("description") or "Deposito").replace("#", "N").replace("/", "-") + ".pdf"
    if not os.path.isfile(depositos_db.absolute_path(deposit["pdf_path"])):
        flash("No se encontró el PDF guardado de ese depósito.", "error")
        return redirect(url_for("carga_datos_depositos", year=deposit["year"], month=deposit["month"]))
    return send_file(
        depositos_db.absolute_path(deposit["pdf_path"]),
        as_attachment=request.args.get("mode") == "download",
        download_name=name,
    )


# Módulos que guardan sus archivos originales vía documents_db (ver ese
# módulo) -- cada entrada define el título/tema/link "volver" que usa la
# página genérica de abajo. EFT junta dos módulos lógicos (PDF de EFT +
# reporte mensual de Cupones) en una sola lista -- son el mismo "cajón" de
# documentos para el usuario, aunque se guardan con distinta granularidad.
_DOCUMENTS_MODULES = {
    "chase": {"title": "Chase Bank", "theme": "carga_chase", "back_endpoint": "chase_historial"},
    "caja": {"title": "Caja", "theme": "carga_caja", "back_endpoint": "carga_datos_caja"},
    "eft": {"title": "EFT y Cupones", "theme": "carga_eft", "back_endpoint": "carga_datos_eft_historial"},
    "gettel_toyota": {"title": "Gettel / Toyota", "theme": "carga_gettel", "back_endpoint": "carga_datos_gettel_historial"},
    "cmv_costo": {"title": "CMV — Costo", "theme": "carga_cmv", "back_endpoint": "carga_datos_cmv_costo_historial"},
    "cmv_ventas": {"title": "CMV — Ventas", "theme": "carga_cmv", "back_endpoint": "carga_datos_cmv_ventas_historial"},
    "lottery_resumen_mensual": {"title": "Lottery — Resumen mensual", "theme": "lottery", "back_endpoint": "carga_datos_lottery_historial"},
    "proveedores": {"title": "Proveedores", "theme": "carga_proveedores", "back_endpoint": "carga_datos_proveedores_historial"},
    "horas_trabajo": {"title": "Horas de Trabajo", "theme": "carga_horas", "back_endpoint": "carga_datos_horas_trabajo_historial"},
    "combustible": {"title": "Combustible", "theme": "fisico", "back_endpoint": "fisico_view"},
}

@app.route("/documentos/<module_key>")
def carga_datos_documentos(module_key):
    """
    Lista genérica de los archivos originales ya subidos para un módulo --
    ver documents_db.py. Un solo template/ruta reusado por EFT, Gettel/
    Toyota, CMV (Costo y Ventas por separado) y el resumen mensual de
    Lottery, en vez de repetir la misma página 5 veces.
    """
    info = _DOCUMENTS_MODULES.get(module_key)
    if info is None:
        flash("Módulo de documentos desconocido.", "error")
        return redirect(url_for("documentos_index"))

    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    docs = documents_db.list_documents(module_key, year, month)
    # Proveedores: mostrar como carpetas por proveedor en vez de una lista
    # plana -- pedido explícito del usuario (2026-09-15): "muestra a los
    # proveedores como si fueran carpetas que expandiendolas ves todos los
    # pdf juntos, no muestres todos como estas haciendo ahora uno encima
    # del otro". `label` ya guarda el nombre del proveedor (ver
    # carga_datos_proveedores_subir) -- se agrupa por ese campo, sin tocar
    # el resto de los módulos (que siguen con la lista plana de siempre).
    supplier_groups = None
    if module_key == "proveedores":
        by_supplier = {}
        for doc in docs:
            key = doc.get("label") or "Sin proveedor identificado"
            by_supplier.setdefault(key, []).append(doc)
        supplier_groups = [
            {"label": label, "docs": by_supplier[label]} for label in sorted(by_supplier)
        ]
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "carga_datos_documentos.html",
        module_key=module_key,
        module_title=info["title"],
        back_url=url_for(info["back_endpoint"]),
        docs=docs,
        supplier_groups=supplier_groups,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY[info["theme"]],
    )


@app.route("/documentos/<module_key>/subir", methods=["POST"])
def carga_datos_documentos_subir(module_key):
    """
    Subida manual de un documento cualquiera a este módulo -- pedido
    explícito del usuario (2026-09-14): "no solo voy a subir los pdf de los
    EFT o de cupones, voy a subir todo tipo de invoice que nos llegan de
    J.H." (ej. facturas de servicios de J.H. Williams -- VPN, Network fee,
    etc. -- que no son ni un EFT ni el reporte mensual de Cupones, pero
    igual hay que poder guardarlas y verlas acá). Genérico para cualquier
    módulo de _DOCUMENTS_MODULES, no solo EFT.
    """
    info = _DOCUMENTS_MODULES.get(module_key)
    if info is None:
        flash("Módulo de documentos desconocido.", "error")
        return redirect(url_for("documentos_index"))

    year = request.form.get("year", type=int) or date.today().year
    month = request.form.get("month", type=int) or date.today().month
    label = (request.form.get("label") or "").strip() or None
    uploads = [f for f in request.files.getlist("documento_files") if f and f.filename]
    if not uploads:
        flash("Seleccioná uno o más archivos.", "error")
        return redirect(url_for("carga_datos_documentos", module_key=module_key, year=year, month=month))

    saved = 0
    for upload in uploads:
        try:
            path, filename = _save_upload_to_workspace(upload)
            documents_db.store_document(module_key, path, filename, year, month, label=label)
            saved += 1
        except Exception as exc:
            print(f"[carga-datos/documentos/{module_key}/subir] {upload.filename}: {exc}")

    if saved:
        flash(f"{saved} documento(s) guardado(s).", "success")
    else:
        flash("No se pudo guardar ningún documento.", "error")
    return redirect(url_for("carga_datos_documentos", module_key=module_key, year=year, month=month))


@app.route("/documentos/archivo/<int:document_id>/descargar")
def carga_datos_documento_descargar(document_id):
    doc = documents_db.get_document(document_id)
    if doc is None:
        flash("Ese documento ya no existe.", "error")
        return redirect(url_for("documentos_index"))
    # Vista previa por default (inline), descarga forzada solo con
    # ?mode=download -- pedido explícito del usuario (2026-09-19), ver
    # templates/_pdf_links.html. Antes esta ruta siempre forzaba la
    # descarga (as_attachment=True) -- ahora es la misma ruta para las dos
    # cosas, el link "Ver" simplemente no manda el parámetro.
    force_download = request.args.get("mode") == "download"
    if not os.path.isfile(doc["stored_path"]):
        flash("El archivo de ese documento ya no está en la carpeta de datos.", "error")
        return redirect(url_for("documentos_index"))
    return send_file(doc["stored_path"], as_attachment=force_download, download_name=doc["filename"])


@app.route("/documentos/archivo/<int:document_id>/eliminar", methods=["POST"])
def carga_datos_documento_eliminar(document_id):
    doc = documents_db.get_document(document_id)
    if doc is None:
        flash("Ese documento ya no existe.", "error")
        return redirect(url_for("documentos_index"))
    documents_db.delete_document(document_id)
    flash("Documento eliminado.", "success")
    # Esta ruta también se usa desde Reporte Diario y Lottery, cuyos módulos
    # no están en _DOCUMENTS_MODULES: se vuelve a la página de donde vino.
    referrer = request.referrer or ""
    if referrer.startswith(request.host_url):
        return redirect(referrer)
    if doc["module"] not in _DOCUMENTS_MODULES:
        return redirect(url_for("documentos_index"))
    return redirect(url_for("carga_datos_documentos", module_key=doc["module"], year=doc["year"], month=doc["month"]))


# ---------------------------------------------------------------------------
# Proveedores -- Carga de Datos (2026-09-14, primer paso, pedido explícito
# del usuario: "quiero que empieces con el modulo de proveedores y donde
# pueda guardar sus facturas"). Reusa el motor de detección/extracción de
# los 32 proveedores + el dinámico tal cual (`extract_invoices_from_pdf`,
# proveedores.py) -- acá solo se guarda el resultado (proveedores_db.py),
# sin escribir ningún Excel Ledger. El PDF original queda en Documentos
# (documents_db, módulo "proveedores") como el resto de los módulos.
# ---------------------------------------------------------------------------

@app.route("/carga-datos/proveedores")
def carga_datos_proveedores():
    _active_job = jobs.get_active_job("proveedores")
    return render_template(
        "carga_datos_proveedores.html",
        resume_job_id=(_active_job["id"] if _active_job else None),
        **THEME_BY_KEY["carga_proveedores"],
    )


@app.route("/carga-datos/proveedores/subir", methods=["POST"])
def carga_datos_proveedores_subir():
    """
    Cada PDF se detecta/extrae con el mismo motor de siempre (sin tocarlo) y
    se guarda en proveedores_db -- aislado por archivo, mismo criterio de
    todo el proyecto (una factura con un problema no tira abajo el resto
    del lote). Duplicado = mismo N° de factura ya guardado para ESE
    proveedor (igual criterio que el Ledger real).
    """
    uploads = [f for f in request.files.getlist("pdf_files") if f and f.filename]
    if not uploads:
        return _error_response("Seleccioná uno o más PDF de factura.")

    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths), kind="proveedores")
    threading.Thread(target=_run_carga_datos_proveedores_job, args=(job_id, paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


def _invoice_ref(invoice_no, filename):
    """
    Lo que identifica una factura en el aviso de la carga: su N° o, si no se
    pudo leer, el N° que trae el nombre del archivo ("Invoice 2035255085
    22.06.2026.pdf") o, si no trae uno solo, el nombre mismo.
    """
    if invoice_no:
        return str(invoice_no)
    found = re.findall(r"(?<![\d.,/])\d{6,}(?![\d.,/])", os.path.splitext(filename)[0])
    return found[0] if len(found) == 1 else f"archivo «{filename}»"


def _invoice_notice_line(text, item):
    """'No se pudo cargar factura de Red Bull (2035255085)' -- una factura por renglón del aviso."""
    ref = _invoice_ref(item.get("invoice_no"), item["filename"])
    if item.get("supplier"):
        return f"{text} de {item['supplier']} ({ref})"
    return f"{text} ({ref}, proveedor no reconocido)"


def _run_carga_datos_proveedores_job(job_id, paths):
    """
    Corre en su propio hilo -- mismo patrón que _run_carga_datos_reporte_diario_job, ver ese docstring.

    Aviso final como lista, una factura por renglón (pedido del usuario
    2026-10-05: "No se pudo cargar factura de Red Bull (N° de invoice)"):
    cada factura de un PDF se carga o se rechaza por su cuenta -- una
    ilegible no frena las demás del mismo PDF ni del lote.
    """
    try:
        saved = []
        duplicates = []
        failed = []
        date_mismatches = []
        lines_saved = 0
        lines_failed = []
        without_invoice = []

        for index, path in enumerate(paths, start=1):
            filename = os.path.basename(path)
            try:
                supplier_key, supplier_label, invoices = extract_invoices_from_pdf(path)
            except _PDF_EXTRACTION_EXCEPTIONS as exc:
                print(f"[carga-datos/proveedores] {filename}: {exc}")
                failed.append({"filename": filename, "supplier": getattr(exc, "supplier_label", None)})
                jobs.update_job(job_id, done=index, total=len(paths))
                continue
            # Gold Coast y Red Bull devuelven aparte cada factura del PDF que no
            # se pudo leer con seguridad (proveedores._ticket_invoices).
            invoices, unreadable = invoice_errors(invoices)
            for bad in unreadable:
                print(f"[carga-datos/proveedores] {filename}, factura {bad.get('invoice_no')}: {bad['error']}")
                failed.append({"filename": filename, "supplier": supplier_label, "invoice_no": bad.get("invoice_no")})
            if not invoices and not unreadable:
                # Un PDF sin ninguna factura (por ejemplo, solo una devolución
                # de Coca-Cola) no es un error, pero tampoco puede desaparecer
                # del lote sin que se note (auditoría 2026-09).
                without_invoice.append({"filename": filename, "supplier": supplier_label})

            # Chequeo de fecha del nombre de archivo (2026-09-15, pedido
            # explícito del usuario -- "al igual que con las facturas de
            # proveedores"): varios proveedores ya nombran el PDF con su
            # propia fecha (Flori-Gas, LMT, SkyHarvest, etc.) -- si el
            # nombre trae una fecha completa y no coincide con la leída de
            # la factura, se rechaza esa factura en vez de guardarla/
            # archivarla bajo un mes que podría no ser el suyo. Un nombre
            # sin fecha completa (la mayoría de los proveedores) no se
            # toca -- nada que cruzar.
            valid_invoices = []
            for invoice in invoices:
                # FPL/Manatee (factura por período): el nombre de archivo
                # trae la fecha de vencimiento/emisión, nunca el fin del
                # período -- el extractor lo marca para no rechazarla.
                if not invoice.get("skip_filename_date_check") and _filename_date_mismatch(
                    filename, invoice["date"]
                ):
                    date_mismatches.append({"filename": filename, "supplier": supplier_label,
                                            "invoice_no": invoice["invoice_no"]})
                    continue
                valid_invoices.append(invoice)

            saved_before = len(saved)
            for invoice in valid_invoices:
                ok = proveedores_db.save_invoice(
                    supplier_key, supplier_label, invoice["date"], invoice["invoice_no"],
                    invoice["amount"], source_filename=filename,
                )
                if ok:
                    saved.append({"filename": filename, "supplier": supplier_label, "date": invoice["date"]})
                else:
                    duplicates.append({"filename": filename, "supplier": supplier_label})

            # Detalle de productos (proveedores_productos.py, pedido explícito
            # del usuario 2026-09-28). Corre también si la factura ya estaba
            # cargada -- volver a subirla completa su detalle sin duplicar
            # nada. Si el detalle no cierra contra la factura no se guarda
            # ningún renglón, pero la factura en sí queda guardada.
            if any("lines" in inv for inv in valid_invoices):
                # Gold Coast y Red Bull: el mismo lector que leyó cada factura ya
                # trae su detalle (o None si los renglones no cerraron).
                for invoice in valid_invoices:
                    if invoice.get("lines"):
                        proveedores_db.replace_invoice_lines(
                            supplier_key, str(invoice["invoice_no"]), invoice["date"], invoice["lines"]
                        )
                        lines_saved += 1
                    else:
                        print(f"[carga-datos/proveedores] detalle de productos de {filename}, factura "
                              f"{invoice['invoice_no']}: {invoice.get('lines_error')}")
                        lines_failed.append({"filename": filename, "supplier": supplier_label,
                                             "invoice_no": invoice["invoice_no"]})
            elif valid_invoices and supplier_key in proveedores_productos.LINE_EXTRACTORS:
                try:
                    detail = proveedores_productos.extract_lines(supplier_key, path, invoices=valid_invoices)
                    invoice = next(
                        (inv for inv in valid_invoices if str(inv["invoice_no"]) == detail["invoice_no"]), None
                    )
                    if invoice is None:
                        raise ValueError("el N° de invoice del detalle no coincide con el de la factura.")
                    proveedores_db.replace_invoice_lines(
                        supplier_key, detail["invoice_no"], invoice["date"], detail["lines"]
                    )
                    lines_saved += 1
                except Exception as exc:
                    print(f"[carga-datos/proveedores] detalle de productos de {filename}: {exc}")
                    lines_failed.append({"filename": filename, "supplier": supplier_label,
                                         "invoice_no": ", ".join(str(inv["invoice_no"]) for inv in valid_invoices)})

            # Solo si se guardó alguna factura de este PDF: volver a subir uno
            # ya cargado no archiva otra copia (auditoría 2026-09).
            if len(saved) > saved_before:
                try:
                    doc_when = valid_invoices[0]["date"]
                    documents_db.store_document(
                        "proveedores", path, filename, doc_when.year, doc_when.month, label=supplier_label
                    )
                except Exception as exc:
                    print(f"[documents_db] no se pudo guardar el documento de Proveedores {filename}: {exc}")
            jobs.update_job(job_id, done=index, total=len(paths))

        parts = []
        if saved:
            parts.append(f"{len(saved)} factura(s) guardada(s)"
                         + (f", {lines_saved} con detalle de productos." if lines_saved else "."))
        elif lines_saved:
            parts.append(f"Detalle de productos guardado en {lines_saved} factura(s).")
        if duplicates:
            parts.append(f"{len(duplicates)} factura(s) ya estaban cargadas y se omitieron.")
        parts.extend(_invoice_notice_line("No se pudo cargar factura", item) for item in failed)
        parts.extend(_invoice_notice_line("No se pudo cargar factura", item)
                     + ": la fecha no coincide con el nombre del archivo" for item in date_mismatches)
        parts.extend(_invoice_notice_line("Sin ninguna factura para guardar (por ejemplo, solo devolución) en el PDF",
                                          item) for item in without_invoice)
        parts.extend(_invoice_notice_line("Sin detalle de productos (los renglones no cerraban): factura", item)
                     for item in lines_failed)
        if not parts:
            parts.append("No se guardó ninguna factura de este lote.")
        level = (
            "success" if (saved and not duplicates and not failed and not date_mismatches and not lines_failed and not without_invoice)
            else ("error" if not saved and not lines_saved else "warning")
        )

        if saved:
            d = saved[0]["date"]
            redirect_url = f"/carga-datos/proveedores/historial?year={d.year}&month={d.month}"
        else:
            redirect_url = "/carga-datos/proveedores/historial"

        jobs.update_job(
            job_id, status="done", done=len(paths), total=len(paths),
            notice="\n".join(parts), notice_level=level, redirect_url=redirect_url,
        )
    except Exception as exc:
        jobs.update_job(job_id, status="error", error=f"Error: {exc}")


@app.route("/carga-datos/proveedores/historial")
def carga_datos_proveedores_historial():
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    groups = proveedores_db.get_month_invoices(year, month)
    # N° de factura -> link al PDF original guardado (2026-09-15, pedido
    # explícito del usuario) -- se cruza por nombre de archivo exacto
    # (proveedores_db.save_invoice y documents_db.store_document guardan el
    # mismo nombre de archivo tal cual se subió), sin filtrar por mes -- la
    # factura puede haberse subido en un mes distinto al de su propia
    # fecha de negocio. Sin dato adjunto (documento ya borrado, o la
    # factura es vieja y se cargó antes de que esto existiera) el N° de
    # factura queda como texto plano, mismo criterio que el link de EFT.
    docs_by_filename = {}
    for doc in documents_db.list_all_documents("proveedores"):
        docs_by_filename.setdefault(doc["filename"], doc)
    for group in groups:
        for inv in group["invoices"]:
            inv["document"] = docs_by_filename.get(inv.get("source_filename"))
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "carga_datos_proveedores_historial.html",
        groups=groups,
        month_total=round(sum(g["total"] for g in groups), 2),
        invoice_count=sum(len(g["invoices"]) for g in groups),
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["carga_proveedores"],
    )


@app.route("/carga-datos/proveedores/factura/<int:invoice_id>/eliminar", methods=["POST"])
def carga_datos_proveedores_eliminar(invoice_id):
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    ok = proveedores_db.delete_invoice(invoice_id)
    flash("Factura eliminada." if ok else "Esa factura ya no existe.", "success" if ok else "error")
    return redirect(url_for("carga_datos_proveedores_historial", year=year, month=month))


# ---------------------------------------------------------------------------
# Proveedores guardados -- una tarjeta por proveedor (2026-09-21, pedido
# explícito del usuario: "cuando entremos a ver guardado, que salgan muchos
# modulos como si fuera el de herramientas... y cuando se entra se va a ver
# los detalles no solo de las facturas que se cargaron, sino tambien va a
# estar conectado eso al detalle del banco donde los pagos a proveedores se
# van a mover al proveedor directamente que sale en el asiento"). Reemplaza
# la vista por mes (carga_datos_proveedores_historial, que sigue existiendo
# tal cual para quien la prefiera) como el link "Ver guardado" de la barra
# lateral.
#
# 2026-09-22, corrección explícita del usuario tras la primera vista de
# esto: (a) la grilla ya NO se filtra a "solo los que tienen algo cargado"
# -- muestra SIEMPRE los ~27 proveedores del registro (mismo criterio que
# la grilla de Herramientas), cada uno con su logo real; (b) el panel de
# "Reglas de pago a proveedores" se sacó de acá -- pasa a tener su propia
# página (carga_datos_proveedores_reglas); (c) el detalle de cada
# proveedor dejó de ser dos tablas sueltas (facturas / pagos) para ser un
# solo "cuadro de cuenta corriente" cronológico con saldo corrido -- ver
# _build_supplier_ledger.
# ---------------------------------------------------------------------------

def _supplier_labels():
    return {entry["key"]: entry["label"] for entry in list_supplier_registry_entries()}


# Productos pasó de Proveedores a Controles (pedido del usuario, 2026-10-02):
# "ahí es donde se va a controlar el cambio de precios de cada producto de
# cada proveedor". Las direcciones viejas redirigen a las nuevas.
@app.route("/carga-datos/proveedores/productos", defaults={"rest": ""})
@app.route("/carga-datos/proveedores/productos/<path:rest>")
def carga_datos_proveedores_productos_viejo(rest):
    if rest.startswith("carpeta/"):
        rest = rest[len("carpeta/"):]
    elif rest and rest != "lista":
        rest = "producto/" + rest
    target = "/controles/productos" + ("/" + quote(rest) if rest else "")
    if request.query_string:
        target += "?" + request.query_string.decode("utf-8", "replace")
    return redirect(target)


@app.route("/controles/productos")
def controles_productos():
    """
    Productos como carpetas (pedido del usuario, 2026-10-02): una por
    proveedor que se lee bien (los que guardan renglones), adentro sus
    facturas por fecha y en cada una los productos contra su compra
    anterior por fecha. La lista completa por producto sigue en /lista.
    """
    folders = proveedores_productos.build_supplier_folders(proveedores_db.get_all_invoice_lines(), _supplier_labels())
    return render_template(
        "controles_productos.html",
        view="suppliers",
        folders=folders,
        **THEME_BY_KEY["carga_proveedores"],
    )


@app.route("/controles/productos/<supplier_key>")
def controles_productos_proveedor(supplier_key):
    invoices = proveedores_productos.build_supplier_invoices(supplier_key, proveedores_db.get_all_invoice_lines())
    if not invoices:
        flash("Ese proveedor no tiene facturas con productos guardados.", "error")
        return redirect(url_for("controles_productos"))
    years = []
    for invoice in invoices:
        year = invoice["invoice_date"][:4]
        if not years or years[-1]["year"] != year:
            years.append({"year": year, "invoices": []})
        years[-1]["invoices"].append(invoice)
    return render_template(
        "controles_productos.html",
        view="invoices",
        supplier_key=supplier_key,
        supplier_label=_supplier_labels().get(supplier_key, supplier_key),
        years=years,
        invoice_count=len(invoices),
        **THEME_BY_KEY["carga_proveedores"],
    )


@app.route("/controles/productos/<supplier_key>/<invoice_date>/<invoice_no>")
def controles_productos_factura(supplier_key, invoice_date, invoice_no):
    found = proveedores_productos.build_invoice_products(
        supplier_key, invoice_date, invoice_no, proveedores_db.get_all_invoice_lines(),
    )
    if found is None:
        flash("Esa factura no tiene productos guardados.", "error")
        return redirect(url_for("controles_productos_proveedor", supplier_key=supplier_key))
    invoice, rows = found
    rows = proveedores_productos.with_departments(rows, cmv_db.get_all_costs())
    return render_template(
        "controles_productos.html",
        view="invoice",
        supplier_key=supplier_key,
        supplier_label=_supplier_labels().get(supplier_key, supplier_key),
        invoice=invoice,
        rows=rows,
        **THEME_BY_KEY["carga_proveedores"],
    )


@app.route("/controles/productos/<supplier_key>/<invoice_date>/<invoice_no>/cambios.<fmt>")
def controles_productos_cambios(supplier_key, invoice_date, invoice_no, fmt):
    """
    Reporte para el manager (pedido del usuario, 2026-10-02): los productos
    de esta factura que cambiaron de costo contra su compra anterior, en PDF
    o Excel, para mandárselo (ver proveedores_productos.price_change_rows).
    """
    if fmt not in ("pdf", "xlsx"):
        abort(404)
    found = proveedores_productos.build_invoice_products(
        supplier_key, invoice_date, invoice_no, proveedores_db.get_all_invoice_lines(),
    )
    if found is None:
        flash("Esa factura no tiene productos guardados.", "error")
        return redirect(url_for("controles_productos_proveedor", supplier_key=supplier_key))
    invoice, rows = found
    changed = proveedores_productos.price_change_rows(rows, cmv_db.get_all_costs())
    label = _supplier_labels().get(supplier_key, supplier_key)
    workspace_dir = tempfile.mkdtemp(prefix="cambios_precio_")
    dest_path = os.path.join(
        workspace_dir,
        f"Price changes {label} {invoice_date[5:7]}-{invoice_date[8:10]}-{invoice_date[0:4]} Invoice {invoice_no}.{fmt}",
    )
    if fmt == "pdf":
        proveedores_productos.build_price_change_pdf(label, invoice, changed, dest_path)
    else:
        proveedores_productos.build_price_change_workbook(label, invoice, changed, dest_path)
    return send_file(dest_path, as_attachment=True, download_name=os.path.basename(dest_path))


@app.route("/controles/productos/lista")
def controles_productos_lista():
    """
    Productos comprados a proveedores (pedido explícito del usuario,
    2026-09-28): un renglón por UPC con su último costo contra el anterior,
    a qué proveedor se compra, y el costo/precio del POS (CMV) para ver el
    margen real y si el costo del POS quedó desactualizado. Se llena solo al
    subir facturas en Proveedores -- ver proveedores_productos.py.
    """
    pos_costs = cmv_db.get_all_costs()
    products = proveedores_productos.build_product_list(
        proveedores_db.get_all_invoice_lines(), _supplier_labels(), pos_costs,
    )
    return render_template(
        "controles_productos_lista.html",
        products=products,
        cost_up=sum(1 for p in products if (p["change"] or 0) > 0.0001),
        cost_down=sum(1 for p in products if (p["change"] or 0) < -0.0001),
        pos_stale=sum(1 for p in products if p["pos_cost_diff"] not in (None, 0)),
        pos_loaded=bool(pos_costs),
        **THEME_BY_KEY["carga_proveedores"],
    )


@app.route("/controles/productos/producto/<product_key>")
def controles_productos_producto(product_key):
    detail = proveedores_productos.build_product_detail(
        product_key, proveedores_db.get_all_invoice_lines(), _supplier_labels(),
        cmv_db.get_all_costs(), cmv_db.get_all_monthly_sales(),
    )
    if detail is None:
        flash("Ese producto no aparece en ninguna factura cargada.", "error")
        return redirect(url_for("controles_productos"))
    return render_template(
        "controles_productos_producto.html",
        month_names=_MONTH_NAMES_ES,
        **detail,
        **THEME_BY_KEY["carga_proveedores"],
    )


@app.route("/carga-datos/proveedores/guardado")
def carga_datos_proveedores_guardado():
    """
    2026-09-22, tres pedidos puntuales del usuario tras ver esto por
    primera vez: (a) separar los proveedores que son en realidad un
    SERVICIO mensual (FPL, Manatee County, Airgas -- "is_service" en el
    registro) de los que traen mercadería para revender, en dos secciones
    de la grilla; (b) poder ocultar/desocultar un proveedor puntual (ver
    proveedores_db.set_supplier_hidden) sin borrar ningún dato; (c) cada
    tarjeta muestra SOLO el nombre -- se sacaron los contadores/totales
    (siguen calculándose para la lógica de "huérfanos" de abajo, pero ya
    no se muestran en pantalla).
    """
    registry = list_supplier_registry_entries()
    invoice_counts = {row["supplier_key"]: row for row in proveedores_db.list_suppliers_with_counts()}
    payment_totals = chase_db.get_supplier_payment_totals()
    hidden_keys = proveedores_db.list_hidden_suppliers()

    all_suppliers = []
    seen_keys = set()
    for entry in registry:
        key = entry["key"]
        seen_keys.add(key)
        all_suppliers.append({
            "key": key,
            "label": entry["label"],
            "logo": entry.get("logo"),
            "is_service": entry.get("is_service", False),
            "hidden": key in hidden_keys,
        })

    # Un supplier_key con datos guardados pero ya sin entrada en el
    # registro (ej. un proveedor dinámico eliminado después de cargarle
    # facturas) igual tiene que poder verse, con el label que ya quedó
    # guardado en su factura como respaldo -- sin esto, esas facturas
    # quedarían huérfanas, invisibles desde cualquier lado.
    for key in (set(invoice_counts) | set(payment_totals)) - seen_keys:
        invoices = invoice_counts.get(key)
        label = invoices["supplier_label"] if invoices else key
        all_suppliers.append({
            "key": key,
            "label": label,
            "logo": None,
            "is_service": False,
            "hidden": key in hidden_keys,
        })
    all_suppliers.sort(key=lambda s: s["label"])

    merchandise = [s for s in all_suppliers if not s["hidden"] and not s["is_service"]]
    services = [s for s in all_suppliers if not s["hidden"] and s["is_service"]]
    hidden = [s for s in all_suppliers if s["hidden"]]

    return render_template(
        "carga_datos_proveedores_guardado.html",
        merchandise_suppliers=merchandise,
        service_suppliers=services,
        hidden_suppliers=hidden,
        **THEME_BY_KEY["carga_proveedores"],
    )


@app.route("/carga-datos/proveedores/guardado/<supplier_key>/ocultar", methods=["POST"])
def carga_datos_proveedores_ocultar(supplier_key):
    """
    Ocultar/mostrar, instantáneo (2026-09-22, pedido explícito del
    usuario: "que no te mande una notificacion... sino que simplemente lo
    oculte rapido"). Ya no hace `flash()`+redirect -- el botón de la
    grilla llama esto por `fetch` y mueve la tarjeta en el DOM al toque
    (ver el <script> de carga_datos_proveedores_guardado.html), sin
    recargar la página ni mostrar ningún aviso.
    """
    if not current_user.is_admin:
        return jsonify({"error": "Solo un administrador puede ocultar/mostrar proveedores."}), 403
    proveedores_db.set_supplier_hidden(supplier_key, True)
    return jsonify({"ok": True})


@app.route("/carga-datos/proveedores/guardado/<supplier_key>/mostrar", methods=["POST"])
def carga_datos_proveedores_mostrar(supplier_key):
    if not current_user.is_admin:
        return jsonify({"error": "Solo un administrador puede ocultar/mostrar proveedores."}), 403
    proveedores_db.set_supplier_hidden(supplier_key, False)
    return jsonify({"ok": True})


@app.route("/carga-datos/proveedores/reglas")
def carga_datos_proveedores_reglas():
    """
    Página propia para las reglas de pago a proveedores (keyword de Chase
    -> proveedor) -- se sacó de la grilla de "Guardado" (2026-09-22,
    pedido explícito: "ahi solo tendrian que ir modulos... asi como en
    herramientas"), sin perder la funcionalidad ni el patrón admin-only ya
    validado.
    """
    return render_template(
        "carga_datos_proveedores_reglas.html",
        pago_rules=proveedores_pago_rules.list_display_rules(),
        supplier_options=list_supplier_registry_entries(),
        **THEME_BY_KEY["carga_proveedores"],
    )


def _supplier_payment_label(description):
    """ "OP" para un pago de Chase a proveedor, "OP Cheque N° 1770" si fue con cheque."""
    from proveedores import _check_number_from_description
    check_no = _check_number_from_description(description or "")
    return f"OP Cheque N° {check_no}" if check_no else "OP"


def _build_supplier_ledger(invoices, payments, credit_memos=None, manual_payments=None, history_payments=None):
    """
    Arma el "cuadro de cuenta corriente" de un proveedor -- pedido
    explícito del usuario (2026-09-22): "que sea como funcionaba la
    planilla de excel, en donde te mostraba con formula si se debia algo
    anteriormente o si estaba en 0". Mezcla facturas (Debe) + pagos de
    Chase ya vinculados (Haber) + credit memos (Haber, ver
    proveedores_db.save_credit_memo) + pagos a mano vía Caja (Haber, ver
    proveedores_db.save_manual_payment) en orden cronológico, con un
    saldo corrido -- misma fórmula que la columna BALANCE del Excel real
    (`=+anterior+DEBE-HABER`) -- y los agrupa por mes, en orden
    ascendente (el más viejo arriba, igual que la planilla real).

    `invoices` y `payments` ya vienen leídos de proveedores_db/chase_db
    (con `document` ya resuelto en cada factura, si corresponde) -- esta
    función es pura, solo mezcla/ordena/suma.
    """
    entries = []
    for inv in invoices:
        entries.append({
            "date": inv["invoice_date"],
            "kind": "invoice",
            "detail": f"Factura {inv['invoice_no']}",
            "document": inv.get("document"),
            "debe": inv["amount"],
            "haber": 0.0,
        })
    for p in payments:
        entries.append({
            "date": p["posting_date"],
            "kind": "payment",
            # Solo "OP" (+ N° de cheque si fue con cheque) -- pedido
            # explícito del usuario (2026-09-23): la descripción completa
            # del banco no se muestra en la planilla, queda solo de tooltip.
            "detail": _supplier_payment_label(p["description"]),
            "bank_description": p["description"],
            "document": None,
            "debe": 0.0,
            # Con el signo real: un débito (negativo) es un pago; un ingreso
            # vinculado (reintegro, reversa) vuelve a sumar deuda en vez de
            # restarla como otro pago (auditoría 2026-09).
            "haber": -p["amount"],
            "manual": p["supplier_source"] == "manual",
        })
    for c in (credit_memos or []):
        label = f"Credit Memo {c['credit_no']}" if c.get("credit_no") else "Credit Memo"
        entries.append({
            "date": c["credit_date"],
            "kind": "credit",
            "detail": label,
            "document": None,
            "debe": 0.0,
            "haber": c["amount"],
        })
    # Pagos de antes de que Chase estuviera cargado, traídos de la planilla
    # Excel real (proveedores_db.replace_history_payments).
    for hp in (history_payments or []):
        entries.append({
            "date": hp["payment_date"],
            "kind": "history_payment",
            "detail": "OP (Excel)",
            "bank_description": hp.get("detail"),
            "document": None,
            "debe": 0.0,
            "haber": hp["amount"],
        })
    for mp in (manual_payments or []):
        label = "Pago a mano (Caja)" + (f" -- {mp['note']}" if mp.get("note") else "")
        entries.append({
            "date": mp["payment_date"],
            "kind": "manual_payment",
            "detail": label,
            "document": None,
            "debe": 0.0,
            "haber": mp["amount"],
        })
    entries.sort(key=lambda e: (e["date"], 0 if e["kind"] == "invoice" else 1))

    months = []
    by_month = {}
    running = 0.0
    for e in entries:
        running = round(running + e["debe"] - e["haber"], 2)
        e["balance"] = running
        month_key = (int(e["date"][0:4]), int(e["date"][5:7]))
        block = by_month.get(month_key)
        if block is None:
            block = {
                "year": month_key[0],
                "month": month_key[1],
                "month_name": _MONTH_NAMES_ES[month_key[1] - 1],
                "entries": [],
                "saldo_inicial": months[-1]["saldo_final"] if months else 0.0,
            }
            by_month[month_key] = block
            months.append(block)
        block["entries"].append(e)
        block["saldo_final"] = running

    return {
        "months": months,
        "saldo_actual": running,
        "invoice_total": round(sum(inv["amount"] for inv in invoices), 2),
        "invoice_count": len(invoices),
        "payment_total": round(sum(-p["amount"] for p in payments), 2),
        "payment_count": len(payments),
        "credit_total": round(sum(c["amount"] for c in (credit_memos or [])), 2),
        "credit_count": len(credit_memos or []),
        "manual_payment_total": round(sum(mp["amount"] for mp in (manual_payments or [])), 2),
        "manual_payment_count": len(manual_payments or []),
        "history_payment_total": round(sum(hp["amount"] for hp in (history_payments or [])), 2),
        "history_payment_count": len(history_payments or []),
    }


@app.route("/carga-datos/proveedores/guardado/<supplier_key>")
def carga_datos_proveedores_guardado_detalle(supplier_key):
    registry_by_key = {entry["key"]: entry for entry in list_supplier_registry_entries()}

    invoices = proveedores_db.get_supplier_invoices(supplier_key)
    docs_by_filename = {}
    for doc in documents_db.list_all_documents("proveedores"):
        docs_by_filename.setdefault(doc["filename"], doc)
    for inv in invoices:
        inv["document"] = docs_by_filename.get(inv.get("source_filename"))

    label = (
        registry_by_key.get(supplier_key, {}).get("label")
        or (invoices[0]["supplier_label"] if invoices else None)
        or supplier_key
    )
    logo = registry_by_key.get(supplier_key, {}).get("logo")
    payments = chase_db.get_supplier_transactions(supplier_key)
    credit_memos = proveedores_db.get_supplier_credit_memos(supplier_key)
    manual_payments = proveedores_db.get_supplier_manual_payments(supplier_key)
    history_payments = proveedores_db.get_supplier_history_payments(supplier_key)
    supplier_settings = proveedores_db.get_supplier_settings(supplier_key)
    is_dynamic_supplier = supplier_key in load_dynamic_suppliers()

    ledger = _build_supplier_ledger(invoices, payments, credit_memos, manual_payments, history_payments)
    today = date.today()
    open_index = None
    if ledger["months"]:
        open_index = next(
            (i for i, block in enumerate(ledger["months"]) if (block["year"], block["month"]) == (today.year, today.month)),
            len(ledger["months"]) - 1,
        )

    return render_template(
        "carga_datos_proveedores_detalle.html",
        supplier_key=supplier_key,
        supplier_label=label,
        supplier_logo=logo,
        months=ledger["months"],
        open_index=open_index,
        saldo_actual=ledger["saldo_actual"],
        invoice_total=ledger["invoice_total"],
        invoice_count=ledger["invoice_count"],
        payment_total=ledger["payment_total"],
        payment_count=ledger["payment_count"],
        credit_memos=credit_memos,
        credit_total=ledger["credit_total"],
        manual_payments=manual_payments,
        manual_payment_total=ledger["manual_payment_total"],
        history_payment_total=ledger["history_payment_total"],
        history_payment_count=ledger["history_payment_count"],
        supplier_settings=supplier_settings,
        is_dynamic_supplier=is_dynamic_supplier,
        today_iso=today.isoformat(),
        **THEME_BY_KEY["carga_proveedores"],
    )


@app.route("/carga-datos/proveedores/guardado/<supplier_key>/credit-memo", methods=["POST"])
def carga_datos_proveedores_credit_memo(supplier_key):
    """
    Credit memos (2026-09-22, pedido explícito del usuario -- "en H.T se
    pueden cargar credits memo que disminuyen lo que hay que pagar de las
    facturas"). Genérico en el backend (cualquier supplier_key) -- qué
    proveedor puede cargar credit memos se decide con
    proveedores_db.get_supplier_settings/set_supplier_settings (botón
    "Configurar" en el detalle, ver carga_datos_proveedores_configuracion)
    en vez de estar fijo en el código. El chequeo se repite acá del lado
    del servidor -- el template ya oculta el form si no está habilitado,
    pero esto evita que alguien lo cargue posteando directo a la ruta.
    """
    if not proveedores_db.get_supplier_settings(supplier_key)["allow_credit_memos"]:
        flash("Este proveedor no tiene habilitados los credit memos -- activalo desde \"Configurar\".", "error")
        return redirect(url_for("carga_datos_proveedores_guardado_detalle", supplier_key=supplier_key))
    registry_by_key = {entry["key"]: entry for entry in list_supplier_registry_entries()}
    label = registry_by_key.get(supplier_key, {}).get("label") or supplier_key
    credit_date = request.form.get("credit_date")
    credit_no = (request.form.get("credit_no") or "").strip() or None
    amount = request.form.get("amount", type=float)
    if not credit_date or not amount:
        flash("Completá la fecha y el monto del credit memo.", "error")
        return redirect(url_for("carga_datos_proveedores_guardado_detalle", supplier_key=supplier_key))
    proveedores_db.save_credit_memo(supplier_key, label, credit_date, credit_no, amount)
    flash("Credit memo guardado -- ya se refleja en el saldo.", "success")
    return redirect(url_for("carga_datos_proveedores_guardado_detalle", supplier_key=supplier_key))


@app.route("/carga-datos/proveedores/credit-memo/<int:credit_id>/eliminar", methods=["POST"])
def carga_datos_proveedores_credit_memo_eliminar(credit_id):
    supplier_key = request.form.get("supplier_key")
    ok = proveedores_db.delete_credit_memo(credit_id)
    flash("Credit memo eliminado." if ok else "Ese credit memo ya no existe.", "success" if ok else "error")
    return redirect(url_for("carga_datos_proveedores_guardado_detalle", supplier_key=supplier_key))


@app.route("/carga-datos/proveedores/guardado/<supplier_key>/pago-manual", methods=["POST"])
def carga_datos_proveedores_pago_manual(supplier_key):
    """
    Pago a mano vía Caja (2026-09-22, pedido explícito del usuario --
    "que se puedan hacer cargas manuales de pago para proveedores como
    Bimbo, Flori gas, Sam's... deben ir conectados a caja y sumarse en la
    columna de gastos el dia que fueran cargados"). Se guarda como un pago
    más del cuadro de cuenta corriente (Haber) Y, en el mismo movimiento,
    como un gasto en efectivo de Caja de esa fecha (caja_db.
    add_expense_item) -- el id de ese gasto queda guardado junto al pago
    para poder borrar los dos juntos si hace falta. Qué proveedor puede
    usar esto se decide con proveedores_db.get_supplier_settings (botón
    "Configurar" del detalle) -- "los demas se pagan por banco, no hace
    falta" (pedido explícito del usuario, 2026-09-22) -- chequeado acá
    también del lado del servidor, no solo ocultando el form.
    """
    if not proveedores_db.get_supplier_settings(supplier_key)["allow_manual_payments"]:
        flash("Este proveedor no tiene habilitados los pagos a mano -- activalo desde \"Configurar\".", "error")
        return redirect(url_for("carga_datos_proveedores_guardado_detalle", supplier_key=supplier_key))
    registry_by_key = {entry["key"]: entry for entry in list_supplier_registry_entries()}
    label = registry_by_key.get(supplier_key, {}).get("label") or supplier_key
    payment_date = request.form.get("payment_date")
    amount = request.form.get("amount", type=float)
    note = (request.form.get("note") or "").strip() or None
    if not payment_date or not amount:
        flash("Completá la fecha y el monto del pago.", "error")
        return redirect(url_for("carga_datos_proveedores_guardado_detalle", supplier_key=supplier_key))
    expense_detail = f"{label} (pago a proveedor, a mano)" + (f" -- {note}" if note else "")
    expense_item_id = caja_db.add_expense_item(payment_date, amount, expense_detail)
    proveedores_db.save_manual_payment(supplier_key, label, payment_date, amount, note=note, caja_expense_item_id=expense_item_id)
    flash("Pago guardado -- se sumó también a los gastos de Caja de ese día.", "success")
    return redirect(url_for("carga_datos_proveedores_guardado_detalle", supplier_key=supplier_key))


@app.route("/carga-datos/proveedores/pago-manual/<int:payment_id>/eliminar", methods=["POST"])
def carga_datos_proveedores_pago_manual_eliminar(payment_id):
    supplier_key = request.form.get("supplier_key")
    row = proveedores_db.delete_manual_payment(payment_id)
    if row and row.get("caja_expense_item_id"):
        caja_db.delete_expense_item(row["caja_expense_item_id"])
    flash("Pago eliminado (también de los gastos de Caja)." if row else "Ese pago ya no existe.", "success" if row else "error")
    return redirect(url_for("carga_datos_proveedores_guardado_detalle", supplier_key=supplier_key))


@app.route("/carga-datos/proveedores/guardado/<supplier_key>/configuracion", methods=["POST"])
def carga_datos_proveedores_configuracion(supplier_key):
    """
    Configuración individual por proveedor (2026-09-22, pedido explícito
    del usuario: "cada proveedor tenga un tipo de configuracion
    individual en la que tocando un boton se pueda agregar de que se le
    hacen pagos en efectivo o recibe credits memo") -- admin-only, mismo
    criterio que ocultar/mostrar. Dos checkboxes; sin marcar = apagado.
    """
    if not _require_admin("Solo un administrador puede cambiar la configuración de un proveedor."):
        return redirect(url_for("carga_datos_proveedores_guardado_detalle", supplier_key=supplier_key))
    allow_manual_payments = request.form.get("allow_manual_payments") == "on"
    allow_credit_memos = request.form.get("allow_credit_memos") == "on"
    proveedores_db.set_supplier_settings(supplier_key, allow_manual_payments, allow_credit_memos)
    return redirect(url_for("carga_datos_proveedores_guardado_detalle", supplier_key=supplier_key))


@app.route("/carga-datos/proveedores/reglas/guardar", methods=["POST"])
def proveedores_pago_reglas_guardar():
    """
    Alta/edición de una regla keyword -> proveedor (proveedores_pago_rules.
    json) -- mismo patrón admin-only que las reglas de Chase. `sheet_name`
    viaja como value real del <select> de proveedor (siempre uno de
    list_supplier_registry_entries -- así una regla nueva nunca puede
    apuntar a un proveedor que no existe).
    """
    if not _require_admin("Solo un administrador puede gestionar las reglas de pago a proveedores."):
        return redirect(url_for("carga_datos_proveedores_reglas"))

    keyword = request.form.get("keyword", "")
    sheet_name = request.form.get("sheet_name", "")
    index = request.form.get("index", "").strip()
    expected_keyword = request.form.get("expected_keyword") or None
    expected_sheet_name = request.form.get("expected_sheet_name") or None

    try:
        if not index:
            proveedores_pago_rules.add_dynamic_rule(keyword, sheet_name)
            flash("Regla creada.", "success")
        else:
            proveedores_pago_rules.edit_dynamic_rule_by_index(
                index, keyword, sheet_name, expected_keyword, expected_sheet_name
            )
            flash("Regla actualizada.", "success")
        _recategorize_all_chase_and_flash()
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("carga_datos_proveedores_reglas"))


@app.route("/carga-datos/proveedores/reglas/eliminar", methods=["POST"])
def proveedores_pago_reglas_eliminar():
    if not _require_admin("Solo un administrador puede gestionar las reglas de pago a proveedores."):
        return redirect(url_for("carga_datos_proveedores_reglas"))

    index = request.form.get("index", "").strip()
    expected_keyword = request.form.get("expected_keyword") or None
    expected_sheet_name = request.form.get("expected_sheet_name") or None

    try:
        proveedores_pago_rules.delete_dynamic_rule_by_index(index, expected_keyword, expected_sheet_name)
        flash("Regla eliminada.", "success")
        _recategorize_all_chase_and_flash()
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("carga_datos_proveedores_reglas"))


@app.route("/proveedores")
def proveedores():
    return render_template("proveedores.html", **THEME_BY_KEY["proveedores"])


def _proveedores_error(message):
    return _error_response(message)


def _group_by_supplier_message(prefix, items, unit_label):
    """
    Arma un aviso corto tipo "No se pudieron cargar 5 factura(s) de
    Colonial, 3 factura(s) de Coca-Cola." -- agrupa por proveedor para que
    el usuario sepa dónde mirar sin tener que abrir el Excel a buscar,
    pero sin nombres de archivo (`items` es la lista `failed`/`unmatched`
    de proveedores.py, cada uno un dict con clave "supplier" o None si no
    se pudo identificar).
    """
    counts = {}
    unknown = 0
    for item in items:
        supplier = item.get("supplier")
        if supplier:
            counts[supplier] = counts.get(supplier, 0) + 1
        else:
            unknown += 1
    parts = [f"{n} {unit_label} de {name}" for name, n in sorted(counts.items())]
    if unknown:
        parts.append(f"{unknown} {unit_label} sin proveedor identificado")
    return f"{prefix}: {', '.join(parts)}."


def _proveedores_success(temp_path, download_name, notices):
    """
    `notices` es una lista de tuplas (level, message) -- level es
    "warning" para algo informativo que no requiere acción (ej. una
    factura que ya estaba cargada, se omite sola) o "error" para algo que
    sí requiere que el usuario cargue esa factura a mano. Se combinan en
    un solo aviso porque _success_response solo lleva uno; el nivel final
    es "error" si alguno de los dos lo es.
    """
    if not notices:
        return _success_response(temp_path, download_name)
    combined = " ".join(message for _level, message in notices)
    worst_level = "error" if any(level == "error" for level, _message in notices) else "warning"
    return _success_response(temp_path, download_name, notice=combined, notice_level=worst_level)


@app.route("/proveedores/facturas", methods=["POST"])
def proveedores_facturas():
    master_upload = request.files.get("master_file")
    pdf_uploads = request.files.getlist("pdf_files")
    if master_upload is None or not master_upload.filename:
        return _proveedores_error("Seleccioná el Excel Ledger.")
    if not pdf_uploads or not any(u.filename for u in pdf_uploads):
        return _proveedores_error("Seleccioná uno o más PDF de facturas.")

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        pdf_paths = _save_uploads_to_workspace(pdf_uploads, workdir=workdir)
        temp_path, summary = append_supplier_invoices(master_path, pdf_paths)
    except Exception as exc:
        return _proveedores_error(f"Error: {exc}")

    notices = []

    duplicate_files = []
    for result in summary["batch_results"]:
        duplicate_files.extend(result.get("duplicates_skipped") or [])
    if duplicate_files:
        notices.append((
            "warning",
            f"{len(duplicate_files)} factura(s) ya estaban cargadas y se omitieron solas.",
        ))

    if summary["failed"]:
        notices.append((
            "error",
            _group_by_supplier_message("No se pudieron cargar", summary["failed"], "factura(s)"),
        ))
        if any(item.get("partial_write") for item in summary["failed"]):
            # Caso raro: la factura falló DESPUÉS de que ya se insertó una
            # fila y se repuntearon fórmulas de RESUMEN COMPRAS -- a
            # diferencia de una falla normal (nada se tocó), acá la hoja del
            # proveedor sí quedó modificada. Avisar explícito para que se
            # revise la hoja completa, no solo se recargue la factura.
            notices.append((
                "error",
                "Al menos una de esas facturas falló después de modificar la hoja del "
                "proveedor (no antes) — revisá esa hoja completa a mano, no le cargues "
                "la factura de nuevo sin mirarla primero.",
            ))

    resumen_warnings = summary.get("resumen_warnings") or []
    if resumen_warnings:
        if any(item["status"] == "sheet_not_found" for item in resumen_warnings):
            notices.append((
                "warning",
                "Las facturas se cargaron bien, pero no se encontró la hoja RESUMEN COMPRAS "
                "para actualizar el resumen mensual.",
            ))
        else:
            suppliers = ", ".join(sorted({item["supplier"] for item in resumen_warnings if item["supplier"]}))
            notices.append((
                "warning",
                f"{len(resumen_warnings)} factura(s) se cargaron bien, pero no se sumaron en "
                f"RESUMEN COMPRAS ({suppliers}). Revisalo a mano.",
            ))

    return _proveedores_success(temp_path, master_filename, notices)


@app.route("/proveedores/pagos", methods=["POST"])
def proveedores_pagos():
    master_upload = request.files.get("master_file")
    bank_upload = request.files.get("bank_file")
    if master_upload is None or not master_upload.filename:
        return _proveedores_error("Seleccioná el Excel Ledger.")
    if bank_upload is None or not bank_upload.filename:
        return _proveedores_error("Seleccioná el extracto de Chase ya categorizado.")

    try:
        workdir = _new_workspace_dir()
        master_path, master_filename = _save_upload_to_workspace(master_upload, workdir=workdir)
        bank_path, _bank_filename = _save_upload_to_workspace(bank_upload, workdir=workdir)
        temp_path, summary = append_supplier_payments(master_path, bank_path)
    except Exception as exc:
        return _proveedores_error(f"Error: {exc}")

    notices = []

    duplicate_count = sum(len(result.get("duplicates_skipped") or []) for result in summary["batch_results"])
    if duplicate_count:
        notices.append((
            "warning",
            f"{duplicate_count} pago(s) ya estaban cargados y se omitieron solos.",
        ))

    if summary["unmatched"]:
        notices.append((
            "error",
            _group_by_supplier_message("No se pudieron cargar", summary["unmatched"], "pago(s)"),
        ))

    return _proveedores_success(temp_path, master_filename, notices)


def _dynamic_wizard_admin_error():
    return jsonify({"error": "Solo un administrador puede agregar proveedores nuevos."}), 403


@app.route("/proveedores/nuevo")
def proveedores_nuevo():
    if not current_user.is_admin:
        flash("Solo un administrador puede agregar proveedores nuevos.", "error")
        return redirect(url_for("carga_datos_proveedores_guardado"))
    return render_template(
        "proveedores_nuevo.html",
        field_labels=DYNAMIC_FIELD_LABELS,
        fields=DYNAMIC_FIELDS,
        **THEME_BY_KEY["proveedores"],
    )


@app.route("/proveedores/nuevo/analizar", methods=["POST"])
def proveedores_nuevo_analizar():
    """
    Paso sin estado del asistente: recibe la factura de muestra + los 3
    valores tipeados por el usuario (siempre los 3, en cada llamada) + lo
    que ya se desambiguó en vueltas anteriores, y devuelve el análisis de
    nuevo -- ver proveedores_dynamic_extractors.analyze_sample. El servidor
    no guarda nada entre llamadas; el navegador mantiene el PDF y las
    elecciones ya hechas y las reenvía cada vez.
    """
    if not current_user.is_admin:
        return _dynamic_wizard_admin_error()

    sample_upload = request.files.get("sample_pdf")
    if sample_upload is None or not sample_upload.filename:
        return jsonify({"error": "Subí una factura de muestra en PDF."}), 400

    sample_values = {
        "invoice_no": request.form.get("sample_invoice_no", ""),
        "date": request.form.get("sample_date", ""),
        "amount": request.form.get("sample_amount", ""),
    }
    try:
        chosen_occurrence_index = json.loads(request.form.get("chosen_occurrence_index_json") or "{}")
    except (TypeError, ValueError):
        chosen_occurrence_index = {}

    try:
        sample_path, _filename = _save_upload_to_workspace(sample_upload)
        result = analyze_dynamic_sample(sample_path, sample_values, chosen_occurrence_index)
    except DYNAMIC_PDF_READ_EXCEPTIONS as exc:
        return jsonify({"error": f"No se pudo leer el PDF: {exc}"}), 400

    return jsonify(result)


@app.route("/proveedores/nuevo/probar", methods=["POST"])
def proveedores_nuevo_probar():
    """
    Corre la regla ya resuelta (todavía sin guardar) contra una SEGUNDA
    factura de muestra real, como dry run -- no persiste nada, solo
    devuelve lo que se leería para que el usuario lo compare a ojo contra
    esa factura antes de guardar de verdad.
    """
    if not current_user.is_admin:
        return _dynamic_wizard_admin_error()

    sample_upload = request.files.get("sample_pdf")
    if sample_upload is None or not sample_upload.filename:
        return jsonify({"error": "Subí una segunda factura de muestra en PDF."}), 400

    try:
        resolved_fields = json.loads(request.form.get("resolved_fields_json") or "{}")
        rule_fields = build_dynamic_rule_fields(resolved_fields)
    except (TypeError, ValueError) as exc:
        return jsonify({"error": str(exc)}), 400

    try:
        sample_path, _filename = _save_upload_to_workspace(sample_upload)
        extracted = extract_with_dynamic_rule(sample_path, {"fields": rule_fields})
    except DYNAMIC_PDF_READ_EXCEPTIONS as exc:
        return jsonify({"error": str(exc)}), 400

    return jsonify(
        {
            "invoice_no": extracted["invoice_no"],
            "date": extracted["date"].strftime("%d/%m/%Y"),
            "amount": f'{extracted["amount"]:.2f}',
        }
    )


@app.route("/proveedores/nuevo/guardar", methods=["POST"])
def proveedores_nuevo_guardar():
    if not _require_admin("Solo un administrador puede agregar proveedores nuevos."):
        return redirect(url_for("proveedores_nuevo"))

    label = request.form.get("label", "").strip()
    sheet_name = request.form.get("sheet_name", "").strip()
    resumen_label = request.form.get("resumen_label", "").strip()
    detect_keyword = request.form.get("detect_keyword", "").strip()

    try:
        resolved_fields = json.loads(request.form.get("resolved_fields_json") or "{}")
        rule_fields = build_dynamic_rule_fields(resolved_fields)
        clave = add_dynamic_supplier(
            {
                "label": label,
                "sheet_name": sheet_name,
                "resumen_label": resumen_label,
                "detect_keyword": detect_keyword,
                "fields": rule_fields,
            },
            created_by=current_user.id,
        )
    except (TypeError, ValueError) as exc:
        flash(str(exc), "error")
        return redirect(url_for("proveedores_nuevo"))

    flash(
        f"Proveedor \"{label}\" agregado. Ya podés cargar sus facturas desde \"Cargar Facturas\".",
        "success",
    )
    return redirect(url_for("proveedores_nuevo"))


@app.route("/proveedores/nuevo/eliminar", methods=["POST"])
def proveedores_nuevo_eliminar():
    """
    Eliminar un proveedor agregado dinámicamente (2026-09-22) -- llamado
    tanto desde "Configurar" en el detalle del proveedor (caso normal)
    como, si hiciera falta, desde cualquier otro lado que postee a esta
    misma ruta. Siempre vuelve a la Planilla -- el detalle del proveedor
    ya borrado no tiene sentido seguir mostrándolo.
    """
    if not _require_admin("Solo un administrador puede eliminar un proveedor."):
        return redirect(url_for("carga_datos_proveedores_guardado"))

    clave = request.form.get("clave", "").strip()
    try:
        delete_dynamic_supplier(clave)
        flash("Proveedor eliminado.", "success")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("carga_datos_proveedores_guardado"))


@app.route("/balance-mensual")
def balance_mensual():
    # Ya no está en TOOLS (ver esa lista, comentario "no confirmado por el
    # usuario") -- el tema del módulo se pasa directo acá en vez de por
    # THEME_BY_KEY (que ya no tiene esta clave) para que la ruta siga
    # andando si alguien entra por URL directa.
    return render_template("balance_mensual.html", accent="#78350F", accent_soft="#F5E4CE")


@app.route("/balance-mensual/procesar", methods=["POST"])
def balance_mensual_procesar():
    balance_upload = request.files.get("balance_file")
    mayor_uploads = request.files.getlist("mayor_files")
    if balance_upload is None or not balance_upload.filename:
        return _error_response("Seleccioná el Excel de Balance del mes.")
    if not any(upload and upload.filename for upload in mayor_uploads):
        return _error_response("Seleccioná uno o más Mayores para reemplazar.")

    try:
        workdir = _new_workspace_dir()
        balance_path, balance_filename = _save_upload_to_workspace(balance_upload, workdir=workdir)
        mayor_paths = _save_uploads_to_workspace(mayor_uploads, workdir=workdir)
        temp_path, summary = replace_mayor_sheets(balance_path, mayor_paths)
    except Exception as exc:
        return _error_response(f"Error: {exc}")

    if temp_path is None:
        return _error_response("Ninguno de los Mayores subidos corresponde a una hoja soportada todavía.")

    notice_parts = []
    if summary["unsupported"]:
        names = ", ".join(sorted(set(summary["unsupported"])))
        notice_parts.append(f"Hoja(s) todavía sin soporte automático (cargalas a mano): {names}.")
    if summary["unmatched"]:
        notice_parts.append(
            f"{len(summary['unmatched'])} Mayor(es) subido(s) no corresponden a ninguna cuenta conocida -- revisá que sean los archivos correctos."
        )

    return _success_response(
        temp_path, balance_filename, notice=" ".join(notice_parts) or None, notice_level="warning"
    )


def _defer_reload_while_jobs_run():
    """
    Con el reloader de debug, guardar un .py reinicia el servidor y corta la
    carga en segundo plano que esté corriendo: pasó el 2026-10-06 con 30
    Reportes Diarios de septiembre ("se cortaba la carga y tenía que volverla
    a empezar diciéndome que el servidor se había caído"). Acá el reinicio
    espera a que no quede ninguna carga corriendo (jobs.has_running_jobs);
    mientras tanto la página sigue andando con el código de antes. Toca una
    parte interna de Werkzeug (ReloaderLoop.trigger_reload): si cambia en
    otra versión, se avisa y se reinicia como siempre.
    """
    try:
        from werkzeug import _reloader

        loops = {_reloader.ReloaderLoop, *_reloader.reloader_loops.values()}
    except (ImportError, AttributeError) as exc:
        print(f"[webapp] el reinicio por cambios de código no espera a las cargas ({exc})")
        return

    def waiting(original):
        def trigger_reload(self, filename):
            if jobs.has_running_jobs():
                print(f" * Cambió {os.path.basename(filename)}: el servidor se reinicia cuando termine la carga en curso.")
                while jobs.has_running_jobs():
                    time.sleep(2)
            original(self, filename)
        return trigger_reload

    for loop in loops:
        if "trigger_reload" in vars(loop):
            loop.trigger_reload = waiting(vars(loop)["trigger_reload"])


if __name__ == "__main__":
    # Render (y cualquier plataforma similar) fija la variable de entorno
    # PORT y espera que el proceso escuche en 0.0.0.0 -- escuchar solo en
    # 127.0.0.1 (el default de Flask sin "host" explícito) deja el server
    # arrancado pero inalcanzable desde afuera del contenedor. En la PC del
    # usuario, sin PORT seteada, esto sigue exactamente igual que siempre
    # (127.0.0.1:5000).
    port_env = os.environ.get("PORT")
    port = int(port_env) if port_env else 5000
    host = "0.0.0.0" if port_env else "127.0.0.1"

    # El modo debug (reloader + debugger interactivo de Werkzeug) prendido
    # por default preserva el flujo de desarrollo local de siempre -- ahí no
    # es un riesgo real porque solo escucha en 127.0.0.1 (nadie fuera de esta
    # PC llega a la página del debugger, que permite correr código Python
    # arbitrario desde el navegador). Pero una vez que el host pasa a
    # 0.0.0.0 (Render) ese mismo debugger quedaría expuesto a cualquiera, así
    # que ahí el default cambia a apagado -- BRADENTON_DEBUG sigue pudiendo
    # forzarlo a "1" a mano si hiciera falta debuggear ahí puntualmente.
    debug_mode = os.environ.get("BRADENTON_DEBUG", "1" if host == "127.0.0.1" else "0") == "1"
    # threaded=True -- necesario para las cargas en segundo plano (ver
    # jobs.py/CLAUDE.md): mientras un trabajo pesado corre en su propio
    # hilo, el navegador sondea /jobs/<id>/status en paralelo -- sin esto,
    # el servidor de desarrollo de Flask atiende una sola request a la vez
    # y ese sondeo quedaría trabado detrás de la carga que se supone que
    # tiene que poder consultar mientras corre.
    #
    # También en ::1 (IPv6 de esta misma PC), 2026-10-02: el navegador
    # resuelve "localhost" primero a ::1 y, como el servidor escuchaba solo
    # en 127.0.0.1, cada pedido esperaba ~300 ms antes de reintentar por IPv4
    # (medido: 330 ms contra 30 ms por pedido). Eso se pagaba en cada página,
    # cada guardado y cada sondeo de progreso. Solo en la PC (no en Render) y
    # solo en el proceso que atiende de verdad (con el reloader de debug, el
    # proceso padre solo vigila archivos). Si la PC no tiene IPv6, se sigue
    # igual que antes.
    if host == "127.0.0.1" and (not debug_mode or os.environ.get("WERKZEUG_RUN_MAIN") == "true"):
        try:
            from werkzeug.serving import make_server

            _ipv6_server = make_server("::1", port, app, threaded=True)
            threading.Thread(target=_ipv6_server.serve_forever, daemon=True, name="ipv6-loopback").start()
        except OSError as exc:
            print(f"[webapp] sin escucha en ::1 ({exc}); localhost va a responder más lento")
    if debug_mode and os.environ.get("WERKZEUG_RUN_MAIN") == "true":
        _defer_reload_while_jobs_run()
    app.run(debug=debug_mode, host=host, port=port, threaded=True)
