"""
BradentonApp — versión web (Flask).

Primer paso de la migración de escritorio (Tkinter) a web, para poder
correr como servicio en Render y entrar en Toolbox. Reusa la lógica de
negocio ya extraída a módulos sin dependencia de Tkinter (chase_rules.py,
cmv_costo.py, etc.) — nunca reimplementa esa lógica acá.

Un módulo por vez: hoy solo está Chase Bank. El resto se va sumando
igual que el desktop, probando cada uno antes de seguir con el próximo.
"""

import json
import os
import re
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from urllib.parse import quote

from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, session, url_for
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
from chase_rules import (
    add_dynamic_rule as add_chase_rule,
    build_chase_export_workbook,
    delete_dynamic_rule_by_index as delete_chase_custom_rule,
    delete_master_rule_by_index as delete_chase_master_rule,
    edit_dynamic_rule_by_index as edit_chase_custom_rule,
    edit_master_rule_by_index as edit_chase_master_rule,
    extract_chase_transactions,
    list_display_rules as list_chase_display_rules,
)
import caja_db
from caja import build_caja_export_pdf, build_caja_export_workbook, build_month_report_from_db as build_caja_month_report
from cmv_costo import _consolidate_department_files, update_master_costo_todos_bulk
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
from proveedores import _PDF_EXTRACTION_EXCEPTIONS
from proveedores_dynamic_extractors import (
    FIELD_LABELS as DYNAMIC_FIELD_LABELS,
    FIELDS as DYNAMIC_FIELDS,
    PDF_READ_EXCEPTIONS as DYNAMIC_PDF_READ_EXCEPTIONS,
    add_dynamic_supplier,
    analyze_sample as analyze_dynamic_sample,
    build_rule_fields as build_dynamic_rule_fields,
    delete_dynamic_supplier,
    extract_with_dynamic_rule,
    list_dynamic_suppliers_display,
)
from reporte_diario import (
    build_store_info_export_pdf,
    build_store_info_export_workbook,
    extract_department_sales_for_day,
    extract_lottery_department_fields_from_pdf,
    extract_lottery_receipt_fields_from_sales_report,
    extract_store_info_for_day,
    group_department_sales,
    process_lottery,
    process_reporte_diario,
    process_store_info,
)
import reportes_db
import lottery_db
import gettel_db
import cmv_db
import documents_db
import proveedores_db
import horas_trabajo_db
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


@app.before_request
def require_login():
    if request.endpoint in ("login", "static") or request.endpoint is None:
        return None
    if not current_user.is_authenticated:
        return redirect(url_for("login"))
    return None


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


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("carga_datos_index"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = auth.verify_user(username, password)
        if user is None:
            flash("Usuario o contraseña incorrectos.", "error")
        else:
            remember = bool(request.form.get("remember"))
            login_user(WebUser(username, user.get("is_admin", False)), remember=remember)
            return redirect(url_for("carga_datos_index"))

    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


@app.route("/perfil/password", methods=["GET", "POST"])
@login_required
def perfil_password():
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
                flash("Contraseña actualizada correctamente.", "success")
            except ValueError as exc:
                flash(str(exc), "error")
        return redirect(url_for("perfil_password"))
    return render_template("perfil_password.html")


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
                    auth.set_password(request.form.get("username", "").strip(), new_password)
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
        "label": "Reportes Diarios",
        "url": "/carga-datos/reporte-diario",
        "description": "Subí el PDF de cierre diario — Departamentos y Store Info quedan guardados solos, día por día.",
        "accent": "#0284C7",
        "accent_soft": "#D7EFFB",
    },
    {
        "key": "carga_lottery",
        "code": "LT",
        "icon": _ICON_TICKET,
        "label": "Lottery",
        "url": "/carga-datos/lottery",
        "description": "Subí el Daily Sales Report — bloques de 7 días con Subtotal y Debito calculados solos, igual que el Excel.",
        "accent": "#0284C7",
        "accent_soft": "#D7EFFB",
    },
    {
        "key": "carga_chase",
        "code": "CH",
        "icon": _ICON_BANK,
        "label": "Chase Bank",
        "url": "/carga-datos/chase",
        "description": "Subí el extracto de Chase — cada movimiento queda categorizado y guardado solo, sin generar ningún Excel.",
        "accent": "#16A34A",
        "accent_soft": "#DCF3E3",
    },
    {
        "key": "carga_eft",
        "code": "EFT",
        "icon": _ICON_EXCHANGE,
        "label": "EFT y Cupones",
        "url": "/carga-datos/eft",
        "description": "Subí el PDF de EFT y el reporte mensual de Cupones — se cruzan solos por DDC, sin generar ningún Excel.",
        "accent": "#3B5BDB",
        "accent_soft": "#DDE3FA",
    },
    {
        "key": "carga_caja",
        "code": "CJ",
        "icon": _ICON_REGISTER,
        "label": "Caja",
        "url": "/carga-datos/caja",
        "description": "Depósitos, Food Truck/Ice y Lottery del mes, completados solos con lo que ya guardaron Chase y Lottery — sin subir nada nuevo.",
        "accent": "#EA580C",
        "accent_soft": "#FCE3D2",
    },
    {
        "key": "carga_gettel",
        "code": "GT",
        "icon": _ICON_CAR,
        "label": "Gettel / Toyota",
        "url": "/carga-datos/gettel",
        "description": "Subí el Excel o PDF de cupones de Gettel/Toyota — Monto y Galones por día quedan guardados solos, sin generar ningún Excel.",
        "accent": "#0D9488",
        "accent_soft": "#D6F1EE",
    },
    {
        "key": "carga_cmv",
        "code": "CMV",
        "icon": _ICON_COINS,
        "label": "CMV",
        "url": "/carga-datos/cmv",
        "description": "Costo por UPC y ventas mensuales por departamento — guardados solos, sin generar ningún Excel.",
        "accent": "#7C3AED",
        "accent_soft": "#E9E0FC",
    },
    {
        "key": "carga_proveedores",
        "code": "PR",
        "icon": _ICON_TRUCK,
        "label": "Proveedores",
        "url": "/carga-datos/proveedores",
        "description": "Subí las facturas de compra — se guardan solas por proveedor, sin generar ningún Excel.",
        "accent": "#DB2777",
        "accent_soft": "#FBD9EA",
    },
    {
        "key": "carga_horas",
        "code": "HT",
        "icon": _ICON_CLOCK,
        "label": "Horas de Trabajo",
        "url": "/carga-datos/horas-trabajo",
        "description": "Subí el reporte semanal de Clock In/Out — horas por empleado, sueldo y descuentos calculados solos, sin generar ningún Excel.",
        "accent": "#0891B2",
        "accent_soft": "#D3F0F4",
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
    return {"SEARCH_INDEX": CONTROLS + CARGA_DATOS_TOOLS, "today": date.today()}


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
    if len(query) < 2:
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

_REPORTE_DIARIO_FILENAME_DATE_RE = re.compile(r"(\d{1,2})[.\-](\d{1,2})\s*$")
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
    """
    if hasattr(target_date, "date") and not isinstance(target_date, date):
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
    Ya no hay que elegir un lado al entrar -- pedido explícito del usuario
    (2026-09-12): "ya solo queda lo de cargar datos en esa pagina no vamos a
    necesitar los excels, ya estan guardado el proyecto donde estan los
    excels" (la copia congelada, ver CLAUDE.md). Carga de Datos pasa a ser
    el destino directo; Excels/Controles siguen andando igual para lo que
    todavía no se convirtió, alcanzables por búsqueda o por URL directa.
    """
    return redirect(url_for("carga_datos_index"))


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
    upload_tools = [tool for tool in CARGA_DATOS_TOOLS if tool["key"] != "carga_caja"]
    return render_template("carga_datos_index.html", tools=upload_tools)


@app.route("/carga-datos/controles")
def carga_datos_controles():
    """
    Scaffold vacío, mismo criterio que /controles cuando arrancó sin ningún
    módulo -- Carga de Datos pasa a tener su propia pareja Herramientas/
    Controles (pedido explícito del usuario 2026-09-11), sin mezclar con la
    de Excels. Sin módulos propios todavía.
    """
    return render_template("carga_datos_controles.html")


@app.route("/carga-datos/reporte-diario")
def carga_datos_reporte_diario():
    """
    Carga directa de Reporte Diario del lado Carga de Datos -- solo PDF, sin
    ningún campo de Excel (pedido explícito del usuario 2026-09-11, parte del
    plan de dejar este lado autosuficiente para subir datos sin depender de
    Herramientas). Ver carga_datos_reporte_diario_subir más abajo.
    """
    return render_template("carga_datos_reporte_diario.html", **THEME_BY_KEY["reporte"])


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
    """
    try:
        days_complete = set()
        days_partial = set()
        days_subtotal_mismatch = set()
        files_unreadable = 0
        date_mismatches = 0
        first_date = None

        for index, pdf_path in enumerate(pdf_paths, start=1):
            filename = os.path.basename(pdf_path)
            filename_day_month = _reporte_filename_day_month(filename)
            day_date = None
            got_departments = False
            got_store_info = False
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
                    reportes_db.upsert_store_info(candidate_date, result["fields"], source="ocr", pdf_filename=pdf_relpath)
                    day_date = candidate_date
                    got_store_info = True
                except Exception as exc:
                    if "no coincide con la fecha del nombre" in str(exc):
                        file_had_mismatch = True
                    print(f"[carga-datos/reporte-diario] store info de {pdf_path}: {exc}")

            if not file_had_mismatch:
                try:
                    fields = extract_lottery_department_fields_from_pdf(pdf_path)
                    _check_filename_date(fields["report_date"])
                    lottery_relpath = lottery_db.store_pdf_copy(fields["report_date"], pdf_path, filename)
                    lottery_db.upsert_department_fields(
                        fields["report_date"], fields["online_count"], fields["online_net_sales"],
                        fields["skoff_count"], fields["skoff_net_sales"], source="ocr", pdf_filename=lottery_relpath,
                    )
                except Exception as exc:
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
                elif got_departments or got_store_info:
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

        if not parts:
            notice, level = "No se pudo guardar nada de este lote.", "error"
        else:
            notice = " ".join(parts)
            level = "warning" if (days_partial or files_unreadable or days_subtotal_mismatch) else "success"

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
    job_id = jobs.create_job(len(pdf_paths))
    threading.Thread(target=_run_carga_datos_reporte_diario_job, args=(job_id, pdf_paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(pdf_paths)})


@app.route("/carga-datos/lottery")
def carga_datos_lottery():
    """Carga directa de Lottery (Daily Sales Report) del lado Carga de Datos -- solo PDF."""
    return render_template("carga_datos_lottery.html", **THEME_BY_KEY["lottery"])


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
    job_id = jobs.create_job(len(pdf_paths))
    threading.Thread(target=_run_carga_datos_lottery_job, args=(job_id, pdf_paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(pdf_paths)})


def _run_carga_datos_lottery_job(job_id, pdf_paths):
    """Corre en su propio hilo -- mismo patrón que _run_carga_datos_reporte_diario_job, ver ese docstring."""
    try:
        saved_dates = []
        failed = 0
        date_mismatches = 0
        for index, pdf_path in enumerate(pdf_paths, start=1):
            filename = os.path.basename(pdf_path)
            try:
                fields = extract_lottery_receipt_fields_from_sales_report(pdf_path)
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
        if not parts:
            notice, level = "No se pudo guardar nada de este lote.", "error"
        else:
            notice, level = " ".join(parts), ("warning" if failed else "success")

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


# Etiquetas de las 11 columnas que suma la fila Subtotal -- pedido
# explícito del usuario (2026-09-16, segunda tanda): "los totales al final
# del bloque... que te muestren la sumatoria de valores que hacen para
# llegar a ese número". Mismo orden que lottery_db._SUBTOTAL_SUM_FIELDS.
_LOTTERY_SUBTOTAL_LABELS = {
    "online_count": "Count (ONLINE)",
    "online_net_sales": "Sales $ (ONLINE)",
    "sales": "Sales (Terminal)",
    "pagos": "Pagos (Terminal)",
    "comis": "Comis (Terminal)",
    "prize_free_plays": "Prize FP (Terminal)",
    "total_comm": "Total Comm (Terminal)",
    "pays_units": "Pays U (SKOFF)",
    "pays_amount": "Pays $ (SKOFF)",
    "skoff_sales_amount": "Sales Amt (SKOFF)",
    "sales_comm": "Sales Comm (SKOFF)",
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
                ("Sales $ (ONLINE, Subtotal)", subtotal.get("online_net_sales")),
                ("Sales (Terminal, Subtotal) — resta", -subtotal["sales"] if subtotal.get("sales") is not None else None),
            ],
            "sales": [  # F =+F+G+I+K+10
                ("Sales (Terminal, Subtotal)", subtotal.get("sales")),
                ("Pagos (Terminal, Subtotal)", subtotal.get("pagos")),
                ("Comis (Terminal, Subtotal)", subtotal.get("comis")),
                ("Prize FP (Terminal, Subtotal)", subtotal.get("prize_free_plays")),
                ("Cargo fijo", 10),
            ],
            "pays_amount": [  # Q
                ("Pays $ (SKOFF, Subtotal)", subtotal.get("pays_amount")),
                ("Sales Amt (SKOFF, Subtotal)", subtotal.get("skoff_sales_amount")),
                ("Sales Comm (SKOFF, Subtotal)", subtotal.get("sales_comm")),
            ],
            "net_debit": [  # V =+F+Q (de la fila Debito)
                ("Sales (F, Debito)", debito.get("sales")),
                ("Pays $ (Q, Debito)", debito.get("pays_amount")),
            ],
        }
        for d in block["days"]:
            d["cuenta_final_breakdown"] = [  # X =-G-Q
                ("Pagos (G), resta", -d["pagos"] if d.get("pagos") is not None else None),
                ("Pays $ (Q), resta", -d["pays_amount"] if d.get("pays_amount") is not None else None),
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
        flash(message, "success")
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


@app.route("/carga-datos/lottery/documentos")
def carga_datos_lottery_documentos():
    """PDFs diarios ya guardados este mes -- ver lottery_db.get_month_pdf_list."""
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    pdfs = lottery_db.get_month_pdf_list(year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "lottery_documentos.html",
        pdfs=pdfs,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["lottery"],
    )


@app.route("/carga-datos/lottery/resumen-mensual")
def carga_datos_lottery_resumen_mensual():
    """
    El PDF de resumen mensual de Lottery -- solo para guardarlo y poder
    verlo después, no se lee ni se procesa (a diferencia del cuadro
    semanal). Reusa documents_db.py -- mismo patrón que EFT/Gettel/CMV.
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    docs = documents_db.list_documents("lottery_resumen_mensual", year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "lottery_resumen_mensual.html",
        docs=docs,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["lottery"],
    )


@app.route("/carga-datos/lottery/resumen-mensual/subir", methods=["POST"])
def carga_datos_lottery_resumen_mensual_subir():
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    upload = request.files.get("resumen_file")
    if not year or not month or not (1 <= month <= 12):
        flash("Elegí a qué mes corresponde este resumen.", "error")
        return redirect(url_for("carga_datos_lottery_resumen_mensual"))
    if upload is None or not upload.filename:
        flash("Seleccioná el PDF del resumen mensual.", "error")
        return redirect(url_for("carga_datos_lottery_resumen_mensual", year=year, month=month))

    path, filename = _save_upload_to_workspace(upload)
    documents_db.store_document("lottery_resumen_mensual", path, filename, year, month)
    flash("Resumen mensual guardado.", "success")
    return redirect(url_for("carga_datos_lottery_resumen_mensual", year=year, month=month))


@app.route("/controles")
def controles():
    return render_template("controles_index.html", controls=CONTROLS)


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
                result = check_store_info_monthly(cierre_path, pdf_path)
                department_result = check_department_sales_monthly(ventas_path, pdf_path)
            except Exception as exc:
                flash(f"Error: {exc}", "error")
                result = None
                department_result = None
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

    inserted, updated = chase_db.upsert_transactions(rows, source_filename=filename)
    skipped = total_rows - len(rows)
    uncategorized = sum(1 for row in rows if not row["detalle"])

    parts = [f"{len(rows)} movimiento(s) guardado(s) ({inserted} nuevo(s), {updated} actualizado(s))."]
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
        known_details=chase_db.list_known_details(),
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
    """
    posting_date = request.form.get("posting_date", "").strip()
    description = request.form.get("description", "")
    amount_raw = request.form.get("amount", "").strip()
    detalle = request.form.get("detalle", "").strip()
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)

    try:
        amount = float(amount_raw)
    except ValueError:
        flash("No se pudo identificar el movimiento (monto inválido).", "error")
        return redirect(url_for("chase_historial", year=year, month=month))

    ok = chase_db.set_manual_detalle(posting_date, description, amount, detalle)
    if ok:
        flash("Detalle guardado." if detalle else "Detalle borrado.", "success")
    else:
        flash("No se encontró ese movimiento -- puede que ya no esté guardado.", "error")
    return redirect(url_for("chase_historial", year=year, month=month))


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
            reportes_db.upsert_store_info(result["date"], result["fields"], source="ocr", pdf_filename=pdf_relpath)
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

    return render_template(
        "reporte_historial.html",
        overview=overview,
        department_totals=department_totals,
        department_groups=department_groups,
        department_unmatched=department_unmatched,
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
        if None not in (row.get("total_fuel"), row.get("non_fuel_total"), row.get("desc_otros"), row.get("tax_collect")):
            row["total_sales"] = round(
                row["total_fuel"] + row["non_fuel_total"] + row["desc_otros"] + row["tax_collect"] - gettel_amount, 2
            )

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
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "reporte_store_info_historial.html",
        store_info_rows=store_info_rows,
        store_info_totals=_store_info_totals(store_info_rows),
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
    if not store_info_rows:
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
    if not store_info_rows:
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


@app.route("/reporte/documentos")
def reporte_documentos():
    """PDFs de cierre diario ya guardados este mes -- ver CLAUDE.md, barra lateral por módulo."""
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    pdfs = reportes_db.get_month_pdf_list(year, month)
    for pdf in pdfs:
        pdf["filename"] = os.path.basename(pdf["pdf_filename"]) if pdf["pdf_filename"] else None
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "reporte_documentos.html",
        pdfs=pdfs,
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


@app.route("/carga-datos/reporte-diario/resumen-mensual")
def carga_datos_reporte_diario_resumen_mensual():
    """
    El reporte mensual de Reporte Diario -- el PDF en sí solo se guarda
    para poder verlo después, nunca se lee ni se procesa (a diferencia del
    PDF de cierre diario, uno por día). Pedido explícito del usuario
    (2026-09-19): poder subirlo aparte, en un apartado extra -- mismo
    patrón ya usado para el resumen mensual de Lottery (documents_db.py).

    Por ahora es solo el cajón de archivos (subir/ver/eliminar) -- pedido
    explícito del usuario (2026-09-19), revirtiendo el intento de esta
    misma sesión de mostrar acá la tabla de Store Info del mes: "quiero
    que lo dejes vacio de momento ya vamos a trabajar en eso mas tarde
    cuando tenga una idea de como implementarlo".
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    docs = documents_db.list_documents("reporte_diario_resumen_mensual", year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "reporte_diario_resumen_mensual.html",
        docs=docs,
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        **THEME_BY_KEY["reporte"],
    )


@app.route("/carga-datos/reporte-diario/resumen-mensual/subir", methods=["POST"])
def carga_datos_reporte_diario_resumen_mensual_subir():
    year = request.form.get("year", type=int)
    month = request.form.get("month", type=int)
    upload = request.files.get("resumen_file")
    if not year or not month or not (1 <= month <= 12):
        flash("Elegí a qué mes corresponde este resumen.", "error")
        return redirect(url_for("carga_datos_reporte_diario_resumen_mensual"))
    if upload is None or not upload.filename:
        flash("Seleccioná el archivo del resumen mensual.", "error")
        return redirect(url_for("carga_datos_reporte_diario_resumen_mensual", year=year, month=month))

    path, filename = _save_upload_to_workspace(upload)
    documents_db.store_document("reporte_diario_resumen_mensual", path, filename, year, month)
    flash("Resumen mensual guardado.", "success")
    return redirect(url_for("carga_datos_reporte_diario_resumen_mensual", year=year, month=month))


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
            count = int(float(count_raw or 0))
            amount = float(amount_raw or 0)
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


@app.route("/reporte/dia/<report_date>/store-info", methods=["POST"])
def reporte_dia_store_info(report_date):
    credit_terms_raw = request.form.get("credit_terms", "")
    try:
        credit_terms = [float(v.strip()) for v in credit_terms_raw.split(",") if v.strip()]
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
    return render_template("carga_datos_eft.html", **THEME_BY_KEY["carga_eft"])


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
    job_id = jobs.create_job(len(pdf_paths))
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

    inserted, updated = eft_db.upsert_cupones(records, source_filename=filename)
    try:
        today = date.today()
        documents_db.store_document("eft", monthly_path, filename, today.year, today.month, label="Reporte mensual de Cupones")
    except Exception as exc:
        print(f"[documents_db] no se pudo guardar el reporte mensual de Cupones {filename}: {exc}")
    flash(f"{len(records)} cupón(es) guardado(s) ({inserted} nuevo(s), {updated} actualizado(s)).", "success")
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
    Historial COMPLETO de Cupones -- histórico desde inicios de 2026, nunca
    filtrado por mes (agrupado por mes solo como separador visual, ver
    eft_db.get_cupones_grouped_by_month), en su propio apartado separado de
    los EFT -- pedido explícito del usuario (2026-09-18), ver el comentario
    de carga_datos_eft_historial más arriba.
    """
    cupon_groups = eft_db.get_cupones_grouped_by_month()
    for group in cupon_groups:
        group["label"] = f"{_MONTH_NAMES_ES[group['month'] - 1]} {group['year']}" if group["month"] else "Sin fecha"

    return render_template(
        "carga_datos_eft_cupones_historial.html",
        cupon_groups=cupon_groups,
        **THEME_BY_KEY["carga_eft"],
    )


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
    como el cuadro del excel exactamente igual". No hay nada que subir acá
    -- Chase y Lottery ya se cargan por sus propios módulos, este reporte
    solo cruza lo que ya está guardado (ver caja.build_month_report_from_db).
    """
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int) or today.month
    if not (1 <= month <= 12):
        month = today.month

    report = build_caja_month_report(year, month)
    prev_month, prev_year = (12, year - 1) if month == 1 else (month - 1, year)
    next_month, next_year = (1, year + 1) if month == 12 else (month + 1, year)

    return render_template(
        "carga_datos_caja_historial.html",
        report=report,
        expense_items=caja_db.get_month_expense_items(year, month),
        year=year,
        month=month,
        month_name=_MONTH_NAMES_ES[month - 1],
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        today_iso=today.isoformat(),
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

    caja_db.add_expense_item(report_date, amount, detail)
    flash("Gasto agregado.", "success")
    return redirect(url_for("carga_datos_caja", year=year, month=month))


@app.route("/carga-datos/caja/gastos/<int:item_id>/eliminar", methods=["POST"])
def carga_datos_caja_gastos_eliminar(item_id):
    caja_db.delete_expense_item(item_id)
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
    """
    return render_template("carga_datos_gettel.html", **THEME_BY_KEY["carga_gettel"])


@app.route("/carga-datos/gettel/subir", methods=["POST"])
def carga_datos_gettel_subir():
    uploads = request.files.getlist("source_files")
    if not uploads or not any(u.filename for u in uploads):
        return _error_response("Seleccioná uno o más Excel/PDF de cupones de Gettel y/o Toyota.")

    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths))
    threading.Thread(target=_run_carga_datos_gettel_job, args=(job_id, paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


def _run_carga_datos_gettel_job(job_id, paths):
    """Corre en su propio hilo -- mismo patrón que _run_carga_datos_reporte_diario_job, ver ese docstring."""
    try:
        days_gettel = set()
        days_toyota = set()
        files_failed = 0
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
                    totals_by_date, _diagnostics = summarize_pdf_report(path)
                    if not totals_by_date:
                        raise ValueError("No se pudo leer ninguna fila del reporte.")
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

        if not parts:
            notice, level = "No se pudo guardar nada de este lote.", "error"
        else:
            notice, level = " ".join(parts), ("warning" if files_failed else "success")

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
    return render_template("carga_datos_horas_trabajo.html", **THEME_BY_KEY["carga_horas"])


@app.route("/carga-datos/horas-trabajo/subir", methods=["POST"])
def carga_datos_horas_trabajo_subir():
    uploads = request.files.getlist("report_files")
    if not uploads or not any(u.filename for u in uploads):
        return _error_response("Seleccioná uno o más PDF de Clock In/Out.")

    paths = _save_uploads_to_workspace(uploads)
    job_id = jobs.create_job(len(paths))
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
                data = extract_hours_report(path)
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
                unresolved_total.extend(data["unresolved_employees"])
                if first_date is None or data["report_date"] < first_date:
                    first_date = data["report_date"]
            except Exception as exc:
                print(f"[carga-datos/horas-trabajo] {path}: {exc}")
                files_failed += 1
            jobs.update_job(job_id, done=index, total=len(paths))

        parts = []
        if weeks_saved:
            parts.append(f"{weeks_saved} semana(s) guardada(s).")
        if unresolved_total:
            parts.append(
                f"{len(unresolved_total)} empleado(s) sin su Total legible -- agregalos a mano desde el cuadro de esa semana."
            )
        if files_failed:
            parts.append(f"{files_failed} archivo(s) no se pudieron leer.")

        if not parts:
            notice, level = "No se pudo guardar nada de este lote.", "error"
        else:
            notice, level = " ".join(parts), ("warning" if (files_failed or unresolved_total) else "success")

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
            employees.append({**emp, "total_pay": total_pay})
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
        hours = float((request.form.get("hours") or "0").strip())
        rate = float((request.form.get("rate") or "0").strip())
        deduct = float((request.form.get("deduct") or "0").strip())
    except ValueError:
        flash("No se pudo guardar: revisá que horas/tarifa/descuento sean números válidos.", "error")
        return redirect(url_for("carga_datos_horas_trabajo_historial", year=year, month=month))

    horas_trabajo_db.update_employee(employee_id, employee_name=name or None, hours=hours, rate=rate, deduct=deduct)
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
        hours = float((request.form.get("hours") or "0").strip())
        rate_raw = (request.form.get("rate") or "").strip()
        rate = float(rate_raw) if rate_raw else horas_trabajo_db.DEFAULT_HOURLY_RATE
        deduct = float((request.form.get("deduct") or "0").strip())
    except ValueError:
        flash("No se pudo agregar: revisá que horas/tarifa/descuento sean números válidos.", "error")
        return redirect(url_for("carga_datos_horas_trabajo_historial", year=year, month=month))

    horas_trabajo_db.add_employee(week_id, name, hours=hours, rate=rate, deduct=deduct)
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

    paths = _save_uploads_to_workspace(uploads)
    try:
        combined, file_stats, failed_files = _consolidate_department_files(paths)
    except ValueError as exc:
        flash(f"Error: {exc}", "error")
        return redirect(url_for("carga_datos_cmv"))

    summary = cmv_db.replace_costs_for_departments(combined.to_dict("records"))

    today = date.today()
    for path in paths:
        try:
            documents_db.store_document("cmv_costo", path, os.path.basename(path), today.year, today.month)
        except Exception as exc:
            print(f"[documents_db] no se pudo guardar el documento de CMV Costo {path}: {exc}")

    parts = [f"{summary['departments']} departamento(s), {summary['rows']} producto(s) guardados."]
    if summary["price_changes"]:
        parts.append(f"{summary['price_changes']} cambio(s) de precio detectado(s).")
    if failed_files:
        parts.append(f"{failed_files} archivo(s) no se pudieron leer.")
    flash(" ".join(parts), "warning" if failed_files else "success")
    return redirect(url_for("carga_datos_cmv_costo_historial"))


@app.route("/carga-datos/cmv/costo/historial")
def carga_datos_cmv_costo_historial():
    departments = cmv_db.list_departments()
    selected = request.args.get("dept") or (departments[0]["dept_name"] if departments else None)
    rows = cmv_db.get_costs_by_department(selected) if selected else []
    return render_template(
        "carga_datos_cmv_costo_historial.html",
        departments=departments,
        selected=selected,
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

    rows_by_dept = {}
    unmapped_departments = set()
    files_failed = 0
    for path in paths:
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
        for record in frame.to_dict("records"):
            dept_name = _resolve_sheet_name(record.get("Dept Name"))
            if dept_name is None:
                unmapped_departments.add((record.get("Dept Name") or "").strip() or "(sin nombre)")
                continue
            rows_by_dept.setdefault(dept_name, []).append(
                {
                    "upc": record.get("UPC"),
                    "name": record.get("Name"),
                    "count": record.get("Count"),
                    "amount": record.get("Retail/Amount"),
                }
            )

    for dept_name, rows in rows_by_dept.items():
        cmv_db.replace_month_department_sales(year, month, dept_name, rows)

    parts = []
    if rows_by_dept:
        parts.append(
            f"{len(rows_by_dept)} departamento(s), "
            f"{sum(len(r) for r in rows_by_dept.values())} producto(s) guardados."
        )
    if unmapped_departments:
        parts.append(f"{len(unmapped_departments)} departamento(s) sin hoja conocida, no se guardaron.")
    if files_failed:
        parts.append(f"{files_failed} archivo(s) no se pudieron leer.")
    if not parts:
        flash("No se pudo guardar nada de este lote.", "error")
    else:
        flash(" ".join(parts), "warning" if (unmapped_departments or files_failed) else "success")

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
}


@app.route("/carga-datos/documentos/<module_key>")
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
        return redirect(url_for("carga_datos_index"))

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


@app.route("/carga-datos/documentos/<module_key>/subir", methods=["POST"])
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
        return redirect(url_for("carga_datos_index"))

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


@app.route("/carga-datos/documentos/<int:document_id>/descargar")
def carga_datos_documento_descargar(document_id):
    doc = documents_db.get_document(document_id)
    if doc is None:
        flash("Ese documento ya no existe.", "error")
        return redirect(url_for("carga_datos_index"))
    # Vista previa por default (inline), descarga forzada solo con
    # ?mode=download -- pedido explícito del usuario (2026-09-19), ver
    # templates/_pdf_links.html. Antes esta ruta siempre forzaba la
    # descarga (as_attachment=True) -- ahora es la misma ruta para las dos
    # cosas, el link "Ver" simplemente no manda el parámetro.
    force_download = request.args.get("mode") == "download"
    return send_file(doc["stored_path"], as_attachment=force_download, download_name=doc["filename"])


@app.route("/carga-datos/documentos/<int:document_id>/eliminar", methods=["POST"])
def carga_datos_documento_eliminar(document_id):
    doc = documents_db.get_document(document_id)
    if doc is None:
        flash("Ese documento ya no existe.", "error")
        return redirect(url_for("carga_datos_index"))
    documents_db.delete_document(document_id)
    flash("Documento eliminado.", "success")
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
    return render_template("carga_datos_proveedores.html", **THEME_BY_KEY["carga_proveedores"])


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
    job_id = jobs.create_job(len(paths))
    threading.Thread(target=_run_carga_datos_proveedores_job, args=(job_id, paths), daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(paths)})


def _run_carga_datos_proveedores_job(job_id, paths):
    """Corre en su propio hilo -- mismo patrón que _run_carga_datos_reporte_diario_job, ver ese docstring."""
    try:
        saved = []
        duplicates = []
        failed = []
        date_mismatches = []

        for index, path in enumerate(paths, start=1):
            filename = os.path.basename(path)
            try:
                supplier_key, supplier_label, invoices = extract_invoices_from_pdf(path)
            except _PDF_EXTRACTION_EXCEPTIONS as exc:
                failed.append({"filename": filename, "error": str(exc), "supplier": None})
                jobs.update_job(job_id, done=index, total=len(paths))
                continue

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
                if _filename_date_mismatch(filename, invoice["date"]):
                    date_mismatches.append({"filename": filename, "supplier": supplier_label})
                    continue
                valid_invoices.append(invoice)

            for invoice in valid_invoices:
                ok = proveedores_db.save_invoice(
                    supplier_key, supplier_label, invoice["date"], invoice["invoice_no"],
                    invoice["amount"], source_filename=filename,
                )
                if ok:
                    saved.append({"filename": filename, "supplier": supplier_label, "date": invoice["date"]})
                else:
                    duplicates.append({"filename": filename, "supplier": supplier_label})

            if valid_invoices:
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
            parts.append(f"{len(saved)} factura(s) guardada(s).")
        if duplicates:
            parts.append(f"{len(duplicates)} factura(s) ya estaban cargadas y se omitieron.")
        if date_mismatches:
            parts.append(_group_by_supplier_message(
                "Rechazadas (la fecha no coincide con el nombre del archivo)", date_mismatches, "factura(s)"
            ))
        if failed:
            parts.append(_group_by_supplier_message("No se pudieron cargar", failed, "factura(s)"))
        if not parts:
            parts.append("No se guardó ninguna factura de este lote.")
        level = (
            "success" if (saved and not duplicates and not failed and not date_mismatches)
            else ("error" if not saved else "warning")
        )

        if saved:
            d = saved[0]["date"]
            redirect_url = f"/carga-datos/proveedores/historial?year={d.year}&month={d.month}"
        else:
            redirect_url = "/carga-datos/proveedores/historial"

        jobs.update_job(
            job_id, status="done", done=len(paths), total=len(paths),
            notice=" ".join(parts), notice_level=level, redirect_url=redirect_url,
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
        return redirect(url_for("proveedores"))
    return render_template(
        "proveedores_nuevo.html",
        dynamic_suppliers=list_dynamic_suppliers_display(),
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
    if not _require_admin("Solo un administrador puede agregar proveedores nuevos."):
        return redirect(url_for("proveedores_nuevo"))

    clave = request.form.get("clave", "").strip()
    try:
        delete_dynamic_supplier(clave)
        flash("Proveedor eliminado.", "success")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("proveedores_nuevo"))


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
    app.run(debug=debug_mode, host=host, port=port, threaded=True)
