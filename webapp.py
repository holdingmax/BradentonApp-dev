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

from chase_rules import process_chase_categorization
from cmv_costo import update_master_costo_todos_bulk
from gettel_toyota_parser import (
    merge_gettel_toyota_into_master,
    merge_gettel_toyota_pdf_into_master,
    process_gettel_pagos,
)
from monthly_sales import process_monthly_sales
from reporte_diario import process_lottery, process_reporte_diario, process_store_info

app = Flask(__name__)
app.secret_key = os.urandom(24)

TOOLS = [
    {
        "label": "Chase Bank",
        "url": "/chase",
        "description": "Categoriza movimientos bancarios contra las reglas de Detalle.",
    },
    {
        "label": "CMV",
        "url": "/cmv",
        "description": "Costo por UPC (COSTO.TODOS) y ventas del POS por departamento.",
    },
    {
        "label": "Gettel / Toyota",
        "url": "/gettel",
        "description": "Cupones diarios (Excel o PDF) y pagos hacia el master Cierre.",
    },
    {
        "label": "Reporte Diario",
        "url": "/reporte",
        "description": "Ventas por Departamento y Store Info desde el PDF de cierre diario.",
    },
    {
        "label": "Lottery",
        "url": "/lottery",
        "description": "Daily Sales Report y PDF Diario hacia el Excel de Lottery.",
    },
]


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
        return render_template("chase.html")

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


@app.route("/cmv")
def cmv():
    return render_template("cmv.html")


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
    return render_template("gettel.html")


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
    return render_template("reporte.html")


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
    return render_template("lottery.html")


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


if __name__ == "__main__":
    app.run(debug=True, port=5000)
