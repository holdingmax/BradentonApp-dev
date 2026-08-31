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
from datetime import datetime

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from chase_rules import process_chase_categorization
from cmv_costo import update_master_costo_todos_bulk
from monthly_sales import process_monthly_sales

app = Flask(__name__)
app.secret_key = os.urandom(24)

TOOLS = [
    {"label": "Chase Bank", "url": "/chase"},
    {"label": "CMV", "url": "/cmv"},
]


def _new_workspace_dir():
    return tempfile.mkdtemp(prefix="bradenton_web_")


def _save_upload_to_workspace(upload, workdir=None):
    """Save an uploaded file to a fresh (or given) temp dir; return its local path."""
    filename = secure_filename(upload.filename)
    if not filename:
        raise ValueError("Nombre de archivo inválido.")
    workdir = workdir or _new_workspace_dir()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_path = os.path.join(workdir, f"{stamp}_{filename}")
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


if __name__ == "__main__":
    app.run(debug=True, port=5000)
