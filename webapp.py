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

app = Flask(__name__)
app.secret_key = os.urandom(24)

TOOLS = [
    {"label": "Chase Bank", "url": "/chase"},
]


def _save_upload_to_workspace(upload):
    """Save an uploaded file to a fresh temp dir; return its local path."""
    filename = secure_filename(upload.filename)
    if not filename:
        raise ValueError("Nombre de archivo inválido.")
    workdir = tempfile.mkdtemp(prefix="bradenton_web_")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_path = os.path.join(workdir, f"{stamp}_{filename}")
    upload.save(temp_path)
    return temp_path, filename


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


if __name__ == "__main__":
    app.run(debug=True, port=5000)
