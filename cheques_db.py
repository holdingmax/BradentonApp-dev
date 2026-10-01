"""
Cheques propios guardados (Chase Bank -> Cheques) -- capa de datos pura, sin
conocimiento de PDFs/OCR (eso vive en cheques.py).

Cada cheque guarda:
- su N° (puede faltar si el OCR no lo pudo leer -- se completa a mano),
- un PDF de una sola página con el cheque ya recortado y derecho (lo que se
  abre al hacer click en la lista) -- vacío si el cheque se cargó a mano,
- el PDF original tal cual se subió (factura del proveedor con el cheque
  adjunto, o el cheque suelto) -- compartido si un mismo PDF trae varios.

Si el cheque ya llegó al banco NO se guarda acá: se cruza en el momento contra
chase_db (movimientos "CHECK {n}"), así una carga nueva del extracto lo marca
como cobrado sin tocar nada de esta base.
"""

import os
import re
import sqlite3
import uuid
from datetime import datetime

import documents_db

_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_data")
_DB_PATH = os.path.join(_BASE_DIR, "cheques.db")
_FILES_DIR = os.path.join(_BASE_DIR, "cheques")


def _connect():
    os.makedirs(_BASE_DIR, exist_ok=True)
    conn = sqlite3.connect(_DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cheques (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            check_number INTEGER,
            number_source TEXT,
            check_pdf TEXT NOT NULL,
            original_pdf TEXT NOT NULL,
            source_filename TEXT,
            page_index INTEGER,
            uploaded_at TEXT
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_cheques_number ON cheques (check_number)")
    existing = {row[1] for row in conn.execute("PRAGMA table_info(cheques)")}
    for column, ddl in _EXTRA_COLUMNS:
        if column not in existing:
            conn.execute(f"ALTER TABLE cheques ADD COLUMN {column} {ddl}")
    conn.commit()
    return conn


# Columnas agregadas cuando la lista pasó a ser el cuadro de control único
# (mismo diseño que "Cheques for pays suppliers. Control.xls"). auto_* se
# completa al subir el PDF leyendo la factura del proveedor que viene con el
# cheque; manual_* es lo que el usuario tipea a mano y siempre gana. Un cheque
# cargado a mano sin PDF guarda check_pdf/original_pdf vacíos.
_EXTRA_COLUMNS = (
    ("auto_supplier_key", "TEXT"),
    ("auto_supplier_label", "TEXT"),
    ("auto_invoice_no", "TEXT"),
    ("auto_invoice_date", "TEXT"),
    ("auto_amount", "REAL"),
    ("manual_date", "TEXT"),
    ("manual_amount", "REAL"),
    ("manual_supplier", "TEXT"),
    ("manual_invoice", "TEXT"),
    ("manual_debit_date", "TEXT"),
    ("voided", "INTEGER DEFAULT 0"),
)
_AUTO_FIELDS = ("auto_supplier_key", "auto_supplier_label", "auto_invoice_no", "auto_invoice_date", "auto_amount")
_MANUAL_FIELDS = ("manual_date", "manual_amount", "manual_supplier", "manual_invoice", "manual_debit_date", "voided")


def _now():
    return datetime.now().isoformat(timespec="seconds")


def absolute_path(relpath):
    # Las filas guardadas desde Windows traen "\": se parte por los dos
    # separadores para que también se resuelvan en Linux (auditoría 2026-09).
    return os.path.join(_FILES_DIR, *[part for part in re.split(r"[\\/]", relpath or "") if part])


def store_original(source_path, original_filename):
    """Copia el PDF subido una sola vez; devuelve su ruta relativa (para varios cheques del mismo PDF)."""
    # Guardado de documentos apagado (documents_db.GUARDAR_DOCUMENTOS): '' es
    # el mismo valor que ya usa una fila cargada a mano sin PDF.
    if not documents_db.GUARDAR_DOCUMENTOS:
        return ""
    folder = uuid.uuid4().hex[:12]
    rel = f"originales/{folder}/{os.path.basename(original_filename)}"
    dest = absolute_path(rel)
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    with open(source_path, "rb") as src, open(dest, "wb") as dst:
        dst.write(src.read())
    return rel


def find_by_number(check_number):
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM cheques WHERE check_number = ?", (check_number,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def find_by_source(source_filename, page_index):
    """Fila ya cargada desde la misma página del mismo PDF (para un cheque sin N° legible)."""
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT * FROM cheques WHERE source_filename = ? AND page_index = ?",
            (source_filename, page_index),
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def _save_check_image(image):
    if not documents_db.GUARDAR_DOCUMENTOS:
        return ""
    rel = f"cheques/{uuid.uuid4().hex[:12]}.pdf"
    dest = absolute_path(rel)
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    image.convert("RGB").save(dest, "PDF", resolution=200)
    return rel


def add_check(check_number, image, original_rel, source_filename, page_index, auto=None):
    """
    Guarda el cheque (imagen PIL ya derecha) como PDF de una página. Si ya
    existe una fila de ese N° cargada a mano sin PDF, se le adjunta el PDF a
    esa fila en vez de crear otra. `auto` = datos leídos de la factura del
    proveedor (ver _AUTO_FIELDS). Devuelve el id.
    """
    auto = auto or {}
    rel = _save_check_image(image)
    auto_values = tuple(auto.get(k) for k in _AUTO_FIELDS)
    conn = _connect()
    try:
        existing = None
        if check_number is not None:
            existing = conn.execute(
                "SELECT id FROM cheques WHERE check_number = ? AND check_pdf = ''", (check_number,)
            ).fetchone()
        if existing:
            conn.execute(
                f"""
                UPDATE cheques SET number_source = 'ocr', check_pdf = ?, original_pdf = ?,
                    source_filename = ?, page_index = ?, uploaded_at = ?,
                    {", ".join(f"{k} = ?" for k in _AUTO_FIELDS)}
                WHERE id = ?
                """,
                (rel, original_rel, source_filename, page_index, _now()) + auto_values + (existing["id"],),
            )
            conn.commit()
            return existing["id"]
        cur = conn.execute(
            f"""
            INSERT INTO cheques (check_number, number_source, check_pdf, original_pdf,
                                 source_filename, page_index, uploaded_at, {", ".join(_AUTO_FIELDS)})
            VALUES (?, ?, ?, ?, ?, ?, ?, {", ".join("?" for _ in _AUTO_FIELDS)})
            """,
            (check_number, "ocr" if check_number is not None else None, rel, original_rel,
             source_filename, page_index, _now()) + auto_values,
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def set_auto_fields(check_id, auto):
    conn = _connect()
    try:
        conn.execute(
            f"UPDATE cheques SET {', '.join(f'{k} = ?' for k in _AUTO_FIELDS)} WHERE id = ?",
            tuple(auto.get(k) for k in _AUTO_FIELDS) + (check_id,),
        )
        conn.commit()
    finally:
        conn.close()


def save_manual(check_id, check_number, fields):
    """
    Carga/corrección a mano de una fila del cuadro. `check_id` None = cheque
    que todavía no tiene fila propia (nuevo, o que hasta ahora solo existía en
    Chase) -- si ya hay una fila con ese N°, ValueError. Un valor None en
    `fields` vuelve ese dato a lo automático. Devuelve el id.
    """
    values = tuple(fields.get(k) for k in _MANUAL_FIELDS)
    conn = _connect()
    try:
        if check_number is not None:
            other = conn.execute(
                "SELECT id FROM cheques WHERE check_number = ? AND id != ?",
                (check_number, check_id if check_id is not None else -1),
            ).fetchone()
            # Alta (check_id None) con un N° que ya tiene fila: se rechaza en
            # vez de editar esa fila, porque el form de alta llega con los
            # demás campos vacíos y pisaría con None lo cargado a mano (y
            # desanularía el cheque). Auditoría 2026-09, cheques_db.py:184.
            if other and check_id is None:
                raise ValueError(
                    f"El cheque N° {check_number} ya está en el cuadro: editalo desde su fila."
                )
            if other:
                raise ValueError(f"El cheque N° {check_number} ya está cargado.")
        if check_id is None:
            cur = conn.execute(
                f"""
                INSERT INTO cheques (check_number, number_source, check_pdf, original_pdf, uploaded_at,
                                     {", ".join(_MANUAL_FIELDS)})
                VALUES (?, 'manual', '', '', ?, {", ".join("?" for _ in _MANUAL_FIELDS)})
                """,
                (check_number, _now()) + values,
            )
            conn.commit()
            return cur.lastrowid
        current = conn.execute("SELECT check_number FROM cheques WHERE id = ?", (check_id,)).fetchone()
        if current is None:
            raise ValueError("Ese cheque ya no existe.")
        number_sql, extra = "", ()
        if check_number != current["check_number"]:
            number_sql, extra = "check_number = ?, number_source = 'manual', ", (check_number,)
        conn.execute(
            f"UPDATE cheques SET {number_sql}{', '.join(f'{k} = ?' for k in _MANUAL_FIELDS)} WHERE id = ?",
            extra + values + (check_id,),
        )
        conn.commit()
        return check_id
    finally:
        conn.close()


def list_checks():
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM cheques ORDER BY check_number IS NULL, check_number, id"
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_check(check_id):
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM cheques WHERE id = ?", (check_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def delete_check(check_id):
    """Borra el cheque y su PDF; el original solo si ningún otro cheque lo usa."""
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM cheques WHERE id = ?", (check_id,)).fetchone()
        if not row:
            return False
        conn.execute("DELETE FROM cheques WHERE id = ?", (check_id,))
        still_used = conn.execute(
            "SELECT 1 FROM cheques WHERE original_pdf = ? LIMIT 1", (row["original_pdf"],)
        ).fetchone()
        conn.commit()
    finally:
        conn.close()
    paths = [p for p in [row["check_pdf"]] + ([] if still_used else [row["original_pdf"]]) if p]
    for rel in paths:
        try:
            os.remove(absolute_path(rel))
        except OSError:
            pass
    if not still_used and row["original_pdf"]:
        try:
            os.rmdir(os.path.dirname(absolute_path(row["original_pdf"])))
        except OSError:
            pass
    return True
