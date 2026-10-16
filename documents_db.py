"""
Almacenamiento genérico de documentos originales (PDF/Excel ya subidos) para
cualquier módulo de Carga de Datos que todavía no tuviera su propio sistema
de archivos -- EFT, Gettel/Toyota, CMV, y el resumen mensual de Lottery.

Pedido explícito del usuario (2026-09-14): "en estas sub categorias es donde
se van a almacenar... los pdf que se suban de los eft... gettel y toyota
junto como con cmv no usan pdf pero se usan excels para completarlos, esos
tambien se deberian guardar y poder abrirse desde ahi" -- y en la misma
línea de "vamos a necesitar pensar en hacer una base de datos grande...
imagina que tenemos cosas en el drive desde hace 3 años" (visión de
reemplazar al Drive): en vez de repetir la tabla `caja_attachments` una vez
por módulo, esto es UN SOLO punto de guardado reusable -- cada archivo
original nunca se lee/parsea desde acá (el procesamiento real ya pasó antes
de llamar a esto), queda solo para poder abrirlo de nuevo más adelante.

reportes_data/documents.db (mismo directorio gitignored de siempre):
- documents: un renglón por archivo, con `module` (clave corta del módulo
  dueño -- "eft", "gettel_toyota", "cmv_costo", "cmv_ventas",
  "lottery_resumen_mensual", etc.), el período (year/month, mes contable al
  que pertenece el archivo -- no necesariamente la fecha de hoy) y una
  `label` opcional (contexto corto, ej. "RCV-20135" para un EFT) para
  encontrarlo más fácil en la lista.

Índice en (module, year, month) -- pensado para escalar a años de archivos
sin que listar "los documentos de este módulo este mes" se vuelva lento
(ver CLAUDE.md, "pensar en una base de datos grande y confiable").
"""

import app_paths
import os
import re
from datetime import datetime

import db

_BASE_DIR = app_paths.DATA_DIR
_DB_PATH = os.path.join(_BASE_DIR, "documents.db")
_FILES_DIR = os.path.join(_BASE_DIR, "documents")

# Interruptor único del guardado de archivos originales -- pedido explícito
# del usuario (2026-09-28), tras la reunión con el jefe: la página tiene que
# extraer los datos de los PDF/Excel y guardar solo esos datos; el archivo
# físico no hace falta. En False no se copia ningún archivo (acá ni en
# reportes_db, lottery_db, cheques_db y Depósitos, que leen esta misma
# constante), las listas de documentos vuelven vacías (así desaparecen solos
# los links "Ver PDF") y la sección Documentos queda oculta. Los archivos ya
# guardados en reportes_data/ quedan en el disco, sin tocar. Volver a True
# reactiva todo tal cual estaba.
GUARDAR_DOCUMENTOS = False


def _init_sqlite(conn):
    # timeout=30 + WAL -- ver reportes_db.py: necesario desde que las cargas
    # en segundo plano (jobs.py) pueden escribir de verdad en paralelo.
    conn.execute("PRAGMA journal_mode=WAL")
    _ensure_schema(conn)


def _connect():
    # Sin DATABASE_URL, la base SQLite de siempre; con ella, el esquema documents de
    # PostgreSQL (db.py; tablas de migraciones/versions/0009_documents.py).
    return db.connect("documents", _DB_PATH, _init_sqlite)


def _ensure_schema(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            module TEXT NOT NULL,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            label TEXT,
            filename TEXT NOT NULL,
            stored_path TEXT NOT NULL,
            uploaded_at TEXT
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_documents_module_period ON documents (module, year, month)"
    )
    conn.commit()


def _now():
    return datetime.now().isoformat(timespec="seconds")


def store_document(module, source_path, original_filename, year, month, label=None):
    """
    Copia un archivo ya guardado en el workspace temporal de la request
    (nunca el original del usuario) a reportes_data/documents/{module}/
    {año}/{mes}/ -- una subcarpeta con timestamp por archivo evita
    colisiones sin tocar el nombre original (mismo criterio que
    reportes_db.store_pdf_copy/caja_db.store_attachment).

    Nunca debe poder romper el flujo que ya guardó los datos reales -- el
    caller la envuelve en su propio try/except, igual que el resto de los
    guardados-espejo del proyecto.
    """
    if not GUARDAR_DOCUMENTOS:
        return None
    dest_dir = os.path.join(_FILES_DIR, module, f"{year:04d}", f"{month:02d}")
    file_dir = os.path.join(dest_dir, datetime.now().strftime("%Y%m%d%H%M%S%f"))
    os.makedirs(file_dir, exist_ok=True)
    dest_path = os.path.join(file_dir, os.path.basename(original_filename))
    with open(source_path, "rb") as src, open(dest_path, "wb") as dst:
        dst.write(src.read())

    now = _now()
    conn = _connect()
    try:
        new_id = conn.execute(
            """
            INSERT INTO documents (module, year, month, label, filename, stored_path, uploaded_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            RETURNING id
            """,
            (module, year, month, label, os.path.basename(original_filename), _portable_relpath(dest_path), now),
        ).fetchone()[0]
        conn.commit()
        return new_id
    finally:
        conn.close()


def list_documents(module, year, month):
    if not GUARDAR_DOCUMENTOS:
        return []
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM documents WHERE module = ? AND year = ? AND month = ? ORDER BY uploaded_at ASC, id",
            (module, year, month),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def list_all_documents(module):
    """
    Todos los documentos guardados de un módulo, sin importar el mes --
    usado para el cruce factura<->documento de EFT (ver webapp.py), donde
    la factura de J.H. Williams puede haberse subido en un mes distinto al
    del EFT que la paga.
    """
    if not GUARDAR_DOCUMENTOS:
        return []
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM documents WHERE module = ? ORDER BY uploaded_at ASC, year, month, id", (module,)
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def search_documents(query, limit=20):
    """
    Busca por nombre de archivo o etiqueta en TODOS los módulos -- pedido
    explícito del usuario (2026-09-16): "quiero que la barra de busqueda
    sirva para encontrar tanto como los modulos, como PDF por su nombre, y
    excels tambien por el nombre". Sin filtrar por módulo/mes -- el punto es
    encontrar un archivo sin tener que saber de antemano dónde quedó
    guardado.
    """
    if not GUARDAR_DOCUMENTOS:
        return []
    like = f"%{query}%"
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM documents WHERE LOWER(filename) LIKE LOWER(?) OR LOWER(label) LIKE LOWER(?) "
            "ORDER BY uploaded_at DESC, id LIMIT ?",
            (like, like, limit),
        ).fetchall()
    finally:
        conn.close()
    return [dict(row) for row in rows]


def _portable_relpath(path):
    """Ruta relativa a _FILES_DIR con "/", para que la base sirva igual en Windows y en Linux."""
    return os.path.relpath(path, _FILES_DIR).replace(os.sep, "/")


def resolve_stored_path(stored):
    """
    Ruta real de un documento guardado. Las filas viejas tienen la ruta
    absoluta de Windows (C:/BradentonApp/reportes_data/documents/... con barras invertidas): se toma lo que
    viene después de la carpeta de documentos y se resuelve contra la de esta
    máquina, así sigue andando si la carpeta se mueve o en Linux/Docker
    (auditoría 2026-09, webapp.py:5907).
    """
    if not stored:
        return stored
    if os.path.isfile(stored):
        return stored
    parts = [part for part in re.split(r"[\\/]", stored) if part]
    marker = os.path.basename(_FILES_DIR)
    if marker in parts:
        parts = parts[len(parts) - parts[::-1].index(marker):]
    return os.path.join(_FILES_DIR, *parts)


def get_document(document_id):
    if not GUARDAR_DOCUMENTOS:
        return None
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM documents WHERE id = ?", (document_id,)).fetchone()
    finally:
        conn.close()
    if row is None:
        return None
    doc = dict(row)
    doc["stored_path"] = resolve_stored_path(doc["stored_path"])
    return doc


def delete_document(document_id):
    doc = get_document(document_id)
    conn = _connect()
    try:
        conn.execute("DELETE FROM documents WHERE id = ?", (document_id,))
        conn.commit()
    finally:
        conn.close()
    if doc:
        try:
            os.remove(doc["stored_path"])
        except OSError:
            pass
