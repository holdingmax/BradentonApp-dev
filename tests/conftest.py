"""
Las pruebas nunca tocan los datos reales: antes de importar cualquier módulo
de la app, la carpeta de datos pasa a una carpeta temporal (SQLite vacío, las
reglas sembradas desde el repo) y se saca DATABASE_URL para no apuntar a
ningún Postgres.
"""

import atexit
import os
import shutil
import tempfile

_DATA_DIR = tempfile.mkdtemp(prefix="bradenton_pruebas_")
os.environ["BRADENTON_DATA_DIR"] = _DATA_DIR
os.environ.pop("DATABASE_URL", None)
atexit.register(shutil.rmtree, _DATA_DIR, ignore_errors=True)
