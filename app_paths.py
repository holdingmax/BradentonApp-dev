"""
Dónde guarda la app todo lo que no es código: las bases SQLite, los usuarios,
la clave de sesión y las reglas que se editan desde la página.

En Render el disco del servicio se borra con cada deploy o reinicio (queda el
código, se pierde lo guardado). Por eso la carpeta de datos es configurable:
con la variable de entorno BRADENTON_DATA_DIR apuntando al disco persistente
de Render (por ejemplo /var/data), todo lo guardado vive ahí y sobrevive a los
deploys. Sin la variable (la PC del usuario) todo queda donde estuvo siempre:
reportes_data/ y los archivos de configuración en la carpeta del código.
"""

import os
import shutil

APP_DIR = os.path.dirname(os.path.abspath(__file__))

_DATA_DIR_ENV = os.environ.get("BRADENTON_DATA_DIR", "").strip()

# Bases SQLite (y las carpetas de PDFs que guardaban algunos módulos).
DATA_DIR = os.path.abspath(_DATA_DIR_ENV) if _DATA_DIR_ENV else os.path.join(APP_DIR, "reportes_data")


def config_file(filename, seed_from_repo=False):
    """
    Ruta de un archivo de configuración que la página escribe (users.json,
    reglas de Chase, etc.). Sin BRADENTON_DATA_DIR es el de la carpeta del
    código, como siempre. Con la variable vive en la carpeta de datos. Con
    seed_from_repo (las reglas, que vienen cargadas en git) la primera vez se
    copia ahí la versión del repo, y desde entonces manda la del disco, así lo
    que se edite desde la página no se pierde en el próximo deploy. Usuarios y
    clave de sesión nunca se copian del repo.
    """
    repo_path = os.path.join(APP_DIR, filename)
    if not _DATA_DIR_ENV:
        return repo_path
    os.makedirs(DATA_DIR, exist_ok=True)
    data_path = os.path.join(DATA_DIR, filename)
    if seed_from_repo and not os.path.exists(data_path) and os.path.isfile(repo_path):
        shutil.copy2(repo_path, data_path)
    return data_path
