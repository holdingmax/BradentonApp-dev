"""
Respaldo de la base PostgreSQL (sección 9 de la guía; punto 8 de
PLAN_POSTGRES.md).

    python respaldo_postgres.py                     respalda la base de DATABASE_URL
    python respaldo_postgres.py --conservar 30      y deja solo los 30 más nuevos de esa base
    python respaldo_postgres.py --restaurar ARCHIVO restaura un respaldo (pide confirmación)

- La base es la de DATABASE_URL (variable de entorno o archivo .env): para
  respaldar testing o producción se usa la URL de ese ambiente. La URL nunca
  se muestra (tiene la contraseña); solo la base y el servidor.
- El archivo va a la carpeta backups/ (está en .gitignore, no viaja a
  GitHub) con la base, la fecha y la hora en el nombre:
  backups/bradenton-dev-db_20261016_1530.dump
- Formato "custom" de pg_dump, sin dueños ni permisos: se puede restaurar
  en otra base (por ejemplo, producción en una de prueba).
- Al terminar, lo verifica abriéndolo con pg_restore --list y avisa si
  salió bien (con el tamaño y la cantidad de tablas) o si falló.

Qué hace falta en la PC: pg_dump y pg_restore de una versión igual o más
nueva que la del servidor. Si no están instalados (las "Command Line Tools"
del instalador de PostgreSQL para Windows), el script usa los de la imagen
de Docker postgres:17 (Docker Desktop tiene que estar abierto).

Para que corra solo todas las noches en Windows (Programador de tareas):
    schtasks /Create /SC DAILY /ST 23:00 /TN "Respaldo Bradenton" /TR "python C:\\BradentonApp\\respaldo_postgres.py --conservar 30"
"""

import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime

import db

APP_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.join(APP_DIR, "backups")
DOCKER_IMAGE = "postgres:17"
LOCAL_HOSTS = {"localhost", "127.0.0.1", "::1"}


def _fail(message):
    print(f"FALLÓ EL RESPALDO: {message}")
    sys.exit(1)


def _target():
    import psycopg

    if not db.uses_postgres():
        db.load_dotenv()
    url = db.database_url()
    if not url:
        _fail("falta DATABASE_URL (variable de entorno o archivo .env).")
    info = psycopg.conninfo.conninfo_to_dict(url)
    return url, info.get("dbname") or "base", info.get("host") or "localhost"


def _tool(name, url):
    """El comando para correr pg_dump/pg_restore: el instalado o el de Docker."""
    if shutil.which(name):
        return [name], url
    if not shutil.which("docker"):
        _fail(f"no está {name} ni Docker. Instalá las Command Line Tools de PostgreSQL o abrí Docker Desktop.")
    # Desde adentro del contenedor, el Postgres de esta PC se ve como host.docker.internal.
    import psycopg

    info = psycopg.conninfo.conninfo_to_dict(url)
    if (info.get("host") or "localhost") in LOCAL_HOSTS:
        info["host"] = "host.docker.internal"
        url = psycopg.conninfo.make_conninfo(**info)
    return ["docker", "run", "--rm", "-i", DOCKER_IMAGE, name], url


def backup(keep):
    url, dbname, host = _target()
    os.makedirs(BACKUP_DIR, exist_ok=True)
    safe = re.sub(r"[^\w.-]", "_", dbname)
    path = os.path.join(BACKUP_DIR, f"{safe}_{datetime.now():%Y%m%d_%H%M}.dump")
    print(f"Respaldando la base {dbname} ({host}) en {os.path.relpath(path, APP_DIR)} ...")
    command, url = _tool("pg_dump", url)
    with open(path, "wb") as out:
        done = subprocess.run(command + ["--format=custom", "--no-owner", "--no-acl", f"--dbname={url}"],
                              stdout=out, stderr=subprocess.PIPE)
    if done.returncode != 0:
        os.remove(path)
        _fail(done.stderr.decode("utf-8", "replace").strip() or f"pg_dump terminó con código {done.returncode}")

    # Verificación: el archivo se puede abrir y trae las tablas.
    command, _ = _tool("pg_restore", url)
    with open(path, "rb") as dump:
        listing = subprocess.run(command + ["--list"], stdin=dump, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if listing.returncode != 0:
        _fail(f"el archivo quedó ilegible: {listing.stderr.decode('utf-8', 'replace').strip()}")
    tables = sum(1 for line in listing.stdout.decode("utf-8", "replace").splitlines() if " TABLE DATA " in line)
    size = os.path.getsize(path)
    print(f"Respaldo OK: {os.path.basename(path)} ({size / 1024:,.0f} KB, {tables} tablas con datos).")

    if keep:
        old = sorted(glob.glob(os.path.join(BACKUP_DIR, f"{safe}_*.dump")))[:-keep]
        for name in old:
            os.remove(name)
        if old:
            print(f"Se borraron {len(old)} respaldos viejos de {dbname} (quedan los {keep} más nuevos).")


def restore(path):
    url, dbname, host = _target()
    if not os.path.isfile(path):
        _fail(f"no existe {path}")
    answer = input(f"Esto REEMPLAZA los datos de la base {dbname} ({host}) por los de {os.path.basename(path)}. "
                   f"Escribí el nombre de la base para seguir: ")
    if answer.strip() != dbname:
        print("Cancelado: no se tocó nada.")
        return
    command, url = _tool("pg_restore", url)
    with open(path, "rb") as dump:
        done = subprocess.run(command + ["--clean", "--if-exists", "--no-owner", "--no-acl", f"--dbname={url}"],
                              stdin=dump, stderr=subprocess.PIPE)
    if done.returncode != 0:
        _fail(done.stderr.decode("utf-8", "replace").strip())
    print(f"Restaurado {os.path.basename(path)} en {dbname}.")


def main():
    parser = argparse.ArgumentParser(description="Respaldo de la base PostgreSQL de BradentonApp.")
    parser.add_argument("--conservar", type=int, default=0, metavar="N", help="deja solo los N respaldos más nuevos de esa base")
    parser.add_argument("--restaurar", metavar="ARCHIVO", help="restaura un respaldo en la base de DATABASE_URL")
    args = parser.parse_args()
    if args.restaurar:
        restore(args.restaurar)
    else:
        backup(args.conservar)


if __name__ == "__main__":
    main()
