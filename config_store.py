"""
La configuración que la app guarda y que no son datos del negocio (pase a
PostgreSQL, punto 5 de PLAN_POSTGRES.md): usuarios, clave de sesión, reglas
de Chase y de pagos de proveedores, proveedores sin código, el límite de la
alerta de Caja y las marcas de Sale Tax del Cierre.

Cada módulo lee y guarda su configuración entera con read(nombre) y
write(nombre, contenido), con la misma forma que tenía su archivo JSON:
- sin DATABASE_URL (la PC), es el archivo de siempre;
- con DATABASE_URL, son las tablas del esquema config de PostgreSQL
  (migraciones/versions/0023_config.py), así en Render no queda nada en el
  disco del servicio, que se borra en cada deploy.

read devuelve None si nunca se guardó (como un archivo que no existe). Si
existe pero no se puede leer da error, nunca "vacío": unos usuarios
ilegibles no pueden tomarse como "no hay usuarios" (A1 de
AUDITORIA_2026-10.md). Las reglas que vienen en el repo son la semilla: la
primera lectura las copia, en el disco o en la base, y desde ahí manda lo
que se edite en la página.
"""

import json
import os

import app_paths
import db

USERS = "users.json"
CHASE_RULES = "chase_rules.json"
CHASE_MASTER_RULES = "chase_master_rules.json"
PAGO_RULES = "proveedores_pago_rules.json"
DYNAMIC_SUPPLIERS = "proveedores_dynamic_suppliers.json"
CONTROLES_RAPIDOS = "controles_rapidos.json"
SALE_TAX = "control_cierre_sale_tax.json"

# Las que tienen semilla en el repo (las reglas) y las que vivían en la
# carpeta de datos en vez de la del código.
_SEEDED = {CHASE_RULES, CHASE_MASTER_RULES, PAGO_RULES, DYNAMIC_SUPPLIERS}
_IN_DATA_DIR = {CONTROLES_RAPIDOS, SALE_TAX}
NAMES = (USERS, CHASE_RULES, CHASE_MASTER_RULES, PAGO_RULES, DYNAMIC_SUPPLIERS, CONTROLES_RAPIDOS, SALE_TAX)

_CHASE_KIND = {CHASE_MASTER_RULES: "maestra", CHASE_RULES: "personalizada"}


class Unreadable(ValueError):
    """La configuración existe pero no se puede leer."""


def read(name):
    if db.uses_postgres():
        return _pg_read(name)
    return _file_read(name)


def write(name, payload):
    if db.uses_postgres():
        _pg_write(name, payload)
    else:
        _file_write(name, payload)


# --- Archivos (la PC) -------------------------------------------------------

def file_path(name, data_dir=None):
    if name not in NAMES:
        raise ValueError(f"Configuración desconocida: {name}")
    if name in _IN_DATA_DIR:
        return os.path.join(data_dir or app_paths.DATA_DIR, name)
    return app_paths.config_file(name, seed_from_repo=name in _SEEDED)


def load_json_file(path):
    """El contenido de un JSON, None si no existe; Unreadable si está roto."""
    if not os.path.isfile(path):
        return None
    try:
        with open(path, encoding="utf-8-sig") as handle:  # -sig: tolera el BOM que deja PowerShell
            return json.load(handle)
    except (OSError, ValueError) as exc:
        raise Unreadable(f"No se pudo leer {os.path.basename(path)}: {exc}") from exc


def _file_read(name):
    return load_json_file(file_path(name))


def _file_write(name, payload):
    path = file_path(name)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    # Se escribe aparte y se reemplaza: un corte a mitad no deja el archivo roto.
    temp_path = f"{path}.tmp"
    with open(temp_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, ensure_ascii=False)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp_path, path)


def secret_key():
    """La clave de sesión de Flask: se crea una vez y queda guardada (así no se
    desloguea a todos en cada reinicio ni en cada deploy)."""
    if db.uses_postgres():
        conn = db.connect("config", None)
        try:
            with conn:
                conn.execute("INSERT INTO secretos (nombre, valor) VALUES (?, ?) ON CONFLICT (nombre) DO NOTHING",
                             ("flask_secret_key", os.urandom(32)))
                return bytes(conn.execute("SELECT valor FROM secretos WHERE nombre = ?", ("flask_secret_key",)).fetchone()[0])
        finally:
            conn.close()
    path = app_paths.config_file(".flask_secret_key")
    if os.path.isfile(path):
        with open(path, "rb") as handle:
            key = handle.read()
        if key:
            return key
    key = os.urandom(32)
    with open(path, "wb") as handle:
        handle.write(key)
    return key


# --- PostgreSQL (Render) ----------------------------------------------------

def _seed_from_repo(name):
    """La versión del repo de una regla (la semilla), o None."""
    if name not in _SEEDED:
        return None
    return load_json_file(os.path.join(app_paths.APP_DIR, name))


def _pg_read(name):
    if name not in NAMES:
        raise ValueError(f"Configuración desconocida: {name}")
    conn = db.connect("config", None)
    try:
        saved = conn.execute("SELECT 1 FROM guardados WHERE nombre = ?", (name,)).fetchone() is not None
        payload = _pg_load(conn, name) if saved else None
    finally:
        conn.close()
    if payload is None and not saved:
        seed = _seed_from_repo(name)
        if seed is not None:
            _pg_write(name, seed)
            return seed
    return payload


def _pg_load(conn, name):
    if name == USERS:
        rows = conn.execute("SELECT username, password_hash, is_admin, must_change_password FROM usuarios ORDER BY username").fetchall()
        return {r["username"]: {"password_hash": r["password_hash"], "is_admin": bool(r["is_admin"]),
                                "must_change_password": bool(r["must_change_password"])} for r in rows}
    if name in _CHASE_KIND:
        rows = conn.execute("SELECT keyword, detail FROM reglas_chase WHERE tipo = ? ORDER BY orden", (_CHASE_KIND[name],)).fetchall()
        return {"rules": [{"keyword": r["keyword"], "detail": r["detail"]} for r in rows]}
    if name == PAGO_RULES:
        rows = conn.execute("SELECT keyword, sheet_name FROM reglas_pago_proveedores ORDER BY orden").fetchall()
        return {"rules": [{"keyword": r["keyword"], "sheet_name": r["sheet_name"]} for r in rows]}
    if name == DYNAMIC_SUPPLIERS:
        rows = conn.execute("SELECT key, definicion FROM proveedores_dinamicos ORDER BY orden").fetchall()
        return {"suppliers": {r["key"]: r["definicion"] for r in rows}}
    if name == CONTROLES_RAPIDOS:
        row = conn.execute("SELECT valor FROM ajustes WHERE nombre = ?", ("controles_rapidos",)).fetchone()
        return row["valor"] if row else {}
    if name == SALE_TAX:
        rows = conn.execute("SELECT year, month, asiento, registrado_en FROM sale_tax_registrado ORDER BY year, month").fetchall()
        return {f"{r['year']:04d}-{r['month']:02d}": {"asiento": r["asiento"] or "", "marked_at": r["registrado_en"]} for r in rows}
    raise ValueError(f"Configuración desconocida: {name}")


def _pg_write(name, payload):
    if name not in NAMES:
        raise ValueError(f"Configuración desconocida: {name}")
    conn = db.connect("config", None)
    try:
        with conn:  # todo o nada
            _pg_store(conn, name, payload)
            conn.execute("INSERT INTO guardados (nombre) VALUES (?) ON CONFLICT (nombre) DO NOTHING", (name,))
    finally:
        conn.close()


def _rules(payload):
    if isinstance(payload, dict):
        payload = payload.get("rules", [])
    return payload if isinstance(payload, list) else []


def _pg_store(conn, name, payload):
    if name == USERS:
        users = payload or {}
        # Se borran los que ya no están y se actualizan los demás (así no se
        # pierde la fecha de alta de cada usuario).
        for (username,) in conn.execute("SELECT username FROM usuarios").fetchall():
            if username not in users:
                conn.execute("DELETE FROM usuarios WHERE username = ?", (username,))
        for username, data in users.items():
            conn.execute(
                "INSERT INTO usuarios (username, password_hash, is_admin, must_change_password) VALUES (?, ?, ?, ?) "
                "ON CONFLICT (username) DO UPDATE SET password_hash = excluded.password_hash, is_admin = excluded.is_admin, "
                "must_change_password = excluded.must_change_password, actualizado = now()",
                (username, data["password_hash"], bool(data.get("is_admin")), bool(data.get("must_change_password"))),
            )
    elif name in _CHASE_KIND:
        kind = _CHASE_KIND[name]
        conn.execute("DELETE FROM reglas_chase WHERE tipo = ?", (kind,))
        for order, rule in enumerate(_rules(payload)):
            conn.execute("INSERT INTO reglas_chase (tipo, orden, keyword, detail) VALUES (?, ?, ?, ?)",
                         (kind, order, rule.get("keyword"), rule.get("detail")))
    elif name == PAGO_RULES:
        conn.execute("DELETE FROM reglas_pago_proveedores")
        for order, rule in enumerate(_rules(payload)):
            conn.execute("INSERT INTO reglas_pago_proveedores (orden, keyword, sheet_name) VALUES (?, ?, ?)",
                         (order, rule.get("keyword"), rule.get("sheet_name")))
    elif name == DYNAMIC_SUPPLIERS:
        suppliers = (payload or {}).get("suppliers", {}) if isinstance(payload, dict) else {}
        conn.execute("DELETE FROM proveedores_dinamicos")
        for order, (key, definition) in enumerate(suppliers.items()):
            conn.execute("INSERT INTO proveedores_dinamicos (key, orden, definicion) VALUES (?, ?, ?)",
                         (key, order, json.dumps(definition, ensure_ascii=False)))
    elif name == CONTROLES_RAPIDOS:
        conn.execute("INSERT INTO ajustes (nombre, valor) VALUES (?, ?) ON CONFLICT (nombre) DO UPDATE SET valor = excluded.valor, actualizado = now()",
                     ("controles_rapidos", json.dumps(payload or {}, ensure_ascii=False)))
    elif name == SALE_TAX:
        conn.execute("DELETE FROM sale_tax_registrado")
        for key, record in (payload or {}).items():
            conn.execute("INSERT INTO sale_tax_registrado (year, month, asiento, registrado_en) VALUES (?, ?, ?, ?)",
                         (int(key[:4]), int(key[5:7]), (record or {}).get("asiento") or "", (record or {}).get("marked_at")))
