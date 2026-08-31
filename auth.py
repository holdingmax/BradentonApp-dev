"""
User accounts: JSON-backed store with hashed passwords (werkzeug.security).

Pure data layer — permission checks (who's allowed to create/reset/delete
users) live in webapp.py's routes, not here.
"""

import json
import os

from werkzeug.security import check_password_hash, generate_password_hash

USERS_FILENAME = "users.json"


def _users_file_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), USERS_FILENAME)


def load_users():
    """Return {username: {"password_hash": str, "is_admin": bool}}."""
    path = _users_file_path()
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, encoding="utf-8") as handle:
            return json.load(handle)
    except (OSError, json.JSONDecodeError):
        return {}


def save_users(users):
    """Persist the full user dict atomically (temp file + os.replace)."""
    path = _users_file_path()
    temp_path = f"{path}.tmp"
    with open(temp_path, "w", encoding="utf-8") as handle:
        json.dump(users, handle, indent=2, ensure_ascii=False)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp_path, path)


def get_user(username):
    return load_users().get(username)


def list_users():
    """Return [{"username": str, "is_admin": bool}, ...], sorted by username."""
    users = load_users()
    return [
        {"username": name, "is_admin": data.get("is_admin", False)}
        for name, data in sorted(users.items())
    ]


def verify_user(username, password):
    """Return the user dict on a correct username/password match, else None."""
    user = get_user(username)
    if user is None or not password:
        return None
    if not check_password_hash(user["password_hash"], password):
        return None
    return user


def create_user(username, password, is_admin=False):
    username = username.strip()
    if not username:
        raise ValueError("El usuario no puede estar vacío.")
    if not password:
        raise ValueError("La contraseña no puede estar vacía.")
    users = load_users()
    if username in users:
        raise ValueError(f'Ya existe un usuario "{username}".')
    users[username] = {
        "password_hash": generate_password_hash(password),
        "is_admin": bool(is_admin),
    }
    save_users(users)


def set_password(username, new_password):
    if not new_password:
        raise ValueError("La contraseña no puede estar vacía.")
    users = load_users()
    if username not in users:
        raise ValueError(f'No existe el usuario "{username}".')
    users[username]["password_hash"] = generate_password_hash(new_password)
    save_users(users)


def delete_user(username, current_username=None):
    """Remove a user — refuses to delete yourself or the last remaining admin."""
    users = load_users()
    if username not in users:
        raise ValueError(f'No existe el usuario "{username}".')
    if username == current_username:
        raise ValueError("No podés eliminar tu propia cuenta mientras estás conectado con ella.")
    other_admins = [
        name for name, data in users.items()
        if data.get("is_admin") and name != username
    ]
    if users[username].get("is_admin") and not other_admins:
        raise ValueError("No se puede eliminar el último administrador.")
    del users[username]
    save_users(users)
