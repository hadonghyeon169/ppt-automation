import functools
from flask import session, redirect, url_for, request, current_app, g
from werkzeug.security import generate_password_hash, check_password_hash

from . import repo


def seed_admin_user(db_path, username, password):
    if repo.count_users(db_path) == 0:
        repo.create_user(db_path, username, generate_password_hash(password), role="admin")


def verify_login(db_path, username, password):
    user = repo.get_user_by_username(db_path, username)
    if not user:
        return None
    if not check_password_hash(user["password_hash"], password):
        return None
    return user


def login_required(view):
    @functools.wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("auth.login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


def current_user():
    if not session.get("user_id"):
        return None
    if "user" not in g:
        g.user = {"id": session["user_id"], "username": session.get("username"), "role": session.get("role")}
    return g.user
