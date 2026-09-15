import functools
from flask import session, redirect, url_for, request, current_app, g
from werkzeug.security import generate_password_hash, check_password_hash

from . import repo


def seed_admin_user(db_path, username, password):
    if repo.count_users(db_path) == 0:
        repo.create_user(db_path, username, generate_password_hash(password), role="admin")


def reset_admin_password(db_path, username, password):
    """ADMIN_USERNAME/ADMIN_PASSWORD 환경변수로 지정된 계정의 비밀번호를 강제로
    재설정한다 (없으면 새로 만든다). seed_admin_user와 달리 이미 사용자가 있어도
    동작한다 — FORCE_ADMIN_RESET=1 환경변수가 켜져 있을 때만 앱 시작 시 호출되는
    명시적 1회성 재설정 경로다 (평소엔 절대 자동으로 비밀번호를 덮어쓰지 않는다)."""
    user = repo.get_user_by_username(db_path, username)
    pw_hash = generate_password_hash(password)
    if user:
        repo.update_user_password(db_path, user["id"], pw_hash)
    else:
        repo.create_user(db_path, username, pw_hash, role="admin")


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
