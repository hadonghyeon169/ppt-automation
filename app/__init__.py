import os
import logging
from flask import Flask

from .config import Config
from . import db as db_module
from .auth import seed_admin_user, reset_admin_password


def create_app():
    logging.basicConfig(level=logging.INFO)
    cfg = Config()
    cfg.ensure_dirs()

    app = Flask(__name__)
    app.config["SECRET_KEY"] = cfg.SECRET_KEY
    app.config["MAX_CONTENT_LENGTH"] = cfg.MAX_CONTENT_LENGTH
    app.config["DB_PATH"] = cfg.DB_PATH
    app.config["UPLOAD_DIR"] = cfg.UPLOAD_DIR
    app.config["PROJECTS_DIR"] = cfg.PROJECTS_DIR
    app.config["LANGUAGES"] = cfg.LANGUAGES
    app.config["ANTHROPIC_MODEL"] = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-5")
    app.config["OPENAI_MODEL"] = os.environ.get("OPENAI_MODEL", "gpt-4o")

    db_module.init_db(cfg.DB_PATH)
    seed_admin_user(cfg.DB_PATH, cfg.ADMIN_USERNAME, cfg.ADMIN_PASSWORD)
    # FORCE_ADMIN_RESET=1일 때만 명시적으로 비밀번호를 재설정한다 (평소엔 seed만 하고
    # 기존 계정은 절대 건드리지 않음 — 배포마다 비밀번호가 초기화되면 안 되므로).
    if os.environ.get("FORCE_ADMIN_RESET") == "1":
        reset_admin_password(cfg.DB_PATH, cfg.ADMIN_USERNAME, cfg.ADMIN_PASSWORD)
        logging.getLogger(__name__).warning(
            "FORCE_ADMIN_RESET=1 감지 — 관리자 계정 '%s'의 비밀번호를 재설정했습니다. "
            "재설정 후에는 이 환경변수를 꼭 꺼주세요.", cfg.ADMIN_USERNAME,
        )

    from .routers.auth_routes import bp as auth_bp
    from .routers.projects import bp as projects_bp
    from .routers.settings_routes import bp as settings_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(projects_bp)
    app.register_blueprint(settings_bp)

    @app.teardown_appcontext
    def close_db(exception=None):
        from flask import g
        conn = getattr(g, "_db_conn", None)
        if conn is not None:
            conn.close()

    return app
