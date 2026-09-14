import os
import logging
from flask import Flask

from .config import Config
from . import db as db_module
from .auth import seed_admin_user


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
