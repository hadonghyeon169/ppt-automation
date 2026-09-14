from flask import Blueprint, render_template, request, redirect, url_for, current_app, flash

from .. import repo
from ..auth import login_required

bp = Blueprint("settings", __name__, url_prefix="/settings")

KEY_FIELDS = [
    ("anthropic_api_key", "Claude API 키 (번역용)"),
    ("openai_api_key", "OpenAI(GPT) API 키 (번역 검수용)"),
    ("typecast_api_key", "타입캐스트 API 키 (음성 생성용)"),
]


def _mask(value):
    if not value:
        return ""
    if len(value) <= 8:
        return "*" * len(value)
    return value[:4] + "..." + value[-4:]


@bp.route("/", methods=["GET", "POST"])
@login_required
def index():
    db_path = current_app.config["DB_PATH"]
    if request.method == "POST":
        for key_name, _ in KEY_FIELDS:
            value = request.form.get(key_name, "").strip()
            if value:  # 빈 값이면 기존 키 유지 (실수로 지우는 것 방지)
                repo.set_api_key(db_path, key_name, value)
        flash("API 키가 저장되었습니다.", "success")
        return redirect(url_for("settings.index"))

    current_keys = repo.get_api_keys(db_path)
    masked = {k: _mask(current_keys.get(k, "")) for k, _ in KEY_FIELDS}
    return render_template("settings.html", key_fields=KEY_FIELDS, masked=masked)
