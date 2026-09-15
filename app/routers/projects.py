import os
import json
import threading
import uuid

from flask import (
    Blueprint, render_template, request, redirect, url_for, current_app,
    send_file, abort, jsonify, flash,
)
from werkzeug.utils import secure_filename

from .. import repo
from ..auth import login_required, current_user
from ..services import pipeline_runner, tts_typecast

bp = Blueprint("projects", __name__)

ALLOWED_EXT = {".pptx"}


def _allowed(filename):
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXT


@bp.route("/")
@login_required
def dashboard():
    db_path = current_app.config["DB_PATH"]
    projects = repo.list_projects(db_path)
    return render_template("dashboard.html", projects=projects, languages=current_app.config["LANGUAGES"])


@bp.route("/projects/new", methods=["GET", "POST"])
@login_required
def new_project():
    languages = current_app.config["LANGUAGES"]
    if request.method == "POST":
        file = request.files.get("pptx_file")
        ref_file = request.files.get("reference_file")
        name = request.form.get("name", "").strip()
        target_lang = request.form.get("target_lang")

        if not file or not file.filename or not _allowed(file.filename):
            flash("PPTX 파일을 선택해주세요.", "error")
            return redirect(url_for("projects.new_project"))
        if target_lang not in languages:
            flash("목표 언어를 선택해주세요.", "error")
            return redirect(url_for("projects.new_project"))

        upload_dir = current_app.config["UPLOAD_DIR"]
        token = uuid.uuid4().hex[:10]
        safe_name = secure_filename(file.filename)
        original_path = os.path.join(upload_dir, f"{token}_{safe_name}")
        file.save(original_path)

        reference_path = None
        if ref_file and ref_file.filename and _allowed(ref_file.filename):
            ref_safe = secure_filename(ref_file.filename)
            reference_path = os.path.join(upload_dir, f"{token}_ref_{ref_safe}")
            ref_file.save(reference_path)

        db_path = current_app.config["DB_PATH"]
        user = current_user()
        pid = repo.create_project(
            db_path, name or safe_name, target_lang, safe_name, original_path,
            user["username"] if user else "unknown", reference_path=reference_path,
        )
        repo.add_event(db_path, pid, "프로젝트 생성됨")
        return redirect(url_for("projects.detail", project_id=pid))

    return render_template("new_project.html", languages=languages)


@bp.route("/projects/<int:project_id>")
@login_required
def detail(project_id):
    db_path = current_app.config["DB_PATH"]
    project = repo.get_project(db_path, project_id)
    if not project:
        abort(404)
    flags = repo.list_review_flags(db_path, project_id)
    events = repo.list_events(db_path, project_id, limit=50)
    audio_assets = repo.list_audio_assets(db_path, project_id)
    slide_texts = repo.list_slide_texts(db_path, project_id)
    lang_meta = current_app.config["LANGUAGES"].get(project["target_lang"], {})
    projects_dir = current_app.config["PROJECTS_DIR"]
    preview_translated = _list_preview_files(projects_dir, project_id, "preview_translated")
    preview_final = _list_preview_files(projects_dir, project_id, "preview_final")
    try:
        failed_batches_count = len(json.loads(project.get("failed_batches_json") or "[]"))
    except (TypeError, ValueError):
        failed_batches_count = 0
    return render_template(
        "project_detail.html", project=project, flags=flags, events=events,
        audio_assets=audio_assets, slide_texts=slide_texts, lang_meta=lang_meta,
        preview_translated=preview_translated, preview_final=preview_final,
        failed_batches_count=failed_batches_count,
    )


@bp.route("/projects/<int:project_id>/status.json")
@login_required
def status_json(project_id):
    db_path = current_app.config["DB_PATH"]
    project = repo.get_project(db_path, project_id)
    if not project:
        abort(404)
    flags = repo.list_review_flags(db_path, project_id)
    return jsonify({
        "stage": project["stage"], "progress": project["progress"],
        "status_message": project["status_message"],
        "error_flags": sum(1 for f in flags if f["severity"] == "error" and not f["resolved"]),
        "warning_flags": sum(1 for f in flags if f["severity"] == "warning" and not f["resolved"]),
    })


@bp.route("/projects/<int:project_id>/translate", methods=["POST"])
@login_required
def start_translation(project_id):
    db_path = current_app.config["DB_PATH"]
    projects_dir = current_app.config["PROJECTS_DIR"]
    languages = current_app.config["LANGUAGES"]
    model = current_app.config["ANTHROPIC_MODEL"]

    t = threading.Thread(
        target=pipeline_runner.run_translation_stage,
        args=(db_path, projects_dir, project_id, languages, model),
        daemon=True,
    )
    t.start()
    flash("번역을 시작했습니다. 진행 상황은 아래에서 확인할 수 있습니다.", "success")
    return redirect(url_for("projects.detail", project_id=project_id))


@bp.route("/projects/<int:project_id>/translate/retry-failed", methods=["POST"])
@login_required
def retry_translation_failed(project_id):
    db_path = current_app.config["DB_PATH"]
    projects_dir = current_app.config["PROJECTS_DIR"]
    languages = current_app.config["LANGUAGES"]
    model = current_app.config["ANTHROPIC_MODEL"]

    t = threading.Thread(
        target=pipeline_runner.retry_failed_translation_batches,
        args=(db_path, projects_dir, project_id, languages, model),
        daemon=True,
    )
    t.start()
    flash("실패했던 배치만 재시도합니다. (성공한 배치는 다시 호출하지 않습니다)", "success")
    return redirect(url_for("projects.detail", project_id=project_id))


@bp.route("/projects/<int:project_id>/translation/confirm", methods=["POST"])
@login_required
def confirm_translation(project_id):
    db_path = current_app.config["DB_PATH"]
    force = request.form.get("force") == "1"
    flags = repo.list_review_flags(db_path, project_id, stage="translation")
    unresolved_errors = [f for f in flags if f["severity"] == "error" and not f["resolved"]]
    if unresolved_errors and not force:
        flash(f"해결되지 않은 오류 {len(unresolved_errors)}건이 있습니다. 확인 후 진행하거나 강제 진행을 선택하세요.", "error")
        return redirect(url_for("projects.detail", project_id=project_id))

    repo.update_project(db_path, project_id, stage="tts_pending",
                         status_message="번역 승인 완료 — 음성 생성 대기 중")
    repo.add_event(db_path, project_id, "번역 검수 승인 (사람 확인 완료)")
    flash("번역이 승인되었습니다. 다음 단계(음성 생성)로 진행할 수 있습니다.", "success")
    return redirect(url_for("projects.detail", project_id=project_id))


@bp.route("/projects/<int:project_id>/tts", methods=["GET", "POST"])
@login_required
def tts_setup(project_id):
    db_path = current_app.config["DB_PATH"]
    project = repo.get_project(db_path, project_id)
    if not project:
        abort(404)

    keys = pipeline_runner.get_effective_keys(db_path)
    voices, voice_error = [], None
    if keys["typecast"]:
        try:
            voices = tts_typecast.list_voices(keys["typecast"])
        except Exception as e:
            voice_error = str(e)

    if request.method == "POST":
        voice_id = request.form.get("voice_id", "").strip()
        audio_format = request.form.get("audio_format", "wav")
        if not voice_id:
            flash("보이스를 선택해주세요.", "error")
            return redirect(url_for("projects.tts_setup", project_id=project_id))

        projects_dir = current_app.config["PROJECTS_DIR"]
        t = threading.Thread(
            target=pipeline_runner.run_tts_stage,
            args=(db_path, projects_dir, project_id, voice_id),
            kwargs={"audio_format": audio_format, "languages": current_app.config["LANGUAGES"]},
            daemon=True,
        )
        t.start()
        flash("음성 생성을 시작했습니다.", "success")
        return redirect(url_for("projects.detail", project_id=project_id))

    return render_template("tts_setup.html", project=project, voices=voices, voice_error=voice_error)


@bp.route("/projects/<int:project_id>/tts/retry-failed", methods=["POST"])
@login_required
def tts_retry_failed(project_id):
    db_path = current_app.config["DB_PATH"]
    projects_dir = current_app.config["PROJECTS_DIR"]
    assets = repo.list_audio_assets(db_path, project_id)
    done = [a for a in assets if a["status"] == "done" and a.get("voice_id")]
    if not done:
        flash("이전에 성공한 오디오가 없어 사용할 보이스를 알 수 없습니다. 음성 생성을 처음부터 다시 진행해주세요.", "error")
        return redirect(url_for("projects.detail", project_id=project_id))
    voice_id = done[-1]["voice_id"]
    ext = os.path.splitext(done[-1]["file_path"] or "")[1].lstrip(".") or "wav"

    t = threading.Thread(
        target=pipeline_runner.run_tts_stage,
        args=(db_path, projects_dir, project_id, voice_id),
        kwargs={"audio_format": ext, "only_failed": True, "languages": current_app.config["LANGUAGES"]},
        daemon=True,
    )
    t.start()
    flash("실패한 슬라이드만 재시도합니다. (이미 성공한 오디오는 다시 만들지 않습니다)", "success")
    return redirect(url_for("projects.detail", project_id=project_id))


@bp.route("/projects/<int:project_id>/final/run", methods=["POST"])
@login_required
def run_final_review(project_id):
    db_path = current_app.config["DB_PATH"]
    projects_dir = current_app.config["PROJECTS_DIR"]
    languages = current_app.config["LANGUAGES"]
    t = threading.Thread(
        target=pipeline_runner.run_final_review_stage,
        args=(db_path, projects_dir, project_id, languages),
        daemon=True,
    )
    t.start()
    flash("최종 검수(AI 자동 비교)를 시작했습니다.", "success")
    return redirect(url_for("projects.detail", project_id=project_id))


@bp.route("/projects/<int:project_id>/final/confirm", methods=["POST"])
@login_required
def confirm_final(project_id):
    db_path = current_app.config["DB_PATH"]
    force = request.form.get("force") == "1"
    flags = repo.list_review_flags(db_path, project_id, stage="final")
    unresolved_errors = [f for f in flags if f["severity"] == "error" and not f["resolved"]]
    if unresolved_errors and not force:
        flash(f"최종 검수에서 해결되지 않은 오류 {len(unresolved_errors)}건이 있습니다. PPT가 100% 일치하지 않을 수 있습니다.", "error")
        return redirect(url_for("projects.detail", project_id=project_id))

    repo.update_project(db_path, project_id, stage="completed", status_message="최종 승인 완료")
    repo.add_event(db_path, project_id, "최종 승인 완료 (사람 확인 완료)")
    flash("최종 승인이 완료되었습니다.", "success")
    return redirect(url_for("projects.detail", project_id=project_id))


@bp.route("/projects/<int:project_id>/flags/<int:flag_id>/resolve", methods=["POST"])
@login_required
def resolve_flag(project_id, flag_id):
    db_path = current_app.config["DB_PATH"]
    repo.resolve_review_flag(db_path, flag_id, project_id)
    return redirect(request.referrer or url_for("projects.detail", project_id=project_id))


@bp.route("/projects/<int:project_id>/delete", methods=["POST"])
@login_required
def delete_project(project_id):
    db_path = current_app.config["DB_PATH"]
    repo.delete_project(db_path, project_id)
    flash("프로젝트를 삭제했습니다.", "success")
    return redirect(url_for("projects.dashboard"))


# ── 파일 다운로드/미리보기 ───────────────────────────────────────────────
@bp.route("/projects/<int:project_id>/download/<which>")
@login_required
def download(project_id, which):
    db_path = current_app.config["DB_PATH"]
    project = repo.get_project(db_path, project_id)
    if not project:
        abort(404)
    path_map = {
        "original": project["original_path"],
        "translated": project["translated_path"],
        "final": project["final_path"],
    }
    path = path_map.get(which)
    if not path or not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=True)


@bp.route("/projects/<int:project_id>/preview/<stage>/<filename>")
@login_required
def preview_image(project_id, stage, filename):
    projects_dir = current_app.config["PROJECTS_DIR"]
    safe_stage = secure_filename(stage)
    safe_file = secure_filename(filename)
    path = os.path.join(projects_dir, str(project_id), safe_stage, safe_file)
    if not os.path.exists(path):
        abort(404)
    return send_file(path)


@bp.route("/projects/<int:project_id>/audio/<int:asset_id>")
@login_required
def audio_file(project_id, asset_id):
    db_path = current_app.config["DB_PATH"]
    assets = repo.list_audio_assets(db_path, project_id)
    asset = next((a for a in assets if a["id"] == asset_id), None)
    if not asset or not asset["file_path"] or not os.path.exists(asset["file_path"]):
        abort(404)
    return send_file(asset["file_path"])


def _list_preview_files(projects_dir, project_id, stage_dir):
    d = os.path.join(projects_dir, str(project_id), stage_dir)
    if not os.path.isdir(d):
        return []
    return sorted(f for f in os.listdir(d) if f.endswith(".png"))
