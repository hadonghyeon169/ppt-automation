"""
데이터 접근 계층. Flask 요청 컨텍스트 / 백그라운드 스레드 양쪽에서 쓸 수 있도록
매번 db_path를 받아 독립 커넥션으로 동작한다 (스레드 안전).
"""
import json
from . import db


def _conn(db_path):
    return db.get_conn_for_path(db_path)


# ── users ────────────────────────────────────────────────────────────────
def get_user_by_username(db_path, username):
    c = _conn(db_path)
    row = c.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    c.close()
    return dict(row) if row else None


def create_user(db_path, username, password_hash, role="admin"):
    c = _conn(db_path)
    c.execute("INSERT OR IGNORE INTO users (username, password_hash, role) VALUES (?,?,?)",
               (username, password_hash, role))
    c.commit()
    c.close()


def count_users(db_path):
    c = _conn(db_path)
    n = c.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]
    c.close()
    return n


def list_users(db_path):
    c = _conn(db_path)
    rows = c.execute("SELECT id, username, role, created_at FROM users ORDER BY created_at").fetchall()
    c.close()
    return [dict(r) for r in rows]


# ── api settings ─────────────────────────────────────────────────────────
def get_api_keys(db_path):
    c = _conn(db_path)
    rows = c.execute("SELECT key_name, key_value FROM api_settings").fetchall()
    c.close()
    return {r["key_name"]: r["key_value"] for r in rows}


def set_api_key(db_path, key_name, key_value):
    c = _conn(db_path)
    c.execute(
        "INSERT INTO api_settings (key_name, key_value, updated_at) VALUES (?, ?, datetime('now')) "
        "ON CONFLICT(key_name) DO UPDATE SET key_value=excluded.key_value, updated_at=datetime('now')",
        (key_name, key_value),
    )
    c.commit()
    c.close()


# ── projects ─────────────────────────────────────────────────────────────
def create_project(db_path, name, target_lang, original_filename, original_path, created_by, reference_path=None):
    c = _conn(db_path)
    cur = c.execute(
        "INSERT INTO projects (name, target_lang, original_filename, original_path, reference_path, created_by) "
        "VALUES (?,?,?,?,?,?)",
        (name, target_lang, original_filename, original_path, reference_path, created_by),
    )
    c.commit()
    pid = cur.lastrowid
    c.close()
    return pid


def get_project(db_path, project_id):
    c = _conn(db_path)
    row = c.execute("SELECT * FROM projects WHERE id = ?", (project_id,)).fetchone()
    c.close()
    return dict(row) if row else None


def list_projects(db_path):
    c = _conn(db_path)
    rows = c.execute("SELECT * FROM projects ORDER BY created_at DESC").fetchall()
    c.close()
    return [dict(r) for r in rows]


def update_project(db_path, project_id, **fields):
    if not fields:
        return
    set_clause = ", ".join(f"{k} = ?" for k in fields)
    values = list(fields.values())
    c = _conn(db_path)
    c.execute(f"UPDATE projects SET {set_clause}, updated_at = datetime('now') WHERE id = ?",
              (*values, project_id))
    c.commit()
    c.close()


def delete_project(db_path, project_id):
    c = _conn(db_path)
    c.execute("DELETE FROM projects WHERE id = ?", (project_id,))
    c.commit()
    c.close()


# ── slide_texts (번역 스냅샷) ─────────────────────────────────────────────
def replace_slide_texts(db_path, project_id, records):
    c = _conn(db_path)
    c.execute("DELETE FROM slide_texts WHERE project_id = ?", (project_id,))
    for r in records:
        c.execute(
            "INSERT INTO slide_texts (project_id, slide_index, shape_id, shape_name, shape_type, "
            "source_korean, source_chinese, translated_text, is_white_text, is_bold, font_size, applied) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,1)",
            (project_id, r["slide_index"], r["shape_id"], r.get("shape_name"), r.get("shape_type"),
             r.get("source_korean"), r.get("source_chinese"), r.get("translated_text"),
             int(bool(r.get("is_white_text"))), int(bool(r.get("is_bold"))), r.get("font_size")),
        )
    c.commit()
    c.close()


def list_slide_texts(db_path, project_id):
    c = _conn(db_path)
    rows = c.execute("SELECT * FROM slide_texts WHERE project_id = ? ORDER BY slide_index", (project_id,)).fetchall()
    c.close()
    return [dict(r) for r in rows]


# ── review flags ─────────────────────────────────────────────────────────
def add_review_flags(db_path, project_id, stage, flags):
    if not flags:
        return
    c = _conn(db_path)
    for f in flags:
        c.execute(
            "INSERT INTO review_flags (project_id, stage, slide_index, shape_name, source_text, "
            "translated_text, issue, severity) VALUES (?,?,?,?,?,?,?,?)",
            (project_id, stage, f.get("slide_index"), f.get("shape_name"), f.get("source_text"),
             f.get("translated_text"), f.get("issue"), f.get("severity", "warning")),
        )
    c.commit()
    c.close()


def clear_review_flags(db_path, project_id, stage):
    c = _conn(db_path)
    c.execute("DELETE FROM review_flags WHERE project_id = ? AND stage = ?", (project_id, stage))
    c.commit()
    c.close()


def list_review_flags(db_path, project_id, stage=None):
    c = _conn(db_path)
    if stage:
        rows = c.execute("SELECT * FROM review_flags WHERE project_id = ? AND stage = ? ORDER BY severity DESC, id",
                          (project_id, stage)).fetchall()
    else:
        rows = c.execute("SELECT * FROM review_flags WHERE project_id = ? ORDER BY stage, severity DESC, id",
                          (project_id,)).fetchall()
    c.close()
    return [dict(r) for r in rows]


def resolve_review_flag(db_path, flag_id, project_id):
    c = _conn(db_path)
    c.execute("UPDATE review_flags SET resolved = 1 WHERE id = ? AND project_id = ?", (flag_id, project_id))
    c.commit()
    c.close()


# ── audio assets ─────────────────────────────────────────────────────────
def replace_audio_assets_pending(db_path, project_id, slide_scripts):
    c = _conn(db_path)
    c.execute("DELETE FROM audio_assets WHERE project_id = ?", (project_id,))
    for slide_index, script_text in slide_scripts:
        c.execute(
            "INSERT INTO audio_assets (project_id, slide_index, script_text, status) VALUES (?,?,?, 'pending')",
            (project_id, slide_index, script_text),
        )
    c.commit()
    c.close()


def update_audio_asset(db_path, asset_id, **fields):
    set_clause = ", ".join(f"{k} = ?" for k in fields)
    values = list(fields.values())
    c = _conn(db_path)
    c.execute(f"UPDATE audio_assets SET {set_clause} WHERE id = ?", (*values, asset_id))
    c.commit()
    c.close()


def list_audio_assets(db_path, project_id):
    c = _conn(db_path)
    rows = c.execute("SELECT * FROM audio_assets WHERE project_id = ? ORDER BY slide_index", (project_id,)).fetchall()
    c.close()
    return [dict(r) for r in rows]


# ── events (진행 로그) ───────────────────────────────────────────────────
def add_event(db_path, project_id, message, level="info"):
    c = _conn(db_path)
    c.execute("INSERT INTO project_events (project_id, message, level) VALUES (?,?,?)",
              (project_id, message, level))
    c.commit()
    c.close()


def list_events(db_path, project_id, limit=200):
    c = _conn(db_path)
    rows = c.execute("SELECT * FROM project_events WHERE project_id = ? ORDER BY id DESC LIMIT ?",
                      (project_id, limit)).fetchall()
    c.close()
    return [dict(r) for r in rows]
