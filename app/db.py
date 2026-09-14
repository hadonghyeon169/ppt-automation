"""
아주 얇은 sqlite3 래퍼.
외부 ORM 의존성 없이(배포 환경 어디서나 바로 동작하도록) 표준 라이브러리만 사용합니다.
"""
import sqlite3
import os
import threading
from contextlib import contextmanager

_local = threading.local()

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'member',   -- 'admin' | 'member'  (다인원 확장 대비)
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS api_settings (
    key_name TEXT PRIMARY KEY,             -- anthropic_api_key / openai_api_key / typecast_api_key
    key_value TEXT,
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    target_lang TEXT NOT NULL,             -- th-TH / ru-RU / vi-VN / id-ID ...
    stage TEXT NOT NULL DEFAULT 'uploaded',
    -- uploaded -> translating -> review_translation -> tts_pending -> tts_review -> review_final -> completed -> failed
    progress INTEGER NOT NULL DEFAULT 0,   -- 0-100, 현재 단계 내 진행률
    status_message TEXT,
    original_filename TEXT,
    original_path TEXT,
    reference_path TEXT,                   -- 완성본 대조용 (같은 시리즈 기 번역본), 선택
    translated_path TEXT,
    final_path TEXT,
    created_by TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS slide_texts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    slide_index INTEGER NOT NULL,
    shape_id TEXT NOT NULL,                -- 도형 내부 식별용 (슬라이드 인덱스 내 순번)
    shape_name TEXT,
    shape_type TEXT,                       -- A / B / C (스킬 문서 유형)
    source_korean TEXT,
    source_chinese TEXT,
    translated_text TEXT,
    is_white_text INTEGER NOT NULL DEFAULT 0,
    is_bold INTEGER NOT NULL DEFAULT 0,
    font_size REAL,
    applied INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS review_flags (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    stage TEXT NOT NULL,                   -- 'translation' | 'final'
    slide_index INTEGER,
    shape_name TEXT,
    source_text TEXT,
    translated_text TEXT,
    issue TEXT,
    severity TEXT NOT NULL DEFAULT 'warning',  -- 'info' | 'warning' | 'error'
    resolved INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS audio_assets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    slide_index INTEGER NOT NULL,
    script_text TEXT,
    file_path TEXT,
    voice_id TEXT,
    duration_sec REAL,
    status TEXT NOT NULL DEFAULT 'pending',  -- pending / done / error
    error_message TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS project_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    message TEXT NOT NULL,
    level TEXT NOT NULL DEFAULT 'info',
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);
"""


def get_db_path():
    from flask import current_app
    return current_app.config["DB_PATH"]


def get_conn():
    """요청/스레드별로 커넥션을 재사용."""
    if not hasattr(_local, "conn") or _local.conn is None:
        path = get_db_path()
        conn = sqlite3.connect(path, check_same_thread=False, timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        _local.conn = conn
    return _local.conn


def get_conn_for_path(db_path):
    """Flask 앱 컨텍스트 밖(백그라운드 스레드)에서 쓰기 위한 독립 커넥션 생성."""
    conn = sqlite3.connect(db_path, check_same_thread=False, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


def init_db(db_path):
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.executescript(SCHEMA)
    conn.commit()
    conn.close()


@contextmanager
def tx(conn):
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
