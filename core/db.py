"""
db.py - データベース接続とスキーマ管理（SQLite / Supabase(Postgres) 両対応）

バックエンドの自動判定:
  - Streamlit Secrets または環境変数に SUPABASE_DB_URL / DATABASE_URL があれば
    → Postgres（Supabase）を使用（本番・データ永続）
  - なければ → ローカル SQLite（data/app.db）を使用（開発）

テーブル:
  users           : ユーザー（ログイン情報・管理者フラグ）
  usage_log       : AI API呼び出しの利用ログ
  settings        : 管理画面で設定するAPIキー等
  process_history : 経費処理の履歴（行データJSON）
"""
import os
from datetime import datetime
from contextlib import contextmanager
from threading import Lock

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "app.db")
_init_lock = Lock()
_initialized = False


def _pg_url() -> str:
    """Postgres接続URLを Secrets / 環境変数 から取得（無ければ空文字）"""
    # Streamlit Secrets
    try:
        import streamlit as st
        for k in ("SUPABASE_DB_URL", "DATABASE_URL"):
            v = st.secrets.get(k, "")
            if v:
                return str(v)
    except Exception:
        pass
    # 環境変数
    return os.environ.get("SUPABASE_DB_URL") or os.environ.get("DATABASE_URL") or ""


PG_URL = _pg_url()
USE_PG = bool(PG_URL)


# =========================================================
# Postgres用のラッパー（sqlite3風のインターフェースに揃える）
# =========================================================
class _PGConn:
    """psycopg2接続を sqlite3.Connection 風に包む（?プレースホルダ→%s 変換）"""
    def __init__(self, raw):
        self._raw = raw

    def execute(self, sql, params=()):
        sql = sql.replace("?", "%s")
        cur = self._raw.cursor()
        cur.execute(sql, params)
        return cur  # RealDictCursor → fetchone/fetchall が dict を返す

    def executescript(self, sql):
        cur = self._raw.cursor()
        cur.execute(sql)
        return cur

    def commit(self):
        self._raw.commit()

    def rollback(self):
        self._raw.rollback()

    def close(self):
        self._raw.close()


def _ensure_dir():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)


@contextmanager
def get_conn():
    """DB接続。Postgres / SQLite を自動切替。"""
    if USE_PG:
        import psycopg2
        import psycopg2.extras
        raw = psycopg2.connect(PG_URL, cursor_factory=psycopg2.extras.RealDictCursor)
        conn = _PGConn(raw)
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
    else:
        import sqlite3
        _ensure_dir()
        conn = sqlite3.connect(DB_PATH, timeout=10.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()


# =========================================================
# スキーマ定義（バックエンドごと）
# =========================================================
_DDL_SQLITE = """
CREATE TABLE IF NOT EXISTS users (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    username      TEXT UNIQUE NOT NULL,
    display_name  TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    is_admin      INTEGER NOT NULL DEFAULT 0,
    is_active     INTEGER NOT NULL DEFAULT 1,
    created_at    TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS usage_log (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id        INTEGER REFERENCES users(id) ON DELETE SET NULL,
    username       TEXT NOT NULL,
    timestamp      TEXT NOT NULL DEFAULT (datetime('now')),
    provider       TEXT NOT NULL,
    model          TEXT NOT NULL,
    input_tokens   INTEGER NOT NULL DEFAULT 0,
    output_tokens  INTEGER NOT NULL DEFAULT 0,
    cost_usd       REAL NOT NULL DEFAULT 0,
    cost_jpy       REAL NOT NULL DEFAULT 0,
    filename       TEXT,
    success        INTEGER NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_usage_user_time ON usage_log(user_id, timestamp);
CREATE INDEX IF NOT EXISTS idx_usage_time ON usage_log(timestamp);
CREATE TABLE IF NOT EXISTS settings (
    key        TEXT PRIMARY KEY,
    value      TEXT,
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS process_history (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id       INTEGER REFERENCES users(id) ON DELETE CASCADE,
    username      TEXT NOT NULL,
    processed_at  TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    record_count  INTEGER NOT NULL DEFAULT 0,
    income_total  INTEGER NOT NULL DEFAULT 0,
    expense_total INTEGER NOT NULL DEFAULT 0,
    sheet_name    TEXT,
    records_json  TEXT NOT NULL DEFAULT '[]'
);
CREATE INDEX IF NOT EXISTS idx_history_user_time ON process_history(user_id, processed_at DESC);
"""

# Postgres: 時刻は JST の文字列(TEXT)で保持し、SQLiteと同じ "YYYY-MM-DD HH:MM:SS" 形式に揃える
_JST_NOW = "to_char(now() AT TIME ZONE 'Asia/Tokyo', 'YYYY-MM-DD HH24:MI:SS')"
_DDL_PG = f"""
CREATE TABLE IF NOT EXISTS users (
    id            SERIAL PRIMARY KEY,
    username      TEXT UNIQUE NOT NULL,
    display_name  TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    is_admin      INTEGER NOT NULL DEFAULT 0,
    is_active     INTEGER NOT NULL DEFAULT 1,
    created_at    TEXT NOT NULL DEFAULT {_JST_NOW}
);
CREATE TABLE IF NOT EXISTS usage_log (
    id             SERIAL PRIMARY KEY,
    user_id        INTEGER REFERENCES users(id) ON DELETE SET NULL,
    username       TEXT NOT NULL,
    timestamp      TEXT NOT NULL DEFAULT {_JST_NOW},
    provider       TEXT NOT NULL,
    model          TEXT NOT NULL,
    input_tokens   INTEGER NOT NULL DEFAULT 0,
    output_tokens  INTEGER NOT NULL DEFAULT 0,
    cost_usd       DOUBLE PRECISION NOT NULL DEFAULT 0,
    cost_jpy       DOUBLE PRECISION NOT NULL DEFAULT 0,
    filename       TEXT,
    success        INTEGER NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_usage_user_time ON usage_log(user_id, timestamp);
CREATE INDEX IF NOT EXISTS idx_usage_time ON usage_log(timestamp);
CREATE TABLE IF NOT EXISTS settings (
    key        TEXT PRIMARY KEY,
    value      TEXT,
    updated_at TEXT NOT NULL DEFAULT {_JST_NOW}
);
CREATE TABLE IF NOT EXISTS process_history (
    id            SERIAL PRIMARY KEY,
    user_id       INTEGER REFERENCES users(id) ON DELETE CASCADE,
    username      TEXT NOT NULL,
    processed_at  TEXT NOT NULL DEFAULT {_JST_NOW},
    record_count  INTEGER NOT NULL DEFAULT 0,
    income_total  INTEGER NOT NULL DEFAULT 0,
    expense_total INTEGER NOT NULL DEFAULT 0,
    sheet_name    TEXT,
    records_json  TEXT NOT NULL DEFAULT '[]'
);
CREATE INDEX IF NOT EXISTS idx_history_user_time ON process_history(user_id, processed_at DESC);
"""


def init_db():
    """初回起動時にテーブルを作成"""
    global _initialized
    with _init_lock:
        if _initialized:
            return
        with get_conn() as conn:
            conn.executescript(_DDL_PG if USE_PG else _DDL_SQLITE)
        _initialized = True


def db_insert(conn, sql: str, params=()):
    """INSERT して新しい行の id を返す（両バックエンド対応）"""
    if USE_PG:
        sql2 = sql.rstrip().rstrip(";") + " RETURNING id"
        row = conn.execute(sql2, params).fetchone()
        return row["id"] if row else None
    cur = conn.execute(sql, params)
    return cur.lastrowid


# =========================================================
# 設定値（settings テーブル）
# =========================================================
def get_setting(key: str, default: str = "") -> str:
    init_db()
    with get_conn() as conn:
        row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default


def set_setting(key: str, value: str) -> None:
    init_db()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO settings(key, value, updated_at)
            VALUES(?, ?, ?)
            ON CONFLICT(key) DO UPDATE SET
                value = excluded.value,
                updated_at = excluded.updated_at
        """, (key, value, ts))


def delete_setting(key: str) -> None:
    init_db()
    with get_conn() as conn:
        conn.execute("DELETE FROM settings WHERE key=?", (key,))
