"""
db.py - SQLiteデータベース接続とスキーマ管理

テーブル:
  users      : ユーザー（ログイン情報・管理者フラグ）
  usage_log  : AI API呼び出しの利用ログ
  settings   : 管理画面で設定するAPIキー等
"""
import os
import sqlite3
from contextlib import contextmanager
from threading import Lock

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "app.db")
_init_lock = Lock()
_initialized = False


def _ensure_dir():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)


@contextmanager
def get_conn():
    """SQLite接続。呼び出すたびに新しい接続を返す（スレッド安全）"""
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


def init_db():
    """初回起動時にテーブルを作成"""
    global _initialized
    with _init_lock:
        if _initialized:
            return
        with get_conn() as conn:
            conn.executescript("""
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

                CREATE INDEX IF NOT EXISTS idx_usage_user_time
                    ON usage_log(user_id, timestamp);
                CREATE INDEX IF NOT EXISTS idx_usage_time
                    ON usage_log(timestamp);

                CREATE TABLE IF NOT EXISTS settings (
                    key        TEXT PRIMARY KEY,
                    value      TEXT,
                    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
                );
            """)
        _initialized = True


def get_setting(key: str, default: str = "") -> str:
    init_db()
    with get_conn() as conn:
        row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default


def set_setting(key: str, value: str) -> None:
    init_db()
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO settings(key, value, updated_at)
            VALUES(?, ?, datetime('now'))
            ON CONFLICT(key) DO UPDATE SET
                value = excluded.value,
                updated_at = datetime('now')
        """, (key, value))


def delete_setting(key: str) -> None:
    init_db()
    with get_conn() as conn:
        conn.execute("DELETE FROM settings WHERE key=?", (key,))
