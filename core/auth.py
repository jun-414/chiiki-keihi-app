"""
auth.py - パスワードハッシュ化・ログイン認証・ユーザー管理

stdlibのみ使用（hashlib PBKDF2-HMAC-SHA256）
"""
import os
import hmac
import hashlib
import secrets
from typing import Optional

from .db import get_conn, init_db

_ITERATIONS = 200_000
_SALT_BYTES = 16


def _hash_password(password: str, salt: bytes = None) -> str:
    """PBKDF2でパスワードをハッシュ化。'pbkdf2_sha256$iter$salt_hex$hash_hex' 形式"""
    if salt is None:
        salt = secrets.token_bytes(_SALT_BYTES)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, _ITERATIONS)
    return f"pbkdf2_sha256${_ITERATIONS}${salt.hex()}${dk.hex()}"


def _verify_password(password: str, stored: str) -> bool:
    try:
        algo, iter_str, salt_hex, hash_hex = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(hash_hex)
        iterations = int(iter_str)
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
        return hmac.compare_digest(dk, expected)
    except Exception:
        return False


# ===== ユーザー操作 =====

def has_any_user() -> bool:
    init_db()
    with get_conn() as conn:
        row = conn.execute("SELECT COUNT(*) AS c FROM users").fetchone()
        return row["c"] > 0


def create_user(username: str, password: str, display_name: str = "",
                is_admin: bool = False) -> int:
    """ユーザーを作成。既存usernameならValueError"""
    init_db()
    username = (username or "").strip()
    if not username:
        raise ValueError("ユーザー名を入力してください")
    if len(password) < 6:
        raise ValueError("パスワードは6文字以上にしてください")
    pwd_hash = _hash_password(password)
    display = (display_name or username).strip()
    with get_conn() as conn:
        try:
            cur = conn.execute("""
                INSERT INTO users(username, display_name, password_hash, is_admin)
                VALUES(?, ?, ?, ?)
            """, (username, display, pwd_hash, 1 if is_admin else 0))
            return cur.lastrowid
        except Exception as e:
            if "UNIQUE" in str(e):
                raise ValueError(f"ユーザー名「{username}」はすでに使用されています")
            raise


def authenticate(username: str, password: str) -> Optional[dict]:
    """ログイン認証。成功時はuser dict、失敗時はNone"""
    init_db()
    with get_conn() as conn:
        row = conn.execute("""
            SELECT id, username, display_name, password_hash, is_admin, is_active
              FROM users WHERE username=?
        """, (username,)).fetchone()
        if not row:
            return None
        if not row["is_active"]:
            return None
        if not _verify_password(password, row["password_hash"]):
            return None
        return {
            "id": row["id"],
            "username": row["username"],
            "display_name": row["display_name"],
            "is_admin": bool(row["is_admin"]),
        }


def list_users() -> list:
    init_db()
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT id, username, display_name, is_admin, is_active, created_at
              FROM users ORDER BY created_at ASC
        """).fetchall()
        return [dict(r) for r in rows]


def update_user_password(user_id: int, new_password: str) -> None:
    if len(new_password) < 6:
        raise ValueError("パスワードは6文字以上にしてください")
    init_db()
    pwd_hash = _hash_password(new_password)
    with get_conn() as conn:
        conn.execute("UPDATE users SET password_hash=? WHERE id=?", (pwd_hash, user_id))


def update_user_profile(user_id: int,
                        username: str = None,
                        display_name: str = None) -> None:
    """ユーザー名・表示名を更新。usernameは半角英数チェック＋重複チェックあり"""
    init_db()
    fields, params = [], []

    if username is not None:
        username = (username or "").strip()
        if not username:
            raise ValueError("ユーザー名を入力してください")
        # 重複チェック（自分以外）
        with get_conn() as conn:
            row = conn.execute(
                "SELECT id FROM users WHERE username=? AND id<>?",
                (username, user_id),
            ).fetchone()
            if row:
                raise ValueError(f"ユーザー名「{username}」はすでに使用されています")
        fields.append("username=?")
        params.append(username)

    if display_name is not None:
        display_name = (display_name or "").strip()
        if not display_name:
            raise ValueError("表示名を入力してください")
        fields.append("display_name=?")
        params.append(display_name)

    if not fields:
        return
    params.append(user_id)
    with get_conn() as conn:
        conn.execute(f"UPDATE users SET {', '.join(fields)} WHERE id=?", params)


def set_user_admin(user_id: int, is_admin: bool) -> None:
    init_db()
    with get_conn() as conn:
        conn.execute("UPDATE users SET is_admin=? WHERE id=?",
                     (1 if is_admin else 0, user_id))


def set_user_active(user_id: int, is_active: bool) -> None:
    init_db()
    with get_conn() as conn:
        conn.execute("UPDATE users SET is_active=? WHERE id=?",
                     (1 if is_active else 0, user_id))


def delete_user(user_id: int) -> None:
    init_db()
    with get_conn() as conn:
        conn.execute("DELETE FROM users WHERE id=?", (user_id,))
