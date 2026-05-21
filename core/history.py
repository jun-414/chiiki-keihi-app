"""
history.py - 経費処理履歴（出納簿に書き込んだ行データ）の保存・取得

画像は保存せず、行データ（日付・取引先・金額・科目など）のみをJSONで保持。
データ量はごく軽量（1セッション数KB程度）。
"""
import json
from typing import Optional

from .db import get_conn, init_db, db_insert

# 履歴に残すフィールド（内部管理フィールドは除外）
_KEEP_KEYS = ("date", "vendor", "memo", "amount", "kamoku", "jigyo", "_kind")


def save_history(user_id: Optional[int], username: str,
                 records: list, sheet_name: str = "") -> int:
    """
    1回の書き込み処理ぶんの履歴を保存。
    records: write_records相当（dictのリスト）
    Returns: history id
    """
    init_db()
    clean = []
    income_total = 0
    expense_total = 0
    for r in records:
        amount = int(r.get("amount", 0) or 0)
        kind = r.get("_kind", "expense")
        if kind == "income":
            income_total += amount
        else:
            expense_total += amount
        clean.append({k.lstrip("_") if k == "_kind" else k: r.get(k)
                      for k in _KEEP_KEYS})

    records_json = json.dumps(clean, ensure_ascii=False)
    with get_conn() as conn:
        return db_insert(conn, """
            INSERT INTO process_history
                (user_id, username, record_count,
                 income_total, expense_total, sheet_name, records_json)
            VALUES(?, ?, ?, ?, ?, ?, ?)
        """, (user_id, username, len(clean),
              income_total, expense_total, sheet_name, records_json))


def list_history(user_id: int, limit: int = 200) -> list:
    """指定ユーザーの履歴一覧（新しい順、records_jsonは含めない軽量版）"""
    init_db()
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT id, processed_at, record_count, income_total, expense_total, sheet_name
              FROM process_history
              WHERE user_id = ?
              ORDER BY processed_at DESC, id DESC
              LIMIT ?
        """, (user_id, limit)).fetchall()
        return [dict(r) for r in rows]


def get_history_detail(history_id: int, user_id: int = None) -> Optional[dict]:
    """1セッションの詳細（行データ込み）。user_id指定時は本人のものだけ返す"""
    init_db()
    with get_conn() as conn:
        if user_id is not None:
            row = conn.execute("""
                SELECT * FROM process_history WHERE id = ? AND user_id = ?
            """, (history_id, user_id)).fetchone()
        else:
            row = conn.execute(
                "SELECT * FROM process_history WHERE id = ?", (history_id,)
            ).fetchone()
        if not row:
            return None
        d = dict(row)
        try:
            d["records"] = json.loads(d.get("records_json") or "[]")
        except Exception:
            d["records"] = []
        return d


def delete_history(history_id: int, user_id: int = None) -> None:
    init_db()
    with get_conn() as conn:
        if user_id is not None:
            conn.execute(
                "DELETE FROM process_history WHERE id = ? AND user_id = ?",
                (history_id, user_id),
            )
        else:
            conn.execute("DELETE FROM process_history WHERE id = ?", (history_id,))
