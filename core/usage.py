"""
usage.py - AI API利用ログの記録と集計
"""
from typing import Optional

from .db import get_conn, init_db
from .pricing import calc_cost


def log_usage(user_id: Optional[int], username: str,
              provider: str, model: str,
              input_tokens: int, output_tokens: int,
              filename: str = "", success: bool = True,
              usd_jpy: float = 150.0) -> None:
    """1回のAI呼び出しを記録"""
    init_db()
    cost_usd, cost_jpy = calc_cost(provider, model,
                                   input_tokens, output_tokens, usd_jpy)
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO usage_log
                (user_id, username, provider, model,
                 input_tokens, output_tokens, cost_usd, cost_jpy,
                 filename, success)
            VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (user_id, username, provider, model,
              input_tokens, output_tokens, cost_usd, cost_jpy,
              filename, 1 if success else 0))


def total_summary(since: str = "", until: str = "") -> dict:
    """期間内の全体サマリ"""
    init_db()
    where, params = _build_where(since, until)
    with get_conn() as conn:
        row = conn.execute(f"""
            SELECT COUNT(*) AS calls,
                   COALESCE(SUM(input_tokens), 0)  AS input_tokens,
                   COALESCE(SUM(output_tokens), 0) AS output_tokens,
                   COALESCE(SUM(cost_usd), 0)      AS cost_usd,
                   COALESCE(SUM(cost_jpy), 0)      AS cost_jpy,
                   COALESCE(SUM(CASE WHEN success=1 THEN 1 ELSE 0 END), 0) AS success_calls
              FROM usage_log
              {where}
        """, params).fetchone()
        return dict(row)


def per_user_summary(since: str = "", until: str = "") -> list:
    """ユーザー別の集計"""
    init_db()
    where, params = _build_where(since, until)
    with get_conn() as conn:
        rows = conn.execute(f"""
            SELECT username,
                   COUNT(*)                        AS calls,
                   COALESCE(SUM(input_tokens), 0)  AS input_tokens,
                   COALESCE(SUM(output_tokens), 0) AS output_tokens,
                   COALESCE(SUM(cost_usd), 0)      AS cost_usd,
                   COALESCE(SUM(cost_jpy), 0)      AS cost_jpy
              FROM usage_log
              {where}
              GROUP BY username
              ORDER BY cost_jpy DESC
        """, params).fetchall()
        return [dict(r) for r in rows]


def per_day_summary(since: str = "", until: str = "") -> list:
    init_db()
    where, params = _build_where(since, until)
    with get_conn() as conn:
        rows = conn.execute(f"""
            SELECT substr(timestamp, 1, 10) AS day,
                   COUNT(*)                        AS calls,
                   COALESCE(SUM(cost_jpy), 0)      AS cost_jpy
              FROM usage_log
              {where}
              GROUP BY day
              ORDER BY day ASC
        """, params).fetchall()
        return [dict(r) for r in rows]


def user_summary(user_id: int, since: str = "", until: str = "") -> dict:
    """指定ユーザーの期間サマリ"""
    init_db()
    where_clauses = ["user_id = ?"]
    params = [user_id]
    if since:
        where_clauses.append("timestamp >= ?")
        params.append(since)
    if until:
        where_clauses.append("timestamp <= ?")
        params.append(until)
    where = "WHERE " + " AND ".join(where_clauses)
    with get_conn() as conn:
        row = conn.execute(f"""
            SELECT COUNT(*) AS calls,
                   COALESCE(SUM(input_tokens), 0)  AS input_tokens,
                   COALESCE(SUM(output_tokens), 0) AS output_tokens,
                   COALESCE(SUM(cost_usd), 0)      AS cost_usd,
                   COALESCE(SUM(cost_jpy), 0)      AS cost_jpy
              FROM usage_log
              {where}
        """, params).fetchone()
        return dict(row)


def recent_logs(limit: int = 100) -> list:
    init_db()
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT timestamp, username, provider, model,
                   input_tokens, output_tokens, cost_jpy, filename, success
              FROM usage_log
              ORDER BY id DESC
              LIMIT ?
        """, (limit,)).fetchall()
        return [dict(r) for r in rows]


def _build_where(since: str, until: str) -> tuple:
    clauses, params = [], []
    if since:
        clauses.append("timestamp >= ?")
        params.append(since)
    if until:
        clauses.append("timestamp <= ?")
        params.append(until)
    where = "WHERE " + " AND ".join(clauses) if clauses else ""
    return where, params
