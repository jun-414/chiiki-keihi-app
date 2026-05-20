"""
pages/admin.py — 管理者専用ページ（URL: /admin）

タブ:
  - 利用状況: 消費クレジット / ユーザー別 / 日別 / 直近ログ
  - API設定:  Claude / Gemini のAPIキー登録
  - ユーザー管理: 追加 / パスワード変更 / 権限 / 有効化
  - 詳細設定:  為替レート、デフォルトプロバイダ
"""
import os
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

from core.auth import (
    authenticate, has_any_user, create_user, list_users,
    update_user_password, update_user_profile,
    set_user_admin, set_user_active, delete_user,
)
from core.db import get_setting, set_setting, delete_setting
from core.usage import (
    total_summary, per_user_summary, per_day_summary, recent_logs,
)
from core.pricing import DEFAULT_USD_JPY
from core.theme import apply_theme, render_header, render_kpi_strip, render_user_badge


st.set_page_config(page_title="管理画面 | 経費管理", page_icon="⚙️", layout="wide")
apply_theme()


# ===== 認証チェック =====
def _login_form():
    render_header(title="管理画面", subtitle="管理者ログイン", icon="⚙️")
    _l, _c, _r = st.columns([1, 2, 1])
    with _c:
        with st.container(border=True):
            st.markdown("#### 🔐 ログイン")
            with st.form("admin_login"):
                username = st.text_input("ユーザー名")
                password = st.text_input("パスワード", type="password")
                ok = st.form_submit_button("ログイン", type="primary",
                                           use_container_width=True)
        if ok:
            user = authenticate(username, password)
            if user:
                st.session_state["auth_user"] = user
                st.rerun()
            else:
                st.error("ユーザー名またはパスワードが違います")


if not has_any_user():
    render_header(title="管理画面", subtitle="初回セットアップ未完了", icon="⚙️")
    st.warning("まずトップページから初回セットアップを完了してください。")
    st.page_link("app.py", label="← トップに戻る")
    st.stop()

current = st.session_state.get("auth_user")
if not current:
    _login_form()
    st.stop()

if not current.get("is_admin"):
    render_header(title="管理画面", subtitle="アクセス権限なし", icon="⚙️")
    st.error("この画面は管理者のみアクセス可能です。")
    st.page_link("app.py", label="← トップに戻る")
    st.stop()


# ===== ヘッダー =====
render_header(
    title="管理画面",
    subtitle=f"ログイン中: {current['display_name']}（管理者）",
    icon="⚙️",
)

with st.sidebar:
    render_user_badge(current['display_name'], role="管理者")
    if st.button("ログアウト", use_container_width=True):
        st.session_state.pop("auth_user", None)
        st.rerun()
    st.page_link("app.py", label="メイン画面に戻る", icon="🏠")


tab_usage, tab_api, tab_users, tab_settings = st.tabs(
    ["📊 利用状況", "🔑 API設定", "👥 ユーザー管理", "⚙️ 詳細設定"]
)


# =========================================================
# 📊 利用状況
# =========================================================
with tab_usage:
    st.subheader("集計期間")
    today = datetime.now().date()
    col1, col2, col3 = st.columns([2, 2, 3])
    with col1:
        since = st.date_input("開始日", value=today - timedelta(days=30))
    with col2:
        until = st.date_input("終了日", value=today)
    with col3:
        st.write("")
        st.write("")
        if st.button("🔄 更新", use_container_width=True):
            st.rerun()

    since_str = f"{since} 00:00:00"
    until_str = f"{until} 23:59:59"

    # サマリ
    summary = total_summary(since_str, until_str)
    st.subheader("📈 期間内サマリ")
    render_kpi_strip([
        ("API呼び出し", f"{summary['calls']:,} 回",
         "AI Vision 経由の処理件数", "📞", "primary"),
        ("合計コスト", f"¥{summary['cost_jpy']:,.1f}",
         f"≒ ${summary['cost_usd']:,.4f}", "💴", "amber"),
        ("入力トークン", f"{summary['input_tokens']:,}",
         "AIへの送信量", "📥", "blue"),
        ("出力トークン", f"{summary['output_tokens']:,}",
         "AIからの応答量", "📤", "teal"),
    ])

    st.divider()

    # ユーザー別
    st.subheader("👤 ユーザー別")
    users_data = per_user_summary(since_str, until_str)
    if users_data:
        df_users = pd.DataFrame(users_data)
        df_users["cost_jpy"] = df_users["cost_jpy"].round(2)
        df_users["cost_usd"] = df_users["cost_usd"].round(6)
        df_users.columns = ["ユーザー", "回数", "入力tok", "出力tok", "コスト(USD)", "コスト(円)"]
        st.dataframe(df_users, use_container_width=True, hide_index=True)
    else:
        st.info("この期間の利用ログはありません。")

    st.divider()

    # 日別推移
    st.subheader("📅 日別")
    days_data = per_day_summary(since_str, until_str)
    if days_data:
        df_days = pd.DataFrame(days_data)
        df_days["cost_jpy"] = df_days["cost_jpy"].round(2)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**呼び出し回数**")
            st.bar_chart(df_days.set_index("day")["calls"])
        with c2:
            st.markdown("**コスト（円）**")
            st.bar_chart(df_days.set_index("day")["cost_jpy"])
    else:
        st.info("グラフ表示するデータがありません。")

    st.divider()

    # 直近ログ
    st.subheader("📝 直近100件")
    logs = recent_logs(100)
    if logs:
        df_logs = pd.DataFrame(logs)
        df_logs["cost_jpy"] = df_logs["cost_jpy"].round(3)
        df_logs["success"] = df_logs["success"].map({1: "✅", 0: "❌"})
        df_logs.columns = ["日時", "ユーザー", "Provider", "Model",
                           "入力tok", "出力tok", "コスト(円)", "ファイル", "成功"]
        st.dataframe(df_logs, use_container_width=True, hide_index=True)
    else:
        st.info("ログはまだありません。")


# =========================================================
# 🔑 API設定
# =========================================================
with tab_api:
    st.subheader("APIキー設定")
    st.caption(
        "ここで設定したキーはサーバー上のSQLiteに保存され、全ユーザーが共通で利用します。"
        "Streamlit Secrets または環境変数が設定されている場合はそちらが優先されます。"
    )

    # Claude
    st.markdown("### 🔵 Claude (Anthropic)")
    claude_existing = get_setting("api_key_claude", "")
    if claude_existing:
        st.info(f"登録済み: `{claude_existing[:8]}...{claude_existing[-4:]}`")
    new_claude = st.text_input(
        "Claude APIキー",
        type="password",
        placeholder="sk-ant-...",
        key="claude_key_input",
        help="https://console.anthropic.com でキー発行（$5チャージで数百〜千枚相当）",
    )
    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("Claudeキーを保存", type="primary", use_container_width=True):
            if new_claude:
                set_setting("api_key_claude", new_claude)
                st.success("保存しました")
                st.rerun()
            else:
                st.error("キーを入力してください")
    with c2:
        if st.button("Claudeキーを削除", use_container_width=True):
            delete_setting("api_key_claude")
            st.success("削除しました")
            st.rerun()

    st.divider()

    # Gemini
    st.markdown("### 🟢 Gemini (Google)")
    gemini_existing = get_setting("api_key_gemini", "")
    if gemini_existing:
        st.info(f"登録済み: `{gemini_existing[:8]}...{gemini_existing[-4:]}`")
    new_gemini = st.text_input(
        "Gemini APIキー",
        type="password",
        placeholder="AIzaSy...",
        key="gemini_key_input",
        help="https://aistudio.google.com/apikey で無料発行",
    )
    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("Geminiキーを保存", type="primary", use_container_width=True):
            if new_gemini:
                set_setting("api_key_gemini", new_gemini)
                st.success("保存しました")
                st.rerun()
            else:
                st.error("キーを入力してください")
    with c2:
        if st.button("Geminiキーを削除", use_container_width=True):
            delete_setting("api_key_gemini")
            st.success("削除しました")
            st.rerun()


# =========================================================
# 👥 ユーザー管理
# =========================================================
with tab_users:
    st.subheader("ユーザー一覧")
    users = list_users()
    df = pd.DataFrame(users)
    if not df.empty:
        df_disp = df.copy()
        df_disp["is_admin"]  = df_disp["is_admin"].map({1: "✅", 0: ""})
        df_disp["is_active"] = df_disp["is_active"].map({1: "✅", 0: "停止"})
        df_disp = df_disp[["id", "username", "display_name", "is_admin", "is_active", "created_at"]]
        df_disp.columns = ["ID", "ユーザー名", "表示名", "管理者", "有効", "作成日"]
        st.dataframe(df_disp, use_container_width=True, hide_index=True)

    st.divider()

    # 追加
    st.subheader("➕ 新規ユーザー追加")
    with st.form("add_user_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            new_username = st.text_input("ユーザー名（半角英数）")
            new_password = st.text_input("初期パスワード（6文字以上）", type="password")
        with c2:
            new_display  = st.text_input("表示名")
            new_is_admin = st.checkbox("管理者権限を付与")
        ok = st.form_submit_button("ユーザーを追加", type="primary")
    if ok:
        try:
            create_user(new_username, new_password, new_display, is_admin=new_is_admin)
            st.success(f"ユーザー「{new_username}」を作成しました")
            st.rerun()
        except ValueError as e:
            st.error(str(e))

    st.divider()

    # 編集（既存ユーザー）
    if users:
        st.subheader("✏️ 既存ユーザーの編集")
        options = {f"{u['display_name']}（{u['username']}）": u for u in users}
        sel_label = st.selectbox("対象ユーザー", list(options.keys()))
        target = options[sel_label]

        # --- プロフィール変更（ユーザー名・表示名）---
        with st.form(f"profile_form_{target['id']}"):
            st.markdown("**プロフィール変更**")
            pc1, pc2 = st.columns(2)
            with pc1:
                edit_username = st.text_input(
                    "ユーザー名（半角英数）",
                    value=target["username"],
                    key=f"username_{target['id']}",
                )
            with pc2:
                edit_display = st.text_input(
                    "表示名",
                    value=target["display_name"],
                    key=f"display_{target['id']}",
                )
            saved = st.form_submit_button("プロフィールを保存", type="primary")
        if saved:
            try:
                update_user_profile(
                    target["id"],
                    username=edit_username,
                    display_name=edit_display,
                )
                # 自分自身を編集した場合はセッションも更新
                if target["id"] == current["id"]:
                    st.session_state["auth_user"]["username"] = edit_username.strip()
                    st.session_state["auth_user"]["display_name"] = edit_display.strip()
                st.success("プロフィールを更新しました")
                st.rerun()
            except ValueError as e:
                st.error(str(e))

        st.markdown("")

        c1, c2 = st.columns(2)

        # --- パスワード変更 ---
        with c1:
            with st.form(f"pw_form_{target['id']}", clear_on_submit=True):
                st.markdown("**パスワード変更**")
                pw_new = st.text_input(
                    "新パスワード（6文字以上）", type="password",
                    key=f"pw_{target['id']}",
                )
                changed = st.form_submit_button("パスワードを変更", type="primary")
            if changed:
                try:
                    update_user_password(target["id"], pw_new)
                    st.success("パスワードを変更しました")
                except ValueError as e:
                    st.error(str(e))

        # --- 権限・状態 ---
        with c2:
            st.markdown("**権限・状態**")
            if target["id"] == current["id"]:
                st.caption("自分自身の権限変更／停止／削除はできません")
            else:
                _r1, _r2 = st.columns(2)
                with _r1:
                    if target["is_admin"]:
                        if st.button("管理者権限を外す", key=f"unadmin_{target['id']}",
                                     use_container_width=True):
                            set_user_admin(target["id"], False)
                            st.rerun()
                    else:
                        if st.button("管理者にする", key=f"admin_{target['id']}",
                                     use_container_width=True):
                            set_user_admin(target["id"], True)
                            st.rerun()
                with _r2:
                    if target["is_active"]:
                        if st.button("アカウント停止", key=f"deact_{target['id']}",
                                     use_container_width=True):
                            set_user_active(target["id"], False)
                            st.rerun()
                    else:
                        if st.button("アカウント有効化", key=f"act_{target['id']}",
                                     use_container_width=True):
                            set_user_active(target["id"], True)
                            st.rerun()

                if st.button("🗑 ユーザーを削除（取消不可）",
                             key=f"del_{target['id']}", type="secondary",
                             use_container_width=True):
                    delete_user(target["id"])
                    st.success("削除しました")
                    st.rerun()


# =========================================================
# ⚙️ 詳細設定
# =========================================================
with tab_settings:
    st.subheader("AI読み取りの制御")
    st.caption(
        "ユーザー側からはAI関連の操作はできません。ここで一括制御します。"
    )

    _ai_enabled_now = (get_setting("ai_enabled", "1") or "1").strip().lower()
    _is_enabled = _ai_enabled_now not in ("0", "false", "off", "no")

    ai_on = st.toggle(
        "AI読み取りを有効にする",
        value=_is_enabled,
        help="OFFにすると、APIキーが設定されていてもAIを使わず、ルールベース読み取りのみになります。",
    )
    if ai_on != _is_enabled:
        set_setting("ai_enabled", "1" if ai_on else "0")
        st.success("AI読み取り設定を更新しました")
        st.rerun()

    cur_provider = get_setting("default_provider", "claude")
    sel = st.radio(
        "使用するAIプロバイダ",
        ["claude", "gemini"],
        index=0 if cur_provider == "claude" else 1,
        horizontal=True,
        format_func=lambda x: "Claude（推奨・高速・高精度）" if x == "claude" else "Gemini（無料枠あり）",
        disabled=not ai_on,
    )
    if sel != cur_provider:
        set_setting("default_provider", sel)
        st.success("プロバイダを更新しました")
        st.rerun()

    if ai_on:
        st.info(
            f"現在のAI読み取り: **有効** / プロバイダ: **{sel}**\n\n"
            "APIキーは「🔑 API設定」タブで登録してください。"
        )
    else:
        st.warning("現在のAI読み取り: **無効**（全ユーザーがルールベース読み取りになります）")

    st.divider()

    st.subheader("為替レート（コスト計算用）")
    cur_rate = float(get_setting("usd_jpy", str(DEFAULT_USD_JPY)) or DEFAULT_USD_JPY)
    new_rate = st.number_input("1 USD = ", min_value=50.0, max_value=300.0,
                               value=cur_rate, step=0.5)
    if st.button("レートを保存"):
        set_setting("usd_jpy", str(new_rate))
        st.success("保存しました")
        st.rerun()
