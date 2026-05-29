"""
app.py - 地域おこし協力隊 経費管理アプリ v3
起動: streamlit run app.py
"""
import os
import io
import base64
import tempfile
import streamlit as st
import pandas as pd
from datetime import datetime
from PIL import Image as _PILImage

# ファビコン: assets/favicon.png（透過PNG）。無ければ絵文字にフォールバック
_FAVICON_PATH = os.path.join(os.path.dirname(__file__), "assets", "favicon.png")
_PAGE_ICON = _PILImage.open(_FAVICON_PATH) if os.path.exists(_FAVICON_PATH) else "📋"

from core.extract import (
    extract_from_file,
    pdf_to_image_bytes,
    image_to_jpeg_bytes,
    KAMOKU_OPTIONS,
    JIGYO_OPTIONS,
)
from core.excel_writer import write_receipts_to_excel
from core.auth import (
    authenticate, has_any_user, create_user,
)
from core.db import get_setting
from core.usage import log_usage, user_summary
from core.history import list_history, get_history_detail
from core.theme import (
    apply_theme, render_header, render_kpi_strip,
    render_user_badge, render_clickable_image,
    render_ai_reading, update_ai_progress,
)

# =========================================================
# ページ設定
# =========================================================
st.set_page_config(
    page_title="地域おこし 経費管理",
    page_icon=_PAGE_ICON,
    layout="wide",
    initial_sidebar_state="expanded",
)

apply_theme()


# =========================================================
# 認証ゲート
# =========================================================
def _get_api_key_from_env_or_secrets(provider: str) -> tuple:
    """(api_key, source) を返す。優先順位: Secrets → 環境変数 → DB"""
    name = "ANTHROPIC_API_KEY" if provider == "claude" else "GEMINI_API_KEY"
    try:
        v = st.secrets.get(name, "")
        if v:
            return v, "Secrets"
    except Exception:
        pass
    v = os.environ.get(name, "")
    if v:
        return v, "環境変数"
    v = get_setting(f"api_key_{provider}", "")
    if v:
        return v, "管理画面で設定"
    return "", ""


def _hide_sidebar():
    """ログイン/初回セットアップ画面ではサイドバーを完全に非表示"""
    st.markdown(
        '<style>'
        '[data-testid="stSidebar"] {display: none !important;}'
        '[data-testid="stSidebarCollapsedControl"] {display: none !important;}'
        '[data-testid="collapsedControl"] {display: none !important;}'
        'section[data-testid="stMain"] {margin-left: 0 !important;}'
        '</style>',
        unsafe_allow_html=True,
    )


def _login_form():
    _hide_sidebar()
    render_header(
        title="地域おこし協力隊　経費管理システム",
        subtitle="領収書から月次出納簿を自動作成",
        icon="📋",
    )
    _l, _c, _r = st.columns([1, 2, 1])
    with _c:
        with st.container(border=True):
            st.markdown("#### 🔐 ログイン")
            with st.form("login_form"):
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


def _initial_setup_form():
    _hide_sidebar()
    render_header(
        title="地域おこし協力隊　経費管理システム",
        subtitle="初回セットアップ",
        icon="🛠",
    )
    _l, _c, _r = st.columns([1, 2, 1])
    with _c:
        with st.container(border=True):
            st.markdown("#### 管理者アカウントの作成")
            st.caption("最初に登録するユーザーが、システム全体の管理者になります。")
            with st.form("setup_form"):
                username = st.text_input("ユーザー名（半角英数）", value="admin")
                display_name = st.text_input("表示名", value="管理者")
                password = st.text_input("パスワード（6文字以上）", type="password")
                password2 = st.text_input("パスワード（確認）", type="password")
                ok = st.form_submit_button("管理者を作成", type="primary",
                                           use_container_width=True)
        if ok:
            if password != password2:
                st.error("パスワードが一致しません")
                return
            try:
                uid = create_user(username, password, display_name, is_admin=True)
                st.success("管理者を作成しました。続けてログインしてください。")
                st.session_state["auth_user"] = {
                    "id": uid, "username": username,
                    "display_name": display_name, "is_admin": True,
                }
                st.rerun()
            except ValueError as e:
                st.error(str(e))


def require_login() -> dict:
    """未ログインなら入力フォームを表示してstop。ログイン済みならuser dictを返す。"""
    if not has_any_user():
        _initial_setup_form()
        st.stop()
    user = st.session_state.get("auth_user")
    if not user:
        _login_form()
        st.stop()
    return user


# =========================================================
# ヘルパー
# =========================================================
def get_display_image(filepath, ext):
    """ファイルをブラウザ表示用の画像バイトに変換"""
    if ext == '.pdf':
        return pdf_to_image_bytes(filepath, zoom=2.0)
    else:
        return image_to_jpeg_bytes(filepath)


def render_history_section(user_id):
    """過去の処理履歴ビューア（KPIエリア直下・折りたたみ・コンパクト）"""
    import pandas as _pd_h
    with st.expander("📜 過去の処理履歴を見る", expanded=False):
        _hist = list_history(user_id)
        if not _hist:
            st.caption("まだ処理履歴がありません。領収書を出納簿に書き込むと、ここに記録されます。")
            return
        st.caption(
            f"過去 {len(_hist)} 回の処理（新しい順）。「詳細」で内容を表示。"
        )
        # 各行をコンパクトに（余白の大きい st.divider は使わない）
        for _idx, _h in enumerate(_hist):
            _inc = f"　/　収入 ¥{_h['income_total']:,}" if _h['income_total'] else ""
            _hc1, _hc2 = st.columns([5, 1], vertical_alignment="center")
            with _hc1:
                st.markdown(
                    f"<div style='font-size:0.86rem;line-height:1.3;'>"
                    f"🗓 <b>{_h['processed_at']}</b>"
                    f"<span style='color:#677291;'>"
                    f"　{_h['record_count']}件　支出 ¥{_h['expense_total']:,}{_inc}"
                    f"</span></div>",
                    unsafe_allow_html=True,
                )
            with _hc2:
                _opened = st.session_state.get("view_history_id") == _h['id']
                if st.button("閉じる" if _opened else "詳細",
                             key=f"hist_btn_{_h['id']}",
                             use_container_width=True):
                    st.session_state["view_history_id"] = (
                        None if _opened else _h['id']
                    )
                    st.rerun()

            if st.session_state.get("view_history_id") == _h['id']:
                _detail = get_history_detail(_h['id'], user_id)
                if _detail and _detail.get("records"):
                    _drows = []
                    for _rec in _detail["records"]:
                        _kd = _rec.get("kind", "expense")
                        _drows.append({
                            "種別":   "💰収入" if _kd == "income" else "💴支出",
                            "日付":   _rec.get("date", ""),
                            "取引先": _rec.get("vendor", ""),
                            "摘要":   _rec.get("memo", ""),
                            "金額":   int(_rec.get("amount", 0) or 0),
                            "勘定科目": _rec.get("kamoku", ""),
                            "事業名": _rec.get("jigyo", ""),
                        })
                    _ddf = _pd_h.DataFrame(_drows)
                    st.dataframe(
                        _ddf, use_container_width=True, hide_index=True,
                        column_config={
                            "金額": st.column_config.NumberColumn(format="¥%d"),
                        },
                    )
                else:
                    st.info("詳細データが見つかりませんでした。")

            # 行間の細い区切り（最後の行以外）
            if _idx < len(_hist) - 1:
                st.markdown(
                    "<hr style='margin:0.35rem 0;border:none;"
                    "border-top:1px solid #e3e7ef;'>",
                    unsafe_allow_html=True,
                )


def load_template(user_name: str = "", nendo: int = None):
    """
    テンプレート出納簿を読み込み、B1の「氏名：xxx」と年度を差し替えて返す。

    user_name: ログイン中ユーザーの表示名（B1の氏名欄に入る）
    nendo:     年度（B1の「令和X年度」を更新）
    """
    tpl = os.path.join(os.path.dirname(__file__), "templates", "出納簿テンプレート.xlsx")
    if not user_name and nendo is None:
        with open(tpl, "rb") as f:
            return f.read()

    import openpyxl
    from io import BytesIO
    wb = openpyxl.load_workbook(tpl)
    if "出納簿" in wb.sheetnames:
        ws = wb["出納簿"]
        reiwa = (nendo - 2018) if nendo else None
        name  = (user_name or "").strip()
        # 既存タイトルが取れたら令和年だけ差し替え、ダメなら定型で組み立て
        cur = ws.cell(row=1, column=2).value or ""
        import re as _re
        new_title = None
        if cur:
            # 「令和X年度」を差し替え
            if reiwa is not None:
                cur = _re.sub(r"令和[\d０-９]+年度", f"令和{reiwa}年度", cur)
            # 「氏名：YYY」を差し替え（全角・半角コロン両対応）
            if name:
                if _re.search(r"氏名[：:]", cur):
                    cur = _re.sub(r"(氏名[：:])[^）)]*", rf"\1{name}　", cur)
                else:
                    # 末尾に追加
                    cur = cur.rstrip("）)" + "　 ")
                    cur = f"{cur}（氏名：{name}　）"
            new_title = cur
        else:
            _r = reiwa if reiwa is not None else 7
            new_title = f"令和{_r}年度　根室市地域おこし協力隊現金出納簿（氏名：{name}　）"
        ws.cell(row=1, column=2).value = new_title

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =========================================================
# セッション初期化
# =========================================================
def init_session():
    _now = datetime.now()
    _default_nendo = _now.year if _now.month >= 4 else _now.year - 1
    defaults = {
        "phase": "upload",
        "records": [],
        "images": [],
        "excel_images": [],
        "filenames": [],
        "denpyo_bytes": None,
        "result_bytes": None,
        "write_results": [],
        "receipt_sheet_option": "new",   # "new" or 既存シート名
        "receipt_new_sheet_name": "",
        "user_jigyo": JIGYO_OPTIONS[0],
        # 収入エントリ（領収書なしの収入レコードのバッファ）
        "pending_income": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # 年度・月区分は「セッション初回のみ」当日基準で強制セット。
    # （古いセッションに前回値が残っていても、初回マーカーで当日に更新）
    if "_period_init" not in st.session_state:
        st.session_state["_period_init"] = True
        st.session_state["user_nendo"] = _default_nendo
        st.session_state["user_month"] = _now.month


def get_receipt_sheets(excel_bytes: bytes) -> list:
    """Excelファイルから「領収書」を含むシート名一覧を返す"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), read_only=True)
        return [n for n in wb.sheetnames if "領収書" in n]
    except Exception:
        return []


def _capture_font_info(cell):
    """openpyxlセルから書式情報をdictで取得（書き戻しで復元用）"""
    try:
        f = cell.font
        if not f:
            return None
        info = {
            "bold":      bool(f.bold),
            "italic":    bool(f.italic),
            "underline": f.underline if f.underline else None,
        }
        # 何も無効化情報がなければNoneを返す（保存量削減）
        if not info["bold"] and not info["italic"] and not info["underline"]:
            return None
        return info
    except Exception:
        return None


def read_existing_rows(excel_bytes: bytes) -> list:
    """
    出納簿シートから既存データを読み取り、write_single_row互換形式で返す。
    重い処理なので呼び出し元でsession_stateにキャッシュすること。

    各レコードに次のメタも付与:
        _font_info: {列番号: {bold, italic, underline}}  ← 書き戻しでスタイル復元
        _kind:      "income" | "expense"
    """
    rows = []
    try:
        import openpyxl
        from core.excel_writer import detect_data_range, DATA_START_ROW
        # 書式情報を取りたいので read_only=False
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=False)
        if "出納簿" not in wb.sheetnames:
            return rows
        ws = wb["出納簿"]
        # 合計行を除いた末尾までだけ読む
        _data_end, _totals_row = detect_data_range(ws)
        scan_end = (_totals_row - 1) if _totals_row else (_data_end + 50)
        for row_num in range(DATA_START_ROW, scan_end + 1):
            a = ws.cell(row=row_num, column=1).value   # No.
            p = ws.cell(row=row_num, column=16).value  # 取引先
            q = ws.cell(row=row_num, column=17).value  # 収入
            r = ws.cell(row=row_num, column=18).value  # 支出
            # 合計行や数式行はスキップ（P列が "=..." 等）
            if isinstance(p, str) and p.startswith("="):
                continue
            # 取引先・収入・支出のいずれかが入っていればデータ行として扱う
            # （A列のNo.が空でもQ/Rに金額があれば収入/支出として取り込む）
            has_vendor = p is not None and str(p).strip() != ""
            has_income = isinstance(q, (int, float)) and q
            has_expense = isinstance(r, (int, float)) and r
            if not (has_vendor or has_income or has_expense):
                continue
            c  = ws.cell(row=row_num, column=3).value   # 令和年
            e  = ws.cell(row=row_num, column=5).value   # 月
            g  = ws.cell(row=row_num, column=7).value   # 日
            k_cell  = ws.cell(row=row_num, column=11)   # 摘要セル
            k  = k_cell.value
            j  = ws.cell(row=row_num, column=10).value  # 勘定科目
            ii = ws.cell(row=row_num, column=9).value   # 事業名
            try:
                date_str = f"{int(c)+2018}-{int(e):02d}-{int(g):02d}" if c and e and g else ""
            except Exception:
                date_str = ""

            # 収入/支出を区別
            income  = int(q or 0) if isinstance(q, (int, float)) else 0
            expense = int(r or 0) if isinstance(r, (int, float)) else 0
            if income > 0 and expense == 0:
                kind   = "income"
                amount = income
            else:
                kind   = "expense"
                amount = expense or income

            # 主要列のフォント情報を取得
            font_info = {}
            for col in (9, 10, 11, 16):  # I:事業, J:科目, K:摘要, P:取引先
                info = _capture_font_info(ws.cell(row=row_num, column=col))
                if info:
                    font_info[col] = info

            rec = {
                "_type":  "existing",
                "_kind":  kind,
                "date":   date_str,
                "vendor": str(p or ""),
                "memo":   str(k or ""),
                "amount": int(amount),
                "kamoku": str(j or "消耗品"),
                "jigyo":  str(ii or "ミッション活動"),
            }
            if font_info:
                rec["_font_info"] = font_info
            rows.append(rec)
    except Exception:
        pass
    return rows


init_session()

# 認証ゲート（ログイン未済はここで停止）
current_user = require_login()

# ===== AIプロバイダ・APIキーの解決（管理画面で制御） =====
# プロバイダ: 管理画面の「デフォルトAIプロバイダ」設定を使う
ai_provider = (get_setting("default_provider", "claude") or "claude").lower()
if ai_provider not in ("claude", "gemini"):
    ai_provider = "claude"

# APIキー: Secrets → 環境変数 → DB の優先順位（既存ロジック）
ai_api_key, _ai_key_source = _get_api_key_from_env_or_secrets(ai_provider)

# 管理画面の「AI読み取りを有効にする」設定（既定: 有効）
# "0"/"false" のみOFF扱い
_ai_enabled_setting = (get_setting("ai_enabled", "1") or "1").strip().lower()
if _ai_enabled_setting in ("0", "false", "off", "no"):
    ai_api_key = ""  # 強制無効化（APIキーがあっても使わない）


# =========================================================
# サイドバー
# =========================================================
with st.sidebar:
    # ユーザーバッジ
    render_user_badge(
        current_user["display_name"],
        role="管理者" if current_user.get("is_admin") else "一般ユーザー",
    )

    # 管理者だけに管理画面リンクを表示
    if current_user.get("is_admin"):
        st.page_link("pages/admin.py", label="管理画面を開く", icon="⚙️")

    # 現在の処理状況サマリ（簡易ナビ）
    st.subheader("処理状況")
    _phase_now = st.session_state.get("phase", "upload")
    _phase_labels = {
        "upload":  ("①", "アップロード"),
        "review":  ("②", "確認・編集"),
        "order":   ("③", "順番調整"),
        "writing": ("④", "書き込み中"),
        "done":    ("✅", "完了"),
    }
    for _key, (_num, _lbl) in _phase_labels.items():
        _is_active = (_key == _phase_now)
        if _is_active:
            _style = (
                "background: linear-gradient(135deg,"
                " rgba(202,81,66,0.42) 0%,"
                " rgba(214,158,46,0.22) 100%);"
                "color: #fff; font-weight: 700;"
                "border: 1px solid rgba(255,255,255,0.18);"
                "box-shadow: 0 4px 12px rgba(202,81,66,0.28),"
                " inset 0 1px 0 rgba(255,255,255,0.20);"
            )
        else:
            _style = (
                "background: transparent;"
                "color: #97a0b8; font-weight: 500;"
                "border: 1px solid transparent;"
            )
        st.markdown(
            f'<div style="{_style}'
            f'padding:7px 11px;border-radius:9px;font-size:0.88rem;'
            f'margin-bottom:3px;transition:all 0.2s ease;">{_num}　{_lbl}</div>',
            unsafe_allow_html=True,
        )

    st.subheader("操作")
    if st.button("🔄 最初からやり直す", use_container_width=True):
        keep = {"auth_user":  st.session_state.get("auth_user"),
                "user_nendo": st.session_state.get("user_nendo"),
                "user_month": st.session_state.get("user_month"),
                "user_jigyo": st.session_state.get("user_jigyo")}
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        for k, v in keep.items():
            if v is not None:
                st.session_state[k] = v
        st.rerun()

    if st.button("ログアウト", use_container_width=True):
        st.session_state.pop("auth_user", None)
        st.rerun()

    # ステータス（フッター）
    st.divider()
    if ai_api_key:
        st.caption(f"🤖 AI読み取り: 有効（{ai_provider}）")
    else:
        st.caption("🤖 AI読み取り: 未設定")
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        st.caption("🔍 OCR: tesseract利用可能")
    except Exception:
        st.caption("🔍 OCR: Apple Vision使用中")

# サイドバー外: 各フェーズで参照する設定値を session_state から取得
nendo = st.session_state["user_nendo"]
month = st.session_state["user_month"]
default_jigyo = st.session_state["user_jigyo"]
tsuki_kubun = f"{nendo}-{month:02d}" if month >= 4 else f"{nendo + 1}-{month:02d}"


# =========================================================
# タイトル
# =========================================================
render_header(
    title="地域おこし協力隊　経費管理システム",
    subtitle="領収書の読み取り → 確認 → 出納簿への自動転記",
    icon="📋",
)

# ===== KPIストリップ（ユーザー実績・処理履歴ベース） =====
_today_str = datetime.now().strftime("%Y-%m-%d")
_month_str = datetime.now().strftime("%Y-%m")

try:
    _hist_all = list_history(current_user["id"])
except Exception:
    _hist_all = []

_today_recs = sum(h["record_count"] for h in _hist_all
                  if str(h.get("processed_at", "")).startswith(_today_str))
_month_recs = sum(h["record_count"] for h in _hist_all
                  if str(h.get("processed_at", "")).startswith(_month_str))
_total_recs = sum(h["record_count"] for h in _hist_all)
_hist_count = len(_hist_all)

render_kpi_strip([
    ("今日の処理", f"{_today_recs} 件",
     "本日処理した領収書", "📅", "primary"),
    ("今月の処理", f"{_month_recs} 件",
     f"累計 {_total_recs} 件", "🗓", "teal"),
    ("累計実績", f"{_total_recs} 件",
     "あなたの全期間の処理件数", "📊", "blue"),
    ("処理履歴", f"{_hist_count} 回",
     "下の履歴から内容を確認", "📜", "violet"),
])

phase = st.session_state.get("phase", "upload")


# =========================================================
# フェーズ1: アップロード
# =========================================================
if phase == "upload":
    # --- 過去の処理履歴（KPIストリップ直下） ---
    render_history_section(current_user["id"])

    # --- 設定バー（期間・事業を最上部に） ---
    with st.container(border=True):
        st.markdown("##### ⚙️ 処理設定")
        sc_y, sc_m, sc_j = st.columns([1.2, 1.2, 1.6])
        now = datetime.now()
        with sc_y:
            # init_session が session_state にデフォルト値をセット済みなので
            # ここでは index= を渡さない（widget warning 回避）
            st.selectbox(
                "年度",
                list(range(now.year - 2, now.year + 3)),
                format_func=lambda y: f"{y}年度（令和{y - 2018}年度）",
                key="user_nendo",
            )
        with sc_m:
            st.selectbox(
                "月区分",
                list(range(4, 13)) + list(range(1, 4)),
                format_func=lambda m: f"{m}月",
                key="user_month",
            )
        with sc_j:
            _jigyo_opts = JIGYO_OPTIONS + ["✏️ その他（手入力）"]
            _sel_jigyo = st.selectbox(
                "事業名（領収書のデフォルト）",
                _jigyo_opts,
                key="user_jigyo_sel",
            )
            # 「その他」選択時は、同じ位置に空欄の入力欄を出してそのまま打ち込み
            if _sel_jigyo == "✏️ その他（手入力）":
                _custom_jigyo = st.text_input(
                    "事業名（手入力）",
                    key="user_jigyo_custom",
                    placeholder="事業名を入力（空欄のままでもOK）",
                    label_visibility="collapsed",
                )
                # 入力がなければ空欄のまま
                st.session_state["user_jigyo"] = _custom_jigyo.strip()
            else:
                st.session_state["user_jigyo"] = _sel_jigyo

        # 表示用に再取得（widgetが書き戻したsession_stateを反映）
        nendo = st.session_state["user_nendo"]
        month = st.session_state["user_month"]
        default_jigyo = st.session_state["user_jigyo"]
        tsuki_kubun = (f"{nendo}-{month:02d}" if month >= 4
                       else f"{nendo + 1}-{month:02d}")
        _jigyo_disp = default_jigyo if default_jigyo else "（未設定）"
        st.caption(f"📌 月区分タグ: `{tsuki_kubun}`　／　事業: `{_jigyo_disp}`")

    st.markdown("### ファイルのアップロード")
    col_l, col_r = st.columns(2, gap="large")

    with col_l:
        st.markdown("##### ① 出納簿（Excel）")
        denpyo_file = st.file_uploader(
            "出納簿アップロード",
            type=["xlsx"],
            key="denpyo_up",
            label_visibility="collapsed",
            help="既存の出納簿 .xlsx ファイルを指定。なければ下のチェックで新規作成。",
        )
        use_template = False
        if not denpyo_file:
            use_template = st.checkbox(f"テンプレートから新規作成（{nendo}年度）")
        if denpyo_file:
            try:
                import openpyxl
                from io import BytesIO as _BytesIO
                from core.excel_writer import count_filled_rows
                _wb = openpyxl.load_workbook(_BytesIO(denpyo_file.read()))
                _cnt = count_filled_rows(_wb['出納簿'])
                denpyo_file.seek(0)
                st.success(f"✅ {denpyo_file.name}　（現在 {_cnt} 件入力済み）")
            except Exception:
                denpyo_file.seek(0)
                st.success(f"✅ {denpyo_file.name}")
        elif use_template:
            st.info("テンプレートを使用します（0件からスタート）")

    with col_r:
        st.markdown("##### ② 領収書（PDF・JPG・PNG・HEIC / 複数可）")
        receipt_files = st.file_uploader(
            "領収書アップロード",
            type=["pdf", "jpg", "jpeg", "png", "heic", "heif"],
            accept_multiple_files=True,
            key="receipt_up",
            label_visibility="collapsed",
            help="複数ファイルをまとめて選択／ドラッグできます。iPhoneのHEIC写真もそのままOK",
        )
        if receipt_files:
            st.success(f"✅ {len(receipt_files)}件 選択済み")

    # --- 収入を追加（領収書なし） ---
    with st.expander(
        f"💰 収入を追加（領収書なし）　— "
        f"{len(st.session_state.get('pending_income', []))} 件追加済み",
        expanded=False,
    ):
        st.caption("補助金・助成金・返金など、領収書がない収入を追加できます（画像なし）。")

        # 既存の収入エントリを表示・削除
        _incs = st.session_state.get("pending_income", [])
        if _incs:
            for _ii, _ie in enumerate(_incs):
                _r1, _r2 = st.columns([6, 1])
                with _r1:
                    st.markdown(
                        f"📅 **{_ie.get('date','')}**　"
                        f"🏷 {_ie.get('vendor','')}　"
                        f"📝 {_ie.get('memo','')}　"
                        f"**¥{int(_ie.get('amount',0)):,}**　"
                        f"`{_ie.get('kamoku','')}`"
                    )
                with _r2:
                    if st.button("削除", key=f"del_pending_inc_{_ii}",
                                 use_container_width=True):
                        st.session_state["pending_income"].pop(_ii)
                        st.rerun()
            st.divider()

        # 新規追加フォーム
        with st.form("add_income_form", clear_on_submit=True):
            ic1, ic2 = st.columns(2)
            with ic1:
                _inc_date = st.text_input(
                    "📅 日付（YYYY-MM-DD）",
                    value=f"{nendo}-{month:02d}-01" if month >= 4
                          else f"{nendo+1}-{month:02d}-01",
                    placeholder="例: 2026-04-15",
                )
                _inc_vendor = st.text_input("💼 入金元 / 取引先",
                                            placeholder="例: 根室市役所")
            with ic2:
                _inc_amount = st.number_input("💴 金額（円）", min_value=0, step=1000)
                _inc_kamoku = st.selectbox("📂 勘定科目", KAMOKU_OPTIONS)
            _inc_memo = st.text_input("📝 摘要（内容）",
                                      placeholder="例: 4月分活動費補助金")
            _inc_jigyo = st.selectbox("🎯 事業名", JIGYO_OPTIONS,
                                      index=JIGYO_OPTIONS.index(default_jigyo)
                                      if default_jigyo in JIGYO_OPTIONS else 0)
            _add_inc = st.form_submit_button("➕ 収入を追加",
                                             type="primary",
                                             use_container_width=True)
        if _add_inc:
            if not _inc_date or not _inc_vendor or _inc_amount <= 0:
                st.error("日付・入金元・金額は必須です")
            else:
                st.session_state.setdefault("pending_income", []).append({
                    "date":   _inc_date,
                    "vendor": _inc_vendor,
                    "memo":   _inc_memo or _inc_vendor,
                    "amount": int(_inc_amount),
                    "kamoku": _inc_kamoku,
                    "jigyo":  _inc_jigyo,
                })
                st.success("収入を追加しました")
                st.rerun()

    st.divider()

    _has_receipts = bool(receipt_files)
    _has_income   = bool(st.session_state.get("pending_income"))
    _has_denpyo   = bool(denpyo_file or use_template)
    can_start = (_has_receipts or _has_income) and _has_denpyo

    if not _has_denpyo:
        st.info("📂 出納簿ファイル（またはテンプレート）を選択してください")
    elif not (_has_receipts or _has_income):
        st.info("📄 領収書ファイルをアップロードするか、💰 収入を追加してください")

    if can_start and st.button("🚀 読み取り開始", type="primary", use_container_width=True):
        denpyo_bytes = (
            load_template(
                user_name=current_user.get("display_name", ""),
                nendo=nendo,
            )
            if use_template else denpyo_file.read()
        )
        records, images, excel_images, filenames = [], [], [], []
        receipt_files = receipt_files or []

        # マスコットアニメ用スロット（ループ中は触らない → CSSアニメが連続再生）
        anim_slot = st.empty()
        # 進捗専用スロット（こちらだけループ中に更新）
        progress_slot = st.empty()
        if receipt_files:
            _ai_msg = ("AIが領収書を1枚ずつチェック中…" if ai_api_key
                       else "領収書を1枚ずつ確認中…")
            render_ai_reading(anim_slot, current=1, total=len(receipt_files),
                              filename="", done=0, ai_label=_ai_msg)
            update_ai_progress(progress_slot, current=0, total=len(receipt_files),
                               filename="", done=0)
        prog = None  # 旧progress barは使わない

        for i, f in enumerate(receipt_files):
            ext = os.path.splitext(f.name)[1].lower()
            raw = f.read()

            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                tmp.write(raw)
                tmp_path = tmp.name

            try:
                data = extract_from_file(tmp_path, filename=f.name,
                                         ai_api_key=ai_api_key, ai_provider=ai_provider)
                data["jigyo"] = default_jigyo
                data["_kind"] = "expense"   # 領収書 = 支出
                if not data.get("date"):
                    yr = nendo if month >= 4 else nendo + 1
                    data["date"] = f"{yr}-{month:02d}-01"
                data["_confirmed"] = False
                records.append(data)
                filenames.append(f.name)

                # 利用ログ記録（AIを実際に使った場合のみ）
                _u = data.get("_ai_usage")
                if _u:
                    try:
                        _rate = float(get_setting("usd_jpy", "150") or 150)
                    except Exception:
                        _rate = 150.0
                    try:
                        log_usage(
                            user_id=current_user["id"],
                            username=current_user["username"],
                            provider=_u.get("provider", ai_provider),
                            model=_u.get("model", ""),
                            input_tokens=_u.get("input_tokens", 0),
                            output_tokens=_u.get("output_tokens", 0),
                            filename=f.name,
                            success=bool(data.get("amount") or data.get("vendor")),
                            usd_jpy=_rate,
                        )
                    except Exception:
                        pass

                # 表示用画像（PDFも画像化）
                disp_img = get_display_image(tmp_path, ext)
                images.append(disp_img)

                # 出納簿貼り付け用
                if ext == '.pdf':
                    xl_img = pdf_to_image_bytes(tmp_path, zoom=1.5)
                else:
                    xl_img = image_to_jpeg_bytes(tmp_path)
                excel_images.append(xl_img)

            finally:
                os.unlink(tmp_path)

            # 進捗だけ更新（マスコットアニメは触らず連続再生）
            if receipt_files:
                update_ai_progress(
                    progress_slot,
                    current=i + 1,
                    total=len(receipt_files),
                    filename=f.name,
                    done=i + 1,
                )

        # 収入エントリ（領収書なし）を records に追加
        for _ie in st.session_state.get("pending_income", []):
            records.append({
                "_kind":     "income",
                "_confirmed": False,
                "date":      _ie.get("date", ""),
                "vendor":    _ie.get("vendor", ""),
                "memo":      _ie.get("memo", "") or _ie.get("vendor", ""),
                "amount":    int(_ie.get("amount", 0)),
                "kamoku":    _ie.get("kamoku", "消耗品"),
                "jigyo":     _ie.get("jigyo", default_jigyo),
                "warning":   "",
                "tax_rate":  "",
                "_ocr_engine": "手動入力",
            })
            images.append(None)
            excel_images.append(None)
            filenames.append(f"収入_{_ie.get('vendor','')}")
        # バッファクリア
        st.session_state["pending_income"] = []

        # 並び替えはユーザーに任せる（自動ソートはしない）

        # アニメーション撤去
        try:
            anim_slot.empty()
            progress_slot.empty()
        except Exception:
            pass

        st.session_state.update({
            "records": records,
            "images": images,
            "excel_images": excel_images,
            "filenames": filenames,
            "denpyo_bytes": denpyo_bytes,
            "phase": "review",
        })
        st.rerun()


# =========================================================
# フェーズ2: 確認・編集（アコーディオン方式）
# =========================================================
elif phase == "review":
    records   = st.session_state["records"]
    images    = st.session_state["images"]
    filenames = st.session_state["filenames"]
    total = len(records)
    confirmed_count = sum(1 for r in records if r.get("_confirmed"))
    warning_count   = sum(1 for r in records if r.get("warning") and not r.get("_confirmed"))
    pending_count   = total - confirmed_count
    total_amount    = sum(int(r.get("amount", 0) or 0) for r in records)

    # --- ステータスKPIストリップ ---
    render_kpi_strip([
        ("確認待ち", f"{pending_count} 件",
         f"全 {total} 件中", "📝", "primary"),
        ("警告あり", f"{warning_count} 件",
         "日付・金額が読めなかったもの", "⚠️", "amber"),
        ("確認済み", f"{confirmed_count} 件",
         f"進捗 {int(confirmed_count/total*100) if total else 0}%", "✅", "teal"),
        ("合計金額", f"¥{total_amount:,}",
         "今回の領収書の合計", "💴", "blue"),
    ])

    # --- アクションバー（操作） ---
    with st.container(border=True):
        ac1, ac2, ac3 = st.columns([2, 1, 1])
        with ac1:
            st.markdown("##### 📄 領収書の確認・編集")
            st.caption("各カードを開いて内容を確認し、「確定」を押してください。日付の古い順に並んでいます。")
        with ac2:
            if st.button("全件まとめて確定", use_container_width=True):
                for r in st.session_state["records"]:
                    r["_confirmed"] = True
                st.rerun()
        with ac3:
            write_disabled = (confirmed_count == 0)
            if st.button(f"次へ：順番を確認 ({confirmed_count}件) →",
                         type="primary", use_container_width=True,
                         disabled=write_disabled):
                all_records = st.session_state["records"]
                new_items = [
                    {**r, "_type": "new", "_orig_idx": j}
                    for j, r in enumerate(all_records)
                    if r.get("_confirmed")
                ]
                with st.spinner("既存データを読み込み中..."):
                    existing_items = read_existing_rows(
                        st.session_state.get("denpyo_bytes", b"")
                    )
                all_order_items = existing_items + new_items
                st.session_state["all_order_items"] = all_order_items
                st.session_state["order_records"]   = new_items
                st.session_state["phase"] = "order"
                st.rerun()

    # --- 領収書画像タブ設定（アコーディオン外し・月自動検出） ---
    _denpyo_bytes = st.session_state.get("denpyo_bytes")
    _receipt_sheets = get_receipt_sheets(_denpyo_bytes) if _denpyo_bytes else []

    # 領収書から最頻月を検出（YYYY-MM-DD → 月）
    _months_from_records = []
    for _r in records:
        _d = _r.get("date", "")
        if _d and len(_d) >= 7 and "-" in _d:
            try:
                _mm = int(_d.split("-")[1])
                _months_from_records.append(_mm)
            except Exception:
                pass
    if _months_from_records:
        from collections import Counter
        _dominant_month = Counter(_months_from_records).most_common(1)[0][0]
    else:
        _dominant_month = month  # サイドバーの月にフォールバック

    _default_new_name = f"領収書 {_dominant_month}月分"

    # 既存シートで「{_dominant_month}月」を含むものがあれば、それを最優先デフォルトに
    _matching_sheet = None
    for _s in _receipt_sheets:
        if f"{_dominant_month}月" in _s:
            _matching_sheet = _s
            break

    _tab_options = ["🆕 新しいタブを作成"]
    if _receipt_sheets:
        _tab_options += [f"📋 続きへ追加：{n}" for n in _receipt_sheets]

    # デフォルト選択: 一致シート > 既選択 > 新規
    if _matching_sheet:
        _default_idx = next(
            (i for i, t in enumerate(_tab_options) if _matching_sheet in t), 0
        )
        # 初回のみセット（ユーザー操作後は上書きしない）
        if "receipt_sheet_option" not in st.session_state or \
           st.session_state.get("receipt_sheet_option") == "new":
            st.session_state["receipt_sheet_option"] = _matching_sheet
    else:
        _current_opt = st.session_state.get("receipt_sheet_option", "new")
        _default_idx = 0
        for _ti, _to in enumerate(_tab_options):
            if _current_opt != "new" and _current_opt in _to:
                _default_idx = _ti
                break

    with st.container(border=True):
        st.markdown(
            f"##### 📸 領収書画像の貼り付け先　"
            f"<span style='font-size:0.78rem;color:#677291;font-weight:500'>"
            f"領収書の主な月: {_dominant_month}月</span>",
            unsafe_allow_html=True,
        )
        _tab_sel = st.radio(
            "貼り付けるタブを選択",
            _tab_options,
            index=_default_idx,
            key="receipt_tab_radio",
            label_visibility="collapsed",
        )

        if _tab_sel.startswith("🆕"):
            _new_name = st.text_input(
                "新しいタブ名",
                value=st.session_state.get("receipt_new_sheet_name") or _default_new_name,
                key="receipt_new_name_input",
            )
            st.session_state["receipt_sheet_option"]   = "new"
            st.session_state["receipt_new_sheet_name"] = _new_name or _default_new_name
            st.caption(
                f"📝 「{st.session_state['receipt_new_sheet_name']}」"
                "という新しいタブを作成します"
            )
        else:
            _sel_sheet = _tab_sel.replace("📋 続きへ追加：", "")
            st.session_state["receipt_sheet_option"]   = _sel_sheet
            st.session_state["receipt_new_sheet_name"] = ""
            st.caption(f"📝 「{_sel_sheet}」タブの続きに追記します")

    # --- 各領収書のカード（日付順） ---
    for i, (record, img_bytes, filename) in enumerate(zip(records, images, filenames)):
        confirmed = record.get("_confirmed", False)
        amount    = int(record.get("amount", 0))
        vendor    = record.get("vendor", "不明")
        date      = record.get("date", "")
        has_warn  = bool(record.get("warning"))
        is_income = record.get("_kind") == "income"

        # 状態に応じたバッジ
        if confirmed:
            badge = "✅ 確定済"
        elif has_warn:
            badge = "⚠️ 要確認"
        else:
            badge = "📝 未確認"
        kind_tag = "💰 収入" if is_income else "💴 支出"

        label = (f"{badge}　│　{kind_tag}　│　No.{i+1}　"
                 f"{date or '日付なし'}　{vendor}　¥{amount:,}")

        with st.expander(label, expanded=(not confirmed and has_warn)):
            # 状態マーカー（CSS:has()でエキスパンダー外観を色分け）
            _state_class = (
                "recpt-marker-done" if confirmed
                else ("recpt-marker-warn" if has_warn else "recpt-marker-pending")
            )
            st.markdown(
                f'<span class="recpt-marker {_state_class}"></span>',
                unsafe_allow_html=True,
            )

            img_col, form_col = st.columns([1, 1], gap="large")

            # 左: 画像プレビュー（収入は画像なしカード）
            with img_col:
                if is_income:
                    st.markdown(
                        '<div style="background:linear-gradient(135deg,#e7f4ec 0%,#f3faf6 100%);'
                        'border:1px solid #aedab9;border-radius:12px;padding:1.4rem 1.2rem;'
                        'text-align:center;box-shadow:0 2px 8px rgba(47,138,79,0.10);">'
                        '<div style="font-size:2.2rem;margin-bottom:0.4rem;">💰</div>'
                        '<div style="color:#1c5b34;font-weight:700;font-size:1rem;">'
                        '収入レコード</div>'
                        '<div style="color:#5b7361;font-size:0.78rem;margin-top:0.3rem;">'
                        '領収書（画像）はありません</div>'
                        '</div>',
                        unsafe_allow_html=True,
                    )
                elif img_bytes:
                    render_clickable_image(
                        img_bytes,
                        key=f"{i}-{filename}",
                        caption=filename,
                    )
                else:
                    st.info("画像プレビューなし")
                    st.caption(filename)

                # 補助情報
                if has_warn:
                    st.warning(record["warning"])
                if record.get("_fx_info"):
                    st.info(f"💱 {record['_fx_info']}")
                engine = record.get("_ocr_engine", "")
                if engine:
                    eng_icon = "🟢" if "Vision" in engine or "AI" in engine else "🟡"
                    st.caption(f"{eng_icon} 読み取り: {engine}")
                if record.get("_ai_error"):
                    st.error(f"⚠️ AI読み取りエラー: {record['_ai_error']}")

            # 右: 編集フォーム（整理された2行レイアウト）
            with form_col:
                with st.form(key=f"form_{i}_{filename}"):
                    # 種別バッジ
                    if is_income:
                        st.markdown(
                            '<span style="display:inline-block;padding:3px 12px;'
                            'border-radius:999px;background:#def1e3;color:#1c5b34;'
                            'border:1px solid #aedab9;font-weight:700;font-size:0.8rem;">'
                            '💰 収入</span>',
                            unsafe_allow_html=True,
                        )
                    # 1行目: 日付 + 金額
                    r1c1, r1c2 = st.columns([1, 1])
                    with r1c1:
                        date_val = st.text_input(
                            "📅 日付（YYYY-MM-DD）",
                            value=record.get("date", ""),
                            placeholder="例: 2026-04-15",
                        )
                    with r1c2:
                        amount_val = st.number_input(
                            "💴 金額（税込・円）" if not is_income else "💰 収入金額（円）",
                            value=amount, min_value=0, step=100,
                        )

                    # 2行目: 取引先 / 入金元
                    vendor_val = st.text_input(
                        "🏪 取引先" if not is_income else "💼 入金元",
                        value=record.get("vendor", ""),
                    )

                    # 3行目: 摘要
                    memo_val = st.text_input(
                        "📝 摘要（内容・品名）",
                        value=record.get("memo", ""),
                        placeholder="例: 消耗品購入、ガソリン代、宿泊料 など",
                    )

                    # 4行目: 勘定科目 + 事業名
                    r4c1, r4c2 = st.columns([1, 1])
                    with r4c1:
                        kamoku_idx = (KAMOKU_OPTIONS.index(record.get("kamoku", "消耗品"))
                                      if record.get("kamoku") in KAMOKU_OPTIONS else 0)
                        kamoku_val = st.selectbox(
                            "📂 勘定科目",
                            options=KAMOKU_OPTIONS,
                            index=kamoku_idx,
                        )
                    with r4c2:
                        # カスタム事業名（その他で入力した値）も選択肢に含めて保持
                        _jg_cur = record.get("jigyo", "ミッション活動")
                        _jg_opts = (JIGYO_OPTIONS if _jg_cur in JIGYO_OPTIONS
                                    else JIGYO_OPTIONS + [_jg_cur])
                        jigyo_val = st.selectbox(
                            "🎯 事業名",
                            options=_jg_opts,
                            index=_jg_opts.index(_jg_cur),
                        )

                    # 確定ボタン
                    btn_label = "✅ 確定済（編集して再確定）" if confirmed else "✅ この内容で確定"
                    submitted = st.form_submit_button(
                        btn_label,
                        use_container_width=True,
                        type="primary" if not confirmed else "secondary",
                    )
                    if submitted:
                        st.session_state["records"][i].update({
                            "date":     date_val,
                            "vendor":   vendor_val,
                            "memo":     memo_val,
                            "amount":   amount_val,
                            "kamoku":   kamoku_val,
                            "jigyo":    jigyo_val,
                            "_confirmed": True,
                        })
                        st.rerun()

    # --- ボトムアクションバー（スクロール後にも次へ進める） ---
    st.divider()
    _confirmed_now = sum(1 for r in st.session_state["records"] if r.get("_confirmed"))
    bc1, bc2, bc3 = st.columns([1, 1, 1.4])
    with bc1:
        if st.button("← アップロードに戻る",
                     use_container_width=True, key="review_back_bottom"):
            st.session_state["phase"] = "upload"
            st.rerun()
    with bc2:
        if st.button("全件まとめて確定",
                     use_container_width=True, key="review_confirm_all_bottom"):
            for r in st.session_state["records"]:
                r["_confirmed"] = True
            st.rerun()
    with bc3:
        _write_disabled_b = (_confirmed_now == 0)
        if st.button(f"次へ：順番を確認 ({_confirmed_now}件) →",
                     type="primary", use_container_width=True,
                     key="review_next_bottom",
                     disabled=_write_disabled_b):
            all_records = st.session_state["records"]
            new_items = [
                {**r, "_type": "new", "_orig_idx": j}
                for j, r in enumerate(all_records)
                if r.get("_confirmed")
            ]
            with st.spinner("既存データを読み込み中..."):
                existing_items = read_existing_rows(
                    st.session_state.get("denpyo_bytes", b"")
                )
            all_order_items = existing_items + new_items
            st.session_state["all_order_items"] = all_order_items
            st.session_state["order_records"]   = new_items
            st.session_state["phase"] = "order"
            st.rerun()


# =========================================================
# フェーズ2.5: 順番確認・並び替え（既存＋新規を一覧で並び替え）
# =========================================================
elif phase == "order":
    import pandas as _pd
    from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

    # all_order_items: 既存(_type="existing") + 新規(_type="new")
    all_order_items = st.session_state.get("all_order_items", [])
    ex_count  = sum(1 for r in all_order_items if r.get("_type") == "existing")
    new_count = sum(1 for r in all_order_items if r.get("_type") == "new")
    _total_income  = sum(int(r.get("amount", 0) or 0) for r in all_order_items
                         if r.get("_kind") == "income")
    _total_expense = sum(int(r.get("amount", 0) or 0) for r in all_order_items
                         if r.get("_kind") != "income")

    # KPIストリップ
    render_kpi_strip([
        ("件数", f"{len(all_order_items)} 件",
         f"既存 {ex_count} ／ 新規 {new_count}", "📋", "blue"),
        ("収入 合計", f"¥{_total_income:,}",
         "Q列に書き込まれます", "💰", "teal"),
        ("支出 合計", f"¥{_total_expense:,}",
         "R列に書き込まれます", "💴", "amber"),
        ("差引", f"¥{(_total_income - _total_expense):,}",
         "収入 − 支出", "📊", "primary"),
    ])

    # アクションバー
    with st.container(border=True):
        ac1, ac2, ac3 = st.columns([3, 1, 1])
        with ac1:
            st.markdown("##### 📋 書き込み順番の確認・並び替え")
            st.caption(
                "🖱 No.列をつかんで上下にドラッグで並び替え。"
                "☑️ 左端のチェックで複数行を選び、まとめてドラッグもできます。"
                + ("　📂 既存データも含めて並び替え可能（書き込み時は全件を指定順で書き直し）" if ex_count > 0 else "")
            )
        with ac2:
            if st.button("← 戻って編集", use_container_width=True):
                st.session_state["phase"] = "review"
                st.rerun()
        with ac3:
            if st.button("✅ Excelに書き込む", type="primary",
                         use_container_width=True,
                         disabled=(not all_order_items)):
                st.session_state["phase"] = "writing"
                st.rerun()

    if not all_order_items:
        st.warning("書き込むデータがありません")
    else:
        # ===== 出納簿風プレビュー（AgGrid: ドラッグで並び替え） =====
        items = st.session_state.get("all_order_items", all_order_items)
        _rows = []
        for i, r in enumerate(items):
            ds = r.get("date", "")
            try:
                _dt = datetime.strptime(ds, "%Y-%m-%d")
                _disp_date = f"令和{_dt.year - 2018}年{_dt.month}月{_dt.day}日"
            except Exception:
                _disp_date = ds or "—"
            _kind = r.get("_kind", "expense")
            _amt  = int(r.get("amount", 0) or 0)
            _rows.append({
                "_idx":   i,                                # 内部識別用
                "_type":  r.get("_type", "new"),            # 行色分け用
                "_kind":  _kind,                            # 列値振り分け用
                "種別":   "💰 収入" if _kind == "income" else "💴 支出",
                "No.":    i + 1,
                "日付":   _disp_date,
                "事業名": r.get("jigyo", ""),
                "勘定科目": r.get("kamoku", ""),
                "摘要":   r.get("memo", "") or r.get("vendor", ""),
                "取引先": r.get("vendor", ""),
                "収入":   _amt if _kind == "income" else 0,
                "支出":   _amt if _kind != "income" else 0,
            })
        _df = _pd.DataFrame(_rows)

        # 行スタイル（区分 + 種別で色分け）
        row_style_js = JsCode("""
        function(params) {
            const style = {};
            if (params.data._type === 'new') {
                style.backgroundColor = '#fff6e8';
            } else {
                style.backgroundColor = '#f5f7fb';
            }
            if (params.data._kind === 'income') {
                style.borderLeft = '4px solid #2f8a4f';
            }
            return style;
        }
        """)

        # 金額セル: 0は空表示
        amount_formatter = JsCode("""
        function(params) {
            if (!params.value || params.value === 0) return '';
            return '¥' + Number(params.value).toLocaleString();
        }
        """)

        gb = GridOptionsBuilder.from_dataframe(_df)
        gb.configure_default_column(
            resizable=True, sortable=False, editable=False,
            filter=False, suppressMenu=True,
        )
        # 内部列は非表示
        gb.configure_column("_idx",  hide=True)
        gb.configure_column("_type", hide=True)
        gb.configure_column("_kind", hide=True)
        # 表示列（No.列にチェックボックス＋ドラッグハンドルを同居）
        gb.configure_column("No.",     width=120, pinned="left", rowDrag=True,
                            checkboxSelection=True, headerCheckboxSelection=True)
        gb.configure_column("種別",    width=90)
        gb.configure_column("日付",    width=150)
        gb.configure_column("事業名",  width=120)
        gb.configure_column("勘定科目", width=110)
        gb.configure_column("摘要",    flex=1, minWidth=180)
        gb.configure_column("取引先",  width=140)
        gb.configure_column("収入", width=110, type=["numericColumn"],
                            valueFormatter=amount_formatter,
                            cellStyle={"color": "#1c5b34", "fontWeight": "600"})
        gb.configure_column("支出", width=110, type=["numericColumn"],
                            valueFormatter=amount_formatter,
                            cellStyle={"color": "#8a4708", "fontWeight": "600"})
        gb.configure_grid_options(
            rowDragManaged=True,
            rowDragMultiRow=True,        # 複数選択した行をまとめてドラッグ
            rowSelection="multiple",     # 複数行選択を許可
            suppressRowClickSelection=True,  # 選択はチェックボックスのみ（誤選択防止）
            animateRows=True,
            getRowStyle=row_style_js,
            domLayout='normal',
            suppressMovableColumns=True,
            rowHeight=34,
            headerHeight=38,
        )
        grid_options = gb.build()

        _height = min(620, 80 + len(_rows) * 34)
        grid_response = AgGrid(
            _df,
            gridOptions=grid_options,
            height=_height,
            width="100%",
            allow_unsafe_jscode=True,
            update_mode="MODEL_CHANGED",
            theme="balham",
            key=f"order_grid_{len(items)}",  # 件数が変わったら再生成
            reload_data=False,
        )

        # ドラッグ後の新しい順序を取得して反映
        try:
            new_data = grid_response.get("data")
            if new_data is not None and len(new_data) == len(items):
                new_order_idx = [int(x) for x in new_data["_idx"].tolist()]
                if new_order_idx != list(range(len(items))):
                    new_items = [items[i] for i in new_order_idx]
                    st.session_state["all_order_items"] = new_items
                    st.rerun()
        except Exception:
            pass

        # 凡例
        st.caption(
            '<span style="background:#fff6e8;padding:2px 10px;border-radius:4px;'
            'border:1px solid #e8d4a8;color:#8a5a17;font-size:0.8rem;font-weight:600">'
            '🆕 新規追加</span>　'
            '<span style="background:#f5f7fb;padding:2px 10px;border-radius:4px;'
            'border:1px solid #d6dde6;color:#3e4a6a;font-size:0.8rem;font-weight:600">'
            '📂 既存データ</span>　'
            '<span style="color:#677291;font-size:0.78rem;">'
            'No.列をドラッグで並び替え／左端☑️で複数選択してまとめて移動</span>',
            unsafe_allow_html=True,
        )

    # 下部にもう一度書き込みボタン
    if all_order_items:
        st.divider()
        bc1, bc2 = st.columns([1, 1])
        with bc1:
            if st.button("← 戻って編集する", use_container_width=True,
                         key="order_back_bottom"):
                st.session_state["phase"] = "review"
                st.rerun()
        with bc2:
            if st.button("✅ この順番でExcelに書き込む", type="primary",
                         use_container_width=True, key="order_write_bottom"):
                st.session_state["phase"] = "writing"
                st.rerun()


# =========================================================
# フェーズ3: 書き込み処理（非表示で実行）
# =========================================================
elif phase == "writing":
    excel_images = st.session_state["excel_images"]

    with st.spinner("出納簿に書き込み中..."):
        all_order_items = st.session_state.get("all_order_items")

        if all_order_items:
            # 既存データが含まれているか確認 → 含まれる場合は全書き直しモード
            has_existing = any(r.get("_type") == "existing" for r in all_order_items)
            rewrite_all  = has_existing

            # 書き込み用レコード（内部管理フィールドの一部を除去、_typeは残す）
            _drop_keys = {"_confirmed", "_ocr_engine", "_ai_error", "_fx_info",
                          "warning", "_orig_idx"}
            write_records = [
                {k: v for k, v in r.items() if k not in _drop_keys}
                for r in all_order_items
            ]

            # 画像リスト: 新規アイテム(_type="new")のみ、順番どおりに収集
            write_imgs = []
            for rec in all_order_items:
                if rec.get("_type") == "new":
                    orig_idx = rec.get("_orig_idx")
                    if orig_idx is not None and orig_idx < len(excel_images):
                        write_imgs.append((0, excel_images[orig_idx]))  # noは内部で決定

        else:
            # フォールバック（order_records のみの場合）
            order_records = st.session_state.get("order_records", [])
            rewrite_all   = False
            _drop_keys    = {"_confirmed", "_ocr_engine", "_ai_error", "_fx_info",
                             "warning", "_orig_idx", "_type"}
            write_records = [
                {k: v for k, v in r.items() if k not in _drop_keys}
                for r in order_records
            ]
            write_imgs = [
                (i + 1, excel_images[r["_orig_idx"]])
                for i, r in enumerate(order_records)
                if r.get("_orig_idx") is not None and r["_orig_idx"] < len(excel_images)
            ]

        try:
            updated_bytes, results = write_receipts_to_excel(
                st.session_state["denpyo_bytes"],
                write_records,
                write_imgs,
                receipt_sheet_option=st.session_state.get("receipt_sheet_option", "new"),
                new_sheet_name=st.session_state.get("receipt_new_sheet_name", ""),
                skip_sort=True,        # 並び替え済みなので再ソートしない
                rewrite_all=rewrite_all,  # 既存含む場合は全書き直し
            )
            st.session_state["result_bytes"]  = updated_bytes
            st.session_state["write_results"] = results

            # 処理履歴を保存（今回追加した「新規」レコードのみ）
            try:
                from core.history import save_history
                if all_order_items:
                    _new_recs = [r for r in all_order_items
                                 if r.get("_type") == "new"]
                else:
                    _new_recs = st.session_state.get("order_records", [])
                if _new_recs:
                    _sheet = (st.session_state.get("receipt_new_sheet_name")
                              or st.session_state.get("receipt_sheet_option", ""))
                    save_history(
                        user_id=current_user["id"],
                        username=current_user["username"],
                        records=_new_recs,
                        sheet_name=_sheet,
                    )
            except Exception:
                pass

            st.session_state["phase"] = "done"
        except Exception as e:
            st.error(f"エラー: {e}")
            st.session_state["phase"] = "review"
        st.rerun()


# =========================================================
# フェーズ4: 完了・ダウンロード
# =========================================================
elif phase == "done":
    results      = st.session_state.get("write_results", [])
    result_bytes = st.session_state.get("result_bytes")

    added   = [r for r in results if r["status"] == "追加"]
    skipped = [r for r in results if r["status"] == "重複スキップ"]
    errors  = [r for r in results if "エラー" in r.get("status", "")]

    st.success(f"✅ 処理完了！　追加: {len(added)}件　スキップ(重複): {len(skipped)}件")

    df = pd.DataFrame([{
        "No.":   r.get("no", "-"),
        "取引先": r["vendor"],
        "金額":   f"¥{int(r['amount']):,}",
        "結果":   r["status"],
    } for r in results])
    st.dataframe(df, use_container_width=True, hide_index=True)

    if errors:
        for e in errors:
            st.error(f"{e['vendor']}: {e['status']}")

    st.divider()

    if result_bytes:
        fname = f"{nendo}年度出納簿_{tsuki_kubun}.xlsx"
        st.download_button(
            label=f"📥 {fname} をダウンロード",
            data=result_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    st.divider()
    st.markdown("##### やり直す・続ける")
    st.caption("間違いに気付いたら戻って修正できます。修正後は再度 Excel に書き直されます。")
    bc1, bc2, bc3 = st.columns(3)
    with bc1:
        if st.button("← 確認・編集に戻る", use_container_width=True,
                     help="領収書1件ずつの内容を修正できます"):
            st.session_state["phase"] = "review"
            st.rerun()
    with bc2:
        if st.button("← 順番調整に戻る", use_container_width=True,
                     help="書き込み順を入れ替えてやり直します",
                     disabled=not st.session_state.get("all_order_items")):
            st.session_state["phase"] = "order"
            st.rerun()
    with bc3:
        if st.button("📋 続けて処理する（次の月など）", use_container_width=True,
                     type="primary",
                     help="今ダウンロードしたExcelを起点に次の月の処理へ進みます"):
            st.session_state.update({
                "denpyo_bytes": result_bytes,
                "records": [], "images": [], "excel_images": [],
                "filenames": [],
                "all_order_items": [], "order_records": [],
                "result_bytes": None, "write_results": [],
                "phase": "upload",
            })
            st.rerun()
