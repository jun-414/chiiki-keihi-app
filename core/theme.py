"""
theme.py - ダッシュボード型UI（ダークサイドバー + 白基調 + KPIカード）

参考: ダッシュボード型SaaS UI + 行政サイトの信頼感
  - サイドバー: ダークネイビーグラデ（#1c2747 → #0f1830）
  - メイン: 白基調、KPIカード、立体感のあるカード群
  - アクセント: テラコッタ #ca5142（CTA・強調）/ ティール #2c7a7b（補助）

各ページの st.set_page_config 直後に apply_theme() を呼ぶことで適用される。
"""
import base64
import re
import streamlit as st


_CSS = """
<style>
:root {
    /* ===== ブランドカラー ===== */
    --primary:        #ca5142;   /* 温かいテラコッタ（CTA） */
    --primary-hover:  #b54533;
    --primary-light:  #fbecea;
    --primary-soft:   #f5d8d3;

    --accent-1:       #2c7a7b;   /* ティール（成功・補助） */
    --accent-2:       #d69e2e;   /* アンバー（注意・指標） */
    --accent-3:       #4f8edc;   /* ブルー（情報・指標） */

    /* ===== サイドバー（ダーク） ===== */
    --sb-bg-from:     #1c2747;
    --sb-bg-to:       #0f1830;
    --sb-text:        #e7ecf5;
    --sb-text-muted:  #97a0b8;
    --sb-border:      #2a3759;
    --sb-active:      rgba(202, 81, 66, 0.18);

    /* ===== メイン（ライト） ===== */
    --bg:             #f4f6fb;   /* 少し冷たい白でカードが浮かぶ */
    --card:           #ffffff;
    --surface:        #faf7f4;
    --surface-2:      #f1ece6;
    --border:         #e3e7ef;
    --border-strong:  #c7cfdc;

    /* ===== テキスト ===== */
    --text:           #1a2540;
    --text-strong:    #0c142a;
    --text-muted:     #677291;
    --text-on-primary:#ffffff;

    /* ===== ステータス ===== */
    --success:        #2f8a4f;
    --warning:        #b25a14;
    --error:          #c0291f;
    --info:           #1f6fa8;

    /* ===== 影 ===== */
    --shadow-sm: 0 1px 2px rgba(15, 24, 48, 0.05);
    --shadow:    0 2px 8px rgba(15, 24, 48, 0.08), 0 1px 2px rgba(15, 24, 48, 0.04);
    --shadow-md: 0 6px 18px rgba(15, 24, 48, 0.10), 0 2px 4px rgba(15, 24, 48, 0.04);
    --shadow-lg: 0 12px 28px rgba(15, 24, 48, 0.14);
}

/* ===== グローバル ===== */
html, body, .stApp, [class*="css"] {
    font-family: -apple-system, "Hiragino Sans", "Hiragino Kaku Gothic ProN",
                 "Yu Gothic UI", "Meiryo", system-ui, sans-serif;
    color: var(--text);
}
.stApp { background-color: var(--bg); }

.block-container {
    padding-top: 1.2rem;
    padding-bottom: 3rem;
    max-width: 1280px;
}

/* Streamlitのページナビと余白を整理 */
[data-testid="stSidebarNav"] { display: none; }
[data-testid="stHeader"] {
    background: transparent;
    border-bottom: none;
}

/* ===== 見出し ===== */
h1, h2, h3, h4 {
    color: var(--text-strong);
    font-weight: 700;
    letter-spacing: 0.01em;
    line-height: 1.4;
}
h1 { font-size: 1.55rem; margin: 0.2rem 0 1rem 0; }
h2 { font-size: 1.2rem; margin: 1.5rem 0 0.7rem 0; }
h3 { font-size: 1.02rem; margin: 1.1rem 0 0.5rem 0; }
h4 { font-size: 0.92rem; margin: 0.8rem 0 0.4rem 0; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.04em; }

p, li, label { line-height: 1.65; }

/* ===== サイドバー（ダーク） ===== */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, var(--sb-bg-from) 0%, var(--sb-bg-to) 100%) !important;
    border-right: 1px solid var(--sb-border);
}
[data-testid="stSidebar"] * { color: var(--sb-text) !important; }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
[data-testid="stSidebar"] .stMarkdown,
[data-testid="stSidebar"] label { color: var(--sb-text) !important; }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] { color: var(--sb-text-muted) !important; }
[data-testid="stSidebar"] hr { border-color: var(--sb-border); margin: 0.9rem 0; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: var(--sb-text) !important;
    border-bottom: none;
    padding-bottom: 0;
}
[data-testid="stSidebar"] h2 { font-size: 0.82rem; text-transform: uppercase; letter-spacing: 0.08em; color: var(--sb-text-muted) !important; font-weight: 600; margin-top: 1.1rem; }

/* サイドバーの入力欄 */
[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] > div,
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stNumberInput input,
[data-testid="stSidebar"] .stDateInput input {
    background-color: rgba(255,255,255,0.06) !important;
    border: 1px solid var(--sb-border) !important;
    color: var(--sb-text) !important;
    border-radius: 8px;
}
[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] svg { color: var(--sb-text) !important; }

/* サイドバーのボタン */
[data-testid="stSidebar"] .stButton > button {
    background: rgba(255,255,255,0.06);
    color: var(--sb-text) !important;
    border: 1px solid var(--sb-border);
    border-radius: 10px;
    font-weight: 600;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,0.12);
    border-color: var(--primary);
    color: #fff !important;
}

/* サイドバーのアラート（success/warning/info） */
[data-testid="stSidebar"] [data-testid="stAlert"] {
    background: rgba(255,255,255,0.06) !important;
    border: 1px solid var(--sb-border);
    border-radius: 10px;
    color: var(--sb-text) !important;
}
[data-testid="stSidebar"] [data-testid="stAlert"] * { color: var(--sb-text) !important; }

/* サイドバーのページリンク */
[data-testid="stSidebar"] [data-testid="stPageLink"] a {
    background: linear-gradient(135deg, var(--primary) 0%, var(--primary-hover) 100%);
    color: #fff !important;
    border: none;
    border-radius: 10px;
    padding: 0.6rem 0.9rem;
    font-weight: 700;
    box-shadow: var(--shadow-sm);
}
[data-testid="stSidebar"] [data-testid="stPageLink"] a:hover {
    transform: translateY(-1px);
    box-shadow: var(--shadow-md);
}
[data-testid="stSidebar"] [data-testid="stPageLink"] a * { color: #fff !important; }

/* サイドバーのラジオチップ */
[data-testid="stSidebar"] .stRadio [role="radiogroup"] label {
    background: rgba(255,255,255,0.06);
    border: 1px solid var(--sb-border);
    color: var(--sb-text) !important;
}
[data-testid="stSidebar"] .stRadio [role="radiogroup"] label:has(input:checked) {
    background: var(--primary);
    border-color: var(--primary);
    color: #fff !important;
}

/* ===== メイン側ボタン ===== */
.stButton > button {
    border-radius: 10px;
    font-weight: 600;
    padding: 0.55rem 1.2rem;
    transition: all 0.2s ease;
    border: 1px solid var(--border-strong);
    background-color: #fff;
    color: var(--text);
    box-shadow: var(--shadow-sm);
    line-height: 1.3;
}
.stButton > button:hover {
    background-color: var(--primary-light);
    border-color: var(--primary);
    color: var(--primary);
    transform: translateY(-1px);
    box-shadow: var(--shadow);
}
.stButton > button:focus, .stButton > button:focus-visible {
    box-shadow: 0 0 0 3px rgba(202,81,66,0.20);
    outline: none;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg,
        #d96a59 0%,
        var(--primary) 50%,
        #a13e2d 100%);
    color: var(--text-on-primary);
    border: none;
    box-shadow:
        0 4px 14px rgba(202, 81, 66, 0.40),
        inset 0 1px 0 rgba(255, 255, 255, 0.20);
    text-shadow: 0 1px 2px rgba(0, 0, 0, 0.18);
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg,
        #e07868 0%,
        var(--primary-hover) 50%,
        #8b3325 100%);
    color: var(--text-on-primary);
    transform: translateY(-1px);
    box-shadow:
        0 8px 22px rgba(202, 81, 66, 0.48),
        inset 0 1px 0 rgba(255, 255, 255, 0.25);
}
.stButton > button:disabled,
.stButton > button[disabled] {
    opacity: 0.55;
    cursor: not-allowed;
    transform: none !important;
    box-shadow: none !important;
}

/* ===== ダウンロードボタン ===== */
.stDownloadButton > button {
    background: linear-gradient(135deg,
        #3a9698 0%,
        var(--accent-1) 50%,
        #1c5b5c 100%);
    color: #fff;
    border: none;
    border-radius: 10px;
    font-weight: 700;
    padding: 0.55rem 1.2rem;
    box-shadow:
        0 4px 14px rgba(44, 122, 123, 0.38),
        inset 0 1px 0 rgba(255, 255, 255, 0.20);
    text-shadow: 0 1px 2px rgba(0, 0, 0, 0.18);
}
.stDownloadButton > button:hover {
    transform: translateY(-1px);
    box-shadow:
        0 8px 22px rgba(44, 122, 123, 0.45),
        inset 0 1px 0 rgba(255, 255, 255, 0.25);
}

/* ===== 入力フィールド（メイン側） ===== */
.stTextInput input, .stTextArea textarea,
.stNumberInput input, .stDateInput input {
    background-color: #fff !important;
    border-radius: 8px;
    border: 1px solid var(--border-strong);
    color: var(--text) !important;
    padding: 0.5rem 0.75rem;
    box-shadow: var(--shadow-sm);
}
.stSelectbox [data-baseweb="select"] > div,
.stMultiSelect [data-baseweb="select"] > div {
    background-color: #fff !important;
    border-radius: 8px;
    border: 1px solid var(--border-strong);
    color: var(--text) !important;
}
.stTextInput input:focus, .stTextArea textarea:focus,
.stNumberInput input:focus, .stDateInput input:focus {
    border-color: var(--primary);
    box-shadow: 0 0 0 3px rgba(202,81,66,0.15);
}
label { color: var(--text); font-weight: 500; }

/* number_input の +/- ステッパー（base-web由来のダーク色を回避） */
[data-testid="stNumberInputContainer"] {
    background-color: #fff !important;
    border-radius: 8px;
    overflow: hidden;
}
[data-testid="stNumberInputStepDown"],
[data-testid="stNumberInputStepUp"] {
    background-color: var(--surface) !important;
    color: var(--text) !important;
    border-color: var(--border) !important;
}
[data-testid="stNumberInputStepDown"]:hover,
[data-testid="stNumberInputStepUp"]:hover {
    background-color: var(--primary-light) !important;
    color: var(--primary) !important;
}
[data-testid="stNumberInputStepDown"] svg,
[data-testid="stNumberInputStepUp"] svg {
    fill: var(--text) !important;
}
[data-testid="stNumberInputStepDown"]:hover svg,
[data-testid="stNumberInputStepUp"]:hover svg {
    fill: var(--primary) !important;
}

/* ===== 日付ピッカーのポップアップ（カレンダー）===== */
div[data-baseweb="calendar"],
div[data-baseweb="popover"] div[role="dialog"] {
    background-color: #fff !important;
    color: var(--text) !important;
}
div[data-baseweb="calendar"] * {
    color: var(--text) !important;
}
/* カレンダーヘッダー（月年表示）の背景を明るく */
div[data-baseweb="calendar-header"],
div[data-baseweb="calendar"] > div:first-child {
    background-color: var(--surface) !important;
    color: var(--text) !important;
}
/* 曜日ラベル */
div[data-baseweb="calendar"] thead th,
div[data-baseweb="day"] {
    color: var(--text) !important;
}
/* 選択された日付・今日（プライマリ色） */
div[data-baseweb="day"][aria-selected="true"] {
    background-color: var(--primary) !important;
    color: #fff !important;
}

/* ===== ポップオーバー全般（Selectbox/Multiselectのドロップダウン等）===== */
div[data-baseweb="popover"] {
    background-color: #fff !important;
}
ul[role="listbox"],
ul[role="listbox"] li {
    background-color: #fff !important;
    color: var(--text) !important;
}
ul[role="listbox"] li:hover,
ul[role="listbox"] li[aria-selected="true"] {
    background-color: var(--primary-light) !important;
    color: var(--primary) !important;
}

/* ===== トグル ===== */
[data-baseweb="checkbox"] [role="checkbox"] { color: var(--text); }

/* ===== ステータス（alert） ===== */
[data-testid="stAlert"] {
    border-radius: 10px;
    border-left-width: 4px;
    padding: 0.85rem 1rem;
    font-size: 0.94rem;
    box-shadow: var(--shadow-sm);
}

/* ===== エキスパンダー ===== */
div[data-testid="stExpander"] {
    background-color: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    margin-bottom: 10px;
    box-shadow: var(--shadow-sm);
    overflow: hidden;
}
div[data-testid="stExpander"] summary { font-weight: 600; padding: 0.7rem 0.9rem; }
div[data-testid="stExpander"] summary:hover { background-color: var(--surface); }

/* ===== タブ ===== */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    border-bottom: 2px solid var(--border);
    background: transparent;
}
.stTabs [data-baseweb="tab"] {
    background: transparent;
    padding: 0.7rem 1.3rem;
    color: var(--text-muted);
    font-weight: 600;
    border-radius: 8px 8px 0 0;
    transition: all 0.15s;
}
.stTabs [data-baseweb="tab"]:hover {
    color: var(--primary);
    background: var(--primary-light);
}
.stTabs [aria-selected="true"] {
    color: var(--primary) !important;
    border-bottom: 3px solid var(--primary);
    margin-bottom: -2px;
    background: transparent;
}

/* ===== メトリック（KPIカード）— 立体感あり ===== */
[data-testid="stMetric"] {
    background:
        linear-gradient(135deg,
            rgba(251, 236, 234, 0.55) 0%,
            #ffffff 60%) ,
        #ffffff;
    padding: 1.1rem 1.2rem;
    border-radius: 14px;
    border: 1px solid var(--border);
    box-shadow:
        0 4px 14px rgba(15, 24, 48, 0.07),
        inset 0 1px 0 rgba(255, 255, 255, 0.65);
    position: relative;
    overflow: hidden;
    transition: transform 0.2s, box-shadow 0.2s;
}
[data-testid="stMetric"]:hover {
    transform: translateY(-2px);
    box-shadow:
        0 10px 28px rgba(15, 24, 48, 0.12),
        0 2px 6px rgba(15, 24, 48, 0.04);
}
[data-testid="stMetric"]::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 4px;
    background: linear-gradient(90deg,
        var(--primary) 0%,
        var(--accent-2) 50%,
        var(--accent-3) 100%);
}
[data-testid="stMetric"]::after {
    content: "";
    position: absolute;
    top: -40px; right: -30px;
    width: 130px; height: 130px;
    border-radius: 50%;
    background: radial-gradient(circle, rgba(202, 81, 66, 0.10) 0%, transparent 70%);
    pointer-events: none;
}
[data-testid="stMetricLabel"] {
    color: var(--text-muted);
    font-weight: 700;
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}
[data-testid="stMetricValue"] {
    color: var(--text-strong);
    font-weight: 800;
    font-size: 1.75rem !important;
    letter-spacing: -0.01em;
}
[data-testid="stMetricDelta"] { font-size: 0.85rem; }

/* ===== ファイルアップローダー — リッチ ===== */
div[data-testid="stFileUploader"] {
    border: 2px dashed var(--primary);
    border-radius: 14px;
    padding: 0.7rem 1rem 1.2rem 1rem;
    background:
        linear-gradient(135deg, rgba(202,81,66,0.06) 0%, rgba(202,81,66,0.02) 100%),
        var(--card);
    transition: all 0.25s;
    box-shadow: var(--shadow-sm);
}
div[data-testid="stFileUploader"]:hover {
    background:
        linear-gradient(135deg, rgba(202,81,66,0.12) 0%, rgba(202,81,66,0.04) 100%),
        var(--card);
    border-color: var(--primary-hover);
    box-shadow: var(--shadow);
    transform: translateY(-1px);
}
div[data-testid="stFileUploaderDropzone"] { padding: 1.8rem 1rem; }

/* ===== 英語UI文字を日本語化 / 二重表示を解消 ===== */
/* インストラクション内の元テキストはすべて非表示にし、親に1つだけ案内を入れる */
[data-testid="stFileUploaderDropzoneInstructions"] > div > * {
    display: none !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] > div {
    position: relative;
    min-height: 1.4em;
}
[data-testid="stFileUploaderDropzoneInstructions"] > div::before {
    content: "ここにドラッグ";
    font-size: 0.95rem;
    font-weight: 700;
    color: var(--primary);
    line-height: 1.4;
    display: inline-block;
}

/* 「Browse files」ボタン → 「ファイルを選ぶ」 */
[data-testid="stFileUploader"] button,
[data-testid="stFileUploaderDropzone"] button,
[data-testid="stBaseButton-secondary"][data-testid*="stFileUploader"] {
    color: transparent !important;
    position: relative;
    min-width: 110px;
    padding: 0.45rem 0.95rem !important;
}
[data-testid="stFileUploader"] button::after,
[data-testid="stFileUploaderDropzone"] button::after {
    content: "ファイルを選ぶ";
    position: absolute;
    left: 0; right: 0; top: 50%;
    transform: translateY(-50%);
    color: var(--primary) !important;
    font-size: 0.88rem;
    font-weight: 700;
}

/* ===== データテーブル ===== */
[data-testid="stDataFrame"], .stDataFrame {
    border: 1px solid var(--border);
    border-radius: 10px;
    overflow: hidden;
    box-shadow: var(--shadow-sm);
}

/* ===== AgGrid（並び替え表）— ヘッダーグラデ ===== */
.ag-theme-balham,
.ag-theme-balham-dark {
    --ag-border-color: var(--border);
    --ag-border-radius: 12px;
    --ag-cell-horizontal-padding: 12px;
}
.ag-theme-balham .ag-root-wrapper {
    border: 1px solid var(--border) !important;
    border-radius: 12px !important;
    overflow: hidden;
    box-shadow: 0 4px 14px rgba(15, 24, 48, 0.08);
}
.ag-theme-balham .ag-header {
    background: linear-gradient(180deg, #1c2747 0%, #2a3759 100%) !important;
    color: #fff !important;
    border-bottom: 2px solid #ca5142 !important;
}
.ag-theme-balham .ag-header-cell {
    color: #fff !important;
    font-weight: 600;
    font-size: 0.85rem;
}
.ag-theme-balham .ag-header-cell-text {
    color: #fff !important;
}
.ag-theme-balham .ag-row {
    border-bottom: 1px solid rgba(15, 24, 48, 0.06);
}
.ag-theme-balham .ag-row-hover {
    background-color: rgba(202, 81, 66, 0.06) !important;
}

/* ===== container(border=True) — リッチカード（グラデ背景） ===== */
[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 14px !important;
    border-color: var(--border) !important;
    background:
        linear-gradient(135deg,
            rgba(247, 249, 252, 0.85) 0%,
            #ffffff 60%) ,
        #ffffff !important;
    box-shadow:
        0 4px 14px rgba(15, 24, 48, 0.07),
        inset 0 1px 0 rgba(255, 255, 255, 0.65) !important;
    padding: 1rem 1.2rem !important;
    position: relative;
    overflow: hidden;
}
[data-testid="stVerticalBlockBorderWrapper"]::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: linear-gradient(90deg,
        var(--primary) 0%,
        var(--accent-2) 50%,
        var(--accent-3) 100%);
    opacity: 0.85;
    border-radius: 14px 14px 0 0;
}

/* ===== 区切り線 ===== */
hr { border-color: var(--border); margin: 1.2rem 0; }

/* ===== キャプション ===== */
[data-testid="stCaptionContainer"], .caption, small {
    color: var(--text-muted);
    line-height: 1.55;
}
code {
    background: var(--surface-2);
    padding: 1px 6px;
    border-radius: 4px;
    color: var(--primary);
    font-size: 0.92em;
}

/* ===== プログレスバー ===== */
.stProgress > div > div > div > div {
    background: linear-gradient(90deg, var(--primary) 0%, var(--accent-2) 100%);
}

/* ===== ラジオチップ（メイン側） ===== */
.stRadio [role="radiogroup"] { gap: 6px; flex-wrap: wrap; }
.stRadio [role="radiogroup"] label {
    background-color: #fff;
    padding: 0.45rem 0.95rem;
    border-radius: 999px;
    border: 1px solid var(--border-strong);
    margin: 0;
    transition: all 0.15s;
    box-shadow: var(--shadow-sm);
}
.stRadio [role="radiogroup"] label:hover {
    border-color: var(--primary);
    background: var(--primary-light);
    color: var(--primary);
}
.stRadio [role="radiogroup"] label:has(input:checked) {
    border-color: var(--primary);
    background: var(--primary);
    color: #fff;
}
.stRadio [role="radiogroup"] label:has(input:checked) p,
.stRadio [role="radiogroup"] label:has(input:checked) span { color: #fff; }

/* ===== カスタム: 統計バー（KPIストリップ）— グラデーション ===== */
.kpi-strip {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    gap: 10px;
    margin: 0.3rem 0 1rem 0;
}
.kpi-card {
    position: relative;
    background:
        linear-gradient(135deg,
            var(--kpi-accent-soft, var(--primary-light)) 0%,
            #ffffff 65%) ,
        #ffffff;
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 0.6rem 0.8rem;
    box-shadow:
        0 1px 2px rgba(15, 24, 48, 0.06),
        inset 0 1px 0 rgba(255, 255, 255, 0.65);
    transition: transform 0.18s ease, box-shadow 0.18s ease;
    overflow: hidden;
    display: flex;
    align-items: center;
    gap: 0.6rem;
}
.kpi-card:hover {
    transform: translateY(-2px);
    box-shadow:
        0 8px 18px rgba(15, 24, 48, 0.10),
        0 2px 4px rgba(15, 24, 48, 0.05);
}
.kpi-card::before {
    content: "";
    position: absolute; top: 0; left: 0; bottom: 0;
    width: 4px;
    background: linear-gradient(180deg,
        var(--kpi-accent-strong, var(--kpi-accent, var(--primary))) 0%,
        var(--kpi-accent, var(--primary)) 100%);
}
/* 右上のうっすら光るオーナメント（カッコよさ） */
.kpi-card::after {
    content: "";
    position: absolute;
    top: -30px; right: -30px;
    width: 90px; height: 90px;
    border-radius: 50%;
    background: radial-gradient(circle,
        var(--kpi-accent-soft, var(--primary-light)) 0%,
        transparent 70%);
    opacity: 0.55;
    pointer-events: none;
}
.kpi-card .kpi-icon {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 30px; height: 30px;
    border-radius: 8px;
    background: linear-gradient(135deg,
        var(--kpi-accent, var(--primary)) 0%,
        var(--kpi-accent-strong, var(--primary-hover)) 100%);
    color: #ffffff;
    font-size: 0.9rem;
    flex-shrink: 0;
    box-shadow:
        0 3px 8px var(--kpi-accent-glow, rgba(202, 81, 66, 0.28)),
        inset 0 1px 0 rgba(255, 255, 255, 0.30);
    position: relative;
    z-index: 1;
}
.kpi-card .kpi-body { line-height: 1.2; min-width: 0; flex: 1; position: relative; z-index: 1; }
.kpi-card .kpi-label {
    color: var(--text-muted);
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 0.03em;
    margin-bottom: 0.1rem;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.kpi-card .kpi-value {
    color: var(--text-strong);
    font-size: 1.05rem;
    font-weight: 800;
    line-height: 1.2;
    letter-spacing: -0.01em;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.kpi-card .kpi-sub {
    color: var(--text-muted);
    font-size: 0.68rem;
    margin-top: 0.1rem;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

/* バリアント（色違いKPI）— 強アクセント・グロー付き */
.kpi-card.kpi-primary {
    --kpi-accent:        #ca5142;
    --kpi-accent-strong: #a13e2d;
    --kpi-accent-soft:   #fbecea;
    --kpi-accent-glow:   rgba(202, 81, 66, 0.30);
}
.kpi-card.kpi-teal {
    --kpi-accent:        #2c7a7b;
    --kpi-accent-strong: #1c5b5c;
    --kpi-accent-soft:   #d8eded;
    --kpi-accent-glow:   rgba(44, 122, 123, 0.30);
}
.kpi-card.kpi-amber {
    --kpi-accent:        #d69e2e;
    --kpi-accent-strong: #a87521;
    --kpi-accent-soft:   #faedca;
    --kpi-accent-glow:   rgba(214, 158, 46, 0.30);
}
.kpi-card.kpi-blue {
    --kpi-accent:        #4f8edc;
    --kpi-accent-strong: #2d6cb8;
    --kpi-accent-soft:   #dbe7f7;
    --kpi-accent-glow:   rgba(79, 142, 220, 0.30);
}
.kpi-card.kpi-violet {
    --kpi-accent:        #8b5cf6;
    --kpi-accent-strong: #6b3fd8;
    --kpi-accent-soft:   #ede6fb;
    --kpi-accent-glow:   rgba(139, 92, 246, 0.30);
}

/* ===== 画像ライトボックス（クリックで拡大） ===== */
.gov-thumb-wrap {
    margin-bottom: 0.4rem;
}
.gov-thumb-wrap > a.gov-thumb {
    display: block;
    cursor: zoom-in;
    border-radius: 10px;
    overflow: hidden;
    border: 1px solid var(--border);
    background: var(--card);
    box-shadow: var(--shadow-sm);
    transition: transform 0.18s, box-shadow 0.18s, border-color 0.18s;
    position: relative;
}
.gov-thumb-wrap > a.gov-thumb:hover {
    transform: translateY(-1px);
    box-shadow: var(--shadow-md);
    border-color: var(--primary);
}
.gov-thumb-wrap > a.gov-thumb img {
    width: 100%;
    display: block;
}
.gov-thumb-wrap > a.gov-thumb::after {
    content: "🔍 タップで拡大";
    position: absolute;
    top: 10px;
    right: 10px;
    background: rgba(15,24,48,0.78);
    color: #fff;
    font-size: 0.72rem;
    font-weight: 600;
    padding: 4px 10px;
    border-radius: 999px;
    opacity: 0;
    transition: opacity 0.18s;
    pointer-events: none;
}
.gov-thumb-wrap > a.gov-thumb:hover::after { opacity: 1; }
.gov-thumb-wrap .gov-thumb-caption {
    font-size: 0.78rem;
    color: var(--text-muted);
    text-align: center;
    padding: 0.35rem 0.5rem;
    line-height: 1.3;
    word-break: break-all;
}

/* ライトボックス本体 */
a.gov-lightbox {
    display: none;
    position: fixed;
    inset: 0;
    width: 100vw;
    height: 100vh;
    background: rgba(10, 16, 32, 0.92);
    z-index: 99999;
    cursor: zoom-out;
    align-items: center;
    justify-content: center;
    padding: 2rem;
    box-sizing: border-box;
    text-decoration: none;
}
a.gov-lightbox:target { display: flex; }
a.gov-lightbox img {
    max-width: 96vw;
    max-height: 92vh;
    border-radius: 8px;
    box-shadow: 0 25px 60px rgba(0,0,0,0.55);
}
a.gov-lightbox .gov-lightbox-close {
    position: absolute;
    top: 18px;
    right: 22px;
    width: 46px;
    height: 46px;
    border-radius: 50%;
    background: rgba(255,255,255,0.14);
    color: #fff;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-size: 1.6rem;
    font-weight: 300;
    line-height: 1;
    backdrop-filter: blur(6px);
    pointer-events: none;
}
a.gov-lightbox .gov-lightbox-caption {
    position: absolute;
    bottom: 22px;
    left: 50%;
    transform: translateX(-50%);
    color: #d8def0;
    font-size: 0.85rem;
    background: rgba(0,0,0,0.45);
    padding: 6px 14px;
    border-radius: 999px;
    pointer-events: none;
    max-width: 80vw;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

/* ===== 領収書ステータスバッジ ===== */
.gov-status-row {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 0.6rem;
    padding: 0.7rem 0.95rem;
    border-radius: 10px;
    border: 1px solid var(--border);
    background: var(--card);
    box-shadow: var(--shadow-sm);
}
.gov-status-row .badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 4px 12px;
    border-radius: 999px;
    font-size: 0.82rem;
    font-weight: 700;
    white-space: nowrap;
    flex-shrink: 0;
}
.gov-status-row .badge.pending {
    background: #ecf2fa;
    color: #2d4a7a;
    border: 1px solid #c8d6ec;
}
.gov-status-row .badge.warn {
    background: #fff1e0;
    color: #8a4708;
    border: 1px solid #f5d29c;
}
.gov-status-row .badge.done {
    background: #def1e3;
    color: #1c5b34;
    border: 1px solid #aedab9;
}
.gov-status-row .info-grid {
    display: grid;
    grid-template-columns: 92px 1fr auto 110px;
    gap: 8px 14px;
    flex: 1;
    align-items: center;
    min-width: 0;
}
.gov-status-row .info-date {
    color: var(--text-muted);
    font-size: 0.85rem;
    font-weight: 600;
    font-variant-numeric: tabular-nums;
}
.gov-status-row .info-vendor {
    color: var(--text-strong);
    font-weight: 700;
    font-size: 0.95rem;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}
.gov-status-row .info-kamoku {
    color: var(--text-muted);
    font-size: 0.78rem;
    background: var(--surface);
    padding: 2px 8px;
    border-radius: 6px;
    border: 1px solid var(--border);
}
.gov-status-row .info-amount {
    color: var(--text-strong);
    font-weight: 800;
    font-size: 1rem;
    font-variant-numeric: tabular-nums;
    text-align: right;
}

/* 状態別エキスパンダー（中の隠しマーカーで色分け） */
.recpt-marker { display: none; }

div[data-testid="stExpander"]:has(.recpt-marker-pending) {
    border-left: 5px solid #4f8edc !important;
    background: linear-gradient(90deg, #f0f6ff 0%, #fff 70%);
}
div[data-testid="stExpander"]:has(.recpt-marker-pending) summary {
    color: #2d4a7a;
    font-weight: 700;
}
div[data-testid="stExpander"]:has(.recpt-marker-warn) {
    border-left: 5px solid #d69e2e !important;
    background: linear-gradient(90deg, #fff5e0 0%, #fff 70%);
    box-shadow: 0 2px 10px rgba(214,158,46,0.15) !important;
}
div[data-testid="stExpander"]:has(.recpt-marker-warn) summary {
    color: #8a4708;
    font-weight: 700;
}
div[data-testid="stExpander"]:has(.recpt-marker-done) {
    border-left: 5px solid #2f8a4f !important;
    background: linear-gradient(90deg, #ebf7ee 0%, #fff 70%);
    opacity: 0.85;
}
div[data-testid="stExpander"]:has(.recpt-marker-done) summary {
    color: #1c5b34;
}

/* ===== ユーザーバッジ（サイドバー上部）— グラデ・光沢付き ===== */
.user-badge {
    display: flex;
    align-items: center;
    gap: 11px;
    padding: 0.75rem 0.65rem;
    background:
        linear-gradient(135deg,
            rgba(202, 81, 66, 0.14) 0%,
            rgba(255, 255, 255, 0.04) 60%);
    border-radius: 12px;
    border: 1px solid rgba(255, 255, 255, 0.10);
    margin-bottom: 0.7rem;
    box-shadow:
        inset 0 1px 0 rgba(255, 255, 255, 0.08),
        0 2px 8px rgba(0, 0, 0, 0.15);
    position: relative;
    overflow: hidden;
}
.user-badge::before {
    content: "";
    position: absolute;
    top: -20px; right: -20px;
    width: 90px; height: 90px;
    background: radial-gradient(circle, rgba(202, 81, 66, 0.20) 0%, transparent 70%);
    pointer-events: none;
}
.user-badge .avatar {
    width: 38px; height: 38px;
    border-radius: 50%;
    background: linear-gradient(135deg,
        #e07868 0%,
        #ca5142 45%,
        #d69e2e 100%);
    color: #fff;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    font-size: 1.05rem;
    flex-shrink: 0;
    box-shadow:
        0 4px 12px rgba(202, 81, 66, 0.40),
        inset 0 1px 0 rgba(255, 255, 255, 0.30);
    position: relative;
    z-index: 1;
}
.user-badge .meta { line-height: 1.2; position: relative; z-index: 1; }
.user-badge .name { color: #fff; font-weight: 700; font-size: 0.95rem; }
.user-badge .role { color: var(--sb-text-muted); font-size: 0.75rem; margin-top: 2px; }
</style>
"""

# ===== ヘッダーバー（多層グラデーション） =====
_HEADER_TEMPLATE = """
<div style="
    background:
        radial-gradient(ellipse at top right, rgba(214, 158, 46, 0.22) 0%, transparent 55%),
        radial-gradient(ellipse at 60% 120%, rgba(202, 81, 66, 0.40) 0%, transparent 60%),
        linear-gradient(135deg, #1c2747 0%, #243460 55%, #1a2a4d 100%);
    color: #fff;
    padding: 0.85rem 1.2rem;
    border-radius: 14px;
    margin: 0 0 1rem 0;
    display: flex;
    align-items: center;
    gap: 0.85rem;
    box-shadow:
        0 10px 30px rgba(28, 39, 71, 0.22),
        0 2px 6px rgba(28, 39, 71, 0.12),
        inset 0 1px 0 rgba(255, 255, 255, 0.10);
    position: relative;
    overflow: hidden;
    border: 1px solid rgba(255, 255, 255, 0.08);
">
    <div style="
        position: absolute; top: -40px; right: -20px;
        width: 200px; height: 200px;
        background: radial-gradient(circle, rgba(214,158,46,0.18) 0%, transparent 70%);
        pointer-events: none;
    "></div>
    <div style="
        position: absolute; bottom: -30px; left: 30%;
        width: 260px; height: 120px;
        background: radial-gradient(ellipse, rgba(202,81,66,0.25) 0%, transparent 70%);
        pointer-events: none;
    "></div>
    <div style="
        display: inline-flex; align-items: center; justify-content: center;
        width: 40px; height: 40px;
        background: linear-gradient(135deg, rgba(255,255,255,0.18) 0%, rgba(255,255,255,0.05) 100%);
        border: 1px solid rgba(255,255,255,0.15);
        border-radius: 10px;
        font-size: 1.2rem;
        flex-shrink: 0;
        z-index: 1;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.20),
            0 4px 10px rgba(0,0,0,0.18);
    ">{icon}</div>
    <div style="z-index: 1;">
        <div style="font-size: 1.05rem; font-weight: 700; line-height: 1.25; letter-spacing: 0.01em; text-shadow: 0 1px 2px rgba(0,0,0,0.25);">{title}</div>
        <div style="font-size: 0.78rem; opacity: 0.88; margin-top: 2px; font-weight: 400;">{subtitle}</div>
    </div>
</div>
"""

# ===== KPIストリップ =====
_KPI_STRIP_OPEN  = '<div class="kpi-strip">'
_KPI_STRIP_CLOSE = '</div>'

_KPI_CARD = """
<div class="kpi-card kpi-{color}">
    <div class="kpi-icon">{icon}</div>
    <div class="kpi-body">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
</div>
"""

# ===== ユーザーバッジ =====
_USER_BADGE = """
<div class="user-badge">
    <div class="avatar">{initial}</div>
    <div class="meta">
        <div class="name">{name}</div>
        <div class="role">{role}</div>
    </div>
</div>
"""


def apply_theme():
    """共通CSSを適用。各ページの st.set_page_config の直後に呼ぶ。"""
    st.markdown(_CSS, unsafe_allow_html=True)


def render_header(title: str, subtitle: str = "", icon: str = "📋"):
    """公式感のあるタイトルヘッダー帯を描画"""
    st.markdown(
        _HEADER_TEMPLATE.format(icon=icon, title=title, subtitle=subtitle),
        unsafe_allow_html=True,
    )


def render_kpi_strip(items):
    """
    KPIストリップを描画。
    items: [(label, value, sub, icon, color), ...]
        color in {"primary","teal","amber","blue","violet"}
    """
    html = [_KPI_STRIP_OPEN]
    for label, value, sub, icon, color in items:
        html.append(_KPI_CARD.format(
            label=label, value=value, sub=sub or "", icon=icon, color=color,
        ))
    html.append(_KPI_STRIP_CLOSE)
    st.markdown("".join(html), unsafe_allow_html=True)


def render_user_badge(display_name: str, role: str = "ユーザー"):
    """サイドバー用のユーザーバッジ（アバター付き）"""
    initial = (display_name or "?")[0].upper()
    st.markdown(
        _USER_BADGE.format(initial=initial, name=display_name, role=role),
        unsafe_allow_html=True,
    )


_AI_READING_CSS = """
<style>
/* ===== AI読み取りアニメ（高級感ver.） ===== */
@keyframes kk-spin-cw   { to { transform: rotate(360deg); } }
@keyframes kk-spin-ccw  { to { transform: rotate(-360deg); } }
@keyframes kk-scan-move {
    0%, 100% { transform: translateY(0);    opacity: 0.85; }
    15%      { opacity: 1; }
    50%      { transform: translateY(26px); opacity: 1; }
    85%      { opacity: 1; }
}
@keyframes kk-doc-glow {
    0%, 100% { filter: drop-shadow(0 0 3px rgba(202,81,66,0.40)); }
    50%      { filter: drop-shadow(0 0 8px rgba(202,81,66,0.65)); }
}
@keyframes kk-dot-pulse {
    0%, 100% { opacity: 0.35; transform: scale(0.9); }
    50%      { opacity: 1;    transform: scale(1.25); }
}
@keyframes kk-paper-slide {
    0%   { left: 100%; opacity: 0; }
    18%  { left: 60%;  opacity: 1; }
    55%  { left: 38%;  opacity: 1; }
    82%  { left: 16%;  opacity: 1; }
    100% { left: -25%; opacity: 0; }
}
@keyframes kk-stamp {
    0%, 58% { opacity: 0; transform: scale(2.2) rotate(-12deg); }
    72%     { opacity: 1; transform: scale(0.9) rotate(2deg); }
    100%    { opacity: 1; transform: scale(1)   rotate(0deg); }
}
@keyframes kk-progress-shine {
    0%   { background-position: 0% 50%; }
    100% { background-position: 200% 50%; }
}
@keyframes kk-ring-pulse {
    0%   { transform: scale(0.95); opacity: 0.55; }
    100% { transform: scale(1.45); opacity: 0; }
}
@keyframes kk-dot-bounce {
    0%, 100% { transform: translateY(0); opacity: 0.7; }
    50%      { transform: translateY(-2px); opacity: 1; }
}

.ai-reading {
    position: relative;
    display: flex;
    align-items: center;
    gap: 16px;
    padding: 14px 18px;
    border-radius: 14px;
    background:
        radial-gradient(circle at 8% 50%, rgba(28,39,71,0.05) 0%, transparent 55%),
        linear-gradient(135deg, #ffffff 0%, #f7f9fc 100%);
    border: 1px solid rgba(28,39,71,0.10);
    box-shadow:
        0 8px 24px rgba(28,39,71,0.10),
        inset 0 1px 0 rgba(255,255,255,0.85);
    overflow: hidden;
    margin: 0.5rem 0 0.8rem 0;
}
.ai-reading::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg,
        #ca5142 0%, #d69e2e 50%, #2c7a7b 100%);
}

/* AIアイコン領域 */
.ai-reading .ai-char {
    position: relative;
    width: 72px; height: 72px;
    flex-shrink: 0;
    display: flex; align-items: center; justify-content: center;
}
.ai-reading .ai-ring {
    position: absolute;
    inset: 8px;
    border-radius: 50%;
    border: 1px solid rgba(202,81,66,0.30);
    animation: kk-ring-pulse 2.2s ease-out infinite;
    pointer-events: none;
}
.ai-reading .ai-char .kk-svg {
    width: 68px; height: 68px;
    overflow: visible;
}
.ai-reading .kk-svg .kk-ring-outer {
    transform-origin: 50px 50px;
    animation: kk-spin-cw 7s linear infinite;
}
.ai-reading .kk-svg .kk-ring-inner {
    transform-origin: 50px 50px;
    animation: kk-spin-ccw 9s linear infinite;
}
.ai-reading .kk-svg .kk-doc {
    animation: kk-doc-glow 2.4s ease-in-out infinite;
}
.ai-reading .kk-svg .kk-scan {
    animation: kk-scan-move 1.8s ease-in-out infinite;
}
.ai-reading .kk-svg .kk-dot { animation: kk-dot-pulse 1.6s ease-in-out infinite; }
.ai-reading .kk-svg .kk-dot-2 { animation-delay: 0.5s; }
.ai-reading .kk-svg .kk-dot-3 { animation-delay: 1.0s; }

/* 真ん中: 書類が流れる細いベルト */
.ai-reading .scene {
    position: relative;
    width: 110px;
    height: 56px;
    flex-shrink: 0;
    overflow: hidden;
    background:
        linear-gradient(180deg,
            transparent 0%, transparent 80%,
            rgba(28,39,71,0.08) 80%, rgba(28,39,71,0.08) 82%,
            transparent 82%);
    border-radius: 4px;
}
.ai-reading .paper {
    position: absolute;
    top: 14px; left: 50%;
    width: 24px; height: 30px;
    background: linear-gradient(180deg, #ffffff 0%, #f0f3f8 100%);
    border: 1px solid rgba(28,39,71,0.25);
    border-radius: 2px;
    box-shadow: 0 2px 4px rgba(28,39,71,0.12);
    animation: kk-paper-slide 3.2s linear infinite;
    overflow: hidden;
}
.ai-reading .paper::before {
    content: "";
    position: absolute;
    top: 4px; left: 3px; right: 3px; height: 1px;
    background: rgba(28,39,71,0.45);
    box-shadow:
        0 4px 0 0 rgba(28,39,71,0.30),
        0 8px 0 0 rgba(28,39,71,0.30),
        0 12px 0 0 rgba(28,39,71,0.18);
}
.ai-reading .stamp {
    position: absolute;
    top: 18px; left: 26%;
    color: #2c7a7b;
    font-size: 1.1rem;
    font-weight: 900;
    animation: kk-stamp 3.2s ease-in-out infinite;
}

/* テキスト・進捗 */
.ai-reading .info { flex: 1; min-width: 0; }
.ai-reading .info .title {
    font-weight: 700;
    color: #1a2540;
    font-size: 0.92rem;
    letter-spacing: 0.02em;
    display: flex; align-items: center; gap: 6px;
}
.ai-reading .info .title .dot {
    display: inline-block;
    width: 6px; height: 6px;
    border-radius: 50%;
    background: #ca5142;
    animation: kk-dot-bounce 0.9s ease-in-out infinite;
}
.ai-reading .info .file {
    color: #677291;
    font-size: 0.76rem;
    margin-top: 2px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}
.ai-reading .info .bar {
    margin-top: 6px;
    height: 5px;
    background: rgba(28,39,71,0.08);
    border-radius: 99px;
    overflow: hidden;
}
.ai-reading .info .bar > .fill {
    height: 100%;
    border-radius: 99px;
    background: linear-gradient(90deg,
        #ca5142 0%, #d69e2e 50%, #2c7a7b 100%);
    background-size: 200% 100%;
    animation: kk-progress-shine 2.4s linear infinite;
    transition: width 0.35s ease;
}
.ai-reading .info .count {
    font-size: 0.72rem;
    color: #677291;
    margin-top: 3px;
    display: flex; justify-content: space-between;
}
.ai-reading .info .count strong {
    color: #1a2540;
    font-weight: 700;
    font-variant-numeric: tabular-nums;
}

/* 右側: 確認済バッジ（洗練 ver.） */
.ai-reading .stack {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 2px;
    padding: 8px 14px;
    border-radius: 10px;
    background: linear-gradient(135deg, #ffffff 0%, #f7f9fc 100%);
    border: 1px solid rgba(28,39,71,0.10);
    min-width: 70px;
    flex-shrink: 0;
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.85);
}
.ai-reading .stack .pile {
    display: inline-flex;
    align-items: center; justify-content: center;
    width: 22px; height: 22px;
    border-radius: 50%;
    background: linear-gradient(135deg, #2c7a7b 0%, #1c5b5c 100%);
    color: #ffffff;
    font-size: 0.78rem;
    font-weight: 800;
    line-height: 1;
    box-shadow: 0 2px 4px rgba(44,122,123,0.30);
}
.ai-reading .stack .num {
    font-size: 1.0rem;
    font-weight: 800;
    color: #1a2540;
    line-height: 1;
    margin-top: 4px;
    font-variant-numeric: tabular-nums;
    letter-spacing: 0.02em;
}
.ai-reading .stack .lbl {
    font-size: 0.65rem;
    color: #677291;
    font-weight: 600;
    margin-top: 2px;
    letter-spacing: 0.04em;
}

@media (max-width: 720px) {
    .ai-reading .scene { width: 80px; }
    .ai-reading .ai-char { width: 58px; height: 58px; }
    .ai-reading .ai-char .kk-svg { width: 54px; height: 54px; }
}
</style>
"""

# ===== AIスキャナーアイコン（幾何学・洗練デザイン） =====
_MASCOT_SVG = (
    '<svg class="kk-svg" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">'
    '<defs>'
    # 回転する大リングのグラデ（テラコッタ→アンバー→ティール）
    '<linearGradient id="kkRingGrad" x1="0" y1="0" x2="1" y2="1">'
    '<stop offset="0%" stop-color="#ca5142"/>'
    '<stop offset="50%" stop-color="#d69e2e"/>'
    '<stop offset="100%" stop-color="#2c7a7b"/>'
    '</linearGradient>'
    # ドキュメント本体の濃ネイビーグラデ
    '<linearGradient id="kkDocGrad" x1="0" y1="0" x2="0" y2="1">'
    '<stop offset="0%" stop-color="#243460"/>'
    '<stop offset="100%" stop-color="#1c2747"/>'
    '</linearGradient>'
    # 中心のソフトグロー
    '<radialGradient id="kkCenterGlow" cx="0.5" cy="0.5" r="0.5">'
    '<stop offset="0%" stop-color="rgba(202,81,66,0.28)"/>'
    '<stop offset="100%" stop-color="rgba(202,81,66,0)"/>'
    '</radialGradient>'
    # スキャンラインのグラデ（中央が濃い）
    '<linearGradient id="kkScanGrad" x1="0" y1="0" x2="1" y2="0">'
    '<stop offset="0%" stop-color="rgba(202,81,66,0)"/>'
    '<stop offset="50%" stop-color="rgba(202,81,66,0.95)"/>'
    '<stop offset="100%" stop-color="rgba(202,81,66,0)"/>'
    '</linearGradient>'
    '</defs>'
    # センターグロー（一番奥）
    '<circle cx="50" cy="50" r="32" fill="url(#kkCenterGlow)"/>'
    # 外側リング：時計回りに回転（破線の弧）
    '<g class="kk-ring-outer">'
    '<circle cx="50" cy="50" r="42" fill="none" stroke="url(#kkRingGrad)" stroke-width="2" stroke-linecap="round" stroke-dasharray="86 178"/>'
    '<circle cx="50" cy="50" r="42" fill="none" stroke="#ca5142" stroke-width="0.4" opacity="0.18"/>'
    '</g>'
    # 内側リング：反時計回り（細め）
    '<g class="kk-ring-inner">'
    '<circle cx="50" cy="50" r="35" fill="none" stroke="#ca5142" stroke-width="1.2" stroke-linecap="round" stroke-dasharray="44 176" opacity="0.55"/>'
    '</g>'
    # ドキュメント本体（クリップマスクでスキャンラインを内部に閉じ込め）
    '<defs>'
    '<clipPath id="kkDocClip">'
    '<path d="M 38 36 L 56 36 L 62 42 L 62 64 L 38 64 Z"/>'
    '</clipPath>'
    '</defs>'
    '<g class="kk-doc">'
    '<path d="M 38 36 L 56 36 L 62 42 L 62 64 L 38 64 Z" fill="url(#kkDocGrad)"/>'
    '<path d="M 56 36 L 56 42 L 62 42 Z" fill="#3a4a72"/>'
    # 中の罫線（テキスト風）
    '<line x1="42" y1="46" x2="58" y2="46" stroke="#ffffff" stroke-width="0.9" opacity="0.55"/>'
    '<line x1="42" y1="50" x2="56" y2="50" stroke="#ffffff" stroke-width="0.9" opacity="0.55"/>'
    '<line x1="42" y1="54" x2="58" y2="54" stroke="#ffffff" stroke-width="0.9" opacity="0.55"/>'
    '<line x1="42" y1="58" x2="51" y2="58" stroke="#ffffff" stroke-width="0.9" opacity="0.55"/>'
    # スキャンライン（ドキュメント内を上下移動）
    '<g clip-path="url(#kkDocClip)">'
    '<rect class="kk-scan" x="36" y="36" width="28" height="2" fill="url(#kkScanGrad)"/>'
    '</g>'
    '</g>'
    # 小さなドット（角の3点アクセント）
    '<circle class="kk-dot kk-dot-1" cx="22" cy="22" r="1.4" fill="#ca5142"/>'
    '<circle class="kk-dot kk-dot-2" cx="78" cy="78" r="1.4" fill="#2c7a7b"/>'
    '<circle class="kk-dot kk-dot-3" cx="78" cy="22" r="1.2" fill="#d69e2e"/>'
    '</svg>'
)

# インデント無しで一行に並べる（Streamlitマークダウン処理がHTMLとして認識するため）
_AI_READING_HTML = (
    '<div class="ai-reading">'
    '<div class="ai-char">'
    '<div class="ai-ring"></div>'
    '{mascot}'
    '</div>'
    '<div class="scene">'
    '<div class="stamp">✓</div>'
    '<div class="paper"></div>'
    '</div>'
    '<div class="info">'
    '<div class="title"><span class="dot"></span>{title}</div>'
    '<div class="file">{file_line}</div>'
    '<div class="bar"><div class="fill" style="width:{pct}%"></div></div>'
    '<div class="count">'
    '<span>{count_left}</span>'
    '<strong>{current} / {total} 件</strong>'
    '</div>'
    '</div>'
    '<div class="stack">'
    '<div class="pile">✓</div>'
    '<div class="num">{done}</div>'
    '<div class="lbl">確認済</div>'
    '</div>'
    '</div>'
)


# ===== 進捗だけ更新するための軽量HTML（アニメは触らない） =====
_AI_PROGRESS_FLOAT_CSS = """
<style>
.ai-progress-float {
    margin: -6px 0 8px 0;
    padding: 10px 14px;
    border-radius: 10px;
    background: linear-gradient(135deg, #fbecea 0%, #faf7f4 100%);
    border: 1px solid rgba(202,81,66,0.25);
    box-shadow: 0 2px 6px rgba(202,81,66,0.08);
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 0.82rem;
    color: #1a2540;
}
.ai-progress-float .pf-file {
    flex: 1; min-width: 0;
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
    color: #677291;
}
.ai-progress-float .pf-bar {
    flex: 0 0 120px;
    height: 6px;
    background: #f5d8d3;
    border-radius: 99px;
    overflow: hidden;
}
.ai-progress-float .pf-bar > i {
    display: block; height: 100%;
    background: linear-gradient(90deg, #ca5142 0%, #d69e2e 50%, #2c7a7b 100%);
    border-radius: 99px;
    transition: width 0.25s ease;
}
.ai-progress-float .pf-count {
    color: #ca5142;
    font-weight: 800;
    font-variant-numeric: tabular-nums;
    white-space: nowrap;
}
</style>
"""

_AI_PROGRESS_FLOAT_HTML = (
    '<div class="ai-progress-float">'
    '<div class="pf-file">📄 {file_line}</div>'
    '<div class="pf-bar"><i style="width:{pct}%"></i></div>'
    '<div class="pf-count">{done} / {total} 件</div>'
    '</div>'
)


def render_ai_reading(slot, current: int, total: int,
                      filename: str = "", done: int = None,
                      ai_label: str = "AIが領収書を1枚ずつチェック中…"):
    """
    AI読み取り中のかわいいアニメーション（フル）をslot（st.empty()）に描画。
    アニメを途切れさせたくない場合は、ループ中はこれを呼ばずに
    update_ai_progress() で進捗だけ更新する。
    """
    if done is None:
        done = max(0, current - 1)
    pct = int((done / total) * 100) if total else 0
    safe_file = (filename or "").replace("<", "&lt;").replace(">", "&gt;")
    file_line = f"📄 {safe_file} を確認中…" if safe_file else "（準備中）"
    count_left = f"残り {max(0, total - done)} 件"
    html = _AI_READING_CSS + _AI_READING_HTML.format(
        mascot=_MASCOT_SVG,
        title=ai_label,
        file_line=file_line,
        pct=pct,
        count_left=count_left,
        current=min(current, total),
        total=total,
        done=done,
    )
    slot.markdown(html, unsafe_allow_html=True)


def update_ai_progress(slot, current: int, total: int,
                       filename: str = "", done: int = None):
    """
    ループ中の進捗だけを軽量に更新する（マスコットアニメは別スロットで動き続ける）。
    """
    if done is None:
        done = max(0, current)
    pct = int((done / total) * 100) if total else 0
    safe_file = (filename or "").replace("<", "&lt;").replace(">", "&gt;")
    file_line = f"{safe_file} を確認しました" if safe_file else "準備中…"
    html = _AI_PROGRESS_FLOAT_CSS + _AI_PROGRESS_FLOAT_HTML.format(
        file_line=file_line, pct=pct, done=done, total=total,
    )
    slot.markdown(html, unsafe_allow_html=True)


def render_clickable_image(img_bytes: bytes, key: str,
                           caption: str = "", mime: str = "image/jpeg"):
    """
    画像をクリックすると全画面ライトボックスで拡大表示するコンポーネント。
    画像のどこをタップしても拡大される。

    img_bytes: 画像バイト列
    key:       ユニーク識別子（複数画像の区別用、例: f"{i}-{filename}"）
    caption:   キャプション
    mime:      MIMEタイプ（既定 "image/jpeg"）
    """
    if not img_bytes:
        return
    b64 = base64.b64encode(img_bytes).decode("ascii")
    src = f"data:{mime};base64,{b64}"
    safe_key = re.sub(r"[^a-zA-Z0-9_-]", "_", str(key))[:80]
    anchor_id = f"lbox-{safe_key}"
    cap_html = (
        f'<div class="gov-thumb-caption">{caption}</div>'
        if caption else ""
    )
    lbox_cap = (
        f'<span class="gov-lightbox-caption">{caption}</span>'
        if caption else ""
    )
    html = (
        f'<div class="gov-thumb-wrap">'
        f'  <a class="gov-thumb" href="#{anchor_id}">'
        f'    <img src="{src}" alt="{caption}" />'
        f'  </a>'
        f'  {cap_html}'
        f'</div>'
        f'<a class="gov-lightbox" id="{anchor_id}" href="#" aria-label="閉じる">'
        f'  <span class="gov-lightbox-close">×</span>'
        f'  <img src="{src}" alt="{caption}" />'
        f'  {lbox_cap}'
        f'</a>'
    )
    st.markdown(html, unsafe_allow_html=True)
