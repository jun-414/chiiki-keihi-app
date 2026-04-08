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

from core.extract import (
    extract_from_file,
    pdf_to_image_bytes,
    image_to_jpeg_bytes,
    KAMOKU_OPTIONS,
    JIGYO_OPTIONS,
)
from core.excel_writer import write_receipts_to_excel

# =========================================================
# ページ設定
# =========================================================
st.set_page_config(
    page_title="地域おこし 経費管理",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container { padding-top: 1.5rem; }
.receipt-header { font-size: 1rem; font-weight: bold; }
div[data-testid="stExpander"] { border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 6px; }

/* ドラッグ＆ドロップエリアのスタイル */
div[data-testid="stFileUploader"] {
    border: 2.5px dashed #4a90d9;
    border-radius: 14px;
    padding: 0.5rem 1rem 1rem 1rem;
    background: #f4f9ff;
    transition: background 0.2s;
}
div[data-testid="stFileUploader"]:hover {
    background: #e8f2ff;
    border-color: #1a6bbf;
}
div[data-testid="stFileUploader"] label {
    font-size: 0.9rem;
    color: #444;
}
/* アップロードボタンとドロップゾーン内テキスト */
div[data-testid="stFileUploaderDropzone"] {
    padding: 1.5rem 1rem;
}
div[data-testid="stFileUploaderDropzoneInstructions"] > div > span {
    font-size: 1.05rem !important;
    font-weight: 600;
    color: #1a6bbf;
}
div[data-testid="stFileUploaderDropzoneInstructions"] > div > small {
    font-size: 0.85rem;
    color: #666;
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# ヘルパー
# =========================================================
def get_display_image(filepath, ext):
    """ファイルをブラウザ表示用の画像バイトに変換"""
    if ext == '.pdf':
        return pdf_to_image_bytes(filepath, zoom=2.0)
    else:
        return image_to_jpeg_bytes(filepath)


def load_template():
    tpl = os.path.join(os.path.dirname(__file__), "templates", "出納簿テンプレート.xlsx")
    with open(tpl, "rb") as f:
        return f.read()


# =========================================================
# セッション初期化
# =========================================================
def init_session():
    defaults = {
        "phase": "upload",
        "records": [],
        "images": [],
        "excel_images": [],
        "filenames": [],
        "denpyo_bytes": None,
        "result_bytes": None,
        "write_results": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_session()


# =========================================================
# サイドバー
# =========================================================
with st.sidebar:
    st.header("⚙️ 設定")

    now = datetime.now()
    nendo = st.selectbox(
        "年度",
        list(range(now.year - 2, now.year + 3)),
        index=2,
        format_func=lambda y: f"{y}年度（令和{y - 2018}年度）"
    )
    month_opts = list(range(4, 13)) + list(range(1, 4))
    month = st.selectbox("月区分", month_opts, format_func=lambda m: f"{m}月")
    tsuki_kubun = f"{nendo}-{month:02d}" if month >= 4 else f"{nendo + 1}-{month:02d}"
    st.caption(f"月区分: `{tsuki_kubun}`")

    st.divider()
    default_jigyo = st.selectbox("事業名（デフォルト）", JIGYO_OPTIONS)

    st.divider()
    if st.button("🔄 最初からやり直す", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

    # AI設定
    st.divider()
    st.markdown("**🤖 AI読み取り（高精度）**")

    _provider = st.radio(
        "AIサービス",
        ["Gemini（無料）", "Claude（有料・高精度）"],
        index=0,
        horizontal=True,
        help="Geminiは無料枠あり（1日1500回）。Claudeは有料だが日本語精度が高い。",
    )
    ai_provider = "gemini" if "Gemini" in _provider else "claude"

    # ===== APIキーの取得優先順位 =====
    # 1. Streamlit Secrets（クラウド公開時に管理者が設定）
    # 2. 環境変数（.envファイルや起動時に設定）
    # 3. ローカル保存ファイル（過去に入力して保存したもの）
    # 4. サイドバー手動入力
    _secret_key_name = "GEMINI_API_KEY" if ai_provider == "gemini" else "ANTHROPIC_API_KEY"

    _preset_key = ""
    _key_source = ""

    # 1. Streamlit Secrets
    try:
        _preset_key = st.secrets.get(_secret_key_name, "")
        if _preset_key:
            _key_source = "管理者設定済み"
    except Exception:
        pass

    # 2. 環境変数
    if not _preset_key:
        _preset_key = os.environ.get(_secret_key_name, "")
        if _preset_key:
            _key_source = "環境変数"

    # 3. ローカル保存ファイル
    _key_file = os.path.join(os.path.dirname(__file__), f".{ai_provider}_key")
    _saved_key = ""
    if not _preset_key and os.path.exists(_key_file):
        with open(_key_file, "r") as _f:
            _saved_key = _f.read().strip()
        if _saved_key:
            _key_source = "保存済み"

    if _preset_key:
        # 管理者設定or環境変数の場合はキー入力欄を非表示
        ai_api_key = _preset_key
        st.success(f"✅ AI読み取り: 有効（{_provider} / {_key_source}）")
    else:
        # 手動入力
        _placeholder = "AIzaSy..." if ai_provider == "gemini" else "sk-ant-..."
        _help_msg = (
            "aistudio.google.com/apikey で無料取得（Googleアカウントのみ・クレカ不要）"
            if ai_provider == "gemini"
            else "console.anthropic.com でキー取得（クレカ登録必要・約0.1〜0.2円/枚）"
        )
        _input_key = st.text_input(
            "APIキー",
            value=_saved_key,
            type="password",
            placeholder=_placeholder,
            help=_help_msg,
        )
        # 新しいキーが入力されたら保存
        if _input_key and _input_key != _saved_key:
            with open(_key_file, "w") as _f:
                _f.write(_input_key)
        ai_api_key = _input_key

        if ai_api_key:
            _icon = "🟢" if ai_provider == "gemini" else "🔵"
            st.success(f"{_icon} AI読み取り: 有効（{_provider}）")
        else:
            st.caption("未設定の場合はルールベースで読み取り")

    # OCR情報
    st.divider()
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        st.caption("🔍 OCR: 利用可能")
    except Exception:
        st.caption("⚠️ OCR: Apple Vision使用中")


# =========================================================
# タイトル
# =========================================================
st.title("📋 地域おこし協力隊 経費管理")
phase = st.session_state.get("phase", "upload")


# =========================================================
# フェーズ1: アップロード
# =========================================================
if phase == "upload":
    st.markdown("### ファイルのアップロード")
    col_l, col_r = st.columns(2, gap="large")

    with col_l:
        st.markdown("**① 出納簿 Excel**")
        denpyo_file = st.file_uploader(
            "📂 ここにドラッグ＆ドロップ、またはクリックして選択",
            type=["xlsx"],
            key="denpyo_up",
            help="既存の出納簿 .xlsx ファイルをドラッグ＆ドロップするか、クリックして選択してください",
        )
        use_template = False
        if not denpyo_file:
            use_template = st.checkbox(f"テンプレートから新規作成（{nendo}年度）")
        if denpyo_file:
            # 現在の使用件数を表示
            try:
                import openpyxl
                from io import BytesIO as _BytesIO
                from core.excel_writer import count_filled_rows
                _wb = openpyxl.load_workbook(_BytesIO(denpyo_file.read()))
                _cnt = count_filled_rows(_wb['出納簿'])
                denpyo_file.seek(0)  # 読み位置をリセット
                st.success(f"✅ {denpyo_file.name}　（現在 {_cnt} 件入力済み）")
            except Exception:
                denpyo_file.seek(0)
                st.success(f"✅ {denpyo_file.name}")
        elif use_template:
            st.info("テンプレートを使用します（0件からスタート）")

    with col_r:
        st.markdown("**② 領収書ファイル**")
        receipt_files = st.file_uploader(
            "📂 ここにドラッグ＆ドロップ、またはクリックして選択（複数OK）",
            type=["pdf", "jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key="receipt_up",
            help="PDF・JPG・PNG に対応。複数ファイルをまとめてドラッグ＆ドロップできます",
        )
        if receipt_files:
            st.success(f"✅ {len(receipt_files)}件 選択済み")

    st.divider()
    can_start = receipt_files and (denpyo_file or use_template)
    if not can_start:
        st.info("出納簿と領収書をアップロードしてください")

    if can_start and st.button("🚀 読み取り開始", type="primary", use_container_width=True):
        _spin_msg = "🤖 AI読み取り中..." if ai_api_key else "読み取り中..."
        with st.spinner(_spin_msg):
            denpyo_bytes = load_template() if use_template else denpyo_file.read()
            records, images, excel_images, filenames = [], [], [], []
            prog = st.progress(0)

            for i, f in enumerate(receipt_files):
                ext = os.path.splitext(f.name)[1].lower()
                raw = f.read()

                with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                    tmp.write(raw)
                    tmp_path = tmp.name

                try:
                    # 2件目以降はAPI制限対策で少し待つ（無料枠: 15回/分）
                    if i > 0 and ai_api_key:
                        import time as _time
                        _time.sleep(4)
                    data = extract_from_file(tmp_path, filename=f.name,
                                             ai_api_key=ai_api_key, ai_provider=ai_provider)
                    data["jigyo"] = default_jigyo
                    if not data.get("date"):
                        yr = nendo if month >= 4 else nendo + 1
                        data["date"] = f"{yr}-{month:02d}-01"
                    data["_confirmed"] = False
                    records.append(data)
                    filenames.append(f.name)

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

                prog.progress((i + 1) / len(receipt_files))

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

    # --- ヘッダー ---
    hdr_col1, hdr_col2, hdr_col3 = st.columns([3, 1, 1])
    with hdr_col1:
        st.markdown(f"### 📄 領収書の確認 　確認済み **{confirmed_count} / {total}** 件")
    with hdr_col2:
        if st.button("全件確認済みにする", use_container_width=True):
            for r in st.session_state["records"]:
                r["_confirmed"] = True
            st.rerun()
    with hdr_col3:
        write_disabled = (confirmed_count == 0)
        if st.button(f"📥 {confirmed_count}件を出納簿へ書き込む",
                     type="primary", use_container_width=True,
                     disabled=write_disabled):
            st.session_state["phase"] = "writing"
            st.rerun()

    st.divider()

    # --- 各領収書のアコーディオン ---
    for i, (record, img_bytes, filename) in enumerate(zip(records, images, filenames)):
        confirmed = record.get("_confirmed", False)
        amount    = int(record.get("amount", 0))
        vendor    = record.get("vendor", "不明")
        date      = record.get("date", "")
        has_warn  = bool(record.get("warning"))

        icon  = "✅" if confirmed else ("⚠️" if has_warn else "📝")
        label = f"{icon}　No.{i+1}　{date}　{vendor}　¥{amount:,}　[{filename}]"

        with st.expander(label, expanded=(not confirmed and has_warn)):

            img_col, form_col = st.columns([1, 1], gap="large")

            # 左: 画像プレビュー
            with img_col:
                if img_bytes:
                    st.image(img_bytes, use_container_width=True,
                             caption=filename)
                else:
                    st.info("画像プレビューなし")
                    st.caption(filename)

                if has_warn:
                    st.warning(record["warning"])

                # 外貨換算情報の表示
                if record.get("_fx_info"):
                    st.info(f"💱 {record['_fx_info']}")

                # 読み取りエンジン表示（診断用）
                engine = record.get("_ocr_engine", "")
                if engine:
                    color = "🟢" if "Vision" in engine or "AI" in engine else "🟡"
                    st.caption(f"{color} 読み取り: {engine}")
                if record.get("_ai_error"):
                    st.error(f"⚠️ AI読み取りエラー: {record['_ai_error']}")

            # 右: 編集フォーム
            with form_col:
                with st.form(key=f"form_{i}_{filename}"):
                    date_val = st.text_input("📅 日付（YYYY-MM-DD）",
                                             value=record.get("date", ""),
                                             placeholder="例: 2026-04-15")
                    vendor_val = st.text_input("🏪 取引先",
                                              value=record.get("vendor", ""))
                    memo_val = st.text_input("📝 摘要（内容・品名）",
                                             value=record.get("memo", ""),
                                             placeholder="例: 消耗品購入、ガソリン代、宿泊料 など")
                    amount_val = st.number_input("💴 金額（税込・円）",
                                                 value=amount,
                                                 min_value=0,
                                                 step=100)
                    kamoku_idx = KAMOKU_OPTIONS.index(record.get("kamoku", "消耗品")) \
                                 if record.get("kamoku") in KAMOKU_OPTIONS else 0
                    kamoku_val = st.selectbox("📂 勘定科目",
                                             options=KAMOKU_OPTIONS,
                                             index=kamoku_idx)
                    jigyo_idx = JIGYO_OPTIONS.index(record.get("jigyo", "ミッション活動")) \
                                if record.get("jigyo") in JIGYO_OPTIONS else 0
                    jigyo_val = st.selectbox("🎯 事業名",
                                            options=JIGYO_OPTIONS,
                                            index=jigyo_idx)

                    btn_label = "✅ 確定済みに戻す" if confirmed else "✅ 確定"
                    submitted = st.form_submit_button(
                        btn_label,
                        use_container_width=True,
                        type="primary" if not confirmed else "secondary"
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


# =========================================================
# フェーズ3: 書き込み処理（非表示で実行）
# =========================================================
elif phase == "writing":
    records      = st.session_state["records"]
    excel_images = st.session_state["excel_images"]

    with st.spinner("出納簿に書き込み中..."):
        write_records = [r for r in records if r.get("_confirmed")]
        write_imgs    = [(i + 1, excel_images[j])
                         for j, (i, r) in enumerate(zip(range(len(records)), records))
                         if r.get("_confirmed")]
        try:
            updated_bytes, results = write_receipts_to_excel(
                st.session_state["denpyo_bytes"],
                write_records,
                write_imgs,
            )
            st.session_state["result_bytes"]  = updated_bytes
            st.session_state["write_results"] = results
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
    if st.button("📋 続けて処理する（次の月など）", use_container_width=True):
        st.session_state.update({
            "denpyo_bytes": result_bytes,
            "records": [], "images": [], "excel_images": [],
            "filenames": [], "phase": "upload",
        })
        st.rerun()
