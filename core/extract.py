"""
extract.py - 領収書（PDF/画像）からデータを抽出する

OCRエンジン優先順位:
  1. pdfplumber（テキストPDF → 最速・最高精度）
  2. Apple Vision Framework（スキャンPDF・画像 → 無料・日本語高精度）
  3. tesseract（インストール済みの場合のみ）

外貨対応:
  - USD/EUR等を検出 → open.er-api.com（無料・キー不要）でその日のレートを取得して円換算
"""
import os
import re
import unicodedata
import tempfile
import json
import urllib.request
from datetime import datetime
from functools import lru_cache

# HEIC（iPhone標準形式）対応: pillow-heif を PIL に登録（無ければ無効化）
HEIC_SUPPORTED = False
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIC_SUPPORTED = True
except Exception:
    HEIC_SUPPORTED = False


# ===== 為替レート取得（無料API） =====

@lru_cache(maxsize=10)
def get_exchange_rate(currency: str, date_str: str = "") -> float:
    """
    指定通貨→JPYのレートを取得（open.er-api.com 無料・キー不要）
    date_str: "YYYY-MM-DD" 指定で過去レートも取得可能
    Returns: レート（失敗時はフォールバック値）
    """
    FALLBACK = {"USD": 150.0, "EUR": 165.0, "GBP": 190.0, "CAD": 110.0, "AUD": 100.0}
    currency = currency.upper()
    if currency == "JPY":
        return 1.0

    try:
        url = f"https://open.er-api.com/v6/latest/{currency}"
        req = urllib.request.Request(url, headers={"User-Agent": "chiiki-keihi-app/1.0"})
        with urllib.request.urlopen(req, timeout=5) as r:
            data = json.loads(r.read())
            rate = data.get("rates", {}).get("JPY")
            if rate:
                return float(rate)
    except Exception:
        pass

    return FALLBACK.get(currency, 150.0)


def convert_to_jpy(amount: float, currency: str, date_str: str = "") -> tuple:
    """
    外貨金額を円に換算
    Returns: (jpy_amount: int, rate: float, currency: str)
    """
    if currency.upper() == "JPY" or not currency:
        return int(round(amount)), 1.0, "JPY"
    rate = get_exchange_rate(currency.upper(), date_str)
    jpy = int(round(amount * rate))
    return jpy, rate, currency.upper()

# ===== 定数 =====
KAMOKU_OPTIONS = [
    "普通旅費", "消耗品", "燃料費", "印刷製本費", "修繕費",
    "通信費", "広告費", "手数料", "保険料", "住宅借上料",
    "会場借上料", "負担金", "その他", "委託料"
]
JIGYO_OPTIONS = ["ミッション活動", "地域活動", "定住活動"]

# キーワード → 勘定科目
# ベンダー名 → 勘定科目（完全一致・部分一致で確実に判定）
VENDOR_KAMOKU_MAP = {
    # 通信・サブスク
    "Adobe": "通信費", "adobe": "通信費", "Adobe Stock": "通信費",
    "1Password": "通信費", "1password": "通信費",
    "OpenAI": "通信費", "Anthropic": "通信費", "Claude": "通信費",
    "ChatGPT": "通信費", "STUDIO": "通信費", "Vercel": "通信費",
    "POVO": "通信費", "povo": "通信費",
    "NTT": "通信費", "SoftBank": "通信費", "docomo": "通信費",
    "Microsoft": "通信費", "Notion": "通信費",
    # 旅費
    "ANA": "普通旅費", "JAL": "普通旅費", "AIRDO": "普通旅費",
    "根室交通": "普通旅費", "阿寒バス": "普通旅費",
    "トヨタレンタ": "普通旅費", "ニッポンレンタカー": "普通旅費",
    "JR": "普通旅費",
    # 燃料
    "ENEOS": "燃料費", "ホクレン": "燃料費", "IDEMITSU": "燃料費",
    "出光": "燃料費", "コスモ": "燃料費",
    # 消耗品
    "Amazon": "消耗品", "楽天": "消耗品", "ケーズデンキ": "消耗品",
    "ヨドバシ": "消耗品", "ビックカメラ": "消耗品", "コジマ": "消耗品",
    # 広告
    "NexusAd": "広告費",
    # 研修・負担金
    "KREDO": "負担金", "リベシティ": "負担金",
    # 印刷
    "ラクスル": "印刷製本費",
}

# テキスト内キーワード → 勘定科目（ベンダー名で判定できなかった場合のフォールバック）
CATEGORY_RULES = [
    # 燃料（確実なキーワードのみ）
    (["ガソリン", "給油", "軽油", "灯油", "燃料費", "ガソリンスタンド"], "燃料費"),
    # 旅費（確実なキーワードのみ）
    (["航空券", "搭乗", "宿泊料", "旅館", "ホテル代", "乗車券", "高速道路", "フェリー", "新幹線"], "普通旅費"),
    # 通信
    (["通信費", "インターネット利用料", "Wi-Fi", "携帯電話", "月額利用料", "サブスクリプション"], "通信費"),
    # 広告
    (["広告費", "広告宣伝費", "チラシ制作", "ポスター制作"], "広告費"),
    # 研修
    (["研修費", "セミナー参加費", "受講料", "講習料", "受験料", "資格取得"], "負担金"),
    # 印刷
    (["印刷費", "名刺印刷", "製本費"], "印刷製本費"),
    # 委託
    (["業務委託費", "外注費", "制作費"], "委託料"),
    # 修繕
    (["修繕費", "修理代", "補修費"], "修繕費"),
    # 会場
    (["会場費", "会場借上", "施設使用料", "貸し会議室"], "会場借上料"),
    # 保険
    (["保険料", "損害保険", "共済費"], "保険料"),
    # 家賃
    (["家賃", "住宅借上料", "賃料"], "住宅借上料"),
    # 手数料
    (["振込手数料", "事務手数料", "取扱手数料"], "手数料"),
]


def normalize(text: str) -> str:
    return unicodedata.normalize('NFKC', text)


def infer_kamoku(vendor: str, text: str = "") -> str:
    """
    勘定科目を推測
    1. ベンダー名の完全一致・部分一致（確実性が高い）
    2. テキスト内キーワード（フォールバック）
    """
    vendor_n = normalize(vendor)

    # 1. ベンダー名で直接判定（最優先）
    for key, kamoku in VENDOR_KAMOKU_MAP.items():
        if key.lower() in vendor_n.lower():
            return kamoku

    # 2. テキストキーワードで判定（確実なキーワードのみ）
    text_n = normalize(text).upper()
    for keywords, kamoku in CATEGORY_RULES:
        for kw in keywords:
            if kw.upper() in text_n:
                return kamoku

    return "消耗品"


# ===== AI抽出（Gemini Vision / Claude 両対応） =====

# 領収書解析プロンプト（共通）
_AI_PROMPT = """あなたは日本の領収書・レシート解析の専門家です。
この領収書・レシートから5つの情報を読み取り、必ずJSON形式のみで返してください。

【重要】返答はJSON1行のみ。前置き・説明・マークダウン・コードブロック（```）は一切不要。
【注意】画像が横向き・逆さま・斜めに回転していることがあります。その場合は文字の向きを正しく読み取る向きに頭の中で回転させてから、内容を読み取ってください。向きが分からない数字や文字を当てずっぽうで書かないこと。

返答例:
{"vendor": "ホクレン 根室SS", "memo": "ガソリン給油", "date": "2026-03-15", "amount": 8540, "kamoku": "燃料費"}

各フィールドの読み取り方:

- vendor: 領収書・レシートを発行したお店や会社の名前（屋号・店名・会社名）
  例: "コープさっぽろ", "ENEOS 根室SS", "根室交通株式会社", "Amazon.co.jp"
  ※ 「上様」「御中」「様」は取引先ではないので含めない
  ※ 宛名（「〇〇様」）ではなく発行元の名前を読む

- memo: 何を購入・利用したか（品目・サービス内容を簡潔に）
  例: "ガソリン給油", "事務用品購入", "宿泊料", "バス運賃", "ソフトウェアサブスク"
  ※ 品目が多い場合は代表的なもの or「消耗品購入」等でまとめる（30文字以内）

- date: 日付をYYYY-MM-DD形式で。**業種に応じた以下のルール**で1つだけ選ぶ。

  【最優先（業種特例）— 必ずこちらを使う】
  - ★航空券・新幹線・特急券・乗車券・フェリー・高速バスなど予約型の交通機関の場合:
    → **搭乗日／乗車日／出発日** を使う。購入日・予約日・決済日・発行日ではない。
    検出キーワード: 「搭乗日」「ご搭乗日」「乗車日」「ご乗車日」「出発日」
                  「Flight Date」「Departure Date」「Travel Date」「Boarding Date」
    例: 航空券に「予約日 2026/03/01、搭乗日 2026/03/15」とあれば → "2026-03-15"
  - ★ホテル・旅館・民宿などの宿泊施設の場合:
    → **チェックイン日／宿泊日** を使う。予約日・決済日・発行日ではない。
    検出キーワード: 「チェックイン」「ご宿泊日」「宿泊日」「宿泊開始日」
                  「Check-in」「Check in date」「Arrival date」
    例: ホテルの請求書に「予約日 2026/01/10、チェックイン 2026/03/15」 → "2026-03-15"

  【通常の物販・サービス】
  上の業種特例に当てはまらない場合は、決済日／領収日／お支払日を使う。
  優先順:
    1位: 「領収日」「お支払日」「決済日」「ご利用日」「Date of payment」「Payment date」
    2位: 発行日／発行年月日／Issue date／Invoice date
    3位: レシート上に1つだけ書かれている日付

  【絶対に使わない】
  - 有効期限／請求期間／サービス利用期間／Next billing date／Period
  - 印字された印刷日・出力日（Printed on / Print date）が領収日と別にある場合
  - 「〜まで」「〜から」のような期間表現の片端だけ

  ※ 令和表記は西暦に変換: 令和7年3月15日 → "2026-03-15"
  ※ 読み取れない場合は空文字 ""

- amount: **税込の合計支払金額** を**日本円（JPY）の整数**で返してください
  ※ 必ず使う金額（最優先・この順）:
      1位: 「合計」「ご請求金額」「お支払い合計」「お支払金額」「税込合計」「Total」「Amount paid」「Grand Total」
      2位: クレジットカード明細の「ご利用金額」「Charged」「Paid」
  ※ **以下は絶対に金額として使わない（よくある誤り）**:
      - 「お預り」「お預かり」「現計」「Cash」「Tendered」（お客が出した現金）
      - 「お釣り」「釣り」「Change」（おつり）
      - 「小計」（税抜の小計のみ書かれていて「合計」が別にある場合）
      - 「内税」「消費税」「(内 ¥XX)」「Tax」（税額そのもの）
      - 「ポイント利用」「割引」「クーポン」（差し引かれる前後の中間値）
      - 「単価」「個別商品の金額」（合計より下の品目別の金額）
  ※ **「税込」と「内税」の区別に注意**:
      - 「¥1,100（内 ¥100）」と書かれていたら **1100** を返す（1100 + 100 = 1200 ではない）
      - 「(内消費税 ¥100)」は税込合計に **含まれている** ので足してはいけない
  ※ 日本円の場合: ¥8,540 → 8540、合計 19,650円 → 19650
  ※ 外貨（USD/EUR等）の場合: 必ず日本円に換算して返す
    例: $4.39 USD → 659（1USD≈150円で計算）
    例: $22.00 USD → 3300（1USD≈150円で計算）
  ※ 読み取れない場合は 0

- kamoku: 勘定科目を以下の選択肢から**1つだけ**選んで返してください（必須）
  選択肢: 普通旅費 / 消耗品 / 燃料費 / 印刷製本費 / 修繕費 / 通信費 / 広告費 / 手数料 / 保険料 / 住宅借上料 / 会場借上料 / 負担金 / その他 / 委託料
  判定基準:
    - 普通旅費: 航空券、JR、バス、タクシー、レンタカー、宿泊（ホテル・旅館）、高速道路、フェリー
    - 燃料費: ガソリン、軽油、灯油、給油（ENEOS、ホクレン、コスモ、出光 等のSS）
    - 通信費: 携帯料金（NTT、docomo、SoftBank、povo）、インターネット、Wi-Fi、サブスク（Adobe、1Password、OpenAI、Anthropic、ChatGPT、Claude、Microsoft、Notion、STUDIO、Vercel、Apple、Google、Netflix、Spotify）
    - 消耗品: 文具、事務用品、日用品、家電量販店（Amazon、楽天、ケーズデンキ、ヨドバシ、ビックカメラ、コジマ、ホームセンター）
    - 印刷製本費: 印刷、名刺、製本（ラクスル等）
    - 広告費: 広告、チラシ、ポスター、SNS広告
    - 手数料: 振込手数料、事務手数料、取扱手数料、決済手数料
    - 保険料: 損害保険、共済、自動車保険
    - 住宅借上料: 家賃、賃料、礼金
    - 会場借上料: 会場費、貸し会議室、施設使用料
    - 負担金: 研修費、セミナー参加費、講習料、受験料、会費（KREDO、リベシティ等の学習サービス含む）
    - 修繕費: 修繕、修理、補修、車検整備
    - 委託料: 業務委託、外注、制作依頼
  迷ったら "消耗品" を返す（"その他" は明らかにどれにも当てはまらない時のみ）"""


def _parse_ai_json(raw: str) -> dict:
    """AI応答からJSONを抽出してパース"""
    raw = re.sub(r'^```[a-z]*\n?|```$', '', raw.strip(), flags=re.MULTILINE).strip()
    m = re.search(r'\{.*\}', raw, re.DOTALL)
    if m:
        raw = m.group()
    import json as _json
    return _json.loads(raw)


GEMINI_MODEL = "gemini-2.0-flash"


def _gemini_api_call(payload_dict: dict, api_key: str, timeout: int = 30) -> dict:
    """
    Gemini API共通呼び出し
    戻り値: {"data": parsed_json, "usage": {input_tokens, output_tokens, model}}
    """
    import json as _json
    import urllib.request as _req

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={api_key}"
    payload = _json.dumps(payload_dict).encode()

    req = _req.Request(url, data=payload, headers={"content-type": "application/json"})
    with _req.urlopen(req, timeout=timeout) as r:
        result = _json.loads(r.read())
        raw = result["candidates"][0]["content"]["parts"][0]["text"]
        meta = result.get("usageMetadata", {}) or {}
        return {
            "data": _parse_ai_json(raw),
            "usage": {
                "provider": "gemini",
                "model": GEMINI_MODEL,
                "input_tokens":  int(meta.get("promptTokenCount", 0)),
                "output_tokens": int(meta.get("candidatesTokenCount", 0)),
            },
        }


def _extract_with_gemini_vision(img_bytes: bytes, api_key: str) -> dict:
    """
    Gemini Vision APIで画像から直接読み取る（OCR不要・最高精度）
    無料枠: 1日1500回、1分15回
    """
    import base64

    # 画像が大きすぎる場合は圧縮（Gemini推奨: 4MB以下）
    if len(img_bytes) > 3 * 1024 * 1024:
        try:
            from PIL import Image as _PIL
            import io as _io
            img = _PIL.open(_io.BytesIO(img_bytes))
            buf = _io.BytesIO()
            img.save(buf, format="JPEG", quality=60)
            img_bytes = buf.getvalue()
        except Exception:
            pass

    img_b64 = base64.b64encode(img_bytes).decode()
    payload = {
        "contents": [{"parts": [
            {"text": _AI_PROMPT},
            {"inline_data": {"mime_type": "image/jpeg", "data": img_b64}},
        ]}],
        "generationConfig": {"maxOutputTokens": 512, "temperature": 0.1},
    }
    return _gemini_api_call(payload, api_key, timeout=30)


def _extract_with_gemini_text(text: str, api_key: str) -> dict:
    """Gemini APIでテキストから抽出（画像化できない場合のフォールバック）"""
    prompt = _AI_PROMPT.replace("この領収書・レシートの画像から", "以下のOCRテキストから") + f"\n\nOCRテキスト:\n{text[:3000]}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"maxOutputTokens": 512, "temperature": 0.1},
    }
    return _gemini_api_call(payload, api_key, timeout=20)


CLAUDE_MODEL = "claude-haiku-4-5"


def _claude_api_call(payload_dict: dict, api_key: str, timeout: int = 30) -> dict:
    """
    Claude API共通呼び出し
    戻り値: {"data": parsed_json, "usage": {input_tokens, output_tokens, model}}
    """
    import json as _json
    import urllib.request as _req

    payload = _json.dumps(payload_dict).encode()
    req = _req.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }
    )
    with _req.urlopen(req, timeout=timeout) as r:
        result = _json.loads(r.read())
        raw = result["content"][0]["text"]
        usage = result.get("usage", {}) or {}
        return {
            "data": _parse_ai_json(raw),
            "usage": {
                "provider": "claude",
                "model": payload_dict.get("model", CLAUDE_MODEL),
                "input_tokens":  int(usage.get("input_tokens", 0)),
                "output_tokens": int(usage.get("output_tokens", 0)),
            },
        }


def _extract_with_claude_vision(img_bytes: bytes, api_key: str) -> dict:
    """
    Claude Vision APIで画像から直接読み取り（高精度・高速）
    モデル: claude-3-5-haiku（約0.1〜0.2円/枚）
    """
    import base64

    # 画像圧縮（5MB以下推奨）
    if len(img_bytes) > 4 * 1024 * 1024:
        try:
            from PIL import Image as _PIL
            import io as _io
            img = _PIL.open(_io.BytesIO(img_bytes))
            buf = _io.BytesIO()
            img.save(buf, format="JPEG", quality=70)
            img_bytes = buf.getvalue()
        except Exception:
            pass

    img_b64 = base64.b64encode(img_bytes).decode()
    payload = {
        "model": CLAUDE_MODEL,
        "max_tokens": 512,
        "messages": [{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/jpeg",
                        "data": img_b64,
                    }
                },
                {"type": "text", "text": _AI_PROMPT}
            ]
        }]
    }
    return _claude_api_call(payload, api_key, timeout=30)


def _extract_with_claude_text(text: str, api_key: str) -> dict:
    """Claude APIでテキストから抽出（フォールバック）"""
    prompt = _AI_PROMPT.replace("この領収書・レシートの画像から", "以下のOCRテキストから") + f"\n\nOCRテキスト:\n{text[:3000]}"
    payload = {
        "model": CLAUDE_MODEL,
        "max_tokens": 512,
        "messages": [{"role": "user", "content": prompt}],
    }
    return _claude_api_call(payload, api_key, timeout=20)


def extract_with_ai(text: str, api_key: str, provider: str = "gemini",
                    img_bytes: bytes = None) -> dict:
    """
    AIで領収書から構造化データを抽出。
    戻り値: {"data": {vendor, memo, date, amount}, "usage": {provider, model, input_tokens, output_tokens}}
    img_bytes があれば Vision APIで画像を直接解析（最高精度）。
    provider: "gemini"（無料） or "claude"（高速・高精度）
    """
    if provider == "claude":
        if img_bytes:
            return _extract_with_claude_vision(img_bytes, api_key)
        else:
            return _extract_with_claude_text(text, api_key)
    else:
        if img_bytes:
            return _extract_with_gemini_vision(img_bytes, api_key)
        else:
            return _extract_with_gemini_text(text, api_key)


# ===== テキスト抽出関数 =====

def extract_date(text: str) -> str:
    t = normalize(text)
    for pat in [
        r'(20\d{2})[/\-年](\d{1,2})[/\-月](\d{1,2})日?',
        r'(\d{4})\.(\d{1,2})\.(\d{1,2})',
        r'(\d{4})(\d{2})(\d{2})',   # 20260415 形式
    ]:
        for m in re.finditer(pat, t):
            try:
                y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if 2020 <= y <= 2035 and 1 <= mo <= 12 and 1 <= d <= 31:
                    return datetime(y, mo, d).strftime("%Y-%m-%d")
            except ValueError:
                pass
    # 令和
    m = re.search(r'令和\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', t)
    if m:
        try:
            return datetime(int(m.group(1)) + 2018, int(m.group(2)), int(m.group(3))).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return ""


def extract_amount_and_currency(text: str) -> tuple:
    """
    金額と通貨を抽出
    Returns: (amount_in_original_currency: float, currency: str)
    """
    t = normalize(text)

    # ===== 外貨（ドル・ユーロ等）の検出 =====
    # "Total $4.39 USD" / "Amount paid $22.00" / "$20.00 USD" 等
    usd_patterns = [
        r'(?:Total|Amount paid|Amount due|Subtotal)[^\d$]*\$\s*([\d,]+\.?\d*)\s*(?:USD)?',
        r'(?:Total|Amount paid)[^\d$]*\$([\d,]+\.?\d*)',
        r'\$\s*([\d,]+\.?\d*)\s*USD',
        r'USD\s*([\d,]+\.?\d*)',
    ]
    for pat in usd_patterns:
        matches = re.findall(pat, t, re.IGNORECASE)
        if matches:
            try:
                val = float(matches[-1].replace(',', ''))
                if 0.01 <= val <= 9999.99:
                    return val, "USD"
            except ValueError:
                pass

    # EUR / GBP / CAD / AUD
    for symbol, currency in [("€", "EUR"), ("£", "GBP"), ("CA\\$", "CAD"), ("AU\\$", "AUD")]:
        m = re.search(rf'{symbol}\s*([\d,]+\.?\d*)', t)
        if m:
            try:
                val = float(m.group(1).replace(',', ''))
                if 0.01 <= val <= 99999:
                    return val, currency
            except ValueError:
                pass

    # ===== 日本円 =====
    # 優先1: 明示的な合計・請求金額パターン
    jpy_priority = [
        r'(?:税込合計|税込み合計|税込金額|税込)[^\d\n]{0,10}[¥￥]?\s*([\d,]+)',
        r'(?:ご請求金額|お支払合計|お支払い合計|請求金額|御請求金額)[^\d\n]{0,10}[¥￥]?\s*([\d,]+)',
        r'(?:合計金額|ご請求額|お支払額)[^\d\n]{0,10}[¥￥]?\s*([\d,]+)',
        r'合\s*計[^\d\n]{0,10}[¥￥]?\s*([\d,]+)',
        r'合計\s*([\d,]+)円',
        r'([\d,]+)円[（\(]?税込',
    ]
    for pat in jpy_priority:
        matches = re.findall(pat, t)
        if matches:
            try:
                val = int(matches[-1].replace(',', ''))
                if 100 <= val <= 9_999_999:
                    return float(val), "JPY"
            except ValueError:
                pass

    # 優先2: ¥マーク付き
    yen_amounts = []
    for m in re.finditer(r'[¥￥]\s*([\d,]+)', t):
        try:
            val = int(m.group(1).replace(',', ''))
            if 100 <= val <= 9_999_999:
                yen_amounts.append(val)
        except ValueError:
            pass
    if yen_amounts:
        return float(max(yen_amounts)), "JPY"

    # 優先3: "XX,XXX円" パターン
    amounts = []
    for m in re.finditer(r'([\d,]{3,})円', t):
        try:
            val = int(m.group(1).replace(',', ''))
            if 100 <= val <= 9_999_999:
                amounts.append(val)
        except ValueError:
            pass
    if amounts:
        return float(max(amounts)), "JPY"

    # 優先4: 手書き領収書向け - "金額" or "Amount" の直後の行にある数字
    # 例: "金額\nAmount\n様\n19650\n御宿泊代"
    lines = t.split('\n')
    for i, line in enumerate(lines):
        if re.search(r'金額|Amount|合計|Total', line, re.IGNORECASE):
            # 次の数行の中から最初に現れる単独の数字を取得
            for j in range(i + 1, min(i + 6, len(lines))):
                candidate = lines[j].strip()
                m = re.fullmatch(r'[\d,]{3,7}', candidate)
                if m:
                    try:
                        val = int(m.group().replace(',', ''))
                        if 500 <= val <= 999_999:
                            return float(val), "JPY"
                    except ValueError:
                        pass

    # 優先5: 伝票番号・No.などの直後を除いた単独数字（最終手段）
    # 伝票番号・No.・TELの直後は除外する
    skip_next = False
    standalone = []
    for line in lines:
        line = line.strip()
        # 伝票番号・No.・TEL・登録番号の後は除外
        if re.search(r'伝票|BILL NO|No\.|TEL|登録番号|Invoice|Receipt number', line, re.IGNORECASE):
            skip_next = True
            continue
        if skip_next:
            skip_next = False
            continue
        m = re.fullmatch(r'[\d,]{4,6}', line)  # 4〜6桁限定（7桁以上は除外）
        if m:
            try:
                val = int(m.group().replace(',', ''))
                if 500 <= val <= 999_999:
                    standalone.append(val)
            except ValueError:
                pass
    if standalone:
        return float(max(standalone)), "JPY"

    return 0.0, "JPY"


def extract_amount(text: str) -> int:
    """後方互換用: 金額のみ返す（円換算済み）"""
    amount, currency = extract_amount_and_currency(text)
    if currency != "JPY" and amount > 0:
        jpy, _, _ = convert_to_jpy(amount, currency)
        return jpy
    return int(amount)


def extract_vendor(text: str, filename: str = "") -> str:
    t = normalize(text)
    lines = [l.strip() for l in t.split('\n') if l.strip()]

    # 法人名パターン（文書の最初のほう）
    for line in lines[:20]:
        for pat in [
            r'株式会社\s*[\w\s・ー－]{1,20}',
            r'[\w\s・ー－]{1,20}\s*株式会社',
            r'[\w\s・ー－]{1,15}(?:有限会社|合同会社|一般社団法人|NPO法人|公益財団法人)',
        ]:
            m = re.search(pat, line)
            if m:
                name = m.group().strip()
                if 3 < len(name) < 30:
                    return name

    # 既知の主要ベンダー名
    known = [
        "ANA", "JAL", "AIRDO", "Amazon", "OpenAI", "Anthropic",
        "STUDIO", "1Password", "Adobe", "Vercel", "ENEOS", "ホクレン",
        "NexusAd", "楽天", "ケーズデンキ", "根室交通", "道東電子",
        "Microsoft", "Apple", "Google", "Netflix", "Spotify",
    ]
    t_lower = t.lower()
    for v in known:
        if v.lower() in t_lower:
            return v

    # 領収書・支払証明の「上様」「宛名」の次の行が取引先の場合
    for i, line in enumerate(lines[:10]):
        if any(kw in line for kw in ["領収書", "領収証", "Receipt", "RECEIPT"]):
            # 発行元は末尾のほうにある場合が多い
            break

    # ファイル名から推測（最終手段）
    if filename:
        name = os.path.splitext(os.path.basename(filename))[0]
        name = re.sub(r'^[\d_\-]+', '', name).strip()  # 先頭の日付除去
        if name:
            return name

    return "不明"


def extract_memo(text: str, vendor: str = "", kamoku: str = "") -> str:
    """
    摘要（取引内容の説明）を抽出
    優先順位:
    1. 明示的な「摘要」「品名」「商品名」「内容」ラベルの後
    2. 宿泊代・ガソリン代など取引内容を示すキーワードを含む行
    3. 品名リストの最初の1〜2行
    4. ベンダー名 + 勘定科目でフォールバック
    """
    t = normalize(text)
    lines = [l.strip() for l in t.split('\n') if l.strip()]

    # 1. 明示的ラベルの後の内容
    for i, line in enumerate(lines):
        m = re.match(r'^(?:摘要|品名|商品名|内容|サービス内容|件名|商品|品目)[:：\s]*(.*)', line)
        if m:
            rest = m.group(1).strip()
            if rest and len(rest) > 1 and not re.match(r'^[\d¥￥,]+$', rest):
                return rest[:40]
            # 次の行
            if i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                if nxt and len(nxt) > 1 and not re.match(r'^[\d¥￥,.\s]+$', nxt):
                    return nxt[:40]

    # 2. 取引内容を示すキーワードを含む行
    CONTENT_KWS = [
        "宿泊", "ガソリン", "給油", "燃料", "消耗品", "印刷", "郵便",
        "通信", "交通費", "飲食", "食事", "文具", "事務用品", "備品",
        "修繕", "研修", "セミナー", "広告", "委託", "保険",
    ]
    for kw in CONTENT_KWS:
        if kw in t:
            for line in lines:
                if kw in line and 2 < len(line) < 40:
                    return line[:40]

    # 3. 品名リスト：数字や¥を含まない短めの行（上から）
    skipped = {"領収書", "領収証", "Receipt", "RECEIPT", "御中", "様", "合計", "小計",
               "税込", "税抜", "消費税", "Thank", "ありがとう"}
    candidates = []
    for line in lines[2:20]:  # 先頭数行はヘッダー寄りなのでスキップ
        if any(s in line for s in skipped):
            continue
        if re.match(r'^[\d¥￥,.\-/\s]+$', line):  # 数字だけの行除外
            continue
        if 2 < len(line) < 30:
            candidates.append(line)
    if candidates:
        return candidates[0][:40]

    # 4. フォールバック: ベンダー＋勘定科目
    if vendor and vendor != "不明":
        return f"{vendor}　{kamoku}" if kamoku else vendor
    return kamoku or ""


def detect_tax_rate(text: str) -> str:
    t = normalize(text)
    if any(x in t for x in ["8%", "８%", "軽減税率"]):
        return "8%"
    return "10%"


def _guess_date_from_filename(filename: str) -> str:
    m = re.search(r'(20\d{2})(\d{2})(\d{2})', filename)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return ""


# ===== OCRエンジン =====

def ocr_apple_vision(img_path: str) -> str:
    """
    macOS Vision Framework でOCR（完全無料・日本語高精度）
    macOS 13以降で精度が大幅向上
    """
    try:
        import Vision
        from Foundation import NSURL

        input_url = NSURL.fileURLWithPath_(img_path)
        handler = Vision.VNImageRequestHandler.alloc().initWithURL_options_(input_url, {})
        request = Vision.VNRecognizeTextRequest.alloc().init()
        request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
        request.setRecognitionLanguages_(["ja-JP", "en-US"])
        request.setUsesLanguageCorrection_(True)

        handler.performRequests_error_([request], None)

        lines = []
        for obs in (request.results() or []):
            candidates = obs.topCandidates_(1)
            if candidates:
                lines.append(str(candidates[0].string()))

        return "\n".join(lines)
    except Exception:
        return ""


def ocr_tesseract(img_path: str) -> str:
    """tesseract OCR（インストール済みの場合のみ）"""
    try:
        import pytesseract
        from PIL import Image, ImageEnhance, ImageFilter
        img = Image.open(img_path).convert('L')
        img = ImageEnhance.Contrast(img).enhance(2.0)
        img = img.filter(ImageFilter.SHARPEN)
        return pytesseract.image_to_string(img, lang='jpn+eng', config='--psm 3 --oem 3')
    except Exception:
        return ""


def run_ocr(img_path: str) -> tuple:
    """
    利用可能なOCRエンジンでテキスト抽出
    Returns: (text, engine_name)
    """
    # 1. Apple Vision（最優先：無料・高精度）
    text = ocr_apple_vision(img_path)
    if text.strip():
        return text, "Apple Vision"

    # 2. tesseract（インストール済みの場合）
    text = ocr_tesseract(img_path)
    if text.strip():
        return text, "tesseract"

    return "", "なし"


# ===== PDF / 画像の変換 =====

def pdf_to_image_bytes(filepath: str, zoom: float = 2.0) -> bytes:
    """PDFの1ページ目をJPEG画像バイトに変換"""
    try:
        import fitz
        doc = fitz.open(filepath)
        page = doc[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        return pix.tobytes("jpeg")
    except Exception:
        return None


def pdf_to_image_bytes_all_pages(filepath: str, zoom: float = 2.0) -> list:
    """PDFの全ページをJPEG画像バイトのリストにして返す。失敗時は空リスト。"""
    try:
        import fitz
        doc = fitz.open(filepath)
        out = []
        for page in doc:
            try:
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                out.append(pix.tobytes("jpeg"))
            except Exception:
                continue
        return out
    except Exception:
        return []


def pdf_page_count(filepath: str) -> int:
    """PDFのページ数を返す。失敗時は0。"""
    try:
        import fitz
        return len(fitz.open(filepath))
    except Exception:
        return 0


def image_to_jpeg_bytes(filepath: str) -> bytes:
    """画像ファイルをJPEGバイト列に変換（EXIFの回転情報を反映して正立させる）"""
    try:
        from PIL import Image as PILImage, ImageOps
        import io
        img = PILImage.open(filepath)
        # iPhone等の写真はEXIFに回転情報を持つ。これを実ピクセルに適用して正立させる
        # （これをしないと横向き・逆さまのままAIに渡り、読み取りが乱れる）
        try:
            img = ImageOps.exif_transpose(img)
        except Exception:
            pass
        if img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        return buf.getvalue()
    except Exception:
        return None


# ===== 画像の向き自動補正（回転スキャン・横向き写真対策） =====

def _rotate_jpeg_bytes(img_bytes: bytes, deg: int) -> bytes:
    """画像を deg 度(PIL基準=反時計回り正)回転してJPEGバイトで返す。"""
    if not img_bytes or not deg:
        return img_bytes
    try:
        from PIL import Image as _PIL
        import io as _io
        img = _PIL.open(_io.BytesIO(img_bytes))
        out = img.rotate(deg, expand=True).convert("RGB")
        buf = _io.BytesIO()
        out.save(buf, format="JPEG", quality=90)
        return buf.getvalue()
    except Exception:
        return img_bytes


def _parse_upright_index(raw: str):
    """向き判定AIの応答から正立画像の番号(0-3)を取り出す。失敗時None。"""
    if not raw:
        return None
    import json as _json
    cleaned = re.sub(r'^```[a-z]*\n?|```$', '', raw.strip(), flags=re.MULTILINE).strip()
    m = re.search(r'\{.*\}', cleaned, re.DOTALL)
    if m:
        try:
            v = _json.loads(m.group()).get("upright")
            if isinstance(v, (int, float)):
                return int(v)
        except Exception:
            pass
    m2 = re.search(r'[0-3]', cleaned)
    return int(m2.group()) if m2 else None


def _vision_multi_image_raw(thumbs_b64: list, prompt: str, api_key: str,
                            provider: str, max_tokens: int = 40,
                            timeout: int = 30) -> str:
    """複数画像＋プロンプトをVision AIに送り、生のテキスト応答を返す。"""
    import json as _json
    import urllib.request as _req
    if provider == "claude":
        content = []
        for i, b in enumerate(thumbs_b64):
            content.append({"type": "text", "text": f"画像{i}:"})
            content.append({"type": "image", "source": {
                "type": "base64", "media_type": "image/jpeg", "data": b}})
        content.append({"type": "text", "text": prompt})
        payload = {"model": CLAUDE_MODEL, "max_tokens": max_tokens,
                   "messages": [{"role": "user", "content": content}]}
        req = _req.Request("https://api.anthropic.com/v1/messages",
                           data=_json.dumps(payload).encode(),
                           headers={"x-api-key": api_key,
                                    "anthropic-version": "2023-06-01",
                                    "content-type": "application/json"})
        with _req.urlopen(req, timeout=timeout) as r:
            return _json.loads(r.read())["content"][0]["text"]
    else:
        parts = []
        for i, b in enumerate(thumbs_b64):
            parts.append({"text": f"画像{i}:"})
            parts.append({"inline_data": {"mime_type": "image/jpeg", "data": b}})
        parts.append({"text": prompt})
        payload = {"contents": [{"parts": parts}],
                   "generationConfig": {"maxOutputTokens": max_tokens,
                                        "temperature": 0.0}}
        url = (f"https://generativelanguage.googleapis.com/v1beta/models/"
               f"{GEMINI_MODEL}:generateContent?key={api_key}")
        req = _req.Request(url, data=_json.dumps(payload).encode(),
                           headers={"content-type": "application/json"})
        with _req.urlopen(req, timeout=timeout) as r:
            res = _json.loads(r.read())
            return res["candidates"][0]["content"]["parts"][0]["text"]


def detect_upright_rotation(img_bytes: bytes, api_key: str,
                            provider: str = "claude") -> int:
    """
    Vision AIに4方向の縮小画像を見せ、文字が正立して最も読める向きを選ばせる。

    Visionモデルは「回転角の計算」は苦手だが「どれが読めるか」の判断は得意なため、
    回転候補(0/90/180/270)を見せて選択させる方式にしている。
    戻り値: 元画像に適用すべき回転角(PIL基準=反時計回り正, 0/90/180/270)。
            APIキー無し・判定失敗時は 0（回転しない）。
    """
    if not (img_bytes and api_key):
        return 0
    try:
        from PIL import Image as _PIL
        import io as _io
        import base64 as _b64
        img = _PIL.open(_io.BytesIO(img_bytes))
        angles = (0, 90, 180, 270)
        thumbs = []
        for a in angles:
            t = img.rotate(a, expand=True)
            t.thumbnail((560, 560))
            buf = _io.BytesIO()
            t.convert("RGB").save(buf, format="JPEG", quality=75)
            thumbs.append(_b64.b64encode(buf.getvalue()).decode())
        prompt = (
            "以下の画像0〜3は同じ領収書・レシートを回転させたものです。"
            "文字が正立して最も正しく読める画像はどれですか。"
            "番号だけをJSONで返してください。前置きや説明は不要。"
            '例: {"upright": 3}'
        )
        raw = _vision_multi_image_raw(thumbs, prompt, api_key, provider)
        idx = _parse_upright_index(raw)
        if idx is None or not (0 <= idx < 4):
            return 0
        return angles[idx]
    except Exception:
        return 0


def _is_garbled_text(text: str) -> bool:
    """
    スキャナが埋め込んだOCRテキスト層が文字化けゴミかどうかを判定。

    日本語の領収書テキストは必ずある程度ひらがなを含む（「です」「として」
    「円を含みます」「お預り」「釣銭」等）。一方スキャナOCRの文字化けは
    ランダムな漢字ばかりでひらがながほとんど無い、という特徴がある。
    → 漢字が十分あるのにひらがな比率が極端に低ければゴミとみなす。
    """
    s = [c for c in text if not c.isspace()]
    if len(s) < 20:
        return False
    hira  = re.findall(r'[\u3041-\u3093]', text)
    kanji = re.findall(r'[\u4e00-\u9fff]', text)
    return len(kanji) >= 25 and (len(hira) / len(s)) < 0.07



# ===== メイン抽出関数 =====

def extract_from_file(filepath: str, filename: str = None,
                      ai_api_key: str = "", ai_provider: str = "gemini") -> dict:
    """
    ファイルから領収書データを抽出

    ai_api_key:  APIキー（指定時はAIで高精度抽出）
    ai_provider: "gemini"（無料）or "claude"（有料・高精度）
    Returns: dict with keys:
        date, vendor, amount, memo, tax_rate, kamoku, jigyo, warning, _ocr_engine
    """
    ext = os.path.splitext(filepath)[1].lower()
    filename = filename or os.path.basename(filepath)

    text = ""
    ocr_engine = "なし"
    warning = ""
    vision_img_bytes = None  # Gemini Visionに渡す画像

    # ===== ファイルを画像化（Gemini Vision / OCR 共用） =====
    if ext == '.pdf':
        # スキャンPDFは文字が小さいので高解像度で描画（読み取り精度向上）
        vision_img_bytes = pdf_to_image_bytes(filepath, zoom=3.0)
        # テキストPDFならpdfplumberでも取得
        try:
            import pdfplumber
            with pdfplumber.open(filepath) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        text += t + "\n"
            text = text.strip()
            # スキャナが埋め込んだOCR層が文字化けゴミの場合は捨てる
            # （ゴミテキストを使うと誤抽出するため、画像OCRに切り替える）
            if text and _is_garbled_text(text):
                text = ""
            if text:
                ocr_engine = "pdfplumber"
        except Exception:
            pass
    elif ext in ['.jpg', '.jpeg', '.png', '.heic', '.heif', '.bmp', '.tiff']:
        vision_img_bytes = image_to_jpeg_bytes(filepath)

    # ===== 文字の向きを自動補正（回転スキャンPDF・横向き写真対策） =====
    # AIに渡る画像が横向き/逆さまだと誤読み取り（ハルシネーション）が起きるため、
    # Vision AI自身に4方向の縮小画像を見せて正立向きを判定させ、回転してから本抽出に渡す。
    # デジタルPDF（pdfplumberでクリーンな水平テキストが取れている）は正立確実なので
    # 向き判定の余計なAPI呼び出しをスキップする。
    _is_digital_pdf = bool(text)
    _orient_deg = 0  # 適用した回転角（複数ページPDFの2ページ目以降にも同じ角度を適用するため保持）
    if vision_img_bytes and ai_api_key and not _is_digital_pdf:
        try:
            _rot = detect_upright_rotation(vision_img_bytes, ai_api_key, ai_provider)
            if _rot:
                vision_img_bytes = _rotate_jpeg_bytes(vision_img_bytes, _rot)
                _orient_deg = _rot
        except Exception:
            pass

    # ===== Vision AIで直接読み取り（APIキーあり・最高精度） =====
    ai_result = {}
    ai_usage = None
    ai_error = ""
    if ai_api_key and vision_img_bytes:
        try:
            _r = extract_with_ai("", ai_api_key, provider=ai_provider,
                                 img_bytes=vision_img_bytes)
            ai_result = _r.get("data", {}) or {}
            ai_usage  = _r.get("usage")
            label = "Claude Vision" if ai_provider == "claude" else "Gemini Vision"
            ocr_engine = label
        except Exception as e:
            ai_error = str(e)[:120]
            ai_result = {}

    # Vision AIが使えなかった場合はOCR→テキストAI or ルールベース
    if not ai_result:
        # OCRでテキスト取得（まだ取れていない場合）
        if not text and vision_img_bytes:
            with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as tmp:
                tmp.write(vision_img_bytes)
                tmp_path = tmp.name
            try:
                text, ocr_engine = run_ocr(tmp_path)
            finally:
                os.unlink(tmp_path)

        # テキストベースAI
        if ai_api_key and text:
            try:
                _r = extract_with_ai(text, ai_api_key, provider=ai_provider)
                ai_result = _r.get("data", {}) or {}
                ai_usage  = _r.get("usage")
                label = "Claude AI" if ai_provider == "claude" else "Gemini AI"
                ocr_engine += f" + {label}"
            except Exception as e:
                ai_error = str(e)[:120]
                ai_result = {}

    # テキストもAI結果もない場合
    if not text and not ai_result:
        warning = "テキストを読み取れませんでした（手動で入力してください）"
        return {
            "date": _guess_date_from_filename(filename),
            "vendor": re.sub(r'^[\d_\-]+', '', os.path.splitext(filename)[0]).strip() or filename,
            "amount": 0,
            "memo": "",
            "tax_rate": "10%",
            "kamoku": "消耗品",
            "jigyo": "ミッション活動",
            "warning": warning,
            "_ocr_engine": ocr_engine,
        }

    # 各項目を抽出（AI結果を優先、なければルールベース）
    vendor = str(ai_result.get("vendor", "")).strip() or extract_vendor(text, filename)
    date   = str(ai_result.get("date", "")).strip()   or extract_date(text)
    memo   = str(ai_result.get("memo", "")).strip()   or extract_memo(text, vendor, "")

    # 勘定科目: AI判定を最優先、選択肢に無い値ならルールベースで補完
    ai_kamoku = str(ai_result.get("kamoku", "")).strip()
    if ai_kamoku in KAMOKU_OPTIONS:
        kamoku = ai_kamoku
    else:
        kamoku = infer_kamoku(vendor, text)

    # 金額: AI結果優先、外貨チェックはルールベース
    raw_amount, currency = extract_amount_and_currency(text) if text else (0.0, "JPY")
    ai_amount = ai_result.get("amount", 0)
    fx_info = ""
    if currency != "JPY" and raw_amount > 0:
        jpy_amount, rate, cur = convert_to_jpy(raw_amount, currency, date)
        amount = jpy_amount
        fx_info = f"{cur} {raw_amount:.2f} → ¥{jpy_amount:,}（レート: {rate:.2f}円/{cur}）"
    elif ai_amount and int(ai_amount) > 0:
        amount = int(ai_amount)
    else:
        amount = int(raw_amount)

    # 警告組み立て
    warns = []
    if not date:
        date = _guess_date_from_filename(filename)
        warns.append("日付を読み取れませんでした" + (f"（ファイル名から: {date}）" if date else ""))
    if amount == 0:
        warns.append("金額を読み取れませんでした")

    warning = " / ".join(warns)

    return {
        "date":        date,
        "vendor":      vendor,
        "amount":      amount,
        "memo":        memo,
        "tax_rate":    detect_tax_rate(text),
        "kamoku":      kamoku,
        "jigyo":       "ミッション活動",
        "warning":     warning,
        "_ocr_engine": ocr_engine,
        "_raw_text":   text,
        "_fx_info":    fx_info,
        "_currency":   currency,
        "_ai_error":   ai_error,  # AIエラー詳細（診断用）
        "_ai_usage":   ai_usage,  # {provider, model, input_tokens, output_tokens} or None
        "_orient_deg": _orient_deg,  # 1ページ目に適用した回転角（複数ページPDFの2ページ目以降用）
    }
