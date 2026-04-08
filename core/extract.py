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
この領収書・レシートから4つの情報を読み取り、必ずJSON形式のみで返してください。

【重要】返答はJSON1行のみ。前置き・説明・マークダウン・コードブロック（```）は一切不要。

返答例:
{"vendor": "ホクレン 根室SS", "memo": "ガソリン給油", "date": "2026-03-15", "amount": 8540}

各フィールドの読み取り方:
- vendor: 領収書・レシートを発行したお店や会社の名前（屋号・店名・会社名）
  例: "コープさっぽろ", "ENEOS 根室SS", "根室交通株式会社", "Amazon.co.jp"
  ※ 「上様」「御中」「様」は取引先ではないので含めない
  ※ 宛名（「〇〇様」）ではなく発行元の名前を読む

- memo: 何を購入・利用したか（品目・サービス内容を簡潔に）
  例: "ガソリン給油", "事務用品購入", "宿泊料", "バス運賃", "ソフトウェアサブスク"
  ※ 品目が多い場合は代表的なもの or「消耗品購入」等でまとめる（30文字以内）

- date: 領収日・購入日をYYYY-MM-DD形式で
  例: 令和7年3月15日 → "2026-03-15"、2026/03/15 → "2026-03-15"
  ※ 読み取れない場合は空文字 ""

- amount: 税込の合計金額（整数・円記号なし・カンマなし）
  例: ¥8,540 → 8540、合計 19,650円 → 19650
  ※ 税抜・小計ではなく「税込合計」「お支払合計」の金額を使う
  ※ 読み取れない場合は 0"""


def _parse_ai_json(raw: str) -> dict:
    """AI応答からJSONを抽出してパース"""
    raw = re.sub(r'^```[a-z]*\n?|```$', '', raw.strip(), flags=re.MULTILINE).strip()
    m = re.search(r'\{.*\}', raw, re.DOTALL)
    if m:
        raw = m.group()
    import json as _json
    return _json.loads(raw)


def _gemini_api_call(payload_dict: dict, api_key: str, timeout: int = 30) -> dict:
    """
    Gemini API共通呼び出し（429レート制限時は自動リトライ）
    無料枠: 15回/分 → 429が出たら最大3回、間隔を空けてリトライ
    """
    import json as _json
    import urllib.request as _req
    import urllib.error as _err
    import time as _time

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
    payload = _json.dumps(payload_dict).encode()

    req = _req.Request(url, data=payload, headers={"content-type": "application/json"})
    with _req.urlopen(req, timeout=timeout) as r:
        result = _json.loads(r.read())
        raw = result["candidates"][0]["content"]["parts"][0]["text"]
        return _parse_ai_json(raw)


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


def _extract_with_claude(text: str, api_key: str) -> dict:
    """
    Anthropic Claude APIで抽出（有料・高精度）
    """
    import json as _json
    import urllib.request as _req

    prompt = _AI_PROMPT.replace("この領収書・レシートの画像から", "以下のOCRテキストから") + f"\n\nOCRテキスト:\n{text[:3000]}"
    payload = _json.dumps({
        "model": "claude-3-5-haiku-20241022",
        "max_tokens": 512,
        "messages": [{"role": "user", "content": prompt}],
    }).encode()

    req = _req.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }
    )
    with _req.urlopen(req, timeout=20) as r:
        result = _json.loads(r.read())
        raw = result["content"][0]["text"]
        return _parse_ai_json(raw)


def extract_with_ai(text: str, api_key: str, provider: str = "gemini",
                    img_bytes: bytes = None) -> dict:
    """
    AIで領収書から構造化データを抽出。
    img_bytes があれば Vision APIで画像を直接解析（最高精度）。
    provider: "gemini"（デフォルト・無料） or "claude"（有料・高精度）
    """
    if provider == "claude":
        return _extract_with_claude(text, api_key)
    elif img_bytes:
        # Gemini Vision: 画像を直接送信（OCR不要・最高精度）
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


def image_to_jpeg_bytes(filepath: str) -> bytes:
    """画像ファイルをJPEGバイト列に変換"""
    try:
        from PIL import Image as PILImage
        import io
        img = PILImage.open(filepath)
        if img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        return buf.getvalue()
    except Exception:
        return None


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
        vision_img_bytes = pdf_to_image_bytes(filepath, zoom=2.0)
        # テキストPDFならpdfplumberでも取得
        try:
            import pdfplumber
            with pdfplumber.open(filepath) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        text += t + "\n"
            text = text.strip()
            if text:
                ocr_engine = "pdfplumber"
        except Exception:
            pass
    elif ext in ['.jpg', '.jpeg', '.png', '.heic', '.bmp', '.tiff']:
        vision_img_bytes = image_to_jpeg_bytes(filepath)

    # ===== Gemini Visionで直接読み取り（APIキーあり・最高精度） =====
    ai_result = {}
    ai_error = ""
    if ai_api_key and ai_provider == "gemini" and vision_img_bytes:
        try:
            ai_result = extract_with_ai("", ai_api_key, provider="gemini",
                                        img_bytes=vision_img_bytes)
            ocr_engine = "Gemini Vision"
        except Exception as e:
            ai_error = str(e)[:120]
            ai_result = {}

    # Gemini Visionが使えなかった場合はOCR→テキストAI or ルールベース
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

        # Claude or テキストAI
        if ai_api_key and text:
            try:
                ai_result = extract_with_ai(text, ai_api_key, provider=ai_provider)
                label = "Claude AI" if ai_provider == "claude" else "Gemini AI"
                ocr_engine += f" + {label}"
            except Exception:
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
    }
