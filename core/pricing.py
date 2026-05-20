"""
pricing.py - AI APIのトークン単価とコスト計算

価格は2026年時点の公開情報を元にした概算。実勢と異なる場合は更新する。
単位: USD / 1Mトークン
"""

# プロバイダ・モデル別の単価（USD / 1M tokens）
PRICING = {
    "claude": {
        "claude-haiku-4-5":  {"input": 1.0,  "output": 5.0},
        "claude-3-5-haiku":  {"input": 0.8,  "output": 4.0},
        "claude-3-5-sonnet": {"input": 3.0,  "output": 15.0},
        "claude-sonnet-4-6": {"input": 3.0,  "output": 15.0},
        "claude-opus-4-7":   {"input": 15.0, "output": 75.0},
    },
    "gemini": {
        # Gemini Flashは現状無料枠で運用想定 → コスト0で記録
        "gemini-2.0-flash":     {"input": 0.0, "output": 0.0},
        "gemini-1.5-flash":     {"input": 0.0, "output": 0.0},
        "gemini-2.5-flash":     {"input": 0.0, "output": 0.0},
    },
}

# 為替: 円換算用の固定レート（管理画面で上書きするまでのフォールバック）
DEFAULT_USD_JPY = 150.0


def get_unit_price(provider: str, model: str) -> dict:
    """{'input': USD/MTok, 'output': USD/MTok} を返す。未知モデルは0"""
    p = (PRICING.get(provider) or {}).get(model)
    if p:
        return p
    # 未知モデルでもプロバイダのデフォルト的な単価で概算
    if provider == "claude":
        return {"input": 1.0, "output": 5.0}
    return {"input": 0.0, "output": 0.0}


def calc_cost(provider: str, model: str,
              input_tokens: int, output_tokens: int,
              usd_jpy: float = DEFAULT_USD_JPY) -> tuple:
    """
    トークン数から (USD, JPY) を計算
    """
    p = get_unit_price(provider, model)
    usd = (input_tokens / 1_000_000.0) * p["input"] + \
          (output_tokens / 1_000_000.0) * p["output"]
    jpy = usd * usd_jpy
    return round(usd, 6), round(jpy, 4)
