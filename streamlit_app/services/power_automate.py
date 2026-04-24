"""Power Automate Webhook 発火サービス。

ペイロードJSON の power_automate.webhook_key で参照される URL を secrets から引き、
power_automate.payload を JSON ボディとして POST する。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import requests


@dataclass
class PowerAutomateResult:
    ok: bool
    message: str
    response_text: str = ""


def fire(webhook_url: str, payload: dict[str, Any]) -> PowerAutomateResult:
    if not webhook_url:
        return PowerAutomateResult(False, "webhook URL が空です")
    try:
        r = requests.post(webhook_url, json=payload or {}, timeout=15)
        if r.ok:
            return PowerAutomateResult(True, f"発火成功 (HTTP {r.status_code})", r.text[:500])
        return PowerAutomateResult(False, f"HTTP {r.status_code}: {r.text[:200]}")
    except Exception as e:
        return PowerAutomateResult(False, f"エラー: {e}")
