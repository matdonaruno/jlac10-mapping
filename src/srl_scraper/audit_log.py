"""監査ログ

全操作を記録する。JSON ファイルに追記。
"""

import json
import logging
from datetime import datetime, timezone
from pathlib import Path

logger = logging.getLogger(__name__)

AUDIT_LOG_FILE = "audit_log.json"


def audit_log_path(data_dir: Path) -> Path:
    return data_dir / AUDIT_LOG_FILE


def audit_load(data_dir: Path) -> list[dict]:
    """監査ログ読み込み"""
    path = audit_log_path(data_dir)
    if not path.exists():
        return []
    return json.loads(path.read_text(encoding="utf-8"))


def audit_add(data_dir: Path, action: str, detail: dict,
              user: str = "", issue: str = "", hospital: str = "") -> dict:
    """監査ログに1件追加

    action: "ssmix_parse" | "error_detect" | "mapping" | "sop_check" |
            "delivery_export" | "mail_sent" | "db_apply" | "custom_rule_add"
    detail: アクション固有の詳細情報

    Returns: 追加されたログエントリ
    """
    logs = audit_load(data_dir)
    entry = {
        "id": "A" + str(len(logs) + 1).zfill(6),
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "user": user,
        "action": action,
        "issue": issue,
        "hospital": hospital,
        "detail": detail,
    }
    logs.append(entry)

    path = audit_log_path(data_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(logs, ensure_ascii=False, indent=2), encoding="utf-8-sig")

    logger.info("監査ログ追加: [%s] %s - %s", entry["id"], action,
                json.dumps(detail, ensure_ascii=False)[:80])
    return entry


def audit_search(data_dir: Path, action: str = None, issue: str = None,
                 hospital: str = None, limit: int = 50) -> list[dict]:
    """監査ログ検索"""
    logs = audit_load(data_dir)
    results = []
    for log in reversed(logs):  # 新しい順
        if action and log.get("action") != action:
            continue
        if issue and log.get("issue") != issue:
            continue
        if hospital and log.get("hospital") != hospital:
            continue
        results.append(log)
        if len(results) >= limit:
            break
    return results
