"""3社のデータをJLAC10コードをキーに統合し、コード名称を付与する"""

import json
import logging
from datetime import datetime, timezone
from pathlib import Path

from .jslm import build_lookup, decode_jlac10

logger = logging.getLogger(__name__)


def load_latest(output_dir: Path, prefix: str) -> dict | None:
    """latest シンボリックリンクからJSONを読み込む"""
    latest = output_dir / f"{prefix}_tests_latest.json"
    if not latest.exists():
        logger.warning("%s が見つかりません", latest)
        return None
    data = json.loads(latest.read_text(encoding="utf-8"))
    return data


def load_jlac10_lookup(output_dir: Path) -> dict:
    """JLAC10 検索用辞書を読み込む"""
    path = output_dir / "jlac10_lookup.json"
    if not path.exists():
        logger.warning("jlac10_lookup.json が見つかりません（jslm コマンドを先に実行してください）")
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def merge_all(output_dir: Path | None = None) -> dict:
    """SRL・BML・LSI の最新JSONを統合し、JLAC10コードを分解・名称付与"""
    output_dir = output_dir or Path("data")

    sources = {
        "srl": load_latest(output_dir, "srl"),
        "bml": load_latest(output_dir, "bml"),
        "lsi": load_latest(output_dir, "lsi"),
    }

    lookup = load_jlac10_lookup(output_dir)

    # JLAC10をキーに統合
    merged: dict[str, dict] = {}

    for source_key, data in sources.items():
        if data is None:
            continue
        for item in data["items"]:
            jlac10 = item.get("jlac10", "")
            if not jlac10:
                continue

            if jlac10 not in merged:
                # JLAC10 コードを分解して名称付与
                decoded = decode_jlac10(jlac10, lookup) if lookup else {"raw": jlac10, "valid": False}
                analyte_code = jlac10[:5] if len(jlac10) >= 5 else jlac10

                merged[jlac10] = {
                    "jlac10": jlac10,
                    "analyte_code": analyte_code,
                    "jlac10_decoded": decoded,
                    "sources": {},
                }

            merged[jlac10]["sources"][source_key] = {
                "item_name": item.get("item_name", ""),
                "material": item.get("material", ""),
                "method": item.get("method", {}).get("name", "") if isinstance(item.get("method"), dict) else item.get("method", ""),
                "reference_value": item.get("reference_value", ""),
                "detail_url": item.get("detail_url", ""),
            }

    # 統合結果
    now = datetime.now(timezone.utc)
    available = {k: v["metadata"]["total_items"] for k, v in sources.items() if v}

    result = {
        "metadata": {
            "merged_at": now.isoformat(),
            "sources_available": available,
            "total_unique_jlac10": len(merged),
            "jlac10_master_available": bool(lookup),
            "by_source_count": {
                "srl_only": sum(1 for v in merged.values() if set(v["sources"]) == {"srl"}),
                "bml_only": sum(1 for v in merged.values() if set(v["sources"]) == {"bml"}),
                "lsi_only": sum(1 for v in merged.values() if set(v["sources"]) == {"lsi"}),
                "all_three": sum(1 for v in merged.values() if len(v["sources"]) == 3),
                "two_sources": sum(1 for v in merged.values() if len(v["sources"]) == 2),
            },
        },
        "items": sorted(merged.values(), key=lambda x: x["jlac10"]),
    }

    filepath = output_dir / "merged_jlac10.json"
    filepath.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("統合完了: %s (%d件)", filepath, len(merged))

    return result
