"""3社のデータをJLAC10コードをキーに統合し、コード名称を付与する"""

import json
import logging
import re
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

    # JLAC10をキーに統合（valid_15 / valid_17 のみ）
    merged: dict[str, dict] = {}
    # JLAC10なし項目（empty / invalid）
    items_no_jlac: list[dict] = []

    for source_key, data in sources.items():
        if data is None:
            continue
        for item in data["items"]:
            jlac10 = item.get("jlac10", "")
            jlac10_status = item.get("jlac10_status", "")

            # jlac10_status が未設定の場合（後方互換）、コードから判定
            if not jlac10_status:
                if not jlac10:
                    jlac10_status = "empty"
                elif re.match(r"^[0-9A-Za-z]{15}$", jlac10):
                    jlac10_status = "valid_15"
                elif re.match(r"^[0-9A-Za-z]{16,17}$", jlac10):
                    jlac10_status = "valid_17"
                else:
                    jlac10_status = "invalid"

            # empty / invalid → JLAC10なしリストへ
            if jlac10_status in ("empty", "invalid"):
                items_no_jlac.append({
                    "item_name": item.get("item_name", ""),
                    "source": source_key,
                    "detail_url": item.get("detail_url", ""),
                    "jlac10_raw": jlac10,
                    "jlac10_status": jlac10_status,
                })
                continue

            if jlac10 not in merged:
                # JLAC10 コードを分解して名称付与
                decoded = decode_jlac10(jlac10, lookup) if lookup else {"raw": jlac10, "valid": False}
                analyte_code = jlac10[:5] if len(jlac10) >= 5 else jlac10

                merged[jlac10] = {
                    "jlac10": jlac10,
                    "jlac10_status": jlac10_status,
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
            "total_items_no_jlac": len(items_no_jlac),
            "items_no_jlac_by_status": {
                "empty": sum(1 for i in items_no_jlac if i["jlac10_status"] == "empty"),
                "invalid": sum(1 for i in items_no_jlac if i["jlac10_status"] == "invalid"),
            },
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
        "items_no_jlac": items_no_jlac,
    }

    filepath = output_dir / "merged_jlac10.json"
    filepath.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("統合完了: %s (%d件)", filepath, len(merged))

    return result


def apply_mapping_results(
    merged_path: Path,
    mapping_items: list[dict],
    hospital: str = "",
    confirmed_only: bool = True,
) -> dict:
    """確定済みマッピング結果をmerged_jlac10.jsonに追加

    Args:
        merged_path: merged_jlac10.json のパス
        mapping_items: マッピング結果リスト。各要素に
            item_name, jlac10, status(auto/confirmed/candidate/manual),
            matched_name 等を含む
        hospital: 病院名（匿名化推奨）
        confirmed_only: Trueの場合 auto + confirmed のみ適用

    Returns:
        {"added": N, "skipped": M, "new_entries": K}
    """
    logger.debug(
        "apply_mapping_results開始: merged_path=%s, 入力%d件, confirmed_only=%s",
        merged_path.name, len(mapping_items), confirmed_only,
    )
    data = json.loads(merged_path.read_text(encoding="utf-8"))
    items_by_jlac = {it["jlac10"]: it for it in data.get("items", [])}
    logger.debug("既存エントリ数: %d", len(items_by_jlac))

    lookup_path = merged_path.parent / "jlac10_lookup.json"
    lookup = {}
    if lookup_path.exists():
        lookup = json.loads(lookup_path.read_text(encoding="utf-8"))

    added = 0
    skipped = 0
    new_entries = 0

    for item in mapping_items:
        status = item.get("status", "")
        if confirmed_only and status not in ("auto", "confirmed"):
            logger.debug("スキップ (status=%s): '%s'", status, item.get("item_name", ""))
            skipped += 1
            continue

        jlac10 = item.get("jlac10", "").replace("-", "")
        item_name = item.get("item_name", "")
        matched_name = item.get("matched_name", "")

        if not jlac10 or not re.match(r"^[0-9A-Za-z]{15,17}$", jlac10):
            logger.debug("スキップ (JLAC10不正='%s'): '%s'", jlac10, item_name)
            skipped += 1
            continue

        if jlac10 in items_by_jlac:
            logger.debug("既存エントリに追加: jlac10=%s, item='%s'", jlac10, item_name)
            entry = items_by_jlac[jlac10]
        else:
            decoded = decode_jlac10(jlac10, lookup) if lookup else {"raw": jlac10, "valid": False}
            entry = {
                "jlac10": jlac10,
                "jlac10_status": "valid_15" if len(jlac10) == 15 else "valid_17",
                "analyte_code": jlac10[:5],
                "jlac10_decoded": decoded,
                "sources": {},
                "mapping_history": [],
            }
            items_by_jlac[jlac10] = entry
            data["items"].append(entry)
            new_entries += 1
            logger.debug("新規エントリ作成: jlac10=%s, item='%s'", jlac10, item_name)

        if "mapping_history" not in entry:
            entry["mapping_history"] = []

        dup = any(
            h.get("hospital") == hospital and h.get("item_name") == item_name
            for h in entry["mapping_history"]
        )
        if not dup:
            logger.debug("マッピング履歴追加: jlac10=%s, item='%s'", jlac10, item_name)
            entry["mapping_history"].append({
                "hospital": hospital,
                "item_name": item_name,
                "abbreviation": item.get("abbreviation", ""),
                "jlac10_standard_name": matched_name,
            })
            added += 1
        else:
            logger.debug("重複スキップ: jlac10=%s, item='%s', hospital='%s'", jlac10, item_name, hospital)
            skipped += 1

    data["items"].sort(key=lambda x: x.get("jlac10", ""))
    data["metadata"]["total_unique_jlac10"] = len(items_by_jlac)
    data["metadata"]["last_applied"] = datetime.now(timezone.utc).isoformat()

    merged_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    result = {"added": added, "skipped": skipped, "new_entries": new_entries}
    logger.info("DB還元完了: 追加%d件, スキップ%d件, 新規エントリ%d件", added, skipped, new_entries)
    return result
