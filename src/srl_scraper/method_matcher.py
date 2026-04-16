"""測定原理 → JLAC10 測定法コード(3桁) 変換

jlac10_lookup.json の method セクション（360件）から
キーワード辞書を自動生成し、測定原理テキストとマッチングする。

用途:
  - PMDA 添付文書の「測定原理」テキスト → 測定法コード推定
  - SOP の「測定法」テキスト → 測定法コード推定
  - 試薬名からの測定法推定
"""

import json
import logging
import re
from pathlib import Path

logger = logging.getLogger(__name__)


def build_method_keyword_map(lookup_path: Path) -> dict[str, dict]:
    """jlac10_lookup.json の method セクションからキーワード辞書を生成

    Returns:
        {
            "023": {"code": "023", "name": "ビウレット法", "keywords": ["ビウレット", "biuret"]},
            "271": {"code": "271", "name": "可視吸光光度法", "keywords": ["可視吸光", "比色", "酵素法"]},
            ...
        }
    """
    if not lookup_path.exists():
        logger.warning("jlac10_lookup.json が見つかりません: %s", lookup_path)
        return {}

    lookup = json.loads(lookup_path.read_text(encoding="utf-8"))
    methods = lookup.get("method", {})

    keyword_map = {}
    for code, info in methods.items():
        name = info.get("name", "")
        name2 = info.get("name2", "")
        name_en = info.get("name_en", "")

        keywords = set()
        # 日本語名からキーワード生成
        if name:
            keywords.add(name.lower())
            # 括弧内のテキストも独立キーワード
            for m in re.findall(r"[（(]([^）)]+)[）)]", name):
                keywords.add(m.lower())
            # スラッシュ区切り
            for part in name.split("/"):
                part = part.strip()
                if len(part) >= 2:
                    keywords.add(part.lower())
        if name2:
            keywords.add(name2.lower())
        if name_en:
            keywords.add(name_en.lower())
            # 英名の略称
            for m in re.findall(r"[（(]([A-Za-z]+)[）)]", name_en):
                keywords.add(m.lower())

        keyword_map[code] = {
            "code": code,
            "name": name,
            "name2": name2,
            "name_en": name_en,
            "keywords": list(keywords),
        }

    total_keywords = sum(len(v["keywords"]) for v in keyword_map.values())
    logger.debug(
        "キーワード辞書構築: %dコード, キーワード総数=%d",
        len(keyword_map), total_keywords,
    )
    logger.info("測定法キーワード辞書: %d コード", len(keyword_map))
    return keyword_map


# 手動追加キーワード（自動生成で拾えないもの）
MANUAL_KEYWORDS: dict[str, list[str]] = {
    "023": ["ビウレット", "biuret"],
    "062": ["ラテックス凝集比濁", "ラテックス比濁", "la法", "latex"],
    "063": ["ネフェロメトリー", "免疫比朧"],
    "061": ["免疫比濁", "tia", "turbidimetric immunoassay"],
    "116": ["eclia", "電気化学発光"],
    "117": ["clia", "化学発光免疫"],
    "271": ["酵素法", "enzymatic", "比色法", "可視吸光"],
    "841": ["fish", "蛍光in situ"],
    "862": ["リアルタイムpcr", "real-time pcr", "real time pcr"],
    "848": ["ダイレクトシーケンス", "塩基配列決定", "サンガー"],
    "051": ["化学発光酵素免疫", "cleia"],
}


def match_method_code(
    text: str,
    keyword_map: dict,
) -> list[dict]:
    """測定原理/測定法テキストからJLAC10測定法コード候補を返す

    Args:
        text: 測定原理テキスト（PMDA/SOP等から）
        keyword_map: build_method_keyword_map() の出力

    Returns:
        スコア順の候補リスト [{code, name, score, matched_keyword}, ...]
    """
    if not text or not keyword_map:
        logger.debug("match_method_code: 入力テキストまたはkeyword_mapが空")
        return []

    logger.debug("match_method_code: 入力='%s'", text[:50])
    text_lower = text.lower()
    text_norm = re.sub(r"\s+", "", text_lower)

    results = []
    for code, info in keyword_map.items():
        all_keywords = info["keywords"][:]
        # 手動キーワードも追加
        if code in MANUAL_KEYWORDS:
            all_keywords.extend(MANUAL_KEYWORDS[code])

        best_score = 0
        best_keyword = ""
        for kw in all_keywords:
            kw_lower = kw.lower()
            kw_norm = re.sub(r"\s+", "", kw_lower)
            if not kw_norm:
                continue
            if kw_norm in text_norm:
                score = len(kw_norm) / len(text_norm) * 100
                # 長いキーワードほど信頼性が高い
                score = min(score * 3, 100)
                if score > best_score:
                    best_score = score
                    best_keyword = kw

        if best_score > 0:
            results.append({
                "code": code,
                "name": info["name"],
                "score": round(best_score, 1),
                "matched_keyword": best_keyword,
            })

    results.sort(key=lambda x: -x["score"])
    if results:
        for r in results:
            logger.debug(
                "  マッチ: code=%s name='%s' score=%.1f keyword='%s'",
                r["code"], r["name"], r["score"], r["matched_keyword"],
            )
    else:
        logger.debug("  マッチなし: '%s'", text[:50])
    return results


def match_from_sop(
    method_summary: str,
    keyword_map: dict,
) -> list[dict]:
    """SOP の method_summary から測定法コードを推定"""
    return match_method_code(method_summary, keyword_map)


def match_from_reagent(
    principle: str,
    keyword_map: dict,
) -> list[dict]:
    """試薬DB (PMDA添付文書) の principle から測定法コードを推定"""
    return match_method_code(principle, keyword_map)


def build_and_match(
    lookup_path: Path,
    text: str,
    max_results: int = 5,
) -> list[dict]:
    """ワンショット: lookup読み込み→マッチ→結果返却"""
    keyword_map = build_method_keyword_map(lookup_path)
    results = match_method_code(text, keyword_map)
    return results[:max_results]
