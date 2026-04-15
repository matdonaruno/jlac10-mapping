"""JLAC10 あいまい検索エンジン

検索の入口: 検査項目名称（院内名・略称・英名なんでもOK）
検索の出口: マッチした JLAC10 コード一覧（院内/SRL/BML/LSI 別）

検索ロジック:
  1. 全ソースの全名称（SRL名、BML名、LSI名、JLAC10標準名、英名）を
     1つの分析物コードに紐づけてインデックス化
  2. クエリをカタカナ・ひらがな・英字で正規化
  3. 部分一致 + スコアリングで候補を返す
"""

import json
import re
import unicodedata
from pathlib import Path


def _normalize(text: str) -> str:
    """検索用にテキストを正規化する
    - 全角→半角（英数字）
    - カタカナ→ひらがな
    - 小文字化
    - 空白・記号除去
    """
    # 全角英数→半角
    text = unicodedata.normalize("NFKC", text)
    # カタカナ→ひらがな
    result = []
    for ch in text:
        cp = ord(ch)
        if 0x30A1 <= cp <= 0x30F6:  # ァ-ヶ
            result.append(chr(cp - 0x60))
        else:
            result.append(ch)
    text = "".join(result)
    # 小文字化
    text = text.lower()
    # 括弧内も保持しつつ、検索用にスペースや記号を除去しない（部分一致で拾う）
    return text


def _score(query_norm: str, target_norm: str, target_raw: str) -> float:
    """マッチスコアを計算（高いほど良い）"""
    if not query_norm or not target_norm:
        return 0.0

    # 完全一致
    if query_norm == target_norm:
        return 100.0

    # 先頭一致（略称でよくあるパターン）
    if target_norm.startswith(query_norm):
        return 80.0 + len(query_norm) / len(target_norm) * 10

    # 含まれる（部分一致）
    if query_norm in target_norm:
        return 60.0 + len(query_norm) / len(target_norm) * 10

    # クエリの各単語が全て含まれる（AND検索）
    words = query_norm.split()
    if len(words) > 1 and all(w in target_norm for w in words):
        return 50.0 + len(query_norm) / len(target_norm) * 10

    # 英字略称チェック: "TP" → "(TP)" や "TP/" を探す
    if query_norm.isascii() and len(query_norm) <= 5:
        # 括弧内の略称
        patterns = [
            f"({query_norm})",
            f"（{query_norm}）",
            f"{query_norm}/",
            f"/{query_norm}",
            f" {query_norm} ",
            f" {query_norm})",
        ]
        target_lower = target_raw.lower()
        for pat in patterns:
            if pat.lower() in target_lower:
                return 75.0

    return 0.0


class SearchIndex:
    """JLAC10 検索インデックス"""

    def __init__(self):
        self.entries: list[dict] = []
        # 分析物コード → エントリインデックス のマッピング
        self._analyte_map: dict[str, list[int]] = {}

    def build_from_merged(self, merged_path: Path, master_path: Path | None = None):
        """merged_jlac10.json と jlac10_master.json からインデックス構築"""
        merged = json.loads(merged_path.read_text(encoding="utf-8"))

        # JLAC10マスターの分析物名称辞書
        analyte_names: dict[str, dict] = {}
        if master_path and master_path.exists():
            master = json.loads(master_path.read_text(encoding="utf-8"))
            for item in master.get("master", {}).get("analyte", []):
                analyte_names[item["code"]] = item

        for item in merged["items"]:
            jlac10 = item["jlac10"]
            analyte_code = item.get("analyte_code", jlac10[:5])
            decoded = item.get("jlac10_decoded", {})

            # 全名称を収集
            names: list[str] = []

            # JLAC10 標準名称
            if decoded.get("valid"):
                an = decoded.get("analyte", {})
                if an.get("name"):
                    names.append(an["name"])
                if an.get("name_en"):
                    names.append(an["name_en"])

            # マスターから追加名称
            if analyte_code in analyte_names:
                a = analyte_names[analyte_code]
                for key in ("name", "name2", "name_en"):
                    if a.get(key):
                        names.append(a[key])

            # 各ソースの検査項目名称
            for src_key, src in item.get("sources", {}).items():
                if src.get("item_name"):
                    names.append(src["item_name"])

            # 重複除去
            seen = set()
            unique_names = []
            for n in names:
                n_strip = n.strip()
                if n_strip and n_strip not in seen:
                    seen.add(n_strip)
                    unique_names.append(n_strip)

            # 正規化済み名称
            normalized = [_normalize(n) for n in unique_names]

            entry = {
                "jlac10": jlac10,
                "analyte_code": analyte_code,
                "names": unique_names,
                "names_normalized": normalized,
                "decoded": decoded,
                "sources": item.get("sources", {}),
            }

            idx = len(self.entries)
            self.entries.append(entry)

            if analyte_code not in self._analyte_map:
                self._analyte_map[analyte_code] = []
            self._analyte_map[analyte_code].append(idx)

    def search(self, query: str, max_results: int = 20) -> list[dict]:
        """あいまい検索を実行

        Args:
            query: 検索文字列（「TP」「総蛋白」「albumin」等なんでもOK）
            max_results: 最大結果数

        Returns:
            スコア順のマッチ結果リスト
        """
        query_norm = _normalize(query)
        if not query_norm:
            return []

        results = []
        for entry in self.entries:
            best_score = 0.0
            matched_name = ""
            for name, name_norm in zip(entry["names"], entry["names_normalized"]):
                s = _score(query_norm, name_norm, name)
                if s > best_score:
                    best_score = s
                    matched_name = name

            if best_score > 0:
                results.append({
                    "score": best_score,
                    "matched_name": matched_name,
                    "jlac10": entry["jlac10"],
                    "analyte_code": entry["analyte_code"],
                    "all_names": entry["names"],
                    "decoded": entry["decoded"],
                    "sources": entry["sources"],
                })

        results.sort(key=lambda x: (-x["score"], x["jlac10"]))
        return results[:max_results]

    def search_by_analyte(self, analyte_code: str) -> list[dict]:
        """分析物コード(5桁)で検索"""
        indices = self._analyte_map.get(analyte_code, [])
        return [self.entries[i] for i in indices]


def format_results(results: list[dict]) -> str:
    """検索結果をターミナル表示用にフォーマット"""
    if not results:
        return "  該当なし"

    lines = []
    for i, r in enumerate(results, 1):
        lines.append(f"\n{'─' * 60}")
        lines.append(f"  [{i}] {r['matched_name']}  (スコア: {r['score']:.0f})")
        lines.append(f"      JLAC10: {r['jlac10']}")
        lines.append(f"      分析物: {r['analyte_code']}")

        dec = r.get("decoded", {})
        if dec.get("valid"):
            parts = []
            for key in ("analyte", "material", "method"):
                d = dec.get(key, {})
                if d.get("name"):
                    parts.append(f"{key}={d['name']}")
            if parts:
                lines.append(f"      標準: {' / '.join(parts)}")

        lines.append(f"      別名: {', '.join(r['all_names'][:5])}")

        # ソース別コード
        for src_key in ("srl", "bml", "lsi"):
            src = r["sources"].get(src_key)
            if src:
                lines.append(
                    f"      {src_key.upper():3s}: {src['item_name']}"
                    f"  材料={src['material']}"
                    f"  方法={src['method']}"
                )

    return "\n".join(lines)


def build_index(data_dir: Path) -> SearchIndex:
    """data ディレクトリからインデックスを構築"""
    merged_path = data_dir / "merged_jlac10.json"
    master_path = data_dir / "jlac10_master.json"

    if not merged_path.exists():
        raise FileNotFoundError(f"{merged_path} が見つかりません。merge を先に実行してください。")

    index = SearchIndex()
    index.build_from_merged(merged_path, master_path)
    return index
