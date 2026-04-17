"""ベンダー別 Excel プロファイル定義

電子カルテベンダーごとにヘッダー構造が異なるため、
ベンダー名を指定するだけで正しい列マッピングが適用されるようにする。

ベンダー分類:
  依頼/結果 別テーブル: Fujitsu, NEC, IBM
  依頼/結果 一体型:     SSI, SBS, KHI, CSI, NAIS
  細菌検査 別ファイル:   CSI 等
"""

import logging
import re

logger = logging.getLogger(__name__)


# ベンダープロファイル定義
# 各ベンダーの依頼/結果シートの列マッピング
# col_* は列を特定するキーワードリスト（ヘッダーに含まれる文字列で自動検出）
VENDOR_PROFILES: dict[str, dict] = {
    "NEC": {
        "type": "separate",  # 依頼/結果 別テーブル
        "sheets": {
            "依頼": {
                "col_item_name": ["TESTITEMNAME", "検査項目名", "項目名称"],
                "col_short_name": ["TESTITEM SHORTNAME", "SHORTNAME", "略称", "短縮名"],
                "col_item_code": ["TESTITEM CODE", "TESTITEMCODE", "項目コード"],
                # 最終的JLAC10 > JLAC10 CODE > JLAC10 の順で優先
                "col_jlac10": ["最終的JLAC10", "JLAC10 CODE", "JLAC10CODE", "JLAC10", "日立エラー修正後JLAC10"],
                "col_jlac10_name": ["JLAC10 NAME", "JLAC10NAME", "JLAC10標準名称", "JLAC10名称"],
                "col_unit": ["UNIT", "単位"],
                "col_category": ["TEST CATEGORY", "TESTCATEGORY", "検査区分"],
                "col_hospital_code": ["HOSPITAL CODE", "HOSPITALCODE", "施設コード"],
                # LACS CODE は院内コードなので jlac10 には使わない
                "col_lacs_code": ["LACS CODE", "LACSCODE"],
            },
            "結果": {
                "col_item_name": ["TESTITEMNAME", "検査項目名", "項目名称", "結果項目名"],
                "col_short_name": ["TESTITEM SHORTNAME", "SHORTNAME", "略称"],
                "col_item_code": ["TESTITEM CODE", "TESTITEMCODE", "項目コード"],
                "col_jlac10": ["最終的JLAC10", "JLAC10 CODE", "JLAC10CODE", "JLAC10"],
                "col_jlac10_name": ["JLAC10 NAME", "JLAC10NAME", "JLAC10標準名称"],
                "col_unit": ["UNIT", "単位"],
                "col_lacs_code": ["LACS CODE", "LACSCODE"],
            },
        },
    },
    "Fujitsu": {
        "type": "separate",
        "sheets": {
            "依頼": {
                "col_item_name": ["検査項目名", "項目名称", "検査名称", "TESTITEMNAME"],
                "col_short_name": ["略称", "短縮名", "表示名"],
                "col_item_code": ["項目コード", "検査コード", "依頼コード"],
                "col_jlac10": ["JLAC10", "JLACコード", "LACS CODE", "標準コード"],
                "col_unit": ["単位"],
            },
            "結果": {
                "col_item_name": ["結果項目名", "検査項目名", "項目名称"],
                "col_short_name": ["略称", "短縮名"],
                "col_item_code": ["項目コード", "結果コード"],
                "col_jlac10": ["JLAC10", "JLACコード", "LACS CODE"],
                "col_unit": ["単位"],
            },
        },
    },
    "IBM": {
        "type": "separate",
        "sheets": {
            "依頼": {
                "col_item_name": ["検査項目名", "項目名称", "TESTITEMNAME"],
                "col_short_name": ["略称", "短縮名", "SHORTNAME"],
                "col_item_code": ["項目コード", "TESTITEMCODE"],
                "col_jlac10": ["JLAC10", "JLACコード", "LACSCODE"],
                "col_unit": ["単位", "UNIT"],
            },
            "結果": {
                "col_item_name": ["結果項目名", "検査項目名", "項目名称"],
                "col_short_name": ["略称", "短縮名"],
                "col_item_code": ["項目コード", "結果コード"],
                "col_jlac10": ["JLAC10", "JLACコード"],
                "col_unit": ["単位"],
            },
        },
    },
    "SSI": {
        "type": "unified",  # 依頼/結果 一体型
        "sheets": {
            "default": {
                "col_item_name": ["検査項目名", "項目名称", "検査名称"],
                "col_short_name": ["略称", "短縮名", "表示名"],
                "col_item_code": ["項目コード", "検査コード"],
                "col_jlac10": ["JLAC10", "JLACコード", "標準コード"],
                "col_unit": ["単位"],
            },
        },
    },
    "SBS": {
        "type": "unified",
        "sheets": {
            "default": {
                "col_item_name": ["検査項目名", "項目名称"],
                "col_short_name": ["略称", "短縮名"],
                "col_item_code": ["項目コード"],
                "col_jlac10": ["JLAC10", "JLACコード"],
                "col_unit": ["単位"],
            },
        },
    },
    "KHI": {
        "type": "unified",
        "sheets": {
            "default": {
                "col_item_name": ["検査項目名", "項目名称"],
                "col_short_name": ["略称", "短縮名"],
                "col_item_code": ["項目コード"],
                "col_jlac10": ["JLAC10", "JLACコード"],
                "col_unit": ["単位"],
            },
        },
    },
    "CSI": {
        "type": "unified",
        "sheets": {
            "default": {
                "col_item_name": ["検査項目名", "項目名称"],
                "col_short_name": ["略称", "短縮名"],
                "col_item_code": ["項目コード"],
                "col_jlac10": ["JLAC10", "JLACコード"],
                "col_unit": ["単位"],
            },
            "細菌": {
                "col_item_name": ["検査項目名", "項目名称", "菌名"],
                "col_item_code": ["項目コード", "細菌コード"],
                "col_jlac10": ["JLAC10", "JANISコード"],
            },
        },
    },
    "NAIS": {
        "type": "unified",
        "sheets": {
            "default": {
                "col_item_name": ["検査項目名", "項目名称"],
                "col_short_name": ["略称", "短縮名"],
                "col_item_code": ["項目コード"],
                "col_jlac10": ["JLAC10", "JLACコード"],
                "col_unit": ["単位"],
            },
        },
    },
}


def _normalize_header(text: str) -> str:
    """ヘッダーテキストを正規化（改行・空白除去、大文字化）"""
    text = re.sub(r"[\n\r\s]+", "", text)
    return text.upper()


def detect_columns(
    headers: list[str],
    vendor: str | None = None,
    sheet_name: str | None = None,
) -> dict[str, int | None]:
    """ヘッダー行からカラムインデックスを自動検出

    Args:
        headers: ヘッダー行のテキストリスト
        vendor: ベンダー名（指定でプロファイル優先）
        sheet_name: シート名（依頼/結果/細菌等）

    Returns:
        {"item_name": 6, "short_name": 5, "item_code": 1, "jlac10": 8, "unit": 7, ...}
        見つからない列は None
    """
    headers_norm = [_normalize_header(h) for h in headers]
    logger.debug("ヘッダー正規化: %s", headers_norm)

    # ベンダープロファイルからキーワードリストを取得
    keywords = _get_keywords(vendor, sheet_name)
    logger.debug("検出キーワード: vendor=%s, sheet=%s", vendor, sheet_name)

    result: dict[str, int | None] = {}

    for field, kw_list in keywords.items():
        # field名から col_ プレフィックスを除去
        field_clean = field.replace("col_", "")
        found_idx = None

        # まず完全一致を試行
        for kw in kw_list:
            kw_norm = _normalize_header(kw)
            for i, h in enumerate(headers_norm):
                if kw_norm == h:
                    found_idx = i
                    logger.debug("  列検出(完全一致): %s → %d列目 (keyword='%s', header='%s')",
                                 field_clean, i, kw, headers[i])
                    break
            if found_idx is not None:
                break

        # 完全一致がなければ部分一致（長いキーワード優先）
        if found_idx is None:
            kw_sorted = sorted(kw_list, key=lambda k: len(k), reverse=True)
            for kw in kw_sorted:
                kw_norm = _normalize_header(kw)
                for i, h in enumerate(headers_norm):
                    if kw_norm in h and len(kw_norm) >= 4:
                        found_idx = i
                        logger.debug("  列検出(部分一致): %s → %d列目 (keyword='%s', header='%s')",
                                     field_clean, i, kw, headers[i])
                        break
                if found_idx is not None:
                    break

        result[field_clean] = found_idx

    # 必須列のチェック
    if result.get("item_name") is None:
        logger.warning("項目名列が検出できませんでした。ヘッダー: %s", headers)

    return result


def _get_keywords(vendor: str | None, sheet_name: str | None) -> dict[str, list[str]]:
    """ベンダー+シート名からキーワード辞書を取得"""
    if vendor and vendor in VENDOR_PROFILES:
        profile = VENDOR_PROFILES[vendor]
        sheets = profile["sheets"]

        # シート名マッチ
        if sheet_name and sheet_name in sheets:
            return sheets[sheet_name]

        # シート名で部分一致
        if sheet_name:
            for sn, cols in sheets.items():
                if sn in sheet_name or sheet_name in sn:
                    return cols

        # default があればそれ
        if "default" in sheets:
            return sheets["default"]

        # 最初のシートプロファイル
        return next(iter(sheets.values()))

    # ベンダー不明: 汎用キーワード
    return {
        "col_item_name": [
            "TESTITEMNAME", "検査項目名", "項目名称", "検査名称", "項目名",
            "結果項目名", "検査名", "TEST ITEM NAME",
        ],
        "col_short_name": [
            "TESTITEM SHORTNAME", "SHORTNAME", "略称", "短縮名", "表示名",
            "SHORT NAME", "ABBR",
        ],
        "col_item_code": [
            "TESTITEM CODE", "TESTITEMCODE", "項目コード", "検査コード",
            "依頼コード", "結果コード", "ITEM CODE", "TEST CODE",
        ],
        "col_jlac10": [
            "最終的JLAC10", "JLAC10 CODE", "JLAC10CODE", "JLAC10コード",
            "JLAC10", "JLACコード", "標準コード",
        ],
        "col_lacs_code": [
            "LACS CODE", "LACSCODE",
        ],
        "col_jlac10_name": [
            "JLAC10標準名称", "JLAC10名称", "標準名称", "JLAC名称",
        ],
        "col_unit": [
            "UNIT", "単位",
        ],
        "col_category": [
            "TEST CATEGORY", "TESTCATEGORY", "検査区分", "区分", "CATEGORY",
        ],
        "col_hospital_code": [
            "HOSPITAL CODE", "HOSPITALCODE", "施設コード", "病院コード",
        ],
    }


def list_vendors() -> list[str]:
    """登録済みベンダー一覧"""
    return sorted(VENDOR_PROFILES.keys())


def get_vendor_info(vendor: str) -> dict | None:
    """ベンダー情報を取得"""
    return VENDOR_PROFILES.get(vendor)
