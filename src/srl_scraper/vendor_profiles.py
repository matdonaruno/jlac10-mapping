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


# =========================================================================
# CS名（Coding System）→ ベンダー/マスタ種別マッピング
# SSMIX2 の OBX-3 フィールドに含まれる CS名からベン���ーを特定する
# =========================================================================

CS_TO_VENDOR: dict[str, dict] = {
    # 富士通
    "99Z14": {"vendor": "Fujitsu", "type": "検体検査", "master": "ET1/検歴"},
    # NEC
    "99ZTI": {"vendor": "NEC", "type": "検体検査依頼", "master": "検査オーダマスタ"},
    "99ZRD": {"vendor": "NEC", "type": "検体検査結果", "master": "検査結果マスタ"},
    "99ZGP": {"vendor": "NEC", "type": "細菌_塗抹", "master": "一般細菌塗抹"},
    "99ZGC": {"vendor": "NEC", "type": "細菌_培養同定", "master": "一般細菌���養同定"},
    "99ZPP": {"vendor": "NEC", "type": "細菌_抗酸菌塗抹", "master": "抗酸菌塗抹"},
    "99ZPC": {"vendor": "NEC", "type": "細菌_抗酸菌培養同定", "master": "抗酸菌培養同定"},
    "99ZPI": {"vendor": "NEC", "type": "細菌_抗酸菌同定", "master": "抗酸菌同定"},
    "99ZGM": {"vendor": "NEC", "type": "細菌_感受性薬剤", "master": "一般細菌感受性"},
    "99ZPm": {"vendor": "NEC", "type": "細菌_抗酸菌感受性", "master": "抗酸菌感受性"},
    "99DGO": {"vendor": "NEC", "type": "細菌_その他", "master": "その他検査"},
    # SSI
    "99Z18": {"vendor": "SSI", "type": "検体検査_OML01", "master": "依頼結果一体"},
    "99ZER": {"vendor": "SSI", "type": "検体検査_OML11", "master": "依頼結果一体"},
    "99ZS6": {"vendor": "SSI", "type": "細菌検査", "master": "依頼結果一体"},
    # SBS
    "99ZB3": {"vendor": "SBS", "type": "検体検査", "master": "依頼結果一体"},  # NAISも同じ
    "99ZC3": {"vendor": "SBS", "type": "細菌検査依頼", "master": "細菌検査依頼"},
    "99ZC4": {"vendor": "SBS", "type": "細菌_塗抹", "master": "一般細菌塗���"},
    "99ZC5": {"vendor": "SBS", "type": "細菌_感受性", "master": "感受性"},
    "99ZC6": {"vendor": "SBS", "type": "細菌_培養同定", "master": "培養同定"},
    "99ZC8": {"vendor": "SBS", "type": "細菌_その他", "master": "その他"},
    # KHI
    "99KEN": {"vendor": "KHI", "type": "検体検査", "master": "依頼結果一体"},
    # IBM
    "99101": {"vendor": "IBM", "type": "検体検査", "master": "検体検査マスタ"},
    "99104": {"vendor": "IBM", "type": "細菌_塗抹", "master": "一般細菌塗抹"},
    "99105": {"vendor": "IBM", "type": "細菌_感受性", "master": "一般細菌感受性"},
    "99106": {"vendor": "IBM", "type": "細菌_抗酸菌感受性", "master": "抗酸菌感受性"},
    "99107": {"vendor": "IBM", "type": "細菌_培養同定", "master": "培養同定"},
    "99108": {"vendor": "IBM", "type": "細菌検査依頼", "master": "細菌検���依頼"},
    # CSI
    "99TM1": {"vendor": "CSI", "type": "検体検査依頼", "master": "検体検査依頼マスタ"},
    "99RM1": {"vendor": "CSI", "type": "検体検査結果", "master": "検体検査結果マスタ"},
}

# 設定依頼の送付先パターン
DELIVERY_TARGET: dict[str, dict[str, str]] = {
    # vendor: {検査種別: 送付先}
    "Fujitsu": {
        "検体依頼": "病院(ET1画面)",
        "検体結果": "病院(JJマスタツール)",
        "細��依頼": "病院(ET2画面)",
        "細菌結果": "病院(JJマスタツール)",
    },
    "NEC": {
        "検体依頼": "病院(画面)",
        "検体結果": "病院(画面)",
        "細菌依頼": "病院(画面)",
        "細菌結果": "NEC(指定フォーマット)",
    },
    "SSI": {
        "検体": "病院(画面)",
        "細菌JLAC10": "SSI(指定フォーマット)",
        "細菌JANIS": "病院(画面)",
    },
    "SBS": {
        "検体": "病院(画面)",
        "細菌依頼": "病院(画面)",
        "細菌結果": "SBS(登録依頼)",
    },
    "KHI": {"検体": "病院(画面)", "細菌": "出力なし"},
    "IBM": {
        "検体": "IBM or 病院検査科",
        "細菌": "IBM or 病院検査科",
    },
    "CSI": {
        "検体依頼": "病院(画面)",
        "検体結果": "病院(画面)",
        "細菌依頼": "病院(画面)",
        "細菌結果": "CSI(登録依頼)",
    },
    "NAIS": {"検体": "病院(画面)", "細菌": "出力なし"},
}


def detect_vendor_from_cs(cs_name: str) -> dict | None:
    """CS名からベンダーとマスタ種別を特定する

    Args:
        cs_name: SSMIX2 OBX-3 の CS名部分 (例: "99ZTI")

    Returns:
        {"vendor": "NEC", "type": "検体検査依頼", "master": "..."} or None
    """
    cs_upper = cs_name.strip().upper()
    # 完全一致
    if cs_upper in CS_TO_VENDOR:
        return CS_TO_VENDOR[cs_upper]
    # 99Z で始まるNEC系
    if cs_upper.startswith("99Z"):
        return {"vendor": "NEC(推定)", "type": "不明", "master": cs_upper}
    return None


def get_delivery_target(vendor: str, exam_type: str) -> str:
    """ベンダーと検査種別から設定依頼の送付��を取得"""
    targets = DELIVERY_TARGET.get(vendor, {})
    if exam_type in targets:
        return targets[exam_type]
    # 部分一致
    for k, v in targets.items():
        if exam_type in k or k in exam_type:
            return v
    return "不明（過去例を参照）"


def list_vendors() -> list[str]:
    """登録済みベンダー一覧"""
    return sorted(VENDOR_PROFILES.keys())


def get_vendor_info(vendor: str) -> dict | None:
    """ベンダー情報を取得"""
    return VENDOR_PROFILES.get(vendor)
