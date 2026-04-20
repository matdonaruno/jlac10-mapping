"""NCDA マッピング SOP ルールバリデーション

標準作業手順書 (2018年3月改訂版) に基づくJLAC10コードのルールチェック。
違反コードを検出し警告を表示する。

参照: NCDA/6. 標準作業手順書.docx
"""

import logging
import re

logger = logging.getLogger(__name__)


# =========================================================================
# SOP ルール定義
# =========================================================================

# 使用禁止の識別コード
FORBIDDEN_IDENTIFICATION = {
    "1351": "定性は使用しない（SOPルール）",
    "1352": "定量は使用しない（SOPルール）",
    "1353": "半定量は使用しない（SOPルール）",
    "1411": "ウイルス抗原定性は使用しない→1410を使用",
    "1412": "ウイルス抗原半定量は使用しない→1410を使用",
    "1413": "ウイルス抗原定量は使用しない→1410を使用",
    "1491": "ウイルス抗体定性は使用しない→1430を使用",
    "1492": "ウイルス抗体定量は使用しない→1430を使用",
    "1493": "ウイルス抗体半定量は使用しない→1430を使用",
    "1441": "ウイルスDNA定量は使用しない→1440を使用",
    "1453": "ウイルスRNA定量は使用しない→1450を使用",
}

# 使用禁止の材料コード（200-999は使用しない）
def _is_forbidden_material(code: str) -> str | None:
    """材料コードの禁止チェック"""
    if not code or len(code) != 3:
        return None
    try:
        num = int(code)
    except ValueError:
        return None
    if 200 <= num <= 999:
        return f"材料コード{code}は200-999の範囲で使用禁止（SOPルール）"
    return None


# 使用禁止の尿材料コード（001/004/010 以外）
ALLOWED_URINE_MATERIALS = {"001", "004", "010"}
URINE_ANALYTES = {"1A005", "1A006", "1A007", "1A010", "1A015", "1A020",
                  "1A025", "1A030", "1A035", "1A100", "1A105", "1A150", "1A990"}

# 使用禁止の測定法コード
FORBIDDEN_METHOD = {
    "999": "測定法999は使用しない→920(その他)を使用",
}

# 特定分析物に対する測定法ルール
ANALYTE_METHOD_RULES = {
    "1A990": {"forbidden": ["901", "911"], "use": "920", "msg": "尿一般の測定法は920を使用"},
    "1A105": {"forbidden": ["310", "662", "701", "735"], "use": "920", "msg": "尿沈渣の測定法は920を使用"},
    "1A035": {"forbidden": [], "use": "920", "msg": "pH[尿]の測定法は920を使用"},
    "1C025": {"forbidden": [], "use": "920", "msg": "pH[髄液]の測定法は920を使用"},
    "1C035": {"forbidden": ["603"], "use": "920", "msg": "細胞種類[髄液]の測定法は920を使用"},
    "2A160": {"forbidden": ["301", "603", "604", "662"], "use": "309", "msg": "血液像の測定法は309(自動機械法)を使用"},
    "2A170": {"forbidden": ["301", "603", "604", "662"], "use": "309", "msg": "骨髄像の測定法は309(自動機械法)を使用"},
    "2A180": {"forbidden": ["612"], "use": "310", "msg": "ALP染色の測定法は310(鏡検法)を使用"},
}

# 特定分析物に対する材料ルール
ANALYTE_MATERIAL_RULES = {
    "2A160": {"forbidden": ["034"], "use": "019", "msg": "血液像の材料は019(全血添加物入り)を使用"},
    "2A170": {"forbidden": ["049"], "use": "046", "msg": "骨髄像の材料は046(骨髄液)を使用"},
}

# 結果識別共通コードのルール
FORBIDDEN_RESULT_COMMON = {
    "32": "陰性コントロール比は使用しない→21(コントロール値)を使用",
    "33": "陽性コントロール比は使用しない→21(コントロール値)を使用",
}


# =========================================================================
# バリデーション関数
# =========================================================================

def validate_jlac10(jlac10: str) -> list[dict]:
    """JLAC10コードをSOPルールでバリデーション

    Args:
        jlac10: 15桁または17桁のJLAC10コード（ハイフンなし）

    Returns:
        警告リスト。空なら問題なし。
        [{"severity": "warning"|"error", "field": "...", "code": "...", "message": "..."}]
    """
    code = jlac10.replace("-", "")
    warnings = []

    if len(code) < 15:
        return warnings

    analyte = code[0:5]
    identification = code[5:9]
    material = code[9:12]
    method = code[12:15]
    result_id = code[15:17] if len(code) >= 17 else ""

    # 識別コードチェック
    if identification in FORBIDDEN_IDENTIFICATION:
        warnings.append({
            "severity": "warning",
            "field": "identification",
            "code": identification,
            "message": FORBIDDEN_IDENTIFICATION[identification],
        })

    # 材料コードチェック（200-999禁止）
    mat_msg = _is_forbidden_material(material)
    if mat_msg:
        warnings.append({
            "severity": "warning",
            "field": "material",
            "code": material,
            "message": mat_msg,
        })

    # 尿検査の材料チェック
    if analyte in URINE_ANALYTES and material not in ALLOWED_URINE_MATERIALS:
        if material != "099":  # その他は許容
            warnings.append({
                "severity": "warning",
                "field": "material",
                "code": material,
                "message": f"尿検査({analyte})の材料は001/004/010のみ使用可",
            })

    # 測定法コードチェック
    if method in FORBIDDEN_METHOD:
        warnings.append({
            "severity": "warning",
            "field": "method",
            "code": method,
            "message": FORBIDDEN_METHOD[method],
        })

    # 分析物別の測定法ルール
    if analyte in ANALYTE_METHOD_RULES:
        rule = ANALYTE_METHOD_RULES[analyte]
        if method in rule["forbidden"]:
            warnings.append({
                "severity": "warning",
                "field": "method",
                "code": method,
                "message": rule["msg"],
            })
        elif rule["use"] and method != rule["use"] and not rule["forbidden"]:
            # "use"が指定され、forbiddenが空＝この測定法以外は全部NG
            warnings.append({
                "severity": "warning",
                "field": "method",
                "code": method,
                "message": rule["msg"],
            })

    # 分析物別の材料ルール
    if analyte in ANALYTE_MATERIAL_RULES:
        rule = ANALYTE_MATERIAL_RULES[analyte]
        if material in rule["forbidden"]:
            warnings.append({
                "severity": "warning",
                "field": "material",
                "code": material,
                "message": rule["msg"],
            })

    # 結果識別（共通）チェック
    if result_id in FORBIDDEN_RESULT_COMMON:
        warnings.append({
            "severity": "warning",
            "field": "result_identification",
            "code": result_id,
            "message": FORBIDDEN_RESULT_COMMON[result_id],
        })

    return warnings


def validate_batch(items: list[dict]) -> list[dict]:
    """複数項目のバッチバリデーション

    各 item に "jlac10" フィールドがあることを前提。
    結果に "sop_warnings" フィールドを追加して返す。
    """
    for item in items:
        jlac10 = item.get("jlac10", "")
        if jlac10:
            item["sop_warnings"] = validate_jlac10(jlac10)
        else:
            item["sop_warnings"] = []
    return items


# =========================================================================
# JavaScript 用のルール出力（Pages で使う）
# =========================================================================

def export_rules_as_json() -> dict:
    """SOPルールをJSON形式で出力（Pages UIに埋め込み用）"""
    return {
        "forbidden_identification": FORBIDDEN_IDENTIFICATION,
        "forbidden_method": FORBIDDEN_METHOD,
        "forbidden_result_common": FORBIDDEN_RESULT_COMMON,
        "analyte_method_rules": ANALYTE_METHOD_RULES,
        "analyte_material_rules": ANALYTE_MATERIAL_RULES,
        "allowed_urine_materials": list(ALLOWED_URINE_MATERIALS),
        "urine_analytes": list(URINE_ANALYTES),
    }
