"""SSMIX2 メッセージパーサー

HL7v2ベースのSSMIX2標準化ストレージデータをパースし、
JLAC10未設定エラーの検出・マッピング支援を行う。
"""

import logging
import re

from .scraper import classify_jlac10
from .sop_rules import validate_jlac10
from .vendor_profiles import detect_vendor_from_cs

logger = logging.getLogger(__name__)

# メタデータ系CS名（スキップ対象）
METADATA_CS_NAMES = {"99ZEC", "99ZED", "99ZER", "99ZES"}

# メタデータ系ローカルコード（溶血・混濁等）
METADATA_LOCAL_CODES = {"0079702", "0079703"}


def _parse_header(line: str) -> dict:
    """#SSMIX ヘッダー行をパース

    形式: #SSMIX,2.00,施設ID,患者ID,日付,メッセージ種別,検査ID,INS,施設番号,タイムスタンプ
    """
    parts = line.strip().split(",")
    return {
        "version": parts[1] if len(parts) > 1 else "",
        "facility_id": parts[2] if len(parts) > 2 else "",
        "patient_id": "***",  # 個人情報保護: マスク
        "date": parts[4] if len(parts) > 4 else "",
        "message_type": parts[5] if len(parts) > 5 else "",
        "order_id": parts[6] if len(parts) > 6 else "",
        "timestamp": parts[9] if len(parts) > 9 else "",
    }


def _parse_spm(line: str) -> dict:
    """SPM セグメントをパース

    形式: SPM|順番|||JC10材料コード^材料名^JC10^ローカル材料コード^材料名^99Z13||容器コード^容器名^99Z17
    """
    fields = line.split("|")
    sequence = int(fields[1]) if len(fields) > 1 and fields[1].isdigit() else 0

    # SPM-4: 検体材料（4番目のフィールド、0-indexed で [4]）
    material_field = fields[4] if len(fields) > 4 else ""
    mat_components = material_field.split("^")

    material_jc10 = mat_components[0] if len(mat_components) > 0 else ""
    material_name = mat_components[1] if len(mat_components) > 1 else ""
    local_material = mat_components[3] if len(mat_components) > 3 else ""

    # SPM-6: 容器（6番目のフィールド、0-indexed で [6]）
    container_field = fields[6] if len(fields) > 6 else ""
    cont_components = container_field.split("^")

    container_code = cont_components[0] if len(cont_components) > 0 else ""
    container_name = cont_components[1] if len(cont_components) > 1 else ""

    return {
        "sequence": sequence,
        "material_jc10": material_jc10,
        "material_name": material_name,
        "local_material": local_material,
        "container_code": container_code,
        "container_name": container_name,
    }


def _parse_obx(line: str, current_spm_material: str) -> dict:
    """OBX セグメントをパース

    JLAC10あり(6コンポーネント):
      OBX|順番|データ型|JLAC10^略称^JC10^ローカルコード^院内名称^CS名||値|単位^単位^ISO+|基準値|フラグ|||ステータス||R|日時

    JLAC10なし(3コンポーネント):
      OBX|順番|データ型|ローカルコード^名称^CS名||値|単位|...
    """
    fields = line.split("|")

    sequence = int(fields[1]) if len(fields) > 1 and fields[1].isdigit() else 0
    data_type = fields[2] if len(fields) > 2 else ""

    # OBX-3: 検査項目識別子
    obx3 = fields[3] if len(fields) > 3 else ""
    components = obx3.split("^")
    num_components = len(components)

    has_jlac10 = False
    jlac10 = ""
    jlac10_abbr = ""
    local_code = ""
    item_name = ""
    cs_name = ""

    if num_components >= 6 and (len(components) < 3 or components[2] == "JC10"):
        # 6コンポーネント形式: JLAC10^略称^JC10^ローカルコード^院内名称^CS名
        has_jlac10 = True
        jlac10 = components[0]
        jlac10_abbr = components[1] if len(components) > 1 else ""
        local_code = components[3] if len(components) > 3 else ""
        item_name = components[4] if len(components) > 4 else ""
        cs_name = components[5] if len(components) > 5 else ""
    elif num_components >= 3:
        # 3コンポーネント形式: ローカルコード^名称^CS名
        has_jlac10 = False
        local_code = components[0]
        item_name = components[1] if len(components) > 1 else ""
        cs_name = components[2] if len(components) > 2 else ""
    else:
        # 不明形式
        local_code = components[0] if components else ""

    # OBX-5: 値
    value = fields[5] if len(fields) > 5 else ""

    # OBX-6: 単位
    unit_field = fields[6] if len(fields) > 6 else ""
    unit = unit_field.split("^")[0] if unit_field else ""

    # OBX-7: 基準値
    reference_range = fields[7] if len(fields) > 7 else ""

    # OBX-8: 異常フラグ
    abnormal_flag = fields[8] if len(fields) > 8 else ""

    # OBX-11: ステータス
    status = fields[11] if len(fields) > 11 else ""

    # OBX-14: 結果日時 (R の次)
    timestamp = fields[14] if len(fields) > 14 else ""

    # メタデータ判定
    cs_upper = cs_name.strip().upper()
    is_metadata = (
        cs_upper in METADATA_CS_NAMES
        or local_code.strip() in METADATA_LOCAL_CODES
    )

    # ベンダー判定
    vendor_info = detect_vendor_from_cs(cs_name) if cs_name else None
    vendor = vendor_info["vendor"] if vendor_info else ""

    return {
        "sequence": sequence,
        "data_type": data_type,
        "has_jlac10": has_jlac10,
        "jlac10": jlac10,
        "jlac10_abbr": jlac10_abbr,
        "local_code": local_code,
        "item_name": item_name,
        "cs_name": cs_name,
        "value": value,
        "unit": unit,
        "reference_range": reference_range,
        "abnormal_flag": abnormal_flag,
        "status": status,
        "timestamp": timestamp,
        "is_metadata": is_metadata,
        "vendor": vendor,
        "spm_material": current_spm_material,
    }


def _detect_errors(message_index: int, obx_list: list[dict]) -> list[dict]:
    """OBXリストからエラーを検出"""
    errors = []

    for i, obx in enumerate(obx_list):
        # メタデータはスキップ
        if obx["is_metadata"]:
            continue

        # JLAC10未設定
        if not obx["has_jlac10"]:
            errors.append({
                "message_index": message_index,
                "obx_index": i,
                "error_type": "jlac10_missing",
                "local_code": obx["local_code"],
                "item_name": obx["item_name"],
                "detail": f"OBX-3が3コンポーネント形式（JLAC10未設定）: {obx['local_code']}^{obx['item_name']}^{obx['cs_name']}",
                "spm_material": obx["spm_material"],
            })
            continue

        # JLAC10形式チェック
        jlac10_status = classify_jlac10(obx["jlac10"])
        if jlac10_status == "invalid":
            errors.append({
                "message_index": message_index,
                "obx_index": i,
                "error_type": "jlac10_invalid",
                "local_code": obx["local_code"],
                "item_name": obx["item_name"],
                "detail": f"JLAC10形式不正: '{obx['jlac10']}' (期待: 15-17桁英数字)",
                "spm_material": obx["spm_material"],
            })
        elif jlac10_status in ("valid_15", "valid_17"):
            # SOPルールチェック
            sop_warnings = validate_jlac10(obx["jlac10"])
            for w in sop_warnings:
                errors.append({
                    "message_index": message_index,
                    "obx_index": i,
                    "error_type": "sop_violation",
                    "local_code": obx["local_code"],
                    "item_name": obx["item_name"],
                    "detail": f"SOP違反 [{w['field']}={w['code']}]: {w['message']}",
                    "spm_material": obx["spm_material"],
                })

        # 値の異常チェック（data_type=NM なのに数値でない）
        if obx["data_type"] == "NM" and obx["value"]:
            val = obx["value"].strip()
            # 数値として解釈できるか（小数点、符号、指数表記対応）
            if val and not re.match(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$", val):
                errors.append({
                    "message_index": message_index,
                    "obx_index": i,
                    "error_type": "value_anomaly",
                    "local_code": obx["local_code"],
                    "item_name": obx["item_name"],
                    "detail": f"NM型だが非数値: '{val}'",
                    "spm_material": obx["spm_material"],
                })

    return errors


def parse_ssmix(text: str) -> dict:
    """SSMIX2テキスト全体をパース

    Args:
        text: SSMIX2テキストファイルの内容（複数メッセージ可）

    Returns:
        {
            "messages": [{"header": {...}, "specimens": [...], "observations": [...]}, ...],
            "errors": [{...}, ...],
            "summary": {
                "total_messages": int,
                "total_obx": int,
                "jlac10_set": int,
                "jlac10_missing": int,
                "metadata_skipped": int,
                "vendor": str,
                "facility_id": str,
            }
        }
    """
    # #SSMIX で分割してメッセージ単位に
    raw_messages = re.split(r"(?=^#SSMIX)", text, flags=re.MULTILINE)
    raw_messages = [m for m in raw_messages if m.strip().startswith("#SSMIX")]

    messages = []
    all_errors = []
    total_obx = 0
    jlac10_set = 0
    jlac10_missing = 0
    metadata_skipped = 0
    vendors_seen = set()
    facility_id = ""

    for msg_idx, raw_msg in enumerate(raw_messages):
        lines = raw_msg.strip().splitlines()
        if not lines:
            continue

        # ヘッダーパース
        header = _parse_header(lines[0])
        if not facility_id:
            facility_id = header["facility_id"]

        specimens = []
        observations = []
        current_spm_material = ""

        for line in lines[1:]:
            line = line.strip()
            if not line:
                continue

            if line.startswith("SPM|"):
                spm = _parse_spm(line)
                specimens.append(spm)
                current_spm_material = f"{spm['material_jc10']}:{spm['material_name']}"
                logger.debug("SPM: %s", current_spm_material)

            elif line.startswith("OBX|"):
                obx = _parse_obx(line, current_spm_material)
                observations.append(obx)
                total_obx += 1

                if obx["is_metadata"]:
                    metadata_skipped += 1
                elif obx["has_jlac10"]:
                    jlac10_set += 1
                else:
                    jlac10_missing += 1

                if obx["vendor"]:
                    vendors_seen.add(obx["vendor"])

            # OBR, TQ1, MSH 等は参考情報として無視

        messages.append({
            "header": header,
            "specimens": specimens,
            "observations": observations,
        })

        # エラー検出
        msg_errors = _detect_errors(msg_idx, observations)
        all_errors.extend(msg_errors)

    # ベンダー判定（最頻出）
    vendor = ", ".join(sorted(vendors_seen)) if vendors_seen else "不明"

    return {
        "messages": messages,
        "errors": all_errors,
        "summary": {
            "total_messages": len(messages),
            "total_obx": total_obx,
            "jlac10_set": jlac10_set,
            "jlac10_missing": jlac10_missing,
            "metadata_skipped": metadata_skipped,
            "vendor": vendor,
            "facility_id": facility_id,
        },
    }
