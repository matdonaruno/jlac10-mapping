"""院内検査マスタ（Excel/CSV）を統一JSONフォーマットに変換するモジュール"""

import csv
import json
import logging
import re
from datetime import datetime, timezone
from pathlib import Path

from .scraper import classify_jlac10

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# 列指定のパース
# ---------------------------------------------------------------------------

def _parse_column_spec(spec: str) -> int | str:
    """列指定文字列をパースする。

    対応形式:
      - アルファベット列名: "A", "B", "AA" → 0-indexed 整数
      - 数値文字列: "1", "2", "3" → 0-indexed 整数 (1始まりを0始まりに変換)
      - ヘッダ名: "検査項目名称" などの文字列 → そのまま返す

    Returns:
        int (0-indexed列番号) または str (ヘッダ名)
    """
    spec = spec.strip()

    # 数値指定 ("1", "2", "3" ...)
    if spec.isdigit():
        idx = int(spec) - 1
        if idx < 0:
            raise ValueError(f"列番号は1以上を指定してください: {spec}")
        return idx

    # アルファベット列名 ("A", "B", "AA" ...)
    if re.fullmatch(r"[A-Za-z]{1,3}", spec):
        result = 0
        for ch in spec.upper():
            result = result * 26 + (ord(ch) - ord("A") + 1)
        return result - 1  # 0-indexed

    # ヘッダ名（それ以外の文字列）
    return spec


def _resolve_column_index(spec: str, headers: list[str] | None) -> int:
    """列指定を確定した0-indexedの列番号に変換する。

    ヘッダ名指定の場合は headers が必要。
    """
    parsed = _parse_column_spec(spec)
    if isinstance(parsed, int):
        return parsed

    # ヘッダ名で検索
    header_name = parsed
    if headers is None:
        raise ValueError(
            f"ヘッダ名 '{header_name}' で列を指定していますが、"
            "ヘッダ行が見つかりません"
        )
    # 完全一致
    for i, h in enumerate(headers):
        if h.strip() == header_name:
            return i
    # 部分一致（フォールバック）
    for i, h in enumerate(headers):
        if header_name in h.strip():
            logger.warning(
                "列 '%s' を部分一致で解決しました: '%s' (列%d)",
                header_name, h.strip(), i + 1,
            )
            return i
    raise ValueError(
        f"ヘッダ名 '{header_name}' に一致する列が見つかりません。"
        f" 利用可能なヘッダ: {headers}"
    )


# ---------------------------------------------------------------------------
# Excel (.xlsx) 読み込み
# ---------------------------------------------------------------------------

def _read_xlsx(
    filepath: Path,
    sheet_name: str | None,
    skip_rows: int,
) -> tuple[list[list[str]], list[str] | None]:
    """openpyxl で Excel を読み込み、文字列の2次元リストとヘッダを返す。

    Returns:
        (data_rows, header_row)
        skip_rows > 0 の場合、先頭 skip_rows 行のうち最後の行をヘッダとして返す。
    """
    import openpyxl

    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    try:
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise ValueError(
                    f"シート '{sheet_name}' が見つかりません。"
                    f" 利用可能: {wb.sheetnames}"
                )
            ws = wb[sheet_name]
        else:
            ws = wb.active

        all_rows: list[list[str]] = []
        for row in ws.iter_rows(values_only=True):
            all_rows.append([str(cell) if cell is not None else "" for cell in row])
    finally:
        wb.close()

    header_row = None
    if skip_rows > 0 and len(all_rows) >= skip_rows:
        header_row = all_rows[skip_rows - 1]
        data_rows = all_rows[skip_rows:]
    else:
        data_rows = all_rows

    return data_rows, header_row


# ---------------------------------------------------------------------------
# CSV (.csv) 読み込み
# ---------------------------------------------------------------------------

def _read_csv(
    filepath: Path,
    skip_rows: int,
) -> tuple[list[list[str]], list[str] | None]:
    """CSV を読み込み、文字列の2次元リストとヘッダを返す。"""
    all_rows: list[list[str]] = []

    # エンコーディングを自動検出
    for encoding in ("utf-8-sig", "utf-8", "cp932", "shift_jis"):
        try:
            with filepath.open("r", encoding=encoding, newline="") as f:
                reader = csv.reader(f)
                all_rows = [row for row in reader]
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(
            f"ファイル '{filepath}' のエンコーディングを判定できません。"
            " UTF-8, CP932, Shift_JIS を試しました。"
        )

    header_row = None
    if skip_rows > 0 and len(all_rows) >= skip_rows:
        header_row = all_rows[skip_rows - 1]
        data_rows = all_rows[skip_rows:]
    else:
        data_rows = all_rows

    return data_rows, header_row


# ---------------------------------------------------------------------------
# メイン変換関数
# ---------------------------------------------------------------------------

def convert_tabular(
    filepath: Path,
    column_map: dict[str, str],
    hospital: str = "",
    sheet_name: str | None = None,
    skip_rows: int = 1,
    output_path: Path | None = None,
) -> dict:
    """院内検査マスタ (Excel/CSV) を統一JSON形式に変換する。

    Args:
        filepath: 入力ファイルパス (.xlsx または .csv)
        column_map: フィールド名→列指定のマッピング
            必須キー: "item_name", "jlac10"
            任意キー: "abbreviation", "jlac10_standard_name"
            列指定: "A"/"B" (アルファベット), "1"/"2" (数値), "検査名" (ヘッダ名)
        hospital: 病院名
        sheet_name: Excelシート名 (省略で最初のシート)
        skip_rows: スキップするヘッダ行数 (デフォルト1)
        output_path: 出力JSONパス (省略で {入力ファイル名}.json)

    Returns:
        出力JSONと同じ構造の dict
    """
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {filepath}")

    suffix = filepath.suffix.lower()
    if suffix == ".xlsx":
        data_rows, header_row = _read_xlsx(filepath, sheet_name, skip_rows)
    elif suffix == ".csv":
        data_rows, header_row = _read_csv(filepath, skip_rows)
    else:
        raise ValueError(
            f"未対応のファイル形式: {suffix} (.xlsx または .csv に対応)"
        )

    # 列指定を解決
    required_fields = {"item_name", "jlac10"}
    optional_fields = {"abbreviation", "jlac10_standard_name"}
    all_fields = required_fields | optional_fields

    for field in required_fields:
        if field not in column_map:
            raise ValueError(f"必須フィールド '{field}' が column_map にありません")

    resolved: dict[str, int] = {}
    for field, spec in column_map.items():
        if field not in all_fields:
            logger.warning("不明なフィールド '%s' は無視します", field)
            continue
        resolved[field] = _resolve_column_index(spec, header_row)

    # データ変換
    items: list[dict] = []
    skipped = 0

    for row_num, row in enumerate(data_rows, start=skip_rows + 1):
        # 空行スキップ
        if not row or all(cell.strip() == "" for cell in row):
            skipped += 1
            continue

        # item_name が空なら空行扱い
        item_col = resolved["item_name"]
        if item_col >= len(row) or row[item_col].strip() == "":
            skipped += 1
            continue

        item_name = row[item_col].strip()

        # JLAC10 取得・正規化
        jlac10_col = resolved["jlac10"]
        raw_jlac10 = row[jlac10_col].strip() if jlac10_col < len(row) else ""
        jlac10 = raw_jlac10.replace("-", "")
        jlac10_status = classify_jlac10(jlac10)

        # analyte_code: JLAC10の先頭5桁
        analyte_code = jlac10[:5] if len(jlac10) >= 5 else ""

        # 任意フィールド
        abbreviation = ""
        if "abbreviation" in resolved:
            col = resolved["abbreviation"]
            abbreviation = row[col].strip() if col < len(row) else ""

        jlac10_standard_name = ""
        if "jlac10_standard_name" in resolved:
            col = resolved["jlac10_standard_name"]
            jlac10_standard_name = row[col].strip() if col < len(row) else ""

        items.append({
            "hospital": hospital,
            "item_name": item_name,
            "abbreviation": abbreviation,
            "jlac10": jlac10,
            "jlac10_status": jlac10_status,
            "analyte_code": analyte_code,
            "jlac10_standard_name": jlac10_standard_name,
        })

    logger.info(
        "変換完了: %d件 (空行スキップ: %d件, ソース: %s)",
        len(items), skipped, filepath.name,
    )

    result = {
        "metadata": {
            "hospital": hospital,
            "source_file": filepath.name,
            "exported_at": datetime.now(timezone.utc).isoformat(),
            "total_items": len(items),
        },
        "items": items,
    }

    # JSON出力
    if output_path is None:
        output_path = filepath.with_suffix(".json")
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    logger.info("JSON出力: %s", output_path)

    return result
