"""縦積み/Lookup JSON にマッピング履歴を追記する（縦積み自動成長）。

CLI:
  # JANIS 菌名
  uv run python scripts/append_stockpile.py species \
      --inhouse "E.coli (ESBL+)" \
      --name "Escherichia coli" \
      --code 2002

  # JANIS 抗菌薬
  uv run python scripts/append_stockpile.py antibiotics \
      --inhouse "*PIPC 100" \
      --abbr "PIPC" \
      --name "ピペラシリン(PIPC)" \
      --code 1266

  # 細菌検査材料
  uv run python scripts/append_stockpile.py material \
      --name "腹水(検査用)" \
      --code 043 \
      --standard "腹水"

方針:
  - 同じ inhouse が既にある場合は既定でスキップ（--force で強制追記）
  - ファイル書き込みは 一時ファイル経由で atomic replace（中断対策）
  - 追加後、count / version を更新
  - Git 履歴で変更追跡できるため独自バックアップは取らない
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
from datetime import date
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
TODAY = date.today().isoformat()

TARGETS = {
    "species": DATA / "janis_species.json",
    "antibiotics": DATA / "janis_antibiotics.json",
    "material": DATA / "bact_materials.json",
}


def _load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _save_atomic(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(
        prefix=path.stem + ".", suffix=".tmp", dir=str(path.parent)
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.chmod(tmp, 0o644)
        os.replace(tmp, path)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def _zfill(code: str, width: int) -> str:
    s = str(code).strip()
    if not s:
        return ""
    try:
        return str(int(s)).zfill(width)
    except ValueError:
        return s


def _finalize(payload: dict) -> dict:
    payload["version"] = TODAY
    payload["count"] = len(payload["entries"])
    return payload


def _exists(entries: list[dict], key_field: str, value: str) -> bool:
    return any(e.get(key_field) == value for e in entries)


def append_species(inhouse: str, name: str, code: str, note: str, force: bool) -> int:
    path = TARGETS["species"]
    payload = _load(path)
    janis_code = _zfill(code, 4)
    if not force and _exists(payload["entries"], "inhouse", inhouse):
        print(f"[skip] species: inhouse='{inhouse}' は既に存在（--force で強制追記）")
        return 0
    payload["entries"].append({
        "inhouse": inhouse,
        "janis_name": name,
        "janis_code": janis_code,
        "note": note,
    })
    _save_atomic(path, _finalize(payload))
    print(f"[add] species: '{inhouse}' → {name} ({janis_code})")
    return 1


def append_antibiotics(
    inhouse: str, abbr: str, name: str, code: str, note: str, force: bool
) -> int:
    path = TARGETS["antibiotics"]
    payload = _load(path)
    janis_code = _zfill(code, 4)
    if not force and _exists(payload["entries"], "inhouse", inhouse):
        print(f"[skip] antibiotics: inhouse='{inhouse}' は既に存在（--force で強制追記）")
        return 0
    payload["entries"].append({
        "inhouse": inhouse,
        "janis_abbr": abbr,
        "janis_name": name,
        "janis_code": janis_code,
        "note": note,
    })
    _save_atomic(path, _finalize(payload))
    print(f"[add] antibiotics: '{inhouse}' → {name} ({janis_code})")
    return 1


def append_material(name: str, code: str, standard: str, force: bool) -> int:
    path = TARGETS["material"]
    payload = _load(path)
    jlac10_material = _zfill(code, 3) if code.lower() != "xxx" else "xxx"
    if not force and _exists(payload["entries"], "material_name", name):
        print(f"[skip] material: material_name='{name}' は既に存在（--force で強制追記）")
        return 0
    payload["entries"].append({
        "material_name": name,
        "jlac10_material": jlac10_material,
        "jlac10_standard_name": standard,
    })
    _save_atomic(path, _finalize(payload))
    print(f"[add] material: '{name}' → {jlac10_material} ({standard})")
    return 1


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="細菌検査 縦積み/Lookup JSON に追記")
    sub = p.add_subparsers(dest="kind", required=True)

    sp = sub.add_parser("species", help="JANIS 菌名縦積みに追記")
    sp.add_argument("--inhouse", required=True, help="院内表記（表記ゆれそのまま）")
    sp.add_argument("--name", required=True, help="JANIS 菌名")
    sp.add_argument("--code", required=True, help="JANIS 菌名コード（4桁）")
    sp.add_argument("--note", default="", help="備考")
    sp.add_argument("--force", action="store_true", help="既存でも強制追記")

    ab = sub.add_parser("antibiotics", help="JANIS 抗菌薬縦積みに追記")
    ab.add_argument("--inhouse", required=True, help="院内表記")
    ab.add_argument("--abbr", required=True, help="JANIS 抗菌薬略号")
    ab.add_argument("--name", required=True, help="JANIS 抗菌薬名（日本語）")
    ab.add_argument("--code", required=True, help="JANIS 抗菌薬コード（4桁）")
    ab.add_argument("--note", default="", help="備考")
    ab.add_argument("--force", action="store_true", help="既存でも強制追記")

    mt = sub.add_parser("material", help="細菌検査材料 Lookup に追記")
    mt.add_argument("--name", required=True, help="病院側の材料名")
    mt.add_argument("--code", required=True, help="JLAC10材料コード（3桁 or 'xxx'）")
    mt.add_argument("--standard", required=True, help="JLAC10標準名称")
    mt.add_argument("--force", action="store_true", help="既存でも強制追記")

    return p


def main() -> int:
    args = build_parser().parse_args()
    if args.kind == "species":
        added = append_species(args.inhouse, args.name, args.code, args.note, args.force)
    elif args.kind == "antibiotics":
        added = append_antibiotics(
            args.inhouse, args.abbr, args.name, args.code, args.note, args.force
        )
    elif args.kind == "material":
        added = append_material(args.name, args.code, args.standard, args.force)
    else:
        print("不明な種別", file=sys.stderr)
        return 2
    return 0 if added >= 0 else 1


if __name__ == "__main__":
    sys.exit(main())
