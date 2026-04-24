"""縦積み追記ページ。

scripts/append_stockpile.py の CLI 関数を直接呼び出して GUI 化。
data/janis_species.json / janis_antibiotics.json / bact_materials.json に追記する。

注意: 追記は即座にローカルファイルへ書き込まれる。リモートへは git commit/push で反映。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parents[2]
SCRIPTS = ROOT / "scripts"
DATA = ROOT / "data"

if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import append_stockpile as stk  # type: ignore[import-not-found]

st.set_page_config(page_title="Stockpile - NCDA Hub", layout="wide")
st.title("JANIS 縦積み / 細菌材料 追記")
st.caption("表記ゆれを縦積みDBに蓄積。SOP: 病院マスタは修正せず、院内表記をそのまま積む")


def _count(path: Path) -> int:
    if not path.exists():
        return 0
    try:
        return int(json.loads(path.read_text(encoding="utf-8")).get("count", 0))
    except Exception:
        return 0


def _recent(path: Path, n: int = 5) -> list[dict]:
    if not path.exists():
        return []
    try:
        entries = json.loads(path.read_text(encoding="utf-8")).get("entries", [])
        return entries[-n:]
    except Exception:
        return []


TARGETS = {
    "species":     {"label": "🦠 JANIS 菌名",       "path": DATA / "janis_species.json"},
    "antibiotics": {"label": "💊 JANIS 抗菌薬",     "path": DATA / "janis_antibiotics.json"},
    "material":    {"label": "🧫 細菌検査材料",      "path": DATA / "bact_materials.json"},
}

# 概況メトリクス
cols = st.columns(3)
for col, (key, info) in zip(cols, TARGETS.items()):
    with col:
        st.metric(info["label"], f"{_count(info['path']):,} 件")

st.divider()

kind = st.radio(
    "追記対象",
    options=list(TARGETS.keys()),
    format_func=lambda k: TARGETS[k]["label"],
    horizontal=True,
)
target = TARGETS[kind]
st.caption(f"追記先: `{target['path'].relative_to(ROOT)}`")


# ----- 入力フォーム -----
added = None
skipped = None
err_msg = None

if kind == "species":
    with st.form("form_species", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            inhouse = st.text_input("院内表記 (inhouse)", help="表記ゆれ含めそのまま入力")
            name    = st.text_input("JANIS 菌名 (janis_name)", help="基準表記")
        with c2:
            code = st.text_input("JANIS コード (4桁)", help="例: 2002")
            note = st.text_input("備考", value="")
        force = st.checkbox("既存でも強制追記 (--force)")
        submitted = st.form_submit_button("追記実行", type="primary")
    if submitted:
        if not (inhouse and name and code):
            err_msg = "inhouse / name / code は必須です"
        else:
            try:
                n = stk.append_species(inhouse, name, code, note, force)
                added, skipped = (n, 1 - n)
            except Exception as e:
                err_msg = f"書き込み失敗: {e}"

elif kind == "antibiotics":
    with st.form("form_antibiotics", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            inhouse = st.text_input("院内表記 (inhouse)", help="濃度・マーカー混入そのまま")
            abbr    = st.text_input("JANIS 略号 (janis_abbr)", help="例: PIPC")
            name    = st.text_input("JANIS 抗菌薬名 (janis_name)", help="例: ピペラシリン(PIPC)")
        with c2:
            code = st.text_input("JANIS コード (4桁)")
            note = st.text_input("備考", value="")
        force = st.checkbox("既存でも強制追記 (--force)")
        submitted = st.form_submit_button("追記実行", type="primary")
    if submitted:
        if not (inhouse and abbr and name and code):
            err_msg = "inhouse / abbr / name / code は必須です"
        else:
            try:
                n = stk.append_antibiotics(inhouse, abbr, name, code, note, force)
                added, skipped = (n, 1 - n)
            except Exception as e:
                err_msg = f"書き込み失敗: {e}"

else:  # material
    with st.form("form_material", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            name     = st.text_input("材料名", help="病院側の材料名")
            code     = st.text_input("JLAC10 材料コード (3桁 or 'xxx')", help="運用上自由に組合可なら 'xxx'")
        with c2:
            standard = st.text_input("JLAC10 標準名称")
        force = st.checkbox("既存でも強制追記 (--force)")
        submitted = st.form_submit_button("追記実行", type="primary")
    if submitted:
        if not (name and code and standard):
            err_msg = "name / code / standard は必須です"
        else:
            try:
                n = stk.append_material(name, code, standard, force)
                added, skipped = (n, 1 - n)
            except Exception as e:
                err_msg = f"書き込み失敗: {e}"


# ----- 結果表示 -----
if err_msg:
    st.error(err_msg)
elif added == 1:
    st.success(f"✅ 追記しました。現在件数: {_count(target['path']):,}")
elif added == 0:
    st.warning("⚠ 既存と重複のためスキップしました（強制追記は --force）")

# 直近のエントリを表示
st.divider()
st.subheader("直近のエントリ（末尾5件）")
recent = _recent(target["path"], 5)
if recent:
    st.dataframe(recent, use_container_width=True)
else:
    st.caption("エントリなし")

st.caption(
    "追記後、`git diff data/*.json` で変更を確認し、コミットしてください。"
    " CLI 版: `uv run python scripts/append_stockpile.py {species|antibiotics|material} ...`"
)
