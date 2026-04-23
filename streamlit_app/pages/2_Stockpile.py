"""縦積み追記ページ（骨格）。

scripts/append_stockpile.py の GUI化。実装詳細は #23 で対応。
"""
from __future__ import annotations

import streamlit as st

st.set_page_config(page_title="Stockpile - NCDA Hub", layout="wide")
st.title("JANIS 縦積み追記")
st.caption("data/janis_species.json / janis_antibiotics.json / bact_materials.json へ追記")

kind = st.selectbox("追記対象", ["species（JANIS菌名）", "antibiotics（JANIS抗菌薬）", "material（細菌検査材料）"])

st.warning("🚧 追記処理は #23 で実装予定。現状はフォームのみ。")

if kind.startswith("species"):
    st.text_input("院内表記（inhouse）")
    st.text_input("JANIS菌名（janis_name）")
    st.text_input("JANISコード（4桁）")
    st.text_input("備考", value="")
elif kind.startswith("antibiotics"):
    st.text_input("院内表記（inhouse）")
    st.text_input("JANIS略号（janis_abbr）")
    st.text_input("JANIS抗菌薬名（janis_name）")
    st.text_input("JANISコード（4桁）")
    st.text_input("備考", value="")
else:
    st.text_input("材料名")
    st.text_input("JLAC10材料コード（3桁 or 'xxx'）")
    st.text_input("JLAC10標準名称")

st.checkbox("既存と重複しても強制追記（--force）")
st.button("追記実行", disabled=True)
