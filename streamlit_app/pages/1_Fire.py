"""3アクション同時発火ページ（骨格）。

実装は #22 で詳細化。今は骨格のみ。
"""
from __future__ import annotations

import json

import streamlit as st

from streamlit_app.config import load_config

st.set_page_config(page_title="Fire - NCDA Hub", layout="wide")
st.title("3アクション同時発火")
st.caption("Pages 付番結果（ペイロードJSON）→ GitHub Issue + Rocket.Chat + Power Automate")

cfg = load_config()

uploaded = st.file_uploader("ペイロードJSON", type=["json"])
if uploaded is None:
    st.info("Pages 側で「JSONダウンロード」したファイルをアップロードしてください。")
    st.stop()

try:
    payload = json.loads(uploaded.read())
except json.JSONDecodeError as e:
    st.error(f"JSONパース失敗: {e}")
    st.stop()

st.subheader("ペイロード概要")
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("バージョン", payload.get("version", "?"))
with col2:
    hospital = payload.get("hospital", {})
    st.metric("病院", (hospital.get("code", "") + hospital.get("name", "")) or "?")
with col3:
    st.metric("行数", len(payload.get("rows", [])))

with st.expander("マッピング結果テーブル", expanded=True):
    st.dataframe(payload.get("rows", []), use_container_width=True)

tabs = st.tabs(["GitHub プレビュー", "Rocket.Chat プレビュー", "Power Automate"])
with tabs[0]:
    st.markdown(payload.get("templates", {}).get("github_markdown", ""))
with tabs[1]:
    st.code(payload.get("templates", {}).get("rocketchat", ""), language="text")
with tabs[2]:
    pa = payload.get("power_automate", {})
    if pa:
        st.json(pa)
    else:
        st.caption("Power Automate 情報なし")

st.divider()
st.warning("🚧 発火ボタンは #22 で実装予定。現在はプレビューのみ。")

col_a, col_b, col_c = st.columns(3)
with col_a:
    st.button("GitHub Issue へコメント投稿", disabled=True, help="#22で実装")
with col_b:
    st.button("Rocket.Chat へ通知", disabled=True, help="#22で実装")
with col_c:
    st.button("Power Automate Webhook", disabled=True, help="#22で実装")
