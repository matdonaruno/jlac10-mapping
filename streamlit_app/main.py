"""ローカル自動化ハブ（Streamlit）エントリポイント。

起動:
  uv run streamlit run streamlit_app/main.py

秘匿情報: .streamlit/secrets.toml （gitignore対象）
契約スキーマ: docs/payload_schema.json
"""
from __future__ import annotations

import streamlit as st

from streamlit_app.config import load_config

st.set_page_config(
    page_title="NCDA ローカル自動化ハブ",
    page_icon="🧰",
    layout="wide",
)

st.title("NCDA ローカル自動化ハブ")
st.caption("Pages で付番 → JSONダウンロード → このアプリで3アクション発火")

cfg = load_config()

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("GitHub", "設定済み" if cfg.github.ready() else "未設定")
with col2:
    st.metric("Rocket.Chat", "設定済み" if cfg.rocketchat.ready() else "未設定")
with col3:
    n_pa = len(cfg.power_automate.webhooks)
    st.metric("Power Automate", f"{n_pa} Webhook" if n_pa else "未設定")

st.divider()

st.markdown(
    """
    ### 使い方
    1. **左サイドバーのページ** を開く
       - **Fire**: ペイロードJSONをアップロードして3アクション同時発火
       - **Stockpile**: JANIS縦積みに追記（`append_stockpile.py` のGUI）
       - **Settings**: 現在の設定値を確認（値の変更は `.streamlit/secrets.toml` を直接編集）
    2. **初回設定**: `.streamlit/secrets.toml.example` をコピーして `secrets.toml` を作成
    3. **Pages 側**: 付番完了後に JSON をダウンロードしてこのアプリにアップロード

    ### ファイル配置
    - 秘匿情報: `.streamlit/secrets.toml`（このリポジトリから除外）
    - サンプル: `.streamlit/secrets.toml.example`
    - スキーマ: `docs/payload_schema.json`
    """
)

if not (cfg.github.ready() or cfg.rocketchat.ready()):
    st.warning(
        "秘匿情報が未設定です。`.streamlit/secrets.toml.example` を参考に "
        "`.streamlit/secrets.toml` を作成してください。"
    )
