"""設定確認ページ。

値の編集は .streamlit/secrets.toml を直接編集する想定（非エンジニアでも可）。
"""
from __future__ import annotations

from pathlib import Path

import streamlit as st

from streamlit_app.config import load_config

SECRETS_EXAMPLE = Path(".streamlit/secrets.toml.example")

st.set_page_config(page_title="Settings - NCDA Hub", layout="wide")
st.title("設定確認")
st.caption("値の変更は `.streamlit/secrets.toml` を直接編集してください")

cfg = load_config()


def _mask(s: str, visible: int = 4) -> str:
    if not s:
        return "(未設定)"
    if len(s) <= visible:
        return "*" * len(s)
    return s[:visible] + "*" * (len(s) - visible)


st.subheader("GitHub")
col1, col2 = st.columns(2)
with col1:
    st.text(f"token:  {_mask(cfg.github.token)}")
    st.text(f"repo:   {cfg.github.default_repo or '(未設定)'}")
with col2:
    st.metric("状態", "設定済み" if cfg.github.ready() else "未設定")

st.subheader("Rocket.Chat")
col1, col2 = st.columns(2)
with col1:
    st.text(f"url:         {cfg.rocketchat.url or '(未設定)'}")
    st.text(f"auth_token:  {_mask(cfg.rocketchat.auth_token)}")
    st.text(f"user_id:     {_mask(cfg.rocketchat.user_id)}")
    st.text(f"user:        {cfg.rocketchat.user or '(未設定)'}")
    st.text(f"password:    {_mask(cfg.rocketchat.password)}")
with col2:
    st.text(f"default_room:    {cfg.rocketchat.default_room or '(未設定)'}")
    st.text(f"default_mention: {cfg.rocketchat.default_mention}")
    st.metric("状態", "設定済み" if cfg.rocketchat.ready() else "未設定")

st.subheader("Power Automate Webhooks")
if cfg.power_automate.webhooks:
    for k, v in cfg.power_automate.webhooks.items():
        st.text(f"{k}:  {_mask(v, 30)}")
else:
    st.text("(未設定)")

st.divider()
st.markdown("### secrets.toml サンプル")
if SECRETS_EXAMPLE.exists():
    st.code(SECRETS_EXAMPLE.read_text(encoding="utf-8"), language="toml")
else:
    st.caption(f"{SECRETS_EXAMPLE} が見つかりません")
