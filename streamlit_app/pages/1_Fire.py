"""3アクション同時発火ページ。

ペイロードJSON をアップロード → プレビュー → 各アクション個別発火 / 一括発火。
"""
from __future__ import annotations

import json

import streamlit as st

from streamlit_app.config import load_config
from streamlit_app.services import github as gh
from streamlit_app.services import power_automate as pa
from streamlit_app.services import rocketchat as rc

st.set_page_config(page_title="Fire - NCDA Hub", layout="wide")
st.title("3アクション同時発火")
st.caption("ペイロードJSON → GitHub Issue コメント + Rocket.Chat 通知 + Power Automate Webhook")

cfg = load_config()

# ----- ペイロード読込 -----
uploaded = st.file_uploader("ペイロードJSON（Pages から書き出し）", type=["json"])
if uploaded is None:
    st.info("Pages 側で「ペイロードJSON ダウンロード」したファイルをアップロードしてください。")
    st.stop()

try:
    payload = json.loads(uploaded.read())
except json.JSONDecodeError as e:
    st.error(f"JSONパース失敗: {e}")
    st.stop()

# ----- 概要 -----
hospital = payload.get("hospital", {})
hosp_label = (hospital.get("code") or "") + (hospital.get("name") or "")
github_meta = payload.get("github") or {}
templates = payload.get("templates") or {}
pa_meta = payload.get("power_automate") or {}

c1, c2, c3, c4 = st.columns(4)
with c1: st.metric("バージョン", payload.get("version", "?"))
with c2: st.metric("病院", hosp_label or "?")
with c3: st.metric("行数", len(payload.get("rows", [])))
with c4: st.metric("Issue", github_meta.get("issue_number") or "—")

with st.expander("マッピング行（rows）", expanded=False):
    st.dataframe(payload.get("rows", []), use_container_width=True)

# ----- プレビュー -----
tabs = st.tabs(["GitHub プレビュー", "Rocket.Chat プレビュー", "Power Automate"])
github_body = templates.get("github_markdown", "")
rocketchat_body = templates.get("rocketchat", "")

with tabs[0]:
    if github_body:
        st.markdown(github_body)
        github_body_edit = st.text_area("送信本文（編集可）", value=github_body, height=180, key="gh_body")
    else:
        st.caption("templates.github_markdown が空です")
        github_body_edit = ""

with tabs[1]:
    if rocketchat_body:
        st.code(rocketchat_body, language="text")
        rc_body_edit = st.text_area("送信本文（編集可）", value=rocketchat_body, height=120, key="rc_body")
    else:
        st.caption("templates.rocketchat が空です")
        rc_body_edit = ""

with tabs[2]:
    pa_payload = pa_meta.get("payload") or {}
    pa_key = pa_meta.get("webhook_key") or ""
    pa_url = cfg.power_automate.webhooks.get(pa_key, "") if pa_key else ""
    st.text_input("Webhook key", value=pa_key, disabled=True, key="pa_key_disp")
    if pa_key and not pa_url:
        st.warning(f"`secrets.toml` に [power_automate].{pa_key} が未設定")
    elif pa_url:
        st.caption("Webhook URL: 設定済み (マスク)")
    if pa_payload:
        st.json(pa_payload)
    else:
        st.caption("payload 空 → 空のJSONを送信します")

st.divider()

# ----- 接続 readiness -----
col1, col2, col3 = st.columns(3)
with col1:
    st.write("**GitHub**", "✅ Ready" if cfg.github.ready() else "❌ PAT 未設定")
with col2:
    st.write("**Rocket.Chat**", "✅ Ready" if cfg.rocketchat.ready() else "❌ 未設定")
with col3:
    st.write("**Power Automate**", "✅ Ready" if pa_url else "❌ Webhook未設定")

st.divider()

# ----- 発火ボタン群 -----
st.subheader("発火")

repo_default = github_meta.get("repo") or cfg.github.default_repo
issue_no = github_meta.get("issue_number") or 0
room_default = cfg.rocketchat.default_room

col_a, col_b, col_c = st.columns(3)
with col_a:
    repo_in = st.text_input("repo (owner/name)", value=repo_default, key="gh_repo")
    issue_in = st.number_input("Issue番号", value=int(issue_no), step=1, key="gh_issue")
    if st.button("GitHub Issue へ投稿", disabled=not cfg.github.ready(), use_container_width=True):
        with st.spinner("投稿中..."):
            res = gh.post_issue_comment(cfg.github.token, repo_in, int(issue_in), github_body_edit)
        if res.ok:
            st.success(res.message)
            if res.html_url:
                st.markdown(f"[👉 コメントを開く]({res.html_url})")
                st.session_state["last_github_url"] = res.html_url
        else:
            st.error(res.message)

with col_b:
    room_in = st.text_input("ROOM 名 or ID", value=room_default, key="rc_room")
    if st.button("Rocket.Chat へ投稿", disabled=not cfg.rocketchat.ready(), use_container_width=True):
        # Issue投稿後なら、permalink を Rocket.Chat 本文末尾に差し込み
        body_to_send = rc_body_edit
        last_url = st.session_state.get("last_github_url")
        if last_url and last_url not in body_to_send:
            body_to_send = body_to_send + "\n" + last_url
        with st.spinner("投稿中..."):
            res = rc.post(
                url=cfg.rocketchat.url,
                room=room_in,
                text=body_to_send,
                auth_token=cfg.rocketchat.auth_token,
                user_id=cfg.rocketchat.user_id,
                user=cfg.rocketchat.user,
                password=cfg.rocketchat.password,
            )
        if res.ok:
            st.success(res.message)
        else:
            st.error(res.message)

with col_c:
    st.text_input("webhook key", value=pa_key, disabled=True, key="pa_key_show")
    if st.button("Power Automate 発火", disabled=not pa_url, use_container_width=True):
        with st.spinner("発火中..."):
            res = pa.fire(pa_url, pa_payload)
        if res.ok:
            st.success(res.message)
            if res.response_text:
                st.caption(f"レスポンス: {res.response_text}")
        else:
            st.error(res.message)

st.divider()

# ----- 一括発火（GitHub → Rocket.Chat → Power Automate の順） -----
if st.button("🚀 全て順次発火", type="primary", use_container_width=True):
    log = []

    # GitHub
    if cfg.github.ready() and repo_in and issue_in:
        with st.spinner("GitHub 投稿中..."):
            r = gh.post_issue_comment(cfg.github.token, repo_in, int(issue_in), github_body_edit)
        log.append(("GitHub", r.ok, r.message, r.html_url if r.ok else ""))
        gh_url = r.html_url if r.ok else ""
    else:
        log.append(("GitHub", False, "未設定/未指定のためスキップ", ""))
        gh_url = ""

    # Rocket.Chat (gh_url を本文末尾に差し込み)
    if cfg.rocketchat.ready() and room_in:
        body = rc_body_edit
        if gh_url and gh_url not in body:
            body = body + "\n" + gh_url
        with st.spinner("Rocket.Chat 投稿中..."):
            r = rc.post(
                url=cfg.rocketchat.url, room=room_in, text=body,
                auth_token=cfg.rocketchat.auth_token, user_id=cfg.rocketchat.user_id,
                user=cfg.rocketchat.user, password=cfg.rocketchat.password,
            )
        log.append(("Rocket.Chat", r.ok, r.message, ""))
    else:
        log.append(("Rocket.Chat", False, "未設定/未指定のためスキップ", ""))

    # Power Automate
    if pa_url:
        with st.spinner("Power Automate 発火中..."):
            r = pa.fire(pa_url, pa_payload)
        log.append(("Power Automate", r.ok, r.message, ""))
    else:
        log.append(("Power Automate", False, "未設定のためスキップ", ""))

    st.subheader("結果")
    for name, ok, msg, link in log:
        icon = "✅" if ok else "❌"
        st.write(f"{icon} **{name}** — {msg}")
        if link:
            st.markdown(f"  [{link}]({link})")
