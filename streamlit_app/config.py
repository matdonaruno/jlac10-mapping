"""Streamlit アプリ設定読み込み。

優先順位:
  1. st.secrets（`.streamlit/secrets.toml`、Streamlit 標準）
  2. 環境変数（RC_URL, RC_USER, GITHUB_TOKEN 等）
  3. 空（未設定）

非エンジニアは secrets.toml を直接編集する想定。
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Any

import streamlit as st


def _from_secrets(section: str, key: str, default: str = "") -> str:
    try:
        s = st.secrets[section]  # type: ignore[index]
        if isinstance(s, dict) and key in s:
            return str(s[key])
    except (KeyError, FileNotFoundError, Exception):
        pass
    return default


def _from_env(key: str, default: str = "") -> str:
    return os.environ.get(key, default)


@dataclass
class GitHubConfig:
    token: str = ""
    default_repo: str = ""

    def ready(self) -> bool:
        return bool(self.token)


@dataclass
class RocketChatConfig:
    url: str = ""
    auth_token: str = ""
    user_id: str = ""
    user: str = ""
    password: str = ""
    default_room: str = ""
    default_mention: str = "@koakutu"

    def ready(self) -> bool:
        if not self.url:
            return False
        # Token 方式 or user+password 方式のいずれか
        return bool(self.auth_token and self.user_id) or bool(self.user and self.password)


@dataclass
class PowerAutomateConfig:
    webhooks: dict[str, str] = field(default_factory=dict)


@dataclass
class AppConfig:
    github: GitHubConfig
    rocketchat: RocketChatConfig
    power_automate: PowerAutomateConfig


def load_config() -> AppConfig:
    github = GitHubConfig(
        token=_from_secrets("github", "token") or _from_env("GITHUB_TOKEN"),
        default_repo=_from_secrets("github", "default_repo") or _from_env("GITHUB_REPO"),
    )
    rc = RocketChatConfig(
        url=_from_secrets("rocketchat", "url") or _from_env("RC_URL"),
        auth_token=_from_secrets("rocketchat", "auth_token") or _from_env("RC_AUTH_TOKEN"),
        user_id=_from_secrets("rocketchat", "user_id") or _from_env("RC_USER_ID"),
        user=_from_secrets("rocketchat", "user") or _from_env("RC_USER"),
        password=_from_secrets("rocketchat", "password") or _from_env("RC_PASS"),
        default_room=_from_secrets("rocketchat", "default_room") or _from_env("RC_ROOM"),
        default_mention=_from_secrets("rocketchat", "default_mention") or "@koakutu",
    )
    # Power Automate: セクション内の全キーを Webhook URL として扱う
    webhooks: dict[str, str] = {}
    try:
        pa = st.secrets.get("power_automate", {})  # type: ignore[attr-defined]
        if isinstance(pa, dict):
            webhooks = {k: str(v) for k, v in pa.items() if v}
    except Exception:
        pass
    return AppConfig(
        github=github,
        rocketchat=rc,
        power_automate=PowerAutomateConfig(webhooks=webhooks),
    )
