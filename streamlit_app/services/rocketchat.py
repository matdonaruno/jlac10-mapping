"""Rocket.Chat 投稿サービス。

PAT(auth_token + user_id) または user/password のどちらかで動作。
オンプレ自己署名証明書を想定して TLS 検証は無効化。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import requests
from urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)


@dataclass
class RocketChatResult:
    ok: bool
    message: str
    response: dict[str, Any] | None = None


def _login_with_password(url: str, user: str, password: str) -> tuple[str, str]:
    r = requests.post(
        f"{url}/api/v1/login",
        json={"user": user, "password": password},
        verify=False,
        timeout=10,
    )
    r.raise_for_status()
    data = r.json()["data"]
    return data["authToken"], data["userId"]


def _resolve_room(url: str, headers: dict[str, str], name_or_id: str) -> str:
    """rooms.info で roomName→roomId 解決。失敗したら roomId として再試行。"""
    for params in ({"roomName": name_or_id}, {"roomId": name_or_id}):
        r = requests.get(
            f"{url}/api/v1/rooms.info", headers=headers, params=params, verify=False, timeout=10
        )
        if r.ok and r.json().get("success"):
            return r.json()["room"]["_id"]
    raise RuntimeError(f"ROOM '{name_or_id}' が見つかりません（rooms.info 失敗）")


def post(
    url: str,
    room: str,
    text: str,
    *,
    auth_token: str = "",
    user_id: str = "",
    user: str = "",
    password: str = "",
) -> RocketChatResult:
    if not url or not room or not text:
        return RocketChatResult(False, "url / room / text は必須です")
    try:
        if auth_token and user_id:
            token, uid = auth_token, user_id
        elif user and password:
            token, uid = _login_with_password(url, user, password)
        else:
            return RocketChatResult(False, "認証情報が不足（auth_token+user_id または user+password）")

        headers = {"X-Auth-Token": token, "X-User-Id": uid}
        room_id = _resolve_room(url, headers, room)
        r = requests.post(
            f"{url}/api/v1/chat.postMessage",
            headers={**headers, "Content-type": "application/json"},
            json={"roomId": room_id, "text": text},
            verify=False,
            timeout=10,
        )
        r.raise_for_status()
        data = r.json()
        if not data.get("success"):
            return RocketChatResult(False, f"投稿失敗: {data}", data)
        return RocketChatResult(True, f"送信成功 (room={room})", data)
    except Exception as e:
        return RocketChatResult(False, f"エラー: {e}")


def test_connection(url: str, *, auth_token: str = "", user_id: str = "") -> RocketChatResult:
    if not (url and auth_token and user_id):
        return RocketChatResult(False, "url / auth_token / user_id を設定してください")
    try:
        r = requests.get(
            f"{url}/api/v1/me",
            headers={"X-Auth-Token": auth_token, "X-User-Id": user_id},
            verify=False,
            timeout=10,
        )
        if r.ok and r.json().get("success"):
            data = r.json()
            return RocketChatResult(True, f"接続OK: username={data.get('username')}", data)
        return RocketChatResult(False, f"接続失敗: HTTP {r.status_code}")
    except Exception as e:
        return RocketChatResult(False, f"エラー: {e}")
