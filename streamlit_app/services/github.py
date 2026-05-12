"""GitHub Issue コメント投稿サービス（GHE 対応）。

API base URL:
  - GitHub.com:  https://api.github.com (デフォルト)
  - GHE:         https://<GHE_HOST>/api/v3
fine-grained PAT または Classic PAT どちらも `Authorization: Bearer <token>` で動作。
GHE で TLS 自己署名証明書を使っている場合は verify_tls=False を渡す。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import requests
from urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)


DEFAULT_API_BASE = "https://api.github.com"


@dataclass
class GitHubResult:
    ok: bool
    message: str
    html_url: str = ""
    response: dict[str, Any] | None = None


def _headers(token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def post_issue_comment(
    token: str,
    repo: str,
    issue_number: int,
    body: str,
    *,
    api_base: str = DEFAULT_API_BASE,
    verify_tls: bool = True,
) -> GitHubResult:
    if not token:
        return GitHubResult(False, "GitHub PAT 未設定（.streamlit/secrets.toml の [github].token）")
    if not repo or not issue_number or not body:
        return GitHubResult(False, "repo / issue_number / body は必須です")
    url = f"{api_base.rstrip('/')}/repos/{repo}/issues/{int(issue_number)}/comments"
    try:
        r = requests.post(
            url,
            headers=_headers(token),
            json={"body": body},
            timeout=15,
            verify=verify_tls,
        )
        if r.status_code == 201:
            data = r.json()
            return GitHubResult(True, "コメント投稿成功", data.get("html_url", ""), data)
        return GitHubResult(False, f"HTTP {r.status_code}: {r.text[:200]}", "", None)
    except Exception as e:
        return GitHubResult(False, f"エラー: {e}")


def test_token(
    token: str,
    *,
    api_base: str = DEFAULT_API_BASE,
    verify_tls: bool = True,
) -> GitHubResult:
    if not token:
        return GitHubResult(False, "PAT 未設定")
    try:
        r = requests.get(
            f"{api_base.rstrip('/')}/user",
            headers=_headers(token),
            timeout=10,
            verify=verify_tls,
        )
        if r.ok:
            data = r.json()
            return GitHubResult(True, f"接続OK: login={data.get('login')}", "", data)
        return GitHubResult(False, f"HTTP {r.status_code}: {r.text[:200]}")
    except Exception as e:
        return GitHubResult(False, f"エラー: {e}")


def get_issue(
    token: str,
    repo: str,
    issue_number: int,
    *,
    api_base: str = DEFAULT_API_BASE,
    verify_tls: bool = True,
) -> GitHubResult:
    if not token or not repo or not issue_number:
        return GitHubResult(False, "token / repo / issue_number は必須です")
    url = f"{api_base.rstrip('/')}/repos/{repo}/issues/{int(issue_number)}"
    try:
        r = requests.get(url, headers=_headers(token), timeout=10, verify=verify_tls)
        if r.ok:
            data = r.json()
            return GitHubResult(True, f"Issue取得OK: {data.get('title', '')}", data.get("html_url", ""), data)
        return GitHubResult(False, f"HTTP {r.status_code}: {r.text[:200]}")
    except Exception as e:
        return GitHubResult(False, f"エラー: {e}")
