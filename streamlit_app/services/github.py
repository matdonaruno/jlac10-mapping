"""GitHub Issue コメント投稿サービス。

fine-grained PAT (Issues: Write) を想定。
PAT 未設定の場合は ready=False を返し、UIで disabled 表示する。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import requests


@dataclass
class GitHubResult:
    ok: bool
    message: str
    html_url: str = ""
    response: dict[str, Any] | None = None


def post_issue_comment(
    token: str, repo: str, issue_number: int, body: str
) -> GitHubResult:
    if not token:
        return GitHubResult(False, "GitHub PAT 未設定（.streamlit/secrets.toml の [github].token）")
    if not repo or not issue_number or not body:
        return GitHubResult(False, "repo / issue_number / body は必須です")
    url = f"https://api.github.com/repos/{repo}/issues/{int(issue_number)}/comments"
    try:
        r = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Accept": "application/vnd.github+json",
                "X-GitHub-Api-Version": "2022-11-28",
            },
            json={"body": body},
            timeout=15,
        )
        if r.status_code == 201:
            data = r.json()
            return GitHubResult(True, "コメント投稿成功", data.get("html_url", ""), data)
        return GitHubResult(False, f"HTTP {r.status_code}: {r.text[:200]}", "", None)
    except Exception as e:
        return GitHubResult(False, f"エラー: {e}")


def test_token(token: str) -> GitHubResult:
    if not token:
        return GitHubResult(False, "PAT 未設定")
    try:
        r = requests.get(
            "https://api.github.com/user",
            headers={"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"},
            timeout=10,
        )
        if r.ok:
            data = r.json()
            return GitHubResult(True, f"接続OK: login={data.get('login')}", "", data)
        return GitHubResult(False, f"HTTP {r.status_code}: {r.text[:200]}")
    except Exception as e:
        return GitHubResult(False, f"エラー: {e}")
