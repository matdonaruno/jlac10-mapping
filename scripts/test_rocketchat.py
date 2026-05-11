"""Rocket.Chat API テストスクリプト

使い方:
  環境変数で認証情報を渡す（スクリプトに直書きしない）

    export ROCKET_CHAT_URL="https://10.152.72.7"
    export ROCKET_CHAT_USER="your_username"
    export ROCKET_CHAT_PASSWORD="your_password"
    export ROCKET_CHAT_TEST_ROOM="umetani"   # プライベートグループ名
    python scripts/test_rocketchat.py

テスト内容:
  Step 1: ログイン（認証トークン取得）
  Step 2: 参加グループ一覧取得（プライベートグループ）
  Step 3: テストメッセージ投稿（TEST_ROOM が設定されていれば）
"""

import os
import sys
import requests

# =====================================================================
# 設定（環境変数から取得）
# =====================================================================
ROCKET_CHAT_URL = os.environ.get("ROCKET_CHAT_URL", "https://10.152.72.7")
USER = os.environ.get("ROCKET_CHAT_USER", "")
PASSWORD = os.environ.get("ROCKET_CHAT_PASSWORD", "")
TEST_ROOM = os.environ.get("ROCKET_CHAT_TEST_ROOM", "umetani")

# プライベートグループか公開チャンネルか
# /group/<name> ならプライベートグループ → groups.* API
# /channel/<name> なら公開チャンネル → channels.* API
ROOM_TYPE = os.environ.get("ROCKET_CHAT_ROOM_TYPE", "group")  # "group" or "channel"


def main():
    if not USER or not PASSWORD:
        print("環境変数 ROCKET_CHAT_USER / ROCKET_CHAT_PASSWORD を設定してください")
        print("例:")
        print('  export ROCKET_CHAT_USER="your_name"')
        print('  export ROCKET_CHAT_PASSWORD="your_password"')
        sys.exit(1)

    base = ROCKET_CHAT_URL.rstrip("/")
    print(f"Rocket.Chat: {base}")
    print(f"Room type:   {ROOM_TYPE}")
    print(f"Test room:   {TEST_ROOM or '(none)'}")

    # API エンドポイントを room type で切り替え
    list_api = "groups.listAll" if ROOM_TYPE == "group" else "channels.list.joined"
    list_key = "groups" if ROOM_TYPE == "group" else "channels"
    info_api = "groups.info" if ROOM_TYPE == "group" else "channels.info"
    info_key = "group" if ROOM_TYPE == "group" else "channel"

    # Step 1: ログイン
    print("\n--- Step 1: ログイン ---")
    try:
        resp = requests.post(
            f"{base}/api/v1/login",
            json={"user": USER, "password": PASSWORD},
            verify=False,  # オンプレ自己署名証明書対応
            timeout=15,
        )
        print(f"Status: {resp.status_code}")
        if resp.status_code != 200:
            print(f"ログイン失敗: {resp.text[:200]}")
            return

        data = resp.json()
        user_id = data["data"]["userId"]
        auth_token = data["data"]["authToken"]
        print(f"ログイン成功: userId={user_id}")
        print(f"authToken: {auth_token[:20]}...")

    except requests.exceptions.SSLError as e:
        print(f"SSL エラー: {e}")
        print("自己署名証明書の場合 verify=False で回避済みのはずです")
        return
    except Exception as e:
        print(f"接続エラー: {e}")
        print("URLが正しいか、ネットワークに接続できるか確認してください")
        return

    headers = {
        "X-Auth-Token": auth_token,
        "X-User-Id": user_id,
    }

    # Step 2: 参加グループ/チャンネル一覧
    print(f"\n--- Step 2: {list_api} ---")
    try:
        # groups.listAll は管理者権限が必要なので groups.list にフォールバック
        resp = requests.get(
            f"{base}/api/v1/{list_api}",
            headers=headers,
            verify=False,
            timeout=15,
        )
        if resp.status_code != 200 and ROOM_TYPE == "group":
            print(f"  {list_api} 失敗 ({resp.status_code}) → groups.list にフォールバック")
            resp = requests.get(
                f"{base}/api/v1/groups.list",
                headers=headers,
                verify=False,
                timeout=15,
            )

        if resp.status_code == 200:
            rooms = resp.json().get(list_key, [])
            print(f"取得件数: {len(rooms)}")
            for r in rooms[:20]:
                print(f"  {r.get('name')} (id={r.get('_id')})")
        else:
            print(f"一覧取得失敗: {resp.status_code} {resp.text[:200]}")
    except Exception as e:
        print(f"エラー: {e}")

    # Step 3: テストメッセージ投稿
    if TEST_ROOM:
        print(f"\n--- Step 3: テスト投稿 ({TEST_ROOM}) ---")
        try:
            # ルーム情報取得
            resp = requests.get(
                f"{base}/api/v1/{info_api}",
                headers=headers,
                params={"roomName": TEST_ROOM},
                verify=False,
                timeout=15,
            )
            if resp.status_code != 200:
                print(f"ルーム '{TEST_ROOM}' が見つかりません: {resp.status_code} {resp.text[:200]}")
                return

            room_id = resp.json()[info_key]["_id"]
            print(f"  room_id: {room_id}")

            # メッセージ投稿
            message = "JLAC10 MAPPER テストメッセージ（API 接続確認）"
            resp = requests.post(
                f"{base}/api/v1/chat.postMessage",
                headers=headers,
                json={"channel": f"#{TEST_ROOM}" if ROOM_TYPE == "channel" else f"@{TEST_ROOM}", "text": message, "roomId": room_id},
                verify=False,
                timeout=15,
            )
            if resp.status_code == 200 and resp.json().get("success"):
                print(f"投稿成功: {message}")
            else:
                print(f"投稿失敗: {resp.status_code} {resp.text[:300]}")
        except Exception as e:
            print(f"エラー: {e}")
    else:
        print("\n--- Step 3: スキップ（TEST_ROOM 未設定） ---")

    print("\n完了")


if __name__ == "__main__":
    # SSL 警告を抑制（自己署名証明書のため）
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    main()
