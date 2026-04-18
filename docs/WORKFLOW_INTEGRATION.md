# GHE × Microsoft 365 連携設計

## 概要

GHE（オンプレ）と Microsoft 365（クラウド）は直接 API 連携ができないため、
ローカル PC を中継して両者のデータを橋渡しする。

## システム構成

```
GHE (オンプレ)              ローカル PC               Microsoft 365 (クラウド)
├── Issues (エラー管理)      ├── ブラウザ              ├── Outlook (メール)
├── Pages (JLAC10 MAPPER)   │   ├── GHE タブ          ├── OneDrive (ファイル中継)
├── Repository (コード/DB)   │   └── M365 タブ         ├── Power Automate (自動化)
└── Rocket.Chat (チャット)   └── コピペで中継           └── Excel Online (進捗管理)
```

## データフロー

```
マッピング完了
     │
     ▼
Pages で一括生成
├── ① メール本文             → Outlook で送信
├── ② GHE コメント用テキスト  → GHE Issues に貼り付け
├── ③ Rocket.Chat 用テキスト  → Rocket.Chat に貼り付け
└── ④ 中継 JSON              → OneDrive に保存（Power Automate 連携用）
```

## 中継フォルダ構成（OneDrive）

```
OneDrive/
└── JLAC10_WORKFLOW/
    ├── pending/              ← 送信待ち
    │   └── H100_2026-04-18_#42.json
    ├── sent/                 ← 送信済み
    ├── replied/              ← 病院返信あり
    ├── closed/               ← 解決済み
    └── templates/
        ├── mail_template.txt   ← メールテンプレート
        └── contacts.json       ← 病院別宛先設定
```

## 中継 JSON フォーマット

```json
{
  "hospital_name": "北海道がんセンター",
  "hospital_id": "H100",
  "issue_number": 42,
  "issue_url": "https://ghe.ncda.hospital.local/.../issues/42",
  "occurred_at": "2026-04-18",
  "to": "tanaka@hospital.jp",
  "cc": ["suzuki@hospital.jp"],
  "subject": "北海道がんセンター_2026-04-18_#42",
  "mapping_results": [
    {
      "item_code": "1001",
      "item_name": "総蛋白",
      "jlac10": "3A010000002327101",
      "standard_name": "総蛋白_血清_可視吸光度法_定量値"
    }
  ],
  "status": "pending",
  "mail_body": "（生成済みメール本文）",
  "ghe_comment": "マッピング完了、病院へ設定依頼済み",
  "rocketchat_message": "@tanaka マッピング完了しました。..."
}
```

---

## 段階的導入計画

### Step 1: Pages でテキスト一括生成（API不要）

**前提:** ブラウザのみ、コピペで運用

```
Pages SN Error タブ:
  マッピング完了後 → [Generate Workflow Package] ボタン

  入力:
    - 病院名、担当者名
    - イシュー番号
    - マッピング結果（分析結果から自動取得）

  出力（各ボタンで個別コピー）:
    [Copy Email]     → Outlook の新規メールに貼り付けて送信
    [Copy GHE Comment] → GHE Issues のコメント欄に貼り付け
    [Copy Chat Message] → Rocket.Chat に貼り付け
    [Download JSON]  → OneDrive に保存（Step 2 用）
```

**手動ステップ:**
1. Pages でテキスト生成
2. Outlook でメール送信（コピペ）
3. GHE でコメント追加（コピペ）
4. Rocket.Chat でメッセージ投稿（コピペ）

### Step 2: Power Automate でメール送信自動化

**前提:** Power Automate クラウドフロー利用可能

```
フロー: OneDrive 監視 → メール自動送信
  トリガー: OneDrive/JLAC10_WORKFLOW/pending/ にファイル追加
  ステップ:
    1. JSON ファイルを読み込み
    2. メール本文をパース
    3. Outlook コネクタで送信
    4. JSON を sent/ に移動
    5. ステータスを sent に更新
```

**手動ステップ:**
1. Pages でテキスト生成
2. JSON を OneDrive/pending/ に保存 → メール自動送信
3. GHE でコメント追加（コピペ）
4. Rocket.Chat でメッセージ投稿（コピペ）

### Step 3: 返信監視 + 進捗管理

**前提:** Power Automate で Outlook トリガーが使用可能

```
フロー1: 返信監視
  トリガー: Outlook で特定件名パターンのメール受信
  ステップ:
    1. 件名から病院名・イシュー番号を抽出
    2. JSON を replied/ に移動
    3. 進捗 Excel を更新

フロー2: 進捗レポート（週次）
  トリガー: スケジュール（毎週月曜）
  ステップ:
    1. pending/sent/replied/closed の件数集計
    2. Excel に出力
    3. メールでレポート送信
```

### Step 4: GHE API + Rocket.Chat API（完全自動化）

**前提:** GHE API / Rocket.Chat API が利用可能

```
Pages から直接:
  - GHE: イシュー作成、コメント追加、ステータス変更
  - Rocket.Chat: メッセージ投稿、メンション
  - 全てワンクリックで完了
```

---

## 確認事項

| Step | 確認内容 | ステータス |
|------|---------|-----------|
| 1 | メールテンプレートの確定 | TODO |
| 1 | GHE コメントのフォーマット | TODO |
| 1 | Rocket.Chat メッセージのフォーマット | TODO |
| 2 | Power Automate クラウドフロー作成可能か | TODO |
| 2 | Power Automate → Outlook コネクタ利用可能か | TODO |
| 3 | Power Automate → Outlook トリガー（受信監視）可能か | TODO |
| 4 | GHE API アクセス確認 | TODO |
| 4 | Rocket.Chat API / Webhook 確認 | TODO |
