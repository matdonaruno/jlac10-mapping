# SN Error ダッシュボード設計

## 概要

SN Error タブをイシュー番号ベースの案件管理ダッシュボードに変更。
エラー検出から付番完了・出力確認までの全フェーズを1つのタブで管理。

## 案件データ構造（LocalStorage に保存）

```json
{
  "cases": {
    "#4123": {
      "issue_number": 4123,
      "hospital_name": "北海道がんセンター",
      "hospital_id": "H100",
      "contact": "田中様",
      "cc": "鈴木様",
      "created_at": "2026-04-18T10:00:00",
      "updated_at": "2026-04-18T15:30:00",
      "status": "mapping",
      "error_items": [
        {"code": "1001", "name": "総蛋白", "status": "assignable",
         "jlac10": "3A010000002327101", "matched_name": "総蛋白_血清_可視吸光度法_定量値"},
        {"code": "9999", "name": "特殊検査X", "status": "inquiry", "jlac10": "", "matched_name": ""}
      ],
      "history": [
        {"date": "2026-04-18T10:00:00", "action": "created", "note": "エラー起票"},
        {"date": "2026-04-18T12:00:00", "action": "analyzed", "note": "5件中3件付番可能"},
        {"date": "2026-04-18T13:00:00", "action": "mail_sent", "note": "設定依頼メール送信"},
        {"date": "2026-04-18T15:30:00", "action": "inquiry_sent", "note": "2件について問い合わせ"}
      ],
      "workflow": {
        "mail_subject": "北海道がんセンター_2026-04-18_#4123",
        "mail_sent_at": "2026-04-18T13:00:00",
        "ghe_comment_posted": true,
        "rocketchat_posted": true
      }
    }
  }
}
```

## ステータス遷移

```
open → mapping → waiting → verifying → closed
                    ↑          │
                    └── NG ────┘

open:      エラー起票済み、未対応
mapping:   マッピング作業中（問い合わせ含む）
waiting:   病院へ設定依頼済み、返答待ち
verifying: 病院設定完了、出力確認中
closed:    解決
```

## UI レイアウト

```
┌─────────────────────────────────────────────────────────┐
│ SN Error Dashboard                                       │
│                                                          │
│ [+ New Case]  [Import from Excel]                        │
│                                                          │
│ Filter: [All ▼] [Open(3)] [Mapping(2)] [Waiting(5)] ... │
│                                                          │
│ ┌──────────────────────────────────────────────────────┐ │
│ │ #4123 | 北海道がんセンター | mapping | 5件 | 04-18   │ │
│ │ #4120 | 東京大学病院       | waiting | 12件 | 04-15  │ │
│ │ #4115 | 大阪医療センター   | closed  | 3件 | 04-10   │ │
│ └──────────────────────────────────────────────────────┘ │
│                                                          │
│ === Case Detail (#4123) ===                              │
│                                                          │
│ Hospital: 北海道がんセンター  Status: [mapping ▼]        │
│ Contact: 田中様  Issue: #4123  Date: 2026-04-18          │
│                                                          │
│ ┌─ Error Items ─────────────────────────────────────── │
│ │ ✅ 1001 総蛋白     3A010... 付番完了                  │ │
│ │ ✅ 1011 LDH        3B0200... 付番完了                 │ │
│ │ ?? 9999 特殊検査X  -         問い合わせ中              │ │
│ └──────────────────────────────────────────────────────┘ │
│                                                          │
│ ┌─ Actions ─────────────────────────────────────────── │
│ │ [Analyze New Errors]  [Generate Email]                │ │
│ │ [Generate GHE Comment] [Generate Chat Message]        │ │
│ │ [Add Note]  [Change Status]                           │ │
│ └──────────────────────────────────────────────────────┘ │
│                                                          │
│ ┌─ History ─────────────────────────────────────────── │
│ │ 04-18 15:30  問い合わせメール送信（2件）              │ │
│ │ 04-18 13:00  設定依頼メール送信（3件付番完了）        │ │
│ │ 04-18 12:00  エラー分析完了（5件中3件付番可能）       │ │
│ │ 04-18 10:00  エラー起票                               │ │
│ └──────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────┘
```

## LocalStorage キー

```
sn_dashboard_cases    → 案件データ全体 (JSON)
sn_dashboard_settings → 設定（デフォルト病院名等）
```

## 実装方針

### Phase A: 案件管理基盤
- 案件の作成・編集・削除
- ステータス管理
- LocalStorage 永続化
- 案件一覧表示（フィルタ付き）

### Phase B: エラー分析統合
- 既存の analyzeErrors を案件に紐付け
- エラー項目を案件に保存
- 項目ごとのステータス管理

### Phase C: ワークフロー統合
- メール/GHEコメント/Rocket.Chat 生成を案件から実行
- 生成履歴を案件に記録

### Phase D: エクスポート/インポート
- 案件データの JSON エクスポート（バックアップ）
- インポート（PC 移行時）
- GHE Issues との同期（API 確認後）
