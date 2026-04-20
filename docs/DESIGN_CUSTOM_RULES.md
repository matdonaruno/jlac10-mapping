# カスタムルール蓄積 + SOP 改版管理 設計

## 概要

マッピング業務中に発見されるルール（例外、追加禁止、同等測定法等）を
蓄積し、SOP 改版時に反映する仕組み。

## ルールの種類

| 種類 | 説明 | 例 |
|------|------|-----|
| 禁止 | このコード/パターンは使わない | 「材料099はこの検査では使わない」 |
| 推奨 | この場合はこのコードを使う | 「XX検査は測定法920で統一」 |
| 例外 | SOPルールの例外的許容 | 「この病院はベンダー仕様でXXを使用可」 |
| 同等 | 名称は違うが原理同じでOK | 「ラテックス凝集比濁法 = LA法 = コード062」 |
| 新規 | 新しい検査/試薬への対応 | 「XX検査は分析物YYYYYを使用」 |

## データ構造（LocalStorage + JSON エクスポート）

```json
{
  "custom_rules": [
    {
      "id": "R001",
      "type": "forbidden",
      "created_at": "2026-04-20",
      "created_by": "umetani",
      "sop_section": "④-4 (4)材料コード",
      "sop_revision_flag": true,
      "condition": {
        "field": "material",
        "analyte": "1A035",
        "code": "023"
      },
      "message": "pH[尿]の材料は001のみ（血清023は不可）",
      "rationale": "SOPルール④-4(4)に基づく",
      "note": "2026-04 に病院Xでこのパターンのエラーが発生",
      "status": "active"
    },
    {
      "id": "R002",
      "type": "equivalent",
      "created_at": "2026-04-20",
      "sop_section": "④-4 (5)測定法コード",
      "sop_revision_flag": true,
      "condition": {
        "field": "method",
        "method_names": ["ラテックス凝集比濁法", "LA法", "ラテックス比濁法"],
        "code": "062"
      },
      "message": "ラテックス凝集比濁法/LA法/ラテックス比濁法 → 062 で統一",
      "rationale": "名称は異なるが測定原理は同一",
      "note": "",
      "status": "active"
    }
  ]
}
```

## SOP セクション分類

| フラグ値 | SOPセクション |
|---------|--------------|
| `general` | ④-4 (1) 全般 |
| `analyte` | ④-4 (2) 分析物コード |
| `identification` | ④-4 (3) 識別コード |
| `material` | ④-4 (4) 材料コード |
| `method` | ④-4 (5) 測定法コード |
| `result_common` | ④-4 (6) 結果識別（共通） |
| `result_specific` | ④-4 (7) 結果識別（固有） |
| `bacteria` | 細菌検査 |
| `other` | その他 |

## 機能

### 1. ルール登録（Pages + CLI）
- マッピング中に発見 → 「ルール追加」ボタン
- 種類、条件、メッセージ、SOPセクション、根拠を入力
- LocalStorage に保存

### 2. ルール適用
- sopCheckJlac10() にカスタムルールも含めてチェック
- 固定ルール + カスタムルール を統合して評価

### 3. SOP 改版出力
- `sop_revision_flag: true` のルールを抽出
- SOPセクション別にグルーピング
- マークダウンまたは Word で出力
- 改版案として提出

### 4. ルール管理
- 一覧表示 / 編集 / 無効化
- エクスポート / インポート（JSON）
- 登録日、根拠、関連イシュー番号で検索

## 実装計画

Phase A: カスタムルールの CRUD + LocalStorage
Phase B: sopCheckJlac10 にカスタムルール統合
Phase C: SOP改版出力機能
Phase D: Pages UI（ルール管理タブ or Database内）
