# マッピングルール

## 一括マッピング時のデータフロー

```
病院マスター（院内項目）
  │
  ├── 院内測定項目 → NCDA(17桁) でマッピング
  │     SOPと試薬情報で測定法を特定
  │     マッピングSOPのルールに従う
  │
  └── 外注項目 → 外注先が提出したマッピング結果を使用
        ただしNCDAルールとの差異チェックが必要
        差異がある場合はエラーとして検出
```

## 外注マッピングとNCDAの差異

外注先（SRL/BML/LSI）が提出するマッピング結果:
- 外注先独自のルールで JLAC10(15桁) を付番
- NCDA(17桁) のマッピングルールとは異なる場合がある
- **差異をエラーとして検出する機能が必要**

## エラーチェック項目（TODO: マッピングSOP確認後に詳細化）

- [ ] 外注先JLAC10 vs NCDA ルールの差異検出
- [ ] 分析物コードの不一致
- [ ] 材料コードの不一致
- [ ] 測定法コードの不一致
- [ ] 結果識別コード（NCDA 16-17桁目）の妥当性

## マッピングSOP

マッピング作業には社内SOPが存在する。
そのSOPに従ってマッピングルールを実装する。
（SOP内容は確認次第ここに反映）

## 過去マッピング実績の JSON 構造

```json
{
  "metadata": {
    "hospital": "A病院",
    "total_items": 1500
  },
  "items": [
    {
      "hospital": "A病院",
      "item_name": "総蛋白",
      "abbreviation": "TP",
      "jlac10": "3A0100000023271",
      "analyte_code": "3A010",
      "jlac10_standard_name": "総蛋白[定量]"
    }
  ]
}
```

## Office Script

- `scripts/MappingHistoryToJson.ts` — 過去マッピング結果Excel → JSON
- `scripts/ExcelToJlac10Json.ts` — 院内マスターExcel → JSON

CONFIG セクションで列番号と病院名を設定して使用。
