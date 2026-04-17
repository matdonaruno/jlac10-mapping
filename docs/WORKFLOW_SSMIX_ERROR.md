# SSMIX2 エラー対応ワークフロー

## 概要

病院の電子カルテから SSMIX2 標準化ストレージにデータ出力する際、
DB 取り込み前チェックでエラーが発生する。
エラーの多くは JLAC10 コード未設定・不正によるもの。

## SSMIX2 データ構造

```
SSMIX2 標準化ストレージ:
  HL7v2.5 メッセージをファイルとして保存

  ファイル名パターン:
    {患者ID}_{日時}_{メッセージ種別}_{イベント種別}_{オーダーID}

  検査関連メッセージ:
    OML^O33  検体検査オーダー（依頼）
    OUL^R22  検体検査結果

  主要セグメント:
    MSH  メッセージヘッダー
    PID  患者情報
    OBR  検査オーダー情報
    OBX  検査結果（1項目1セグメント）
      OBX-2: データ型（NM=数値, ST=文字列, CE=コード等）
      OBX-3: 検査項目コード（JLAC10）
      OBX-5: 結果値
      OBX-6: 単位
```

## エラーの主な原因

| エラー種別 | 原因 | 対応 |
|-----------|------|------|
| JLAC10 未設定 | OBX-3 が空 or 無効なコード | マッピングツールで符番 |
| JLAC10 形式不正 | 桁数不足、不正文字 | コード修正 |
| 材料コード不正 | JLAC10 の材料部分(10-12桁)が無効 | コード確認・修正 |
| 結果値形式不正 | OBX-2(データ型)と OBX-5(値)の不整合 | データ型を確認 |
| 必須セグメント欠落 | OBR/OBX が不足 | 電子カルテ側設定確認 |
| 文字コード問題 | エンコーディング不正 | 出力設定確認 |

## 現在の運用フロー

```
1. 病院の電子カルテ → SSMIX2 出力
                          │
2. DB 取り込み前チェック
                          │ エラー検出
                          ▼
3. 担当者: GHE にイシュー手動作成
   ├── 病院名
   ├── エラー内容（種別・件数）
   ├── SSMIX2 出力結果をコピペ
   └── 進捗ステータス設定（Open/In Progress/Resolved）
                          │
4. 担当者: Rocket.Chat でイシュー発生を連絡
                          │
5. JLAC10 担当: Rocket.Chat で通知を受ける
   → GHE でイシュー確認
   → エラー内容を確認
   → マッピングツールで検索・符番
   → イシューに結果を記述
   → 解決 → イシュークローズ
```

## マッピングツールでの対応手順

### CLI（Python）

```powershell
# 1. エラー項目を検索
py -m uv run srl-scraper search "エラー項目名"

# 2. 一括で分析（エラーリストがある場合）
py -m uv run srl-scraper -v map error_items.xlsx --col-name B

# 3. 符番結果を DB に蓄積
py -m uv run srl-scraper apply-mapping mapping_result.json --hospital H100
```

### Pages（ブラウザ）

```
1. SN Error タブを開く
2. Paste Text モードで SSMIX2 出力結果を貼り付け
   - HL7v2 OBX セグメント自動検知対応
   - タブ区切り/カンマ区切り/1行1項目にも対応
3. Analyze → 付番可能/要問い合わせに分類
4. 付番不可の項目 → Generate Mail → 問い合わせメール自動生成
```

## 将来の自動化計画（GHE API 確認後）

### Phase 1: イシュー作成の半自動化

```
Pages SN Error タブの分析結果から:
  → イシュー本文のマークダウンを自動生成
  → コピーボタン → GHE の New Issue に貼り付け
  
  イシューテンプレート:
    ## エラー概要
    - 病院: {hospital_name}
    - エラー件数: {count}
    - 分析日: {date}
    
    ## エラー項目一覧
    | 項目コード | 項目名称 | JLAC10候補 | スコア | ステータス |
    |-----------|---------|-----------|--------|----------|
    | 12345     | TP      | 3A010...  | 95     | 付番可能  |
    | 67890     | 特殊X   | -         | -      | 要問合せ  |
    
    ## 問い合わせ事項
    以下の項目について病院に確認が必要:
    - ...
```

### Phase 2: GHE API による完全自動化

```
エラーログ読み込み
  → Pages で分析
  → GHE API でイシュー自動作成
    POST /api/v3/repos/{owner}/{repo}/issues
    {title, body, labels, assignees}
  → ラベル自動付与（エラー種別・病院名・優先度）
  → Rocket.Chat Webhook で通知（将来）
```

### Phase 3: 進捗管理の自動化

```
マッピング完了時:
  → GHE API でイシューにコメント追加
  → ステータス更新（In Progress → Resolved）
  → 関連する符番結果を自動記述
```

## 確認事項（TODO）

- [ ] GHE API アクセス確認: `https://ghe.ncda.hospital.local/api/v3/repos/umetani/jlac10-mapping/contents/excel/hospitals`
- [ ] GHE API でイシュー作成可能か確認
- [ ] Rocket.Chat の Webhook / API 確認
- [ ] エラーイシューのテンプレート確認（現在使用中のフォーマット）

## 関連ツール

| ツール | 用途 |
|--------|------|
| GHE (GitHub Enterprise Server) | イシュー管理、コード管理 |
| Rocket.Chat | チーム内連絡 |
| JLAC10 MAPPER (Pages) | 検索・マッピング・エラー分析 |
| JLAC10 MAPPER (CLI) | データ収集・一括処理 |
