# JLAC10 検査コード符番支援システム

病院の検査マスター符番業務を支援するツール群。  
外注先3社（SRL / BML / LSI メディエンス）の検査情報と JSLM JLAC10 コード表を統合し、  
検査項目名称のあいまい検索・一括マッピング・エラー対応をブラウザ上で行う。

## クイックスタート

### GitHub Pages アプリ（ブラウザで動作、インストール不要）

```
https://{ユーザー名}.github.io/{リポジトリ名}/app/
```

1. **DB管理タブ** → `data/merged_jlac10.json` をダウンロードして読み込み
2. **検索タブ** → 「TP」「アルブミン」「CRP」等で検索
3. **一括マッピングタブ** → 病院マスター Excel を読み込み → 自動マッチ

### データ収集（Python CLI、開発環境で実行）

```bash
# 初回セットアップ
uv sync

# 外注3社 + JLAC10マスター取得
srl-scraper srl
srl-scraper bml
srl-scraper lsi
srl-scraper jslm

# 統合
srl-scraper merge

# あいまい検索（対話モード）
srl-scraper search
```

## 機能一覧

### GitHub Pages アプリ（`app/index.html`）

| タブ | 機能 | 状態 |
|------|------|------|
| 検索 | 項目名・略称であいまい検索 → JLAC10 候補表示 | 実装済 |
| Excel JSON化+統合 | 過去マッピング Excel → JSON → DB統合 | 実装済 |
| 一括マッピング | 病院マスター Excel → 全件自動マッチ → 結果 Excel DL | 実装済 |
| SNエラー対応 | SSMIX エラーリスト分析 → 付番可否判定 → 問い合わせメール生成 | 実装済 |
| DB管理 | JSON 読み込み / 病院データ追加 / 統合DB ダウンロード | 実装済 |

### Python CLI（`srl-scraper`）

| コマンド | 機能 | 状態 |
|---------|------|------|
| `srl-scraper srl` | SRL 検査案内スクレイピング（2,120件） | 実装済 |
| `srl-scraper bml` | BML 検査案内スクレイピング（1,160件） | 実装済 |
| `srl-scraper lsi` | LSI メディエンス検査案内スクレイピング（1,335件） | 実装済 |
| `srl-scraper jslm` | JSLM JLAC10 コード表取得（月次更新） | 実装済 |
| `srl-scraper merge` | 3社 + JLAC10マスターを統合 → `merged_jlac10.json` | 実装済 |
| `srl-scraper search` | あいまい検索（対話モード / CLI） | 実装済 |
| `srl-scraper reagent` | 試薬DB構築（メーカーサイト + PMDA手動追加） | 実装済 |
| `srl-scraper sop` | SOP パーサー（Word / PDF → 測定法・試薬抽出） | 実装済 |
| `srl-scraper check` | 全ソースの更新状況確認 | 実装済 |
| `srl-scraper diff` | 2つの JSON の差分比較 | 実装済 |

### Office Script（`scripts/`）

| スクリプト | 用途 |
|-----------|------|
| `ExcelToJlac10Json.ts` | 院内検査マスター Excel → JSON 変換 |
| `MappingHistoryToJson.ts` | 過去マッピング結果 Excel → JSON 変換 |

## データ構造

### JLAC10 コード（15桁 / 17桁）

```
外注(15桁): 3A010 0000 023 271
            ├───┘ ├──┘ ├─┘ ├─┘
            │     │    │   └── 測定法コード（3桁）
            │     │    └────── 材料コード（3桁）
            │     └─────────── 識別コード（4桁）
            └───────────────── 分析物コード（5桁）

院内NCDA(17桁): 上記15桁 + 結果識別コード（2桁）
```

### 統合 JSON（`merged_jlac10.json`）

```json
{
  "jlac10": "3A0100000023271",
  "analyte_code": "3A010",
  "jlac10_decoded": {
    "valid": true,
    "analyte":        { "code": "3A010", "name": "総蛋白" },
    "identification": { "code": "0000",  "name": "" },
    "material":       { "code": "023",   "name": "血清" },
    "method":         { "code": "271",   "name": "可視吸光光度法" }
  },
  "sources": {
    "srl": { "item_name": "総蛋白(TP)", "method": "Biuret法", ... },
    "bml": { "item_name": "総蛋白(TP)", "method": "比色法(Biuret法)", ... },
    "lsi": { "item_name": "総蛋白(TP)", "method": "ビューレット法", ... }
  },
  "mapping_history": [
    { "hospital": "A病院", "item_name": "総蛋白", "abbreviation": "TP" }
  ]
}
```

### 検索のしくみ

```
入力: 「TP」「総蛋白」「albumin」「HbA1c」等（略称・日本語・英名なんでもOK）
  ↓
全ソースの全名称を検索対象:
  - SRL / BML / LSI の検査項目名称
  - JLAC10 標準名称（日本語・英語）
  - 過去の病院マッピング実績（院内名称・略称）
  ↓
スコアリング:
  完全一致=100 / 先頭一致=80+ / 部分一致=60+ / 略称マッチ=75
  ↓
出力: JLAC10 候補リスト（スコア順）
  各候補にSRL/BML/LSI情報 + 過去の実績付き
```

## 全体アーキテクチャ

```
┌─────────────── 開発環境（Python CLI） ──────────────────┐
│                                                          │
│  SRL/BML/LSI/JSLM → スクレイピング → JSON                │
│  メーカーサイト → 試薬DB                                   │
│  SOP (Word/PDF) → パース → 測定法抽出                     │
│  統合 → merged_jlac10.json                               │
│                                                          │
└──────────────────────┬───────────────────────────────────┘
                       │ git push
                       ▼
┌─────────────── GitHub リポジトリ ────────────────────────┐
│                                                          │
│  data/merged_jlac10.json  (統合DB)                       │
│  app/index.html           (GitHub Pages アプリ)           │
│                                                          │
└──────────────────────┬───────────────────────────────────┘
                       │ GitHub Pages
                       ▼
┌─────────────── 職場 PC（ブラウザのみ） ─────────────────┐
│                                                          │
│  ブラウザでアクセス → インストール不要                      │
│  JSON読み込み → 検索 / 一括マッピング / エラー対応         │
│  結果Excel ダウンロード                                   │
│                                                          │
│  Excel Online + Office Script で JSON 変換               │
│  → GitHub にブラウザからアップロード                       │
│                                                          │
└──────────────────────────────────────────────────────────┘
```

## ディレクトリ構成

```
jlac10-mapping/
├── README.md
├── .gitignore
├── pyproject.toml              # Python 依存関係
├── uv.lock
├── app/
│   └── index.html              # GitHub Pages アプリ（全機能）
├── data/
│   ├── merged_jlac10.json      # 統合DB（2,514 JLAC10）
│   ├── jlac10_master.json      # JSLM コードマスター
│   ├── jlac10_lookup.json      # 検索用辞書
│   ├── srl_tests_latest.json   # SRL データ
│   ├── bml_tests_latest.json   # BML データ
│   ├── lsi_tests_latest.json   # LSI データ
│   ├── sop_parsed.json         # SOP パース結果
│   └── reagents/
│       └── reagent_db.json     # 試薬DB
├── docs/
│   ├── SYSTEM_OVERVIEW.md      # 詳細設計ドキュメント
│   └── MAPPING_RULES.md        # マッピングルール
├── scripts/
│   ├── ExcelToJlac10Json.ts    # Office Script: 院内マスター→JSON
│   └── MappingHistoryToJson.ts # Office Script: 過去実績→JSON
└── src/srl_scraper/            # Python ツール
    ├── cli.py                  # CLI エントリーポイント
    ├── scraper.py              # SRL スクレイパー
    ├── bml.py                  # BML スクレイパー
    ├── lsi.py                  # LSI スクレイパー
    ├── jslm.py                 # JSLM JLAC10 コード表
    ├── merge.py                # 統合 + JLAC10 デコード
    ├── search.py               # あいまい検索エンジン
    ├── reagent.py              # 試薬DB
    ├── sop_parser.py           # SOP パーサー
    └── categories.py           # SRL カテゴリ定義
```

## 定期更新

```bash
# 週次: 更新があるソースだけ取得
srl-scraper srl --check-update
srl-scraper bml --check-update
srl-scraper lsi --check-update
srl-scraper merge

# 月次: JLAC10 コード表
srl-scraper jslm --check-update
srl-scraper merge
```

各コマンドに `--check-update` を付けると、サイトの更新日を確認して  
変更がなければスキップする（サーバーへの負荷を最小限に）。

## サーバー負荷対策

| 対策 | 詳細 |
|-----|------|
| リクエスト間隔 | 全サイト共通 3秒 |
| HTML キャッシュ | 24時間有効 |
| 更新検知 | トップページ1回で判定、変更なしなら即終了 |
| リトライ | 指数バックオフ（5秒→10秒）、最大3回 |
| PMDA | 自動巡回禁止（利用規約準拠）、手動URL指定のみ |

## TODO

- [ ] GitHub Pages 動作確認（職場PC）
- [ ] 過去マッピング Excel の JSON 化・蓄積
- [ ] マッピング結果の DB 還元（確定分を共有DBに差分追加）
- [ ] 試薬→測定法マッピング（測定原理→JLAC10 測定法コード変換）
- [ ] 他試薬メーカー対応（栄研、シスメックス、積水メディカル等）
- [ ] JANIS コード体系（細菌検査）
- [ ] 関連マスター群（親子 / 材料 / 容器 / 基準値 / 装置 / 外注先設定）
- [ ] マッピング SOP に基づくルール実装
- [ ] 外注先マッピング vs NCDA の差異検出
- [ ] Power Automate 連携（GitHub → OneDrive 自動同期）

## ライセンス

業務利用。データソースの利用規約に従うこと。  
特に PMDA 添付文書は自動巡回禁止。
