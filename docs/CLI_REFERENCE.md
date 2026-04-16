# CLI コマンドリファレンス

Windows 環境では `py -m uv run srl-scraper` で実行。
`-v` を付けるとデバッグログが表示される。

---

## セットアップ

```powershell
# 初回のみ: 環境構築
.\scripts\setup_windows.ps1

# または手動:
py -m pip install uv
py -m uv sync
```

---

## 1. 検索（search）

検査項目名・略称であいまい検索。単発またはアシスト対話モード。

```powershell
# 単発検索
py -m uv run srl-scraper search "TP"
py -m uv run srl-scraper search "アルブミン" -n 5

# 対話モード（連続検索、quit で終了）
py -m uv run srl-scraper search

# デバッグ（検索ヒット詳細）
py -m uv run srl-scraper -v search "TP"
```

**出力:**
```
[1] TP  (スコア: 100)
    JLAC10: 3A0100000023271
    分析物: 3A010
    標準: analyte=総蛋白 / material=血清 / method=可視吸光光度法
    SRL: 総蛋白(TP) / BML: 総蛋白(TP) / LSI: 総蛋白(TP)
```

**デバッグログ (-v):**
```
[DEBUG] インデックス構築: エントリ数=2834, 名称総数=9024
[DEBUG] Search 'TP' → 26 hits (top: score=100.0, jlac10=3A0100000023271)
```

---

## 2. Excel/CSV → JSON 変換（convert）

どんな形式の Excel/CSV でも統一 JSON に変換する。

```powershell
# 基本（A列=項目名, C列=JLAC10）
py -m uv run srl-scraper convert input.xlsx --col-item A --col-jlac10 C

# 全オプション指定
py -m uv run srl-scraper convert input.xlsx \
  --col-item A \
  --col-jlac10 C \
  --col-abbr B \
  --col-std-name D \
  --hospital "H001" \
  --sheet "検査マスター" \
  --skip-rows 2 \
  -o output.json

# ヘッダ名で列指定（Excel のヘッダ行テキストで指定）
py -m uv run srl-scraper convert input.xlsx \
  --col-item "検査項目名称" \
  --col-jlac10 "JLAC10コード"

# デバッグ（各行の変換結果を表示）
py -m uv run srl-scraper -v convert input.xlsx --col-item A --col-jlac10 C
```

**デバッグログ (-v):**
```
[DEBUG] ファイル読み込み: input.xlsx (xlsx), 1500行
[DEBUG] 列解決: item_name → 0(A), jlac10 → 2(C)
[DEBUG] Row 2: item_name='総蛋白' jlac10='3A0100000023271' status=valid_15
[DEBUG] Row 3: item_name='アルブミン' jlac10='' status=empty
[INFO]  変換完了: 1498件 (空行スキップ: 2件)
[DEBUG] 内訳: valid_15=1200, valid_17=0, empty=280, invalid=18
```

---

## 3. 一括マッピング（map）

病院マスターの全項目を自動で JLAC10 にマッチング。

```powershell
# 基本（A列=項目名）
py -m uv run srl-scraper map master.xlsx --col-name A

# 閾値変更（auto=80以上, candidate=40以上）
py -m uv run srl-scraper map master.xlsx --col-name A \
  --threshold-auto 80 --threshold-candidate 40

# 出力先指定
py -m uv run srl-scraper map master.xlsx --col-name A -o result.xlsx

# 答え合わせ（JLAC10列を指定すると正解率を計算）
py -m uv run srl-scraper -v map master.xlsx --col-name A --col-jlac10 C

# 病院名指定
py -m uv run srl-scraper map master.xlsx --col-name A --hospital "H001"
```

**出力ファイル:**
- `mapping_result.xlsx` — 色分き Excel（auto=緑, candidate=黄, manual=赤）
- `mapping_result.json` — 同内容の JSON

**デバッグログ (-v):**
```
[DEBUG] マッピング開始: 1500件, auto閾値=90.0, candidate閾値=50.0
[DEBUG] Map[1/1500] '総蛋白' → status=auto, score=100.0, jlac10=3A0100000023271
[DEBUG]   VERIFY: original=3A0100000023271, mapped=3A0100000023271, match=OK
[DEBUG] Map[2/1500] 'ALB' → status=auto, score=95.0, jlac10=3A0150000023271
[DEBUG]   VERIFY: original=3A0150000023271, mapped=3A0150000023271, match=OK
[DEBUG] Map[3/1500] '特殊検査X' → status=manual, score=0.0
[DEBUG]   VERIFY: original=8C5700000070848, mapped=(none), match=MISMATCH
...
[INFO]  マッピング完了: 1500件 (auto=1200, candidate=200, manual=100)
[DEBUG] 答え合わせ: 1350/1500 correct (rate=90.0%)
```

---

## 4. マッピング結果 → DB 還元（apply-mapping）

確定したマッピングを `merged_jlac10.json` に蓄積。次回の検索精度が上がる。

```powershell
# JSON 結果を還元
py -m uv run srl-scraper apply-mapping mapping_result.json --hospital "H001"

# Excel 結果を還元（Status が auto/confirmed の行のみ）
py -m uv run srl-scraper apply-mapping mapping_result.xlsx --hospital "H001"

# 全行を還元（candidate/manual も含む）
py -m uv run srl-scraper apply-mapping mapping_result.json --hospital "H001" --all

# デバッグ
py -m uv run srl-scraper -v apply-mapping mapping_result.json --hospital "H001"
```

**デバッグログ (-v):**
```
[DEBUG] apply_mapping_results開始: 入力1500件, confirmed_only=True
[DEBUG] 既存エントリ数: 2427
[DEBUG] 既存エントリに追加: jlac10=3A0100000023271, item='総蛋白'
[DEBUG] 新規エントリ作成: jlac10=9Z9990000000000, item='特殊項目'
[DEBUG] 重複スキップ: jlac10=3A0100000023271, item='総蛋白', hospital='H001'
[INFO]  DB還元完了: 追加800件, スキップ700件, 新規エントリ50件
```

---

## 5. 外注 vs NCDA 差異チェック（check-ncda）

外注先提出の JLAC10(15桁) と院内 NCDA(17桁) の差異を検出。

```powershell
# 基本（C列=外注JLAC10, D列=NCDA, B列=項目名）
py -m uv run srl-scraper check-ncda input.xlsx \
  --outsource-col C --ncda-col D --name-col B

# 出力先指定
py -m uv run srl-scraper check-ncda input.xlsx \
  --outsource-col C --ncda-col D --name-col B \
  -o ncda_report.xlsx

# デバッグ
py -m uv run srl-scraper -v check-ncda input.xlsx \
  --outsource-col C --ncda-col D --name-col B
```

**出力 Excel:**
- OK行=緑、Warning行=黄、Error行=赤

**デバッグログ (-v):**
```
[DEBUG] Check: '総蛋白' outsource=3A0100000023271(valid_15) ncda=3A010000002327100(valid_17)
[DEBUG]   差異なし
[DEBUG] Check: 'ALB' outsource=3A0150000023271(valid_15) ncda=3A0150000041271XX(valid_17)
[DEBUG]   差異: material(023→041)
[INFO]  完了: 100件 (ok=90, warning=8, error=2)
```

---

## 6. データ収集（srl / bml / lsi / jslm）

外注先・JSLM からデータを自動取得。

```powershell
# 更新確認（アクセス最小限）
py -m uv run srl-scraper check

# 更新がある場合のみ取得
py -m uv run srl-scraper srl --check-update
py -m uv run srl-scraper bml --check-update
py -m uv run srl-scraper lsi --check-update
py -m uv run srl-scraper jslm --check-update

# 統合
py -m uv run srl-scraper merge

# 全部まとめて（週次実行用）
py -m uv run srl-scraper srl --check-update && ^
py -m uv run srl-scraper bml --check-update && ^
py -m uv run srl-scraper lsi --check-update && ^
py -m uv run srl-scraper merge
```

---

## 7. 試薬 DB（reagent）

```powershell
# メーカーサイトから試薬一覧取得（カイノス）
py -m uv run srl-scraper reagent

# PMDA 添付文書を手動追加（ブラウザで URL を確認してから）
py -m uv run srl-scraper reagent --pmda "https://www.pmda.go.jp/PmdaSearch/ivdDetail/..."
```

---

## 8. SOP パース（sop）

```powershell
# 単一ファイル
py -m uv run srl-scraper sop path\to\SOP_TP.docx

# ディレクトリ一括（サブディレクトリ再帰）
py -m uv run srl-scraper sop path\to\sop_folder\
```

---

## 9. 差分比較（diff）

```powershell
py -m uv run srl-scraper diff old_merged.json new_merged.json -o diff_report.json
```

---

## 10. Git 同期（GHE）

```powershell
# 変更を確認
git status

# 更新されたDBをコミット & プッシュ
git add data/merged_jlac10.json
git commit -m "DB更新: マッピング実績追加"
git push

# 他メンバーの変更を取得
git pull
```

---

## ワークフロー例

### A. 日常業務（SNエラー対応）

```powershell
# 1. エラー項目を検索
py -m uv run srl-scraper search "検査項目名"

# 2. 付番 → DB に蓄積
# （Pages UI で操作、または apply-mapping）
```

### B. 新規病院の一括マッピング

```powershell
# 1. マスターを JSON 変換
py -m uv run srl-scraper convert master.xlsx --col-item B --col-jlac10 E --hospital "H001"

# 2. 一括マッピング（答え合わせ付き）
py -m uv run srl-scraper -v map master.xlsx --col-name B --col-jlac10 E --hospital "H001"

# 3. 結果確認 → mapping_result.xlsx を開く

# 4. 確定分を DB に還元
py -m uv run srl-scraper apply-mapping mapping_result.json --hospital "H001"

# 5. Git で共有
git add data/merged_jlac10.json
git commit -m "H001マッピング実績追加"
git push
```

### C. 外注先マッピング検証

```powershell
# 外注先提出のマッピングと NCDA の差異チェック
py -m uv run srl-scraper check-ncda outsource_mapping.xlsx \
  --outsource-col C --ncda-col D --name-col B

# → ncda_report.xlsx で差異確認
```
