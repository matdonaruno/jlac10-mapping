# Power Automate フロー設定手順

## フォルダ構成（OneDrive）

```
OneDrive/
└── JLAC10_WORKFLOW/
    ├── json_input/       ← ここに JSON を置く
    ├── excel_output/     ← 変換された Excel が自動出力
    └── processed/        ← 処理済み JSON が移動
```

## フロー: JSON → Excel 自動変換

### 設定手順

1. `https://make.powerautomate.com` にアクセス
2. 「作成」→「自動化したクラウドフロー」
3. フロー名: `JLAC10_JSON_to_Excel`

### トリガー
- 「OneDrive for Business」→「ファイルが作成されたとき」
- フォルダ: `/JLAC10_WORKFLOW/json_input`

### ステップ

```
1. トリガー: OneDrive「ファイルが作成されたとき」
   フォルダ: /JLAC10_WORKFLOW/json_input

2. ファイルコンテンツの取得
   ファイル: トリガーのファイルID

3. Excel Online (Business)「スクリプトの実行」
   場所: OneDrive for Business
   ドキュメントライブラリ: OneDrive
   ファイル: /JLAC10_WORKFLOW/template.xlsx（空のテンプレート）
   スクリプト: WatchAndConvert
   パラメータ:
     jsonContent: ファイルコンテンツ（式: base64ToString(body('ファイルコンテンツの取得')?['$content']))
     fileName: トリガーのファイル名
     convertType: "auto"

4. OneDrive「ファイルの作成」
   フォルダ: /JLAC10_WORKFLOW/excel_output
   ファイル名: {元ファイル名}.xlsx
   ファイルコンテンツ: スクリプト実行結果

5. OneDrive「ファイルの移動」
   ファイル: トリガーのファイルID
   移動先: /JLAC10_WORKFLOW/processed
```

### 前提条件
- Power Automate プレミアムコネクタ（Excel Online スクリプト実行）が利用可能
- OneDrive for Business が利用可能
- Office Script が有効

### 確認事項
- [ ] Power Automate にログインできるか
- [ ] OneDrive コネクタが使えるか
- [ ] Excel Online「スクリプトの実行」が使えるか（プレミアム）
- [ ] `/JLAC10_WORKFLOW/` フォルダを作成

## 手動運用（Power Automate なしの場合）

1. JSON ファイルを OneDrive にアップロード
2. Excel Online で新しいブックを開く
3. JSON の中身をコピー → A1 に貼り付け
4. 自動化タブ → WatchAndConvert を実行
5. 生成されたシートを「名前を付けて保存」
