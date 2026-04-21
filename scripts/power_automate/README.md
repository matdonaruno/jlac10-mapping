# Power Automate フローテンプレート

## インポート手順

1. `https://make.powerautomate.com` にアクセス
2. 左メニュー「マイフロー」
3. 「インポート」→「パッケージのインポート (zip)」
4. 該当の `.zip` ファイルをアップロード
5. 接続の設定:
   - OneDrive for Business → 自分のアカウントを選択
   - Excel Online (Business) → 自分のアカウントを選択
6. 「インポート」をクリック

## フロー一覧

### 1. JLAC10_JSON_to_Excel
JSON ファイルを OneDrive にアップロードすると自動で Excel に変換。

### 2. JLAC10_Email_Delivery（将来）
マッピング完了後のメール自動送信。

## 手動作成する場合

Power Automate で手動作成する手順は `docs/SETUP_POWER_AUTOMATE.md` を参照。

## 注意事項

- Excel Online「スクリプトの実行」はプレミアムコネクタ
- 組織のライセンスを確認してください
- 接続情報は環境固有のため、インポート後に再設定が必要
