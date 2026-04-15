/**
 * Office Script: 院内検査マスターExcel → JLAC10 JSON 変換
 *
 * 使い方:
 *   1. Excel Online / デスクトップ版 Excel →「自動化」タブ →「新しいスクリプト」
 *   2. このコードを貼り付け
 *   3. CONFIG セクションでシート名・列番号を設定
 *   4. 実行 → "JLAC10_JSON" シートに JSON が出力される
 *
 * 出力 JSON の各項目:
 *   - item_name:       検査項目名称
 *   - jlac10:          JLAC10コード（ハイフンなし15桁）
 *   - jlac10_standard: JLAC10標準名称
 *   - analyte_code:    分析物コード（先頭5桁、あいまい検索用）
 *   - source:          "院内"
 *
 * どのブックでも使えるよう、CONFIG だけ変更すればOK。
 */
function main(workbook: ExcelScript.Workbook) {
  // =====================================================================
  // CONFIG: ここを各ブックに合わせて変更する
  // =====================================================================
  const SHEET_NAME = "検査マスター";     // 対象シート名
  const HEADER_ROW = 1;                  // ヘッダー行（1始まり）
  const DATA_START_ROW = 2;              // データ開始行

  // 列番号（1始まり）— 必要な列を指定
  const COL_ITEM_NAME = 2;              // 検査項目名称
  const COL_JLAC10 = 5;                 // JLAC10コード
  const COL_JLAC10_STANDARD_NAME = 6;   // JLAC10標準名称

  // 追加で取得したい列（任意、不要なら空配列）
  // 例: [{col: 3, key: "department"}, {col: 7, key: "unit"}]
  const EXTRA_COLUMNS: { col: number; key: string }[] = [];

  // =====================================================================
  // 処理開始
  // =====================================================================
  const sheet = workbook.getWorksheet(SHEET_NAME);
  if (!sheet) {
    console.log(`シート "${SHEET_NAME}" が見つかりません`);
    return;
  }

  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    console.log("データがありません");
    return;
  }

  const values = usedRange.getValues();
  const totalRows = values.length;
  const items: Record<string, string>[] = [];

  for (let i = DATA_START_ROW - 1; i < totalRows; i++) {
    const row = values[i];

    const rawJlac10 = String(row[COL_JLAC10 - 1] ?? "").trim();
    const itemName = String(row[COL_ITEM_NAME - 1] ?? "").trim();

    // JLAC10 が空なら除外
    if (!rawJlac10) continue;

    // ハイフン除去して15桁に正規化
    const jlac10 = rawJlac10.replace(/-/g, "");

    // 分析物コード = 先頭5桁
    const analyteCode = jlac10.length >= 5 ? jlac10.substring(0, 5) : jlac10;

    const standardName = String(
      row[COL_JLAC10_STANDARD_NAME - 1] ?? ""
    ).trim();

    const obj: Record<string, string> = {
      item_name: itemName,
      jlac10: jlac10,
      jlac10_standard_name: standardName,
      analyte_code: analyteCode,
      source: "院内",
    };

    // 追加列
    for (const extra of EXTRA_COLUMNS) {
      obj[extra.key] = String(row[extra.col - 1] ?? "").trim();
    }

    items.push(obj);
  }

  // JSON 生成
  const output = {
    metadata: {
      source_workbook: workbook.getName(),
      source_sheet: SHEET_NAME,
      exported_at: new Date().toISOString(),
      total_items: items.length,
    },
    items: items,
  };

  const jsonStr = JSON.stringify(output, null, 2);

  // 出力シート作成
  const outputSheetName = "JLAC10_JSON";
  const existing = workbook.getWorksheet(outputSheetName);
  if (existing) {
    existing.delete();
  }
  const outputSheet = workbook.addWorksheet(outputSheetName);

  // JSON を行ごとにセルへ出力
  const lines = jsonStr.split("\n");
  for (let i = 0; i < lines.length; i++) {
    outputSheet.getRange(`A${i + 1}`).setValue(lines[i]);
  }
  outputSheet.getRange("A:A").getFormat().setColumnWidth(800);

  // サマリー
  outputSheet.getRange("C1").setValue("=== 変換サマリー ===");
  outputSheet.getRange("C2").setValue(`元シート: ${SHEET_NAME}`);
  outputSheet.getRange("C3").setValue(`検査項目数: ${items.length}`);
  outputSheet
    .getRange("C4")
    .setValue(`変換日時: ${new Date().toISOString()}`);
  outputSheet
    .getRange("C5")
    .setValue("このJSONをコピーして merged_jlac10.json と統合して使用");

  outputSheet.activate();
  console.log(
    `変換完了: ${items.length}件 → ${outputSheetName} シートに出力`
  );
}
