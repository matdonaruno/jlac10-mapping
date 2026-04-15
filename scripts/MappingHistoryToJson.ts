/**
 * Office Script: 過去マッピング結果 Excel → JSON 変換
 *
 * 使い方:
 *   1. マッピング結果の Excel を開く
 *   2. 自動化タブ → このスクリプトを実行
 *   3. CONFIG で列番号とシート名を調整
 *   4. "JSON_OUTPUT" シートに JSON が出力される
 *   5. JSON をコピーしてリポジトリにアップロード
 */
function main(workbook: ExcelScript.Workbook) {
  // =====================================================================
  // CONFIG: 各ファイルに合わせて変更
  // =====================================================================
  const HOSPITAL_NAME = "A病院";        // ファイル名から判断して設定
  const SHEET_NAME = "";                // 空ならアクティブシート
  const HEADER_ROW = 1;                 // ヘッダー行
  const DATA_START_ROW = 2;             // データ開始行

  // 列番号（1始まり）— 該当しない場合は 0
  const COL_ITEM_NAME = 1;             // 院内項目名称
  const COL_ABBREVIATION = 2;          // 略称
  const COL_JLAC10 = 3;               // JLAC10コード
  const COL_JLAC10_STANDARD_NAME = 4;  // JLAC10正式名称

  // =====================================================================
  // 処理
  // =====================================================================
  const sheet = SHEET_NAME
    ? workbook.getWorksheet(SHEET_NAME)
    : workbook.getActiveWorksheet();

  if (!sheet) {
    console.log("シートが見つかりません");
    return;
  }

  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    console.log("データがありません");
    return;
  }

  const values = usedRange.getValues();
  const items: Record<string, string>[] = [];

  for (let i = DATA_START_ROW - 1; i < values.length; i++) {
    const row = values[i];

    const itemName = COL_ITEM_NAME > 0
      ? String(row[COL_ITEM_NAME - 1] ?? "").trim() : "";
    const abbreviation = COL_ABBREVIATION > 0
      ? String(row[COL_ABBREVIATION - 1] ?? "").trim() : "";
    const rawJlac10 = COL_JLAC10 > 0
      ? String(row[COL_JLAC10 - 1] ?? "").trim() : "";
    const standardName = COL_JLAC10_STANDARD_NAME > 0
      ? String(row[COL_JLAC10_STANDARD_NAME - 1] ?? "").trim() : "";

    // 項目名もJLAC10も空なら除外
    if (!itemName && !rawJlac10) continue;

    // JLAC10 ハイフン除去
    const jlac10 = rawJlac10.replace(/-/g, "");

    // 分析物コード（先頭5桁）
    const analyteCode = jlac10.length >= 5 ? jlac10.substring(0, 5) : "";

    items.push({
      hospital: HOSPITAL_NAME,
      item_name: itemName,
      abbreviation: abbreviation,
      jlac10: jlac10,
      analyte_code: analyteCode,
      jlac10_standard_name: standardName,
    });
  }

  // JSON 生成
  const output = {
    metadata: {
      hospital: HOSPITAL_NAME,
      source_sheet: sheet.getName(),
      exported_at: new Date().toISOString(),
      total_items: items.length,
    },
    items: items,
  };

  const jsonStr = JSON.stringify(output, null, 2);

  // 出力シート
  const outputSheetName = "JSON_OUTPUT";
  const existing = workbook.getWorksheet(outputSheetName);
  if (existing) existing.delete();
  const outputSheet = workbook.addWorksheet(outputSheetName);

  const lines = jsonStr.split("\n");
  for (let i = 0; i < lines.length; i++) {
    outputSheet.getRange(`A${i + 1}`).setValue(lines[i]);
  }
  outputSheet.getRange("A:A").getFormat().setColumnWidth(800);

  // サマリー
  outputSheet.getRange("C1").setValue("=== 変換サマリー ===");
  outputSheet.getRange("C2").setValue(`病院: ${HOSPITAL_NAME}`);
  outputSheet.getRange("C3").setValue(`項目数: ${items.length}`);
  outputSheet.getRange("C4").setValue(`JLAC10あり: ${items.filter(x => x.jlac10).length}`);
  outputSheet.getRange("C5").setValue(`日時: ${new Date().toISOString()}`);

  outputSheet.activate();
  console.log(`変換完了: ${items.length}件`);
}
