/**
 * Office Script: 汎用 Excel → JSON 変換
 *
 * アクティブシートのデータを JSON に変換。
 * 1行目をヘッダーとして使用。
 *
 * 使い方:
 *   1. 変換したいシートを選択
 *   2. このスクリプトを実行
 *   3. "JSON_OUTPUT" シートに結果が出力される
 */
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const usedRange = sheet.getUsedRange();

  if (!usedRange) {
    console.log("データがありません");
    return;
  }

  const values = usedRange.getValues();
  if (values.length < 2) {
    console.log("ヘッダー行とデータ行が必要です");
    return;
  }

  // ヘッダー
  const headers = values[0].map((h) => String(h ?? "").trim());

  // データ
  const items: Record<string, string | number | boolean>[] = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const isEmpty = row.every((c) => c === "" || c === null || c === undefined);
    if (isEmpty) continue;

    const obj: Record<string, string | number | boolean> = {};
    for (let j = 0; j < headers.length; j++) {
      const key = headers[j];
      if (!key) continue;
      const val = row[j];
      if (val === null || val === undefined || val === "") {
        obj[key] = "";
      } else if (typeof val === "number" || typeof val === "boolean") {
        obj[key] = val;
      } else {
        obj[key] = String(val).trim();
      }
    }
    items.push(obj);
  }

  // JSON 出力
  const output = {
    metadata: {
      source_sheet: sheet.getName(),
      source_workbook: workbook.getName(),
      exported_at: new Date().toISOString(),
      total_items: items.length,
      headers: headers.filter((h) => h),
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
  outputSheet.getRange("A:A").getFormat().setColumnWidth(600);

  // サマリー
  outputSheet.getRange("C1").setValue("=== 変換サマリー ===");
  outputSheet.getRange("C2").setValue(`元シート: ${sheet.getName()}`);
  outputSheet.getRange("C3").setValue(`ヘッダー数: ${headers.filter((h) => h).length}`);
  outputSheet.getRange("C4").setValue(`データ行数: ${items.length}`);
  outputSheet.getRange("C5").setValue(`日時: ${new Date().toISOString()}`);

  outputSheet.activate();
  console.log(`変換完了: ${items.length}件 → JSON_OUTPUT シート`);
}
