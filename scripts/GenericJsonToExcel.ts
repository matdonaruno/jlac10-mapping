/**
 * Office Script: 汎用 JSON → Excel 変換
 *
 * どんな JSON 配列でも Excel テーブルに変換する。
 * キーがヘッダー、値がデータ行になる。
 *
 * 使い方:
 *   1. 新しいブックを開く
 *   2. A1 セルに JSON テキストを貼り付け
 *   3. このスクリプトを実行
 *
 * 対応形式:
 *   { "items": [{...}, {...}] }    → items 配列を変換
 *   { "results": [{...}, {...}] }  → results 配列を変換
 *   [{...}, {...}]                  → そのまま配列を変換
 */
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const jsonText = sheet.getRange("A1").getValue() as string;

  if (!jsonText) {
    console.log("A1 セルに JSON を貼り付けてください");
    return;
  }

  let rawData: unknown;
  try {
    rawData = JSON.parse(jsonText);
  } catch (e) {
    console.log("JSON パースエラー");
    return;
  }

  // 配列を探す
  let items: Record<string, unknown>[] = [];
  if (Array.isArray(rawData)) {
    items = rawData as Record<string, unknown>[];
  } else if (typeof rawData === "object" && rawData !== null) {
    const obj = rawData as Record<string, unknown>;
    // items, results, data, entries などを探す
    for (const key of ["items", "results", "data", "entries", "records"]) {
      if (Array.isArray(obj[key])) {
        items = obj[key] as Record<string, unknown>[];
        break;
      }
    }
    // 見つからなければ最初の配列値を使う
    if (items.length === 0) {
      for (const val of Object.values(obj)) {
        if (Array.isArray(val) && val.length > 0) {
          items = val as Record<string, unknown>[];
          break;
        }
      }
    }
  }

  if (items.length === 0) {
    console.log("変換可能な配列が見つかりません");
    return;
  }

  // ヘッダー（全アイテムのキーを集約）
  const headerSet = new Set<string>();
  for (const item of items) {
    for (const key of Object.keys(item)) {
      headerSet.add(key);
    }
  }
  const headers = Array.from(headerSet);

  // 出力シート
  const ws = workbook.addWorksheet("Data");

  // ヘッダー
  for (let c = 0; c < headers.length; c++) {
    const cell = ws.getRange(`${colLetter(c)}1`);
    cell.setValue(headers[c]);
    cell.getFormat().getFill().setColor("4472C4");
    cell.getFormat().getFont().setColor("FFFFFF");
    cell.getFormat().getFont().setBold(true);
  }

  // データ
  for (let r = 0; r < items.length; r++) {
    const item = items[r];
    for (let c = 0; c < headers.length; c++) {
      const val = item[headers[c]];
      const cell = ws.getRange(`${colLetter(c)}${r + 2}`);
      if (val === null || val === undefined) {
        cell.setValue("");
      } else if (typeof val === "object") {
        cell.setValue(JSON.stringify(val));
      } else {
        cell.setValue(val as string | number | boolean);
      }
    }
  }

  // 列幅自動調整（最大26列まで）
  const maxCols = Math.min(headers.length, 26);
  for (let c = 0; c < maxCols; c++) {
    ws.getRange(`${colLetter(c)}:${colLetter(c)}`).getFormat().autofitColumns();
  }

  // メタデータシート（元データにmetadataがあれば）
  if (typeof rawData === "object" && !Array.isArray(rawData)) {
    const obj = rawData as Record<string, unknown>;
    if (obj.metadata && typeof obj.metadata === "object") {
      const metaWs = workbook.addWorksheet("Metadata");
      const meta = obj.metadata as Record<string, unknown>;
      let row = 1;
      for (const [k, v] of Object.entries(meta)) {
        metaWs.getRange(`A${row}`).setValue(k);
        metaWs.getRange(`A${row}`).getFormat().getFont().setBold(true);
        metaWs.getRange(`B${row}`).setValue(
          typeof v === "object" ? JSON.stringify(v) : String(v ?? "")
        );
        row++;
      }
      metaWs.getRange("A:B").getFormat().autofitColumns();
    }
  }

  sheet.setName("_JSON_INPUT");
  console.log(`変換完了: ${items.length}行 × ${headers.length}列`);
}

function colLetter(idx: number): string {
  let s = "";
  idx++;
  while (idx > 0) {
    idx--;
    s = String.fromCharCode(65 + (idx % 26)) + s;
    idx = Math.floor(idx / 26);
  }
  return s;
}
