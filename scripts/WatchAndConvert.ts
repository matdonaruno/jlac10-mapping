/**
 * Office Script: フォルダ監視型 JSON → Excel 一括変換
 *
 * OneDrive の指定フォルダ内の JSON ファイルを読み込み、
 * Excel に変換して出力フォルダに保存する。
 *
 * Power Automate から呼び出すことを想定:
 *   トリガー: OneDrive「ファイルが作成されたとき」
 *   → このスクリプトを実行（ファイルパスを引数で渡す）
 *
 * 手動実行の場合:
 *   CONFIG のフォルダパスを設定して実行
 *
 * フォルダ構成（OneDrive）:
 *   JLAC10_WORKFLOW/
 *   ├── json_input/       ← JSON を置く
 *   ├── excel_output/     ← 変換された Excel が出力される
 *   └── processed/        ← 処理済み JSON が移動される
 */
function main(
  workbook: ExcelScript.Workbook,
  jsonContent?: string,
  fileName?: string,
  convertType?: string
) {
  // =====================================================================
  // 引数がない場合（手動実行）: A1 セルの JSON を使用
  // 引数がある場合（Power Automate）: 引数の JSON を使用
  // =====================================================================

  let jsonText = jsonContent || "";
  let outputName = fileName || "output";
  let type = convertType || "auto"; // auto / delivery / mapping / generic

  if (!jsonText) {
    // 手動実行: A1 セルから読み取り
    const sheet = workbook.getActiveWorksheet();
    jsonText = sheet.getRange("A1").getValue() as string;
    if (!jsonText) {
      console.log("引数またはA1セルにJSONが必要です");
      return;
    }
  }

  let data: Record<string, unknown>;
  try {
    data = JSON.parse(jsonText);
  } catch (e) {
    console.log("JSONパースエラー: " + (e as Error).message);
    return;
  }

  // 変換タイプの自動判定
  if (type === "auto") {
    type = detectConvertType(data);
  }

  console.log(`変換タイプ: ${type}, ファイル: ${outputName}`);

  // タイプ別変換
  switch (type) {
    case "delivery":
      convertDelivery(workbook, data);
      break;
    case "mapping":
      convertMapping(workbook, data);
      break;
    default:
      convertGeneric(workbook, data);
      break;
  }

  console.log(`変換完了: ${outputName}`);
}

// JSON 内容から変換タイプを自動判定
function detectConvertType(data: Record<string, unknown>): string {
  const items = data.items as Record<string, unknown>[] | undefined;
  const results = data.results as Record<string, unknown>[] | undefined;

  if (results && results.length > 0 && results[0].hasOwnProperty("best_match")) {
    return "mapping";
  }
  if (items && items.length > 0 && items[0].hasOwnProperty("cs_name")) {
    return "delivery";
  }
  return "generic";
}

// ===== 設定依頼変換 =====
function convertDelivery(workbook: ExcelScript.Workbook, data: Record<string, unknown>) {
  const items = (data.items || []) as Record<string, string>[];
  const meta = (data.metadata || {}) as Record<string, string>;
  const vendor = meta.vendor || "";

  // 依頼シート
  const reqItems = items.filter(i => i.usage === "依頼" || i.usage === "依頼,結果");
  if (reqItems.length > 0) {
    createDeliverySheet(workbook, "依頼", reqItems);
  }

  // 結果シート
  const resItems = items.filter(i => i.usage === "結果" || i.usage === "依頼,結果");
  if (resItems.length > 0) {
    createDeliverySheet(workbook, "結果", resItems);
  }

  // JJ シート（富士通のみ）
  if (vendor.toUpperCase().includes("FUJITSU") || vendor === "富士通") {
    createJJSheet(workbook, resItems.length > 0 ? resItems : reqItems);
  }
}

function createDeliverySheet(
  workbook: ExcelScript.Workbook,
  name: string,
  items: Record<string, string>[]
) {
  const ws = workbook.addWorksheet(name);
  const headers = ["ローカルコード", "項目名称", "CS名", "JLAC10", "JLAC10標準名称"];
  setHeaders(ws, headers, "4472C4");

  for (let r = 0; r < items.length; r++) {
    const item = items[r];
    ws.getRange(`A${r + 2}`).setValue(item.local_code || "");
    ws.getRange(`B${r + 2}`).setValue(item.item_name || "");
    ws.getRange(`C${r + 2}`).setValue(item.cs_name || "");
    ws.getRange(`D${r + 2}`).setValue(item.jlac10 || "");
    ws.getRange(`E${r + 2}`).setValue(item.jlac10_standard_name || "");
  }
  ws.getRange("A:E").getFormat().autofitColumns();
}

function createJJSheet(
  workbook: ExcelScript.Workbook,
  items: Record<string, string>[]
) {
  const ws = workbook.addWorksheet("JJ");
  setHeaders(ws, ["HL7KEY", "ICODE", "JJCODE", "JJNAME"], "C00000");

  for (let r = 0; r < items.length; r++) {
    const item = items[r];
    const examType = item.exam_type || "検体";
    let hl7key = item.local_code || "";

    const prefixMap: Record<string, string> = {
      "細菌結果": "OTKKINF", "塗抹": "ER02TMTCD", "同定": "ER02DTKCD",
      "抗酸菌塗抹": "ER03TMTCD", "抗酸菌同定": "ER03DTKCD", "抗菌薬": "KJYKCD",
    };
    if (prefixMap[examType]) hl7key = prefixMap[examType] + hl7key;

    const icode = item.jlac10 || "L";
    let jjname = item.jlac10_standard_name || item.item_name || "";
    const encoder = new TextEncoder();
    while (encoder.encode(jjname).length > 128) jjname = jjname.slice(0, -1);

    ws.getRange(`A${r + 2}`).setValue(hl7key);
    ws.getRange(`B${r + 2}`).setValue(icode);
    ws.getRange(`C${r + 2}`).setValue(item.jlac10 || "");
    ws.getRange(`D${r + 2}`).setValue(jjname);
  }
  ws.getRange("A:D").getFormat().autofitColumns();
}

// ===== マッピング結果変換 =====
function convertMapping(workbook: ExcelScript.Workbook, data: Record<string, unknown>) {
  const results = (data.results || []) as Record<string, unknown>[];
  const meta = (data.metadata || {}) as Record<string, number>;
  const ws = workbook.addWorksheet("マッピング結果");

  const headers = ["ステータス", "項目名", "略称", "元JLAC10", "マッチJLAC10", "マッチ名称", "スコア", "分析物"];
  setHeaders(ws, headers, "4472C4");

  for (let r = 0; r < results.length; r++) {
    const item = results[r];
    const row = r + 2;
    const best = item.best_match as Record<string, unknown> | null;
    const candidates = (item.candidates || []) as Record<string, unknown>[];
    const status = item.status === "auto" ? "自動確定" : item.status === "candidate" ? "要選択" : "手動";

    ws.getRange(`A${row}`).setValue(status);
    ws.getRange(`B${row}`).setValue(String(item.item_name || ""));
    ws.getRange(`C${row}`).setValue(String(item.abbreviation || ""));
    ws.getRange(`D${row}`).setValue(String(item.original_jlac10 || ""));
    ws.getRange(`E${row}`).setValue(best ? String(best.jlac10 || "") : "");
    ws.getRange(`F${row}`).setValue(best ? String(best.matched_name || "") : "");
    ws.getRange(`G${row}`).setValue(best ? Number(best.score || 0) : "");
    ws.getRange(`H${row}`).setValue(best ? String(best.analyte_code || "") : "");

    let fill = "FFC7CE";
    if (item.status === "auto") fill = "C6EFCE";
    else if (item.status === "candidate") fill = "FFEB9C";

    for (let c = 0; c < headers.length; c++) {
      ws.getRange(`${String.fromCharCode(65 + c)}${row}`).getFormat().getFill().setColor(fill);
    }
  }
  ws.getRange("A:H").getFormat().autofitColumns();
}

// ===== 汎用変換 =====
function convertGeneric(workbook: ExcelScript.Workbook, data: Record<string, unknown>) {
  let items: Record<string, unknown>[] = [];

  if (Array.isArray(data)) {
    items = data as Record<string, unknown>[];
  } else {
    for (const key of ["items", "results", "data", "entries", "records"]) {
      if (Array.isArray(data[key])) {
        items = data[key] as Record<string, unknown>[];
        break;
      }
    }
    if (items.length === 0) {
      for (const val of Object.values(data)) {
        if (Array.isArray(val) && val.length > 0) {
          items = val as Record<string, unknown>[];
          break;
        }
      }
    }
  }

  if (items.length === 0) {
    console.log("変換可能な配列なし");
    return;
  }

  const headerSet = new Set<string>();
  for (const item of items) for (const k of Object.keys(item)) headerSet.add(k);
  const headers = Array.from(headerSet);

  const ws = workbook.addWorksheet("Data");
  setHeaders(ws, headers, "4472C4");

  for (let r = 0; r < items.length; r++) {
    const item = items[r];
    for (let c = 0; c < headers.length; c++) {
      const val = item[headers[c]];
      const cell = ws.getRange(`${colLetter(c)}${r + 2}`);
      if (val === null || val === undefined) cell.setValue("");
      else if (typeof val === "object") cell.setValue(JSON.stringify(val));
      else cell.setValue(val as string | number | boolean);
    }
  }

  for (let c = 0; c < Math.min(headers.length, 26); c++) {
    ws.getRange(`${colLetter(c)}:${colLetter(c)}`).getFormat().autofitColumns();
  }
}

// ===== ヘルパー =====
function setHeaders(ws: ExcelScript.Worksheet, headers: string[], color: string) {
  for (let c = 0; c < headers.length; c++) {
    const cell = ws.getRange(`${colLetter(c)}1`);
    cell.setValue(headers[c]);
    cell.getFormat().getFill().setColor(color);
    cell.getFormat().getFont().setColor("FFFFFF");
    cell.getFormat().getFont().setBold(true);
  }
}

function colLetter(idx: number): string {
  let s = "";
  idx++;
  while (idx > 0) { idx--; s = String.fromCharCode(65 + (idx % 26)) + s; idx = Math.floor(idx / 26); }
  return s;
}
