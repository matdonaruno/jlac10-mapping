/**
 * Office Script: 設定依頼 JSON → Excel 変換
 *
 * export-delivery コマンドが出力した JSON を Excel に変換する。
 * 依頼用 / 結果用 / JJ用（富士通）の3シートを1ブックに生成。
 *
 * 使い方:
 *   1. 新しいブックを開く
 *   2. A1 セルに JSON テキストを貼り付け
 *   3. このスクリプトを実行
 *   4. 「依頼」「結果」「JJ」シートが生成される
 *
 * JSON 構造:
 *   { "items": [{ local_code, item_name, cs_name, jlac10, jlac10_standard_name, usage }] }
 */
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const jsonText = sheet.getRange("A1").getValue() as string;

  if (!jsonText) {
    console.log("A1 セルに JSON を貼り付けてください");
    return;
  }

  let data: {
    items: {
      local_code: string;
      item_name: string;
      cs_name: string;
      jlac10: string;
      jlac10_standard_name: string;
      usage: string;
      exam_type?: string;
    }[];
    metadata?: { hospital?: string; issue?: string; vendor?: string };
  };

  try {
    data = JSON.parse(jsonText);
  } catch (e) {
    console.log("JSON パースエラー");
    return;
  }

  const items = data.items || [];
  if (items.length === 0) {
    console.log("items が空です");
    return;
  }

  const vendor = data.metadata?.vendor || "";
  const hospital = data.metadata?.hospital || "";
  const issue = data.metadata?.issue || "";

  // ヘッダースタイル
  const headerColor = "4472C4";

  // --- 依頼シート ---
  const reqItems = items.filter((i) => i.usage === "依頼" || i.usage === "依頼,結果");
  if (reqItems.length > 0) {
    const ws = workbook.addWorksheet("依頼");
    const headers = ["ローカルコード", "項目名称", "CS名", "JLAC10", "JLAC10標準名称"];
    for (let c = 0; c < headers.length; c++) {
      const cell = ws.getRange(`${String.fromCharCode(65 + c)}1`);
      cell.setValue(headers[c]);
      cell.getFormat().getFill().setColor(headerColor);
      cell.getFormat().getFont().setColor("FFFFFF");
      cell.getFormat().getFont().setBold(true);
    }
    for (let r = 0; r < reqItems.length; r++) {
      const item = reqItems[r];
      ws.getRange(`A${r + 2}`).setValue(item.local_code || "");
      ws.getRange(`B${r + 2}`).setValue(item.item_name || "");
      ws.getRange(`C${r + 2}`).setValue(item.cs_name || "");
      ws.getRange(`D${r + 2}`).setValue(item.jlac10 || "");
      ws.getRange(`E${r + 2}`).setValue(item.jlac10_standard_name || "");
    }
    ws.getRange("A:E").getFormat().autofitColumns();
    console.log(`依頼シート: ${reqItems.length}件`);
  }

  // --- 結果シート ---
  const resItems = items.filter((i) => i.usage === "結果" || i.usage === "依頼,結果");
  if (resItems.length > 0) {
    const ws = workbook.addWorksheet("結果");
    const headers = ["ローカルコード", "項目名称", "CS名", "JLAC10", "JLAC10標準名称"];
    for (let c = 0; c < headers.length; c++) {
      const cell = ws.getRange(`${String.fromCharCode(65 + c)}1`);
      cell.setValue(headers[c]);
      cell.getFormat().getFill().setColor(headerColor);
      cell.getFormat().getFont().setColor("FFFFFF");
      cell.getFormat().getFont().setBold(true);
    }
    for (let r = 0; r < resItems.length; r++) {
      const item = resItems[r];
      ws.getRange(`A${r + 2}`).setValue(item.local_code || "");
      ws.getRange(`B${r + 2}`).setValue(item.item_name || "");
      ws.getRange(`C${r + 2}`).setValue(item.cs_name || "");
      ws.getRange(`D${r + 2}`).setValue(item.jlac10 || "");
      ws.getRange(`E${r + 2}`).setValue(item.jlac10_standard_name || "");
    }
    ws.getRange("A:E").getFormat().autofitColumns();
    console.log(`結果シート: ${resItems.length}件`);
  }

  // --- JJ シート（富士通のみ） ---
  if (vendor.toUpperCase().includes("FUJITSU") || vendor === "富士通") {
    const ws = workbook.addWorksheet("JJ");
    const headers = ["HL7KEY", "ICODE", "JJCODE", "JJNAME"];
    for (let c = 0; c < headers.length; c++) {
      const cell = ws.getRange(`${String.fromCharCode(65 + c)}1`);
      cell.setValue(headers[c]);
      cell.getFormat().getFill().setColor("C00000");
      cell.getFormat().getFont().setColor("FFFFFF");
      cell.getFormat().getFont().setBold(true);
    }
    const jjItems = resItems.length > 0 ? resItems : reqItems;
    for (let r = 0; r < jjItems.length; r++) {
      const item = jjItems[r];
      const examType = item.exam_type || "検体";
      let hl7key = item.local_code || "";
      if (examType === "細菌結果") hl7key = "OTKKINF" + hl7key;
      else if (examType === "塗抹") hl7key = "ER02TMTCD" + hl7key;
      else if (examType === "同定") hl7key = "ER02DTKCD" + hl7key;
      else if (examType === "抗酸菌塗抹") hl7key = "ER03TMTCD" + hl7key;
      else if (examType === "抗酸菌同定") hl7key = "ER03DTKCD" + hl7key;
      else if (examType === "抗菌薬") hl7key = "KJYKCD" + hl7key;

      const icode = item.jlac10 ? item.jlac10 : "L";
      let jjname = item.jlac10_standard_name || item.item_name || "";
      // 128Byte制限
      const encoder = new TextEncoder();
      while (encoder.encode(jjname).length > 128) {
        jjname = jjname.slice(0, -1);
      }

      ws.getRange(`A${r + 2}`).setValue(hl7key);
      ws.getRange(`B${r + 2}`).setValue(icode);
      ws.getRange(`C${r + 2}`).setValue(item.jlac10 || "");
      ws.getRange(`D${r + 2}`).setValue(jjname);
    }
    ws.getRange("A:D").getFormat().autofitColumns();
    console.log(`JJシート: ${jjItems.length}件`);
  }

  // 元のシート名を変更
  sheet.setName("_JSON_INPUT");

  console.log(`変換完了 (${hospital} #${issue})`);
}
