/**
 * Office Script: マッピング結果 JSON → Excel 変換
 *
 * map コマンドが出力した JSON を色分き Excel に変換。
 * auto=緑、candidate=黄、manual=赤。
 *
 * 使い方:
 *   1. 新しいブックを開く
 *   2. A1 セルに JSON テキストを貼り付け
 *   3. このスクリプトを実行
 */
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const jsonText = sheet.getRange("A1").getValue() as string;

  if (!jsonText) {
    console.log("A1 セルに JSON を貼り付けてください");
    return;
  }

  let data: {
    metadata: { total: number; auto: number; candidate: number; manual: number };
    results: {
      item_name: string;
      abbreviation: string;
      original_jlac10: string;
      status: string;
      best_match: { jlac10: string; matched_name: string; score: number; analyte_code: string } | null;
      candidates: { jlac10: string; matched_name: string; score: number }[];
    }[];
  };

  try {
    data = JSON.parse(jsonText);
  } catch (e) {
    console.log("JSON パースエラー");
    return;
  }

  const results = data.results || [];
  if (results.length === 0) {
    console.log("results が空です");
    return;
  }

  // 結果シート
  const ws = workbook.addWorksheet("マッピング結果");
  const headers = [
    "ステータス", "項目名", "略称", "元JLAC10",
    "マッチJLAC10", "マッチ名称", "スコア", "分析物", "候補2", "候補3",
  ];

  // ヘッダー
  for (let c = 0; c < headers.length; c++) {
    const cell = ws.getRange(`${String.fromCharCode(65 + c)}1`);
    cell.setValue(headers[c]);
    cell.getFormat().getFill().setColor("4472C4");
    cell.getFormat().getFont().setColor("FFFFFF");
    cell.getFormat().getFont().setBold(true);
  }

  // データ
  for (let r = 0; r < results.length; r++) {
    const item = results[r];
    const row = r + 2;
    const best = item.best_match;
    const status = item.status === "auto" ? "自動確定" : item.status === "candidate" ? "要選択" : "手動";

    ws.getRange(`A${row}`).setValue(status);
    ws.getRange(`B${row}`).setValue(item.item_name || "");
    ws.getRange(`C${row}`).setValue(item.abbreviation || "");
    ws.getRange(`D${row}`).setValue(item.original_jlac10 || "");
    ws.getRange(`E${row}`).setValue(best ? best.jlac10 : "");
    ws.getRange(`F${row}`).setValue(best ? best.matched_name : "");
    ws.getRange(`G${row}`).setValue(best ? best.score : "");
    ws.getRange(`H${row}`).setValue(best ? best.analyte_code : "");
    ws.getRange(`I${row}`).setValue(item.candidates[1] ? item.candidates[1].jlac10 : "");
    ws.getRange(`J${row}`).setValue(item.candidates[2] ? item.candidates[2].jlac10 : "");

    // 色分け
    let fillColor = "FFC7CE"; // manual = 赤
    if (item.status === "auto") fillColor = "C6EFCE";
    else if (item.status === "candidate") fillColor = "FFEB9C";

    for (let c = 0; c < headers.length; c++) {
      ws.getRange(`${String.fromCharCode(65 + c)}${row}`).getFormat().getFill().setColor(fillColor);
    }
  }

  ws.getRange("A:J").getFormat().autofitColumns();

  // サマリーシート
  const meta = data.metadata;
  const sumWs = workbook.addWorksheet("サマリー");
  const summaryData = [
    ["全項目数", meta.total],
    ["自動確定", meta.auto],
    ["要選択", meta.candidate],
    ["手動", meta.manual],
  ];
  for (let i = 0; i < summaryData.length; i++) {
    sumWs.getRange(`A${i + 1}`).setValue(summaryData[i][0] as string);
    sumWs.getRange(`B${i + 1}`).setValue(summaryData[i][1] as number);
    sumWs.getRange(`A${i + 1}`).getFormat().getFont().setBold(true);
  }

  sheet.setName("_JSON_INPUT");
  console.log(`変換完了: ${results.length}件 (auto=${meta.auto}, candidate=${meta.candidate}, manual=${meta.manual})`);
}
