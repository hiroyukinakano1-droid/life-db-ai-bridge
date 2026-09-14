// ==========================================
// ★新規追加：Claude（AI）からの書き込みを受け付けるWeb App橋渡し関数
// 人生DBの任意シートに1行を安全に追記するためのエンドポイント
// ==========================================

// ここに、あなただけが知っている合言葉（トークン）を設定してください。
// 推測されにくい文字列にしてください（例："hiro-lifedb-2026-xyz789"など）。
const ACCESS_TOKEN = "ここに任意の合言葉を設定";

function doGet(e) {
  try {
    // ① トークン確認（合言葉が一致しない場合は拒否）
    if (!e.parameter.token || e.parameter.token !== ACCESS_TOKEN) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: "error", message: "unauthorized" })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // ② パラメータ取得
    //   sheet : シート名（例："01_事実台帳"）またはシートのインデックス番号（0始まり）
    //   data  : 追記する行データ。JSON配列の文字列（例：'["F-010","職歴","...","2026-09-10"]'）
    const sheetParam = e.parameter.sheet;
    const dataParam = e.parameter.data;

    if (!sheetParam || !dataParam) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: "error", message: "missing sheet or data parameter" })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const rowData = JSON.parse(dataParam); // 例: ["F-010", "職歴", "...", "2026-09-10"]

    // ③ 対象シートを取得（シート名 or インデックスどちらでも対応）
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet;
    if (isNaN(sheetParam)) {
      sheet = ss.getSheetByName(sheetParam);
    } else {
      sheet = ss.getSheets()[parseInt(sheetParam, 10)];
    }

    if (!sheet) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: "error", message: "sheet not found: " + sheetParam })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // ④ 行を追記し、見た目を既存の他の行に揃える
    sheet.appendRow(rowData);
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow, 1, 1, rowData.length);
    range.setWrap(true);
    range.setVerticalAlignment("top");
    range.setFontSize(13);

    return ContentService.createTextOutput(
      JSON.stringify({ status: "ok", sheet: sheet.getName(), row: lastRow })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
