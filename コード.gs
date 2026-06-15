/**
 * コード.gs — 今週の良かったこと3つだけ日記 バックエンド
 * =====================================================
 * Google Apps Script に貼り付けて使います。
 *
 * スプレッドシートのヘッダー構成（1行目）:
 *   A列: id  |  B列: good1  |  C列: good2  |  D列: good3
 * =====================================================
 */

// ▼ あなたのGoogleスプレッドシートのIDをここに貼り付けてください ▼
const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE";
// ▲ スプレッドシートのURL: https://docs.google.com/spreadsheets/d/【ここ】/edit ▲

const SHEET_NAME = "diary"; // シート名（後述の手順で作成するシート名と合わせてください）

// =====================================================
// GETリクエスト処理（データ取得）
// =====================================================
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === "read") {
    return handleRead();
  }

  // デフォルトはread
  return handleRead();
}

// =====================================================
// POSTリクエスト処理（作成・更新・削除）
// =====================================================
function doPost(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = params.action;

    if (action === "create") {
      return handleCreate(params);
    } else if (action === "update") {
      return handleUpdate(params);
    } else if (action === "delete") {
      return handleDelete(params);
    } else {
      return jsonResponse({ status: "error", message: "不明なactionです: " + action });
    }
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

// =====================================================
// Read: 全データ取得
// =====================================================
function handleRead() {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      // ヘッダーのみ、またはデータなし
      return jsonResponse({ status: "ok", entries: [] });
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const entries = data
      .filter(row => row[0] !== "") // 空行をスキップ
      .map(row => ({
        id: String(row[0]),
        good1: String(row[1]),
        good2: String(row[2]),
        good3: String(row[3]),
      }));

    return jsonResponse({ status: "ok", entries: entries });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

// =====================================================
// Create: 新規投稿
// =====================================================
function handleCreate(params) {
  const good1 = sanitize(params.good1);
  const good2 = sanitize(params.good2);
  const good3 = sanitize(params.good3);

  if (!good1 || !good2 || !good3) {
    return jsonResponse({ status: "error", message: "3つすべての内容が必要です。" });
  }

  const id = String(new Date().getTime()); // ミリ秒タイムスタンプをIDに使用
  const sheet = getSheet();
  sheet.appendRow([id, good1, good2, good3]);

  return jsonResponse({ status: "ok", id: id });
}

// =====================================================
// Update: 既存投稿の更新
// =====================================================
function handleUpdate(params) {
  const id = sanitize(params.id);
  const good1 = sanitize(params.good1);
  const good2 = sanitize(params.good2);
  const good3 = sanitize(params.good3);

  if (!id || !good1 || !good2 || !good3) {
    return jsonResponse({ status: "error", message: "id と 3つの内容が必要です。" });
  }

  const sheet = getSheet();
  const rowIndex = findRowById(sheet, id);

  if (rowIndex === -1) {
    return jsonResponse({ status: "error", message: "指定されたIDの記録が見つかりません。" });
  }

  sheet.getRange(rowIndex, 2, 1, 3).setValues([[good1, good2, good3]]);
  return jsonResponse({ status: "ok" });
}

// =====================================================
// Delete: 投稿の削除
// =====================================================
function handleDelete(params) {
  const id = sanitize(params.id);

  if (!id) {
    return jsonResponse({ status: "error", message: "削除するIDが指定されていません。" });
  }

  const sheet = getSheet();
  const rowIndex = findRowById(sheet, id);

  if (rowIndex === -1) {
    return jsonResponse({ status: "error", message: "指定されたIDの記録が見つかりません。" });
  }

  sheet.deleteRow(rowIndex);
  return jsonResponse({ status: "ok" });
}

// =====================================================
// ユーティリティ関数
// =====================================================

/**
 * シート取得（なければ作成し、ヘッダーを設定）
 */
function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["id", "good1", "good2", "good3"]);
    sheet.setFrozenRows(1); // ヘッダー行を固定
    // ヘッダー行の書式設定
    const headerRange = sheet.getRange(1, 1, 1, 4);
    headerRange.setBackground("#f3e8f0");
    headerRange.setFontWeight("bold");
  }

  return sheet;
}

/**
 * A列からIDで行インデックス（1-indexed）を探す
 * @returns {number} 行番号（見つからない場合は -1）
 */
function findRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;

  const idColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < idColumn.length; i++) {
    if (String(idColumn[i][0]) === String(id)) {
      return i + 2; // 2行目スタート（1行目はヘッダー）
    }
  }
  return -1;
}

/**
 * JSONレスポンスを返す（CORS対応）
 */
function jsonResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * 文字列のサニタイズ（前後の空白除去・200文字制限）
 */
function sanitize(value) {
  if (value === undefined || value === null) return "";
  return String(value).trim().slice(0, 200);
}
