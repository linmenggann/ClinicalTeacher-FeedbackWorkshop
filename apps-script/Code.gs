/**
 * 臨床教師的心理安全與支持式回饋工作坊 — 線上報名後端
 * 將網頁報名表單資料寫入 Google Sheets 試算表。
 *
 * 試算表：https://docs.google.com/spreadsheets/d/1oghXn_uNIrESKl-i7WZu8XyAyskizCGgtLnzCVdQNTw/edit
 * 分頁名稱：工作坊報名資料
 *
 * ── 部署步驟 ──────────────────────────────────────────────
 * 1. 開啟上方試算表 → 擴充功能 (Extensions) → Apps Script。
 * 2. 將本檔內容貼入 Code.gs，存檔。
 * 3. 點「部署」→「新增部署作業」→ 類型選「網頁應用程式 (Web app)」。
 *    - 執行身分 (Execute as)：我 (你的帳號)
 *    - 具有存取權的使用者 (Who has access)：所有人 (Anyone)
 * 4. 複製產生的 /exec 網址，貼到 index.html 的 SCRIPT_URL 常數。
 * 5. 第一次部署會要求授權，請依指示允許。
 *
 * 試算表第一列表頭 (請手動建立，或執行 setupHeaders 函式自動建立)：
 *   報名時間 | 姓名 | 人事號 | 職稱 | 院區 | 單位 | 連絡電話/分機 | E-mail
 * ──────────────────────────────────────────────────────────
 */

var SHEET_NAME = '工作坊報名資料';
var HEADERS = ['報名時間', '姓名', '人事號', '職稱', '院區', '單位', '連絡電話/分機', 'E-mail'];

/** 接收網頁 POST 並寫入試算表 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // 避免同時寫入造成資料覆蓋
  try {
    var data = JSON.parse(e.postData.contents);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    // 若工作表是空的，先補上表頭
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    sheet.appendRow([
      timestamp,
      data.name  || '',
      data.empno || '',
      data.title || '',
      data.campus || '',
      data.dept  || '',
      data.phone || '',
      data.email || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/** 供瀏覽器直接開啟 /exec 網址時的健康檢查 */
function doGet() {
  return ContentService
    .createTextOutput('工作坊報名後端運作中。')
    .setMimeType(ContentService.MimeType.TEXT);
}

/** 一次性建立表頭（可在 Apps Script 編輯器中手動執行此函式） */
function setupHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}
