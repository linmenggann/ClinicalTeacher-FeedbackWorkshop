/**
 * 臨床教學中的拋與接-教學溝通的心理安全感與回饋技巧工作坊
 * Google Sheets 後端：處理「線上報名」與「當天報到」兩種寫入。
 *
 * 試算表：https://docs.google.com/spreadsheets/d/1oghXn_uNIrESKl-i7WZu8XyAyskizCGgtLnzCVdQNTw/edit
 * 分頁：工作坊報名資料（報名）、工作坊報到資料（報到）
 *
 * ── 部署步驟 ──────────────────────────────────────────────
 * 1. 開啟上方試算表 → 擴充功能 (Extensions) → Apps Script。
 * 2. 將本檔內容貼入 Code.gs，存檔。
 * 3. （可先執行一次 setupHeaders 與 setupCheckinSheet 建立兩個分頁表頭）
 * 4. 點「部署」→「管理部署作業」→ 編輯 → 版本選「新版本」→ 部署。
 *    - 執行身分 (Execute as)：我 (你的帳號)
 *    - 具有存取權的使用者 (Who has access)：所有人 (Anyone)
 * 5. 網址不變；沿用同一組 /exec 網址給 index.html 與 checkin.html。
 *
 * 儀表板 (dashboard.html / checkin-dashboard.html) 直接讀取 GViz，
 * 需將試算表共用設為「知道連結的任何人 → 檢視者」。
 * ──────────────────────────────────────────────────────────
 */

var SHEET_NAME = '工作坊報名資料';
var HEADERS = ['報名時間', '院區', '單位', '姓名', '人事號', '職稱', '連絡電話/分機', 'E-mail'];
var CAPACITY = 50; // 報名人數上限；達上限即停止受理

var CHECKIN_SHEET_NAME = '工作坊報到資料';
var CHECKIN_HEADERS = ['報到時間', '人事號', '院區', '單位', '姓名', '職稱'];

/** 目前已報名人數（不含表頭） */
function countRegistrations(sheet) {
  if (!sheet) return 0;
  return Math.max(0, sheet.getLastRow() - 1);
}

/** 取得分頁，不存在則建立並補上表頭 */
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 以人事號在報名資料中查詢報名者，找到回傳其基本資料，否則 null */
function findRegistrationByEmpno(ss, empno) {
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var target = String(empno).trim();
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][4]).trim() === target) { // 人事號在第 5 欄
      return { campus: values[i][1], dept: values[i][2], name: values[i][3], title: values[i][5] };
    }
  }
  return null;
}

/** 以人事號在報到資料中查詢是否已報到，找到回傳報到時間與姓名，否則 null */
function findCheckinByEmpno(sheet, empno) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  var target = String(empno).trim();
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, CHECKIN_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][1]).trim() === target) { // 人事號在第 2 欄
      return { time: values[i][0], name: values[i][4] };
    }
  }
  return null;
}

/** 接收網頁 POST：依 action 分流為報到 (checkin) 或報名 (預設) */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // 避免同時寫入造成資料覆蓋
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.action === 'checkin') {
      return handleCheckin(ss, data);
    }
    return handleRegister(ss, data);
  } catch (err) {
    return jsonOut({ result: 'error', message: err.message });
  } finally {
    lock.releaseLock();
  }
}

/** 報名寫入 */
function handleRegister(ss, data) {
  var sheet = getOrCreateSheet(ss, SHEET_NAME, HEADERS);

  // 達報名上限即拒絕，避免超收
  if (countRegistrations(sheet) >= CAPACITY) {
    return jsonOut({ result: 'full', capacity: CAPACITY });
  }

  var timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  sheet.appendRow([
    timestamp,
    data.campus || '',
    data.dept  || '',
    data.name  || '',
    data.empno || '',
    data.title || '',
    data.phone || '',
    data.email || ''
  ]);
  return jsonOut({ result: 'success' });
}

/** 報到寫入：驗證人事號是否已報名、是否重複報到 */
function handleCheckin(ss, data) {
  var empno = String(data.empno || '').trim();
  if (!empno) return jsonOut({ result: 'error', message: '請輸入人事號' });

  var reg = findRegistrationByEmpno(ss, empno);
  if (!reg) return jsonOut({ result: 'notfound', empno: empno });

  var csheet = getOrCreateSheet(ss, CHECKIN_SHEET_NAME, CHECKIN_HEADERS);

  var existing = findCheckinByEmpno(csheet, empno);
  if (existing) {
    return jsonOut({ result: 'already', name: reg.name, campus: reg.campus, dept: reg.dept, time: existing.time });
  }

  var timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  csheet.appendRow([timestamp, empno, reg.campus, reg.dept, reg.name, reg.title]);
  return jsonOut({
    result: 'success', empno: empno, name: reg.name,
    campus: reg.campus, dept: reg.dept, title: reg.title, time: timestamp
  });
}

/**
 * 提供報名資料給儀表板 (dashboard.html) 讀取（保留；儀表板現主要改用 GViz）。
 */
function doGet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    var records = [];
    if (sheet && sheet.getLastRow() > 1) {
      var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
      records = values.map(function (r) {
        return {
          time:   r[0], campus: r[1], dept: r[2], name: r[3],
          empno:  r[4], title: r[5], phone: r[6], email: r[7]
        };
      });
    }
    return jsonOut({
      result: 'success', count: records.length, capacity: CAPACITY,
      full: records.length >= CAPACITY, records: records
    });
  } catch (err) {
    return jsonOut({ result: 'error', message: err.message });
  }
}

/** 一次性建立報名分頁表頭 */
function setupHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/** 一次性建立報到分頁與表頭 */
function setupCheckinSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHECKIN_SHEET_NAME) || ss.insertSheet(CHECKIN_SHEET_NAME);
  sheet.getRange(1, 1, 1, CHECKIN_HEADERS.length).setValues([CHECKIN_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}
