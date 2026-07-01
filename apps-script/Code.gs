/**
 * 臨床教學中的拋與接-教學溝通的心理安全感與回饋技巧工作坊
 * Google Sheets 後端：處理「線上報名」、「當天報到/簽退」與「滿意度問卷」寫入。
 *
 * 試算表：https://docs.google.com/spreadsheets/d/1oghXn_uNIrESKl-i7WZu8XyAyskizCGgtLnzCVdQNTw/edit
 * 分頁：
 *   工作坊報名資料      — 報名
 *   工作坊報到/簽退資料 — 當天報到與簽退（同一列，報到時間 / 簽退時間）
 *   工作坊滿意度調查    — 課程滿意度問卷
 *
 * ── 部署步驟 ──────────────────────────────────────────────
 * 1. 開啟上方試算表 → 擴充功能 (Extensions) → Apps Script。
 * 2. 將本檔內容貼入 Code.gs，存檔。
 * 3. （可先執行 setupHeaders / setupCheckinSheet / setupSurveySheet 建立各分頁表頭）
 * 4. 「部署」→「管理部署作業」→ 編輯 → 版本選「新版本」→ 部署。
 *    - 執行身分：我　- 具有存取權的使用者：所有人 (Anyone)
 * 5. 網址不變；沿用同一組 /exec 給 index.html、checkin.html、survey.html。
 *
 * 儀表板 (dashboard.html / checkin-dashboard.html) 直接讀取 GViz，
 * 需將試算表共用設為「知道連結的任何人 → 檢視者」。
 * ──────────────────────────────────────────────────────────
 */

var SHEET_NAME = '工作坊報名資料';
var HEADERS = ['報名時間', '院區', '單位', '姓名', '人事號', '職稱', '連絡電話/分機', 'E-mail'];
var CAPACITY = 50; // 報名人數上限；達上限即停止受理

var CHECKIN_SHEET_NAME = '工作坊報到/簽退資料';
var CHECKIN_HEADERS = ['報到時間', '簽退時間', '院區', '單位', '姓名', '人事號', '職稱'];

var SURVEY_SHEET_NAME = '工作坊滿意度調查';
var SURVEY_HEADERS = ['填答時間', '院區', '單位', '姓名', '人事號', '職稱',
  'Q1整體滿意度', 'Q2主題對教學有幫助', 'Q3教學背後的基本功實用', 'Q4教學溝通困境交流有幫助',
  'Q5講師講授清楚', 'Q6講師引導互動', 'Q7營造心理安全能力', 'Q8運用支持式回饋', 'Q9時間流程安排', 'Q10願意推薦',
  '其他建議'];

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

function nowStamp() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
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

/** 以人事號在報到/簽退資料中查詢該列，回傳列號與報到/簽退時間，否則 null */
function findCheckinRow(sheet, empno) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  var target = String(empno).trim();
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, CHECKIN_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][5]).trim() === target) { // 人事號在第 6 欄
      return { rowIndex: i + 2, checkin: values[i][0], checkout: values[i][1], name: values[i][4] };
    }
  }
  return null;
}

/** 接收網頁 POST：依 action 分流 checkin / checkout / survey / 報名(預設) */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.action === 'checkin')  return handleCheckin(ss, data);
    if (data.action === 'checkout') return handleCheckout(ss, data);
    if (data.action === 'survey')   return handleSurvey(ss, data);
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
  if (countRegistrations(sheet) >= CAPACITY) {
    return jsonOut({ result: 'full', capacity: CAPACITY });
  }
  sheet.appendRow([
    nowStamp(), data.campus || '', data.dept || '', data.name || '',
    data.empno || '', data.title || '', data.phone || '', data.email || ''
  ]);
  return jsonOut({ result: 'success' });
}

/** 報到：驗證人事號、防重複報到 */
function handleCheckin(ss, data) {
  var empno = String(data.empno || '').trim();
  if (!empno) return jsonOut({ result: 'error', message: '請輸入人事號' });

  var reg = findRegistrationByEmpno(ss, empno);
  if (!reg) return jsonOut({ result: 'notfound', empno: empno });

  var csheet = getOrCreateSheet(ss, CHECKIN_SHEET_NAME, CHECKIN_HEADERS);
  var row = findCheckinRow(csheet, empno);
  var ts = nowStamp();

  if (row && row.checkin) {
    return jsonOut({ result: 'already', mode: 'checkin', name: reg.name, campus: reg.campus, dept: reg.dept, time: row.checkin });
  }
  if (row && !row.checkin) {
    // 已先簽退但無報到紀錄 → 補上報到時間
    csheet.getRange(row.rowIndex, 1).setValue(ts);
  } else {
    csheet.appendRow([ts, '', reg.campus, reg.dept, reg.name, empno, reg.title]);
  }
  return jsonOut({ result: 'success', mode: 'checkin', name: reg.name, campus: reg.campus, dept: reg.dept, title: reg.title, time: ts });
}

/** 簽退：驗證人事號、防重複簽退；未報到者仍可簽退（報到時間留空） */
function handleCheckout(ss, data) {
  var empno = String(data.empno || '').trim();
  if (!empno) return jsonOut({ result: 'error', message: '請輸入人事號' });

  var reg = findRegistrationByEmpno(ss, empno);
  if (!reg) return jsonOut({ result: 'notfound', empno: empno });

  var csheet = getOrCreateSheet(ss, CHECKIN_SHEET_NAME, CHECKIN_HEADERS);
  var row = findCheckinRow(csheet, empno);
  var ts = nowStamp();

  if (row) {
    if (row.checkout) {
      return jsonOut({ result: 'already', mode: 'checkout', name: reg.name, campus: reg.campus, dept: reg.dept, time: row.checkout });
    }
    csheet.getRange(row.rowIndex, 2).setValue(ts); // 簽退時間在第 2 欄
  } else {
    csheet.appendRow(['', ts, reg.campus, reg.dept, reg.name, empno, reg.title]);
  }
  return jsonOut({ result: 'success', mode: 'checkout', name: reg.name, campus: reg.campus, dept: reg.dept, title: reg.title, time: ts });
}

/** 滿意度問卷寫入；以人事號帶出院區/單位/姓名/職稱 */
function handleSurvey(ss, data) {
  var sheet = getOrCreateSheet(ss, SURVEY_SHEET_NAME, SURVEY_HEADERS);
  var a = data.answers || {};
  var empno = String(data.empno || '').trim();
  var reg = empno ? findRegistrationByEmpno(ss, empno) : null;
  sheet.appendRow([
    nowStamp(),
    reg ? reg.campus : '', reg ? reg.dept : '', reg ? reg.name : '', empno, reg ? reg.title : '',
    a.q1 || '', a.q2 || '', a.q3 || '', a.q4 || '', a.q5 || '',
    a.q6 || '', a.q7 || '', a.q8 || '', a.q9 || '', a.q10 || '',
    data.comment || ''
  ]);
  return jsonOut({ result: 'success' });
}

/** 提供報名資料給儀表板讀取（保留；儀表板現主要改用 GViz） */
function doGet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    var records = [];
    if (sheet && sheet.getLastRow() > 1) {
      var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
      records = values.map(function (r) {
        return { time: r[0], campus: r[1], dept: r[2], name: r[3], empno: r[4], title: r[5], phone: r[6], email: r[7] };
      });
    }
    return jsonOut({ result: 'success', count: records.length, capacity: CAPACITY, full: records.length >= CAPACITY, records: records });
  } catch (err) {
    return jsonOut({ result: 'error', message: err.message });
  }
}

/** 一次性建立各分頁表頭（可在 Apps Script 編輯器中手動執行） */
function setupHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}
function setupCheckinSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHECKIN_SHEET_NAME) || ss.insertSheet(CHECKIN_SHEET_NAME);
  sheet.getRange(1, 1, 1, CHECKIN_HEADERS.length).setValues([CHECKIN_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}
function setupSurveySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SURVEY_SHEET_NAME) || ss.insertSheet(SURVEY_SHEET_NAME);
  sheet.getRange(1, 1, 1, SURVEY_HEADERS.length).setValues([SURVEY_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}
