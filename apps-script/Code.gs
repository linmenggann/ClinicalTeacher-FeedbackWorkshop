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

/* ══════════════════════════════════════════════════════════
 * 課前提醒信（方案 A：Apps Script 一鍵寄送）
 *
 * 使用方式（在 Apps Script 編輯器中）：
 *   1. 先執行 sendReminderTest()  → 只寄一封給你自己，確認樣式與附件
 *   2. 確認無誤後執行 sendReminderEmails() → 逐一寄給全部報名者
 *
 * 說明：
 *   - 寄件人為執行者的 Google 帳號（linmenggann@gmail.com）
 *   - 附件海報 PDF 由 GitHub Pages 自動抓取，無需上傳 Drive
 *   - 逐一個別寄送，收件人彼此看不到對方信箱
 *   - Gmail 免費帳號每日上限約 100 封；本名單 49 封在額度內
 * ══════════════════════════════════════════════════════════ */

var REMINDER_SUBJECT = '【課前提醒】7/4(六)臨床教學中的拋與接-教學溝通的心理安全感與回饋技巧工作坊';
var REMINDER_SENDER_NAME = '奇美醫院教學部';
var POSTER_PDF_URL = 'https://linmenggann.github.io/ClinicalTeacher-FeedbackWorkshop/' +
  encodeURIComponent('1150704臨床教學中的拋與接v2.pdf');
var POSTER_PDF_NAME = '1150704臨床教學中的拋與接v2.pdf';

/** 信件 HTML 內文（仿院內範本樣式：微軟正黑體、深藍標題、黃底日期強調） */
function buildReminderHtml() {
  var blue = '#121b89', hl = 'background-color:#e9eafc;';
  return '' +
  '<div style="font-family:微軟正黑體,Microsoft JhengHei,sans-serif;color:#1f1f1f;">' +
    '<div><b><span style="font-size:26px;color:' + blue + ';">7/4</span>' +
      '<span style="font-size:18px;color:' + blue + ';">(六)</span>' +
      '<span style="font-size:26px;color:' + blue + ';">課前提醒📢📢</span></b></div>' +
    '<div><b><span style="font-size:24px;color:' + blue + ';' + hl + '">臨床教學中的拋與接-教學溝通的心理安全感與回饋技巧工作坊</span></b></div>' +
    '<br>' +
    '<div><b><span style="font-size:16px;color:' + blue + ';">您好，恭喜您已成功報名7/4(六)【臨床教學中的拋與接-教學溝通的心理安全感與回饋技巧工作坊】，課程相關資訊如下，敬請預留時間準時出席：</span></b></div>' +
    '<br>' +
    '<div><b style="background-color:#fff9dd;">【課程資訊】</b></div>' +
    '<ul>' +
      '<li><b>日期：</b>115 年 <b style="font-size:18px;color:' + blue + ';' + hl + '">7 月 4 日(六)</b></li>' +
      '<li><b>時間：</b>12:50 ~ 16:15（12:30 開始報到）</li>' +
      '<li><b>地點：</b>奇美醫院 第三醫療大樓331、332會議室（710臺南市永康區中華路901號）</li>' +
      '<li><b>學分申請：</b>臨床教師認證：「進階」2 學分、Ac 類「敘事醫學」1 學分</li>' +
    '</ul>' +
    '<div><b style="background-color:#fff9dd;">【注意事項】</b></div>' +
    '<ul>' +
      '<li>若您因故無法參與課程，敬請提前致電予教學部詩芸（分機57439），以利候補同仁遞補參與，謝謝。</li>' +
      '<li>本課程須<b><u>全程參與</u></b>並完成課程當天的<b><u>課後問卷</u></b>，始得核予學分。</li>' +
      '<li>因為課程中會有很多內容素材不會直接在簡報中，且會有需要走動的情境。參與者可自行攜帶手機做內容拍照紀錄。但平板或筆電做筆記的話，估計對個人參與狀態可能會比較干擾，因此較不建議攜帶。</li>' +
    '</ul>' +
    '<br>' +
    '<div style="font-size:12px;">=================================<br>' +
    '奇美醫療財團法人奇美醫院<br>' +
    '教學部 陳詩芸教學行政管理員<br>' +
    '電話：06-2812811分機57439／簡碼：1412<br>' +
    '信箱：b20715@mail.chimei.org.tw<br>' +
    '地址：台南市永康區中華路901號<br>' +
    '=================================</div>' +
  '</div>';
}

/** 純文字備援內文 */
function buildReminderText() {
  return '7/4(六)課前提醒\n臨床教學中的拋與接-教學溝通的心理安全感與回饋技巧工作坊\n\n' +
  '您好，恭喜您已成功報名7/4(六)【臨床教學中的拋與接-教學溝通的心理安全感與回饋技巧工作坊】，課程相關資訊如下，敬請預留時間準時出席：\n\n' +
  '【課程資訊】\n' +
  '日期：115 年 7 月 4 日(六)\n' +
  '時間：12:50 ~ 16:15（12:30 開始報到）\n' +
  '地點：奇美醫院 第三醫療大樓331、332會議室（710臺南市永康區中華路901號）\n' +
  '學分申請：臨床教師認證：「進階」2 學分、Ac 類「敘事醫學」1 學分\n\n' +
  '【注意事項】\n' +
  '．若您因故無法參與課程，敬請提前致電予教學部詩芸（分機57439），以利候補同仁遞補參與，謝謝。\n' +
  '．本課程須全程參與並完成課程當天的課後問卷，始得核予學分。\n' +
  '．因為課程中會有很多內容素材不會直接在簡報中，且會有需要走動的情境。參與者可自行攜帶手機做內容拍照紀錄。但平板或筆電做筆記的話，估計對個人參與狀態可能會比較干擾，因此較不建議攜帶。\n\n' +
  '=================================\n奇美醫療財團法人奇美醫院\n教學部 陳詩芸教學行政管理員\n' +
  '電話：06-2812811分機57439／簡碼：1412\n信箱：b20715@mail.chimei.org.tw\n地址：台南市永康區中華路901號\n=================================';
}

/** 從報名分頁讀取收件人（去重、過濾空白與無效信箱） */
function getReminderRecipients() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var out = [], seen = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  values.forEach(function (r) {
    var name = String(r[3] || '').trim();
    var email = String(r[7] || '').trim();
    if (!email || email.indexOf('@') < 1) return;
    var key = email.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    out.push({ name: name, email: email });
  });
  return out;
}

/** 抓取海報 PDF 附件 */
function getPosterBlob() {
  var res = UrlFetchApp.fetch(POSTER_PDF_URL);
  return res.getBlob().setName(POSTER_PDF_NAME);
}

/** 收件備援：若無法取得執行者信箱時使用 */
var REMINDER_FALLBACK_EMAIL = 'linmenggann@gmail.com';

function getSelfEmail() {
  var me = '';
  try { me = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!me) { try { me = Session.getEffectiveUser().getEmail(); } catch (e) {} }
  return me || REMINDER_FALLBACK_EMAIL;
}

/** 步驟 1：先寄一封測試信給自己，確認樣式與附件 */
function sendReminderTest() {
  var me = getSelfEmail();
  Logger.log('本日剩餘寄信額度：' + MailApp.getRemainingDailyQuota());
  Logger.log('收件地址：' + me);
  var blob = getPosterBlob();
  Logger.log('附件已抓取：' + blob.getName() + '（' + blob.getBytes().length + ' bytes）');
  MailApp.sendEmail({
    to: me,
    subject: '[測試] ' + REMINDER_SUBJECT,
    htmlBody: buildReminderHtml(),
    body: buildReminderText(),
    attachments: [blob],
    name: REMINDER_SENDER_NAME
  });
  Logger.log('✅ 測試信已寄出至：' + me + '（請同時檢查垃圾郵件匣）');
}

/** 補寄對象：修改此地址後執行 sendReminderSupplement（多人可用逗號分隔） */
var SUPPLEMENT_EMAILS = 'B41242@chimei.org.tw';

/** 補寄課前提醒給指定收件人（晚報名者適用） */
function sendReminderSupplement() {
  var blob = getPosterBlob();
  var html = buildReminderHtml();
  var text = buildReminderText();
  var list = SUPPLEMENT_EMAILS.split(',').map(function (s) { return s.trim(); }).filter(String);
  list.forEach(function (email) {
    MailApp.sendEmail({
      to: email,
      subject: REMINDER_SUBJECT,
      htmlBody: html,
      body: text,
      attachments: [blob],
      name: REMINDER_SENDER_NAME
    });
    Logger.log('✅ 已補寄至：' + email);
    Utilities.sleep(500);
  });
}

/** 步驟 2：正式寄送給全部報名者（逐一個別寄送） */
function sendReminderEmails() {
  var recipients = getReminderRecipients();
  if (recipients.length === 0) { Logger.log('查無收件人'); return; }
  var blob = getPosterBlob();
  var html = buildReminderHtml();
  var text = buildReminderText();
  var ok = 0, fail = [];
  recipients.forEach(function (p) {
    try {
      MailApp.sendEmail({
        to: p.email,
        subject: REMINDER_SUBJECT,
        htmlBody: html,
        body: text,
        attachments: [blob],
        name: REMINDER_SENDER_NAME
      });
      ok++;
      Utilities.sleep(500); // 避免觸發速率限制
    } catch (err) {
      fail.push(p.name + ' <' + p.email + '>：' + err.message);
    }
  });
  var report = '寄送完成：成功 ' + ok + ' 封 / 共 ' + recipients.length + ' 位' +
    (fail.length ? '\n失敗清單：\n' + fail.join('\n') : '');
  Logger.log(report);
  // 寄一份結果摘要給自己
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: '[寄送結果] ' + REMINDER_SUBJECT,
    body: report,
    name: REMINDER_SENDER_NAME
  });
}
