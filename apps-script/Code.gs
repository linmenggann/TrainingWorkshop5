/**
 * 教學訓練計畫主持人工作坊 - 報名資料接收
 *
 * 部署步驟：
 *   1. 開啟目標 Google Sheets（本專案的試算表）。
 *   2. 選單「擴充功能 Extensions → Apps Script」開啟編輯器。
 *   3. 將本檔案內容貼入 Code.gs 後儲存。
 *   4. 點「部署 Deploy → 新增部署作業 New deployment」。
 *   5. 類型選「網頁應用程式 Web app」。
 *   6. 執行身分：選擇「我」；存取權限：選擇「任何人 Anyone」。
 *   7. 點「部署」取得「網頁應用程式網址」，貼回 index.html 內的
 *      APPS_SCRIPT_URL 常數。
 *
 *   之後若修改程式碼，請使用「管理部署作業 → 編輯 → 新版本」重新部署，
 *   URL 才會指向最新版。
 */

const SHEET_ID = '1idrIOTSCd1Ev2CHZ0I306N-LyqJViEpxXyQAML-aw6I';
const SHEET_NAME = '工作坊報名資料';

const HEADERS = [
  '報名時間',
  '姓名',
  '機構名',
  '職稱',
  '負責的職類',
  '是否擔任教學訓練計畫主持人',
  '參與方式',
  'Email',
  '聯繫電話'
];

/**
 * 接收前端 POST 請求並寫入試算表。
 * 前端以 text/plain 傳送 JSON 字串（避免 CORS preflight）。
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    const sheet = getOrCreateSheet_();

    sheet.appendRow([
      new Date(),
      payload.name || '',
      payload.org || '',
      payload.title || '',
      payload.profession || '',
      payload.is_host || '',
      payload.mode || '',
      payload.email || '',
      payload.phone || ''
    ]);

    return jsonResponse_({ status: 'ok' });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: String(err) });
  }
}

/**
 * GET 測試端點：部署後於瀏覽器開啟 URL 會看到 OK 訊息。
 */
function doGet() {
  return jsonResponse_({ status: 'ok', message: '教學訓練計畫主持人工作坊 API is running.' });
}

/**
 * 一次性初始化：於 Apps Script 編輯器內手動執行，會建立分頁並寫入表頭、
 * 設定欄寬與凍結首列。已執行過可重複呼叫，不會重複寫入表頭。
 */
function setupSheet() {
  const sheet = getOrCreateSheet_();
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const isEmpty = firstRow.every(function (v) { return v === '' || v === null; });

  if (isEmpty) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#6a11cb');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 1, 170);
  sheet.setColumnWidths(2, HEADERS.length - 1, 160);
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sheet;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
