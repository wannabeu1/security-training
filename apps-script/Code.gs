const SHEET_NAME = '접속기록';
const HEADERS = [
  '서버 기록시각',
  '입력 이름',
  '캠페인',
  '브라우저 기록시각',
  '페이지 URL',
  '이전 페이지',
  '브라우저 정보',
  '언어',
  '공인 IP'
];

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const name = cleanText_(data.name, 50);

    if (name.length < 2) {
      return jsonResponse_({ok: false, error: 'invalid_name'});
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    if (sheet.getLastRow() === 0 || sheet.getRange(1, HEADERS.length).getValue() !== HEADERS[HEADERS.length - 1]) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }

    sheet.appendRow([
      new Date(),
      name,
      cleanText_(data.campaign, 100),
      cleanText_(data.openedAt, 40),
      safeCell_(data.pageUrl, 500),
      safeCell_(data.referrer, 500),
      safeCell_(data.userAgent, 500),
      cleanText_(data.language, 30),
      safeCell_(data.ipAddress, 64)
    ]);

    return jsonResponse_({ok: true});
  } catch (error) {
    console.error(error);
    return jsonResponse_({ok: false, error: 'server_error'});
  } finally {
    try {
      lock.releaseLock();
    } catch (ignored) {}
  }
}

function doGet() {
  return jsonResponse_({ok: true, service: 'security-training-collector'});
}

function cleanText_(value, maxLength) {
  return String(value || '').trim().replace(/\s+/g, ' ').slice(0, maxLength);
}

// 스프레드시트 수식 삽입을 막기 위해 위험한 첫 문자를 이스케이프합니다.
function safeCell_(value, maxLength) {
  const text = cleanText_(value, maxLength);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
