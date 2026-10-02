/**
 * stocks.html 용 Google Sheets 동기화 스크립트
 *
 * 설정 방법
 * 1) Google Sheets에서 새 시트를 만들고 [확장 프로그램 > Apps Script]를 엽니다.
 * 2) 이 코드를 붙여넣고 저장합니다.
 * 3) [배포 > 새 배포 > 유형: 웹 앱] 선택
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 4) 발급된 웹 앱 URL(…/exec)을 stocks.html 상단 입력란에 붙여넣습니다.
 *
 * 전체 상태(JSON)를 'data' 시트의 A열에 40,000자씩 나눠 저장합니다.
 */
const SHEET = 'data';
const CHUNK = 40000;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
}

function doGet(e) {
  const rows = getSheet_().getRange(1, 1, Math.max(getSheet_().getLastRow(), 1), 1).getValues();
  const json = rows.map(r => r[0]).join('');
  const state = json ? JSON.parse(json) : null;
  return ContentService.createTextOutput(JSON.stringify({ state })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (body.action === 'save') {
    const json = JSON.stringify(body.state);
    const sheet = getSheet_();
    sheet.clear();
    const parts = [];
    for (let i = 0; i < json.length; i += CHUNK) parts.push([json.slice(i, i + CHUNK)]);
    sheet.getRange(1, 1, parts.length, 1).setValues(parts);
  }
  return ContentService.createTextOutput('ok');
}
