/**
 * habits.html(습관의 숲) 용 Google Sheets 동기화 스크립트
 *
 * 설정 방법 (처음 한 번만)
 * 1) Google Sheets에서 "습관의 숲" 같은 이름의 새 스프레드시트를 만듭니다. (주식·냉장고 시트와 따로 만드세요)
 * 2) [확장 프로그램 > Apps Script]를 열고 이 코드를 붙여넣어 저장합니다.
 * 3) [배포 > 새 배포 > 유형: 웹 앱] 선택
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 4) 발급된 웹 앱 URL(…/exec)을 habits.html 설정(⚙️) > 기기 간 동기화에 붙여넣습니다.
 * 5) 설정에 나오는 "다른 기기 설정 링크"를 휴대폰에서 열면 같은 시트로 바로 연결됩니다.
 *
 * 전체 상태(JSON)를 'habit-data' 시트의 A열에 40,000자씩 나눠 저장합니다.
 * (코드를 바꾼 뒤에는 새 버전으로 다시 배포해야 합니다)
 */
const SHEET = 'habit-data';
const CHUNK = 40000;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
}

function doGet() {
  const sh = getSheet_();
  const rows = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  const text = rows.map(r => r[0]).join('');
  return ContentService.createTextOutput(JSON.stringify({ state: text ? JSON.parse(text) : null }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (body.action === 'save' && body.state) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const json = JSON.stringify(body.state);
      const sh = getSheet_();
      sh.clear();
      const parts = [];
      for (let i = 0; i < json.length; i += CHUNK) parts.push([json.slice(i, i + CHUNK)]);
      sh.getRange(1, 1, parts.length, 1).setValues(parts);
    } finally { lock.releaseLock(); }
  }
  return ContentService.createTextOutput('ok');
}
