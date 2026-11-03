/**
 * class.html(담임 플랫폼) 용 Google Sheets 백업 스크립트
 *
 * 설정 방법
 * 1) Google Sheets에서 새 시트를 만들고 [확장 프로그램 > Apps Script]를 엽니다.
 * 2) 이 코드를 붙여넣고 저장합니다.
 * 3) [배포 > 새 배포 > 유형: 웹 앱] 선택
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자  (※ 주소를 아는 사람만 접근 가능. 주소를 외부에 공유하지 마세요)
 * 4) 발급된 웹 앱 URL(…/exec)을 class.html 하단 '구글 시트 연결'에 붙여넣습니다.
 *
 * 전체 데이터(JSON)를 'data' 시트 A열에 40,000자씩 나눠 저장합니다.
 * 학생 개인정보가 들어 있으므로 이 시트는 본인 계정에서만 공유하세요.
 */
const SHEET = 'data';
const CHUNK = 40000;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
}

function doGet() {
  const sh = sheet_();
  const rows = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  const text = rows.map(r => r[0]).join('');
  return ContentService.createTextOutput(JSON.stringify({ state: text ? JSON.parse(text) : null }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (body.action !== 'save') return ContentService.createTextOutput('ignored');
  const text = JSON.stringify(body.state);
  const chunks = [];
  for (let i = 0; i < text.length; i += CHUNK) chunks.push([text.slice(i, i + CHUNK)]);
  const sh = sheet_();
  sh.clear();
  sh.getRange(1, 1, chunks.length, 1).setValues(chunks);
  return ContentService.createTextOutput('ok');
}
