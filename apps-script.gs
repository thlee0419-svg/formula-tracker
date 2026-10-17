// 구글 시트 > 확장 프로그램 > Apps Script 에 붙여넣고
// 배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 모든 사용자) 로 배포하세요.
const SHEET = '재료';
const HEAD = ['id', '이름', '수량', '위치', '유통기한', '등록자', '등록일', '상태'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.appendRow(HEAD); }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const items = rows.filter(r => r[7] === '보관중').map(r => ({
    id: r[0], name: r[1], qty: r[2], loc: r[3],
    expiry: r[4] instanceof Date ? Utilities.formatDate(r[4], 'Asia/Seoul', 'yyyy-MM-dd') : String(r[4]),
    by: r[5],
    added: r[6] instanceof Date ? Utilities.formatDate(r[6], 'Asia/Seoul', 'yyyy-MM-dd') : String(r[6]),
  }));
  return out_({ items });
}
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const s = sheet_();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (d.action === 'add') {
      s.appendRow([d.id, d.name, d.qty, d.loc, d.expiry, d.by, d.added || new Date(), '보관중']);
    } else if (d.action === 'done') { // status: 먹음 | 버림
      const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().flat();
      const i = ids.indexOf(d.id);
      if (i > 0) s.getRange(i + 1, 8).setValue(d.status);
    }
  } finally { lock.release(); }
  return out_({ ok: true });
}
