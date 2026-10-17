// 구글 시트 > 확장 프로그램 > Apps Script 에 붙여넣고
// 배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 모든 사용자) 로 배포하세요.
const SHEET = '재료';
const HEAD = ['id', '이름', '수량', '위치', '유통기한', '등록자', '등록일', '상태', '처리일', '버린양(g)'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.appendRow(HEAD); }
  // 예전에 만든 시트에는 처리일/버린양 칸이 없으므로 머리글을 채워 둔다
  if (s.getRange(1, 9).getValue() === '') s.getRange(1, 9, 1, 2).setValues([HEAD.slice(8)]);
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function ymd_(v) {
  return v instanceof Date ? Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd') : String(v);
}
function doGet(e) {
  const rows = sheet_().getDataRange().getValues().slice(1);
  // 버린 음식 기록 (통계용)
  if (e && e.parameter && e.parameter.action === 'waste') {
    const waste = rows.filter(r => r[7] === '버림').map(r => ({
      name: r[1], loc: r[3], by: r[5], added: ymd_(r[6]), date: ymd_(r[8]),
      grams: r[9] === '' ? null : Number(r[9]),
    }));
    return out_({ waste });
  }
  // 먹은 음식 기록 (자주 먹는 재료 통계용)
  if (e && e.parameter && e.parameter.action === 'eaten') {
    const eaten = rows.filter(r => r[7] === '먹음').map(r => ({
      name: r[1], loc: r[3], by: r[5], added: ymd_(r[6]), date: r[8] === '' ? '' : ymd_(r[8]),
    }));
    return out_({ eaten });
  }
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
      if (i > 0) {
        // 상태, 처리일, 버린양(g)을 한 번에 기록
        s.getRange(i + 1, 8, 1, 3).setValues([[d.status, d.date || new Date(), d.grams === '' || d.grams == null ? '' : d.grams]]);
      }
    }
  } finally { lock.release(); }
  return out_({ ok: true });
}
