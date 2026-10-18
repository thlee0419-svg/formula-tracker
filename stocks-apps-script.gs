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
 * 현재가는 'quotes' 시트에서 GOOGLEFINANCE 함수로 조회하고, 구글에서 못 받은 종목(코스닥 등)은 야후 파이낸스로 보충합니다. (코드를 바꾼 뒤에는 새 버전으로 다시 배포해야 합니다)
 */
const SHEET = 'data';
const CHUNK = 40000;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'quotes') return json_(getQuotes_((p.t || '').split(',').filter(Boolean), p.rate === '1'));
  const rows = getSheet_().getRange(1, 1, Math.max(getSheet_().getLastRow(), 1), 1).getValues();
  const text = rows.map(r => r[0]).join('');
  const state = text ? JSON.parse(text) : null;
  return json_({ state });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 현재가 조회: 'quotes' 시트에 GOOGLEFINANCE 수식을 넣고 계산 결과를 읽어 온다.
 * 구글 파이낸스 시세는 약 15~20분 지연될 수 있다.
 * 티커 예) 한국: KRX:005930 / 미국: NVDA 또는 NASDAQ:NVDA
 */
function getQuotes_(symbols, withRate) {
  const SAFE = /^[A-Za-z0-9.:_-]{1,24}$/;               // 수식 주입 방지: 안전한 문자만 허용
  const list = symbols.filter(t => SAFE.test(t)).slice(0, 60);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('quotes') || ss.insertSheet('quotes');
  sh.clear();
  const rows = list.length + (withRate ? 1 : 0);
  if (!rows) return { quotes: {}, rate: null };
  sh.getRange(1, 1, rows, 1).setNumberFormat('@').setValues(list.map(t => [t]).concat(withRate ? [['USDKRW']] : []));
  sh.getRange(1, 2, rows, 1).setFormulas(
    list.map((t, i) => ['=GOOGLEFINANCE(A' + (i + 1) + ',"price")']).concat(withRate ? [['=GOOGLEFINANCE("CURRENCY:USDKRW")']] : []));
  let values = [];
  for (let tries = 0; tries < 8; tries++) {              // 계산이 끝날 때까지 잠깐씩 기다림
    SpreadsheetApp.flush();
    values = sh.getRange(1, 2, rows, 1).getValues().map(r => r[0]);
    if (!values.some(v => v === 'Loading...' || v === '')) break;
    Utilities.sleep(700);
  }
  const num = v => (typeof v === 'number' && isFinite(v) && v > 0) ? v : null;
  const quotes = {};
  list.forEach((t, i) => quotes[t] = num(values[i]));
  fillFromYahoo_(quotes);                                // 구글에서 못 받은 종목(코스닥 등)은 야후로 보충
  return { quotes, rate: withRate ? num(values[list.length]) : null };
}

/** 구글 파이낸스에서 값이 없는 종목만 야후 파이낸스(비공식)로 조회한다. */
function yahooSymbol_(sym) {
  if (sym.indexOf('KRX:') === 0) return sym.slice(4) + '.KS';        // 코스피
  if (sym.indexOf('KOSDAQ:') === 0) return sym.slice(7) + '.KQ';     // 코스닥
  const t = sym.indexOf(':') >= 0 ? sym.split(':')[1] : sym;         // NASDAQ:NVDA -> NVDA
  return t.replace(/\./g, '-');                                      // BRK.B -> BRK-B
}
function fillFromYahoo_(quotes) {
  const missing = Object.keys(quotes).filter(t => quotes[t] === null);
  if (!missing.length) return;
  const reqs = missing.map(t => ({
    url: 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(yahooSymbol_(t)) + '?interval=1d&range=1d',
    headers: { 'User-Agent': 'Mozilla/5.0' },
    muteHttpExceptions: true,
  }));
  let resps = [];
  try { resps = UrlFetchApp.fetchAll(reqs); } catch (err) { return; }   // 야후가 막혀도 나머지 결과는 그대로 돌려줌
  resps.forEach((r, i) => {
    try {
      if (r.getResponseCode() !== 200) return;
      const meta = JSON.parse(r.getContentText()).chart.result[0].meta;
      const price = meta.regularMarketPrice;
      if (typeof price === 'number' && isFinite(price) && price > 0) quotes[missing[i]] = price;
    } catch (err) {}
  });
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
