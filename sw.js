// 내 주식 대시보드 서비스 워커: 오프라인에서도 앱이 열리도록 화면 파일을 보관한다.
// 투자 데이터(구글 시트 통신)는 건드리지 않는다.
const CACHE = 'stocks-v1';
const CHART_URL = 'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js';
const SHELL = ['./stocks.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png', CHART_URL];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(SHELL.map(u => c.add(u).catch(() => {})))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET') return;
  const isChart = req.url === CHART_URL;
  const isShell = url.origin === location.origin && /\/(stocks\.html|manifest\.webmanifest|icons\/[^/]+)$/.test(url.pathname);
  if (!isChart && !isShell) return;               // 시트 통신 등은 그대로 통과
  if (isChart) {                                   // 차트 라이브러리: 캐시 우선
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return res; })));
    return;
  }
  e.respondWith(fetch(req, { cache: 'no-cache' }).then(res => {  // 화면 파일: 네트워크 우선(브라우저 임시 캐시도 건너뜀), 오프라인이면 캐시
    const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return res;
  }).catch(() => caches.match(req, { ignoreSearch: true })));
});
