// ===== Service Worker: يجلب آخر نسخة دائماً من الشبكة، ويرجع للكاش عند انقطاع الاتصال =====
// غيّر رقم النسخة عند كل تحديث لتنظيف الكاش القديم
const VERSION = 'v16';
const CACHE = 'supervision-app-' + VERSION;

self.addEventListener('install', () => self.skipWaiting());

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  // نتعامل مع ملفات التطبيق فقط؛ الخطوط والمكتبات وروابط Google Sheets تمر كما هي
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;

  e.respondWith(
    fetch(req, { cache: 'no-store' })
      .then(res => {
        if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
        return res;
      })
      .catch(() => caches.match(req).then(r => r || caches.match('./index.html')))
  );
});
