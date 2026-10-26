// ===== Service Worker: يجلب آخر نسخة دائماً من الشبكة، ويرجع للكاش عند انقطاع الاتصال =====
// غيّر رقم النسخة عند كل تحديث لتنظيف الكاش القديم وتطبيق التعديلات البرمجية فوراً
const VERSION = 'v17';
const CACHE = 'supervision-app-' + VERSION;

// تثبيت الـ Service Worker فوراً دون انتظار
self.addEventListener('install', () => self.skipWaiting());

// تفعيل وتنظيف ملفات الكاش القديمة
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// اعتراض الطلبات لجلب الملفات من الشبكة أولاً، وإن فشل (أوفلاين) يجلبها من الكاش
self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  
  // نتعامل مع ملفات التطبيق المحلية فقط؛ الخطوط والمكتبات الخارجية وروابط Google Sheets تمر كما هي
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;

  e.respondWith(
    fetch(req, { cache: 'no-store' })
      .then(res => {
        // تحديث الكاش بالنسخة الأحدث من الشبكة
        if (res && res.ok) { 
            const copy = res.clone(); 
            caches.open(CACHE).then(c => c.put(req, copy)); 
        }
        return res;
      })
      .catch(() => {
          // في حال انقطاع الإنترنت، جلب الملف من الكاش
          return caches.match(req).then(r => r || caches.match('./index.html'));
      })
  );
});
