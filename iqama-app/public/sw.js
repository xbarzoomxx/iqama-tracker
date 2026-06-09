const CACHE_NAME = "iqama-app-v1";
const STATIC_ASSETS = ["/", "/index.html"];

// تثبيت: تخزين الملفات الأساسية
self.addEventListener("install", (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS))
  );
  self.skipWaiting();
});

// تفعيل: حذف الكاش القديم
self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// جلب: شبكة أولاً ثم كاش (مناسب للتطبيقات الديناميكية)
self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  // Firebase و APIs: شبكة فقط
  if (e.request.url.includes("firebase") || e.request.url.includes("googleapis")) return;

  e.respondWith(
    fetch(e.request)
      .then((res) => {
        const clone = res.clone();
        caches.open(CACHE_NAME).then((c) => c.put(e.request, clone));
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
