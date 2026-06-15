// Safe no-cache service worker for B-Four Ops.
// Version: 2026-06-15-render-fix-excel-rate-engine-fix
// It intentionally avoids caching financial/job-card data so role changes and live deployments show immediately.
self.addEventListener('install', event => {
  self.skipWaiting();
});
self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.map(key => caches.delete(key)));
    await self.clients.claim();
  })());
});
self.addEventListener('fetch', event => {
  event.respondWith(fetch(event.request, { cache: 'no-store' }));
});
