// Safe no-cache service worker for B-Four Ops.
// It intentionally avoids caching financial/job-card data so role changes and live deployments show immediately.
self.addEventListener('install', event => self.skipWaiting());
self.addEventListener('activate', event => event.waitUntil(self.clients.claim()));
self.addEventListener('fetch', event => {
  event.respondWith(fetch(event.request));
});
