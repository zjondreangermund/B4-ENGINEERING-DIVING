// Safe no-cache service worker for B-Four Ops.
// This avoids stale/corrupt cached responses and lets the live Railway app serve files normally.

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    if (self.caches && caches.keys) {
      const keys = await caches.keys();
      await Promise.all(keys.map(key => caches.delete(key)));
    }
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  event.respondWith(
    fetch(event.request).catch(() => new Response('Offline', {
      status: 503,
      headers: { 'Content-Type': 'text/plain' }
    }))
  );
});
