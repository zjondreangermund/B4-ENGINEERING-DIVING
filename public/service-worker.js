// Service worker disabled for stable presentations.
// No caching, no injection, no fetch interception.
self.addEventListener('install', event => self.skipWaiting());
self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.map(key => caches.delete(key)));
    if (self.registration && self.registration.unregister) await self.registration.unregister();
  })());
});
