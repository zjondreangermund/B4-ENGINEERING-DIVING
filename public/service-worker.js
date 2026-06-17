// B-Four Ops service worker disabled. No cache and no HTML injection.
self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil((async()=>{
  const keys = await caches.keys();
  await Promise.all(keys.map(k => caches.delete(k)));
  if (self.registration && self.registration.unregister) await self.registration.unregister();
  const clients = await self.clients.matchAll({type:'window', includeUncontrolled:true});
  for (const c of clients) c.navigate(c.url);
})()));
