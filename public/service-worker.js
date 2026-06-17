// B-Four Ops service worker: no caching, only loads stable filter helper.
const CACHE_NAME='b4-no-cache-filter-helper-v22';

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

const HELPER_TAG = '<scr' + 'ipt src="/filter-fix.js?v=22"></scr' + 'ipt>';

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);
  if (req.mode === 'navigate' || url.pathname === '/' || url.pathname.endsWith('/index.html')) {
    event.respondWith((async () => {
      const res = await fetch(req, { cache: 'no-store' });
      const type = res.headers.get('content-type') || '';
      if (!type.includes('text/html')) return res;
      let html = await res.text();
      if (!html.includes('filter-fix.js')) html = html.replace('</body>', HELPER_TAG + '\n</body>');
      return new Response(html, {
        headers: {
          'content-type': 'text/html; charset=utf-8',
          'cache-control': 'no-store'
        }
      });
    })());
    return;
  }
  event.respondWith(fetch(req, { cache: 'no-store' }));
});
