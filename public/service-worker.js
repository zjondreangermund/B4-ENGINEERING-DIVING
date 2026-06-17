// B-Four Ops service worker v14
self.addEventListener("install", e=>self.skipWaiting());
self.addEventListener("activate", e=>self.clients.claim());
