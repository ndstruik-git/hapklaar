/* Hapklaar service worker: the app opens instantly and offline. Data goes through the page itself. */
const VERSION = 'hapklaar-2026.10.05';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './apple-touch-icon.png', './favicon-32.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin === location.origin) {
    // Serve from cache straight away, refresh the cache in the background.
    const key = req.mode === 'navigate' ? './index.html' : req;
    e.respondWith(caches.open(VERSION).then(async c => {
      const hit = await c.match(key, { ignoreSearch: true });
      const net = fetch(req).then(r => { if (r.ok) c.put(key, r.clone()); return r; }).catch(() => hit);
      return hit || net;
    }));
    return;
  }
  if (/^(fonts\.googleapis\.com|fonts\.gstatic\.com|cdnjs\.cloudflare\.com)$/.test(url.hostname)) {
    e.respondWith(caches.open(VERSION + '-ext').then(async c => {
      const hit = await c.match(req);
      if (hit) return hit;
      const r = await fetch(req);
      if (r.ok || r.type === 'opaque') c.put(req, r.clone());
      return r;
    }));
  }
  // Everything else (the Apps Script API) goes straight to the network.
});
