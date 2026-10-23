/* Wednesday 9–10 service worker: makes the site installable as an app. Pages always come from the network first (so updates
   show straight away) and fall back to the last copy when offline. Data calls (the sheet, the backend, /api) are left alone. */
const CACHE = 'w910-v1';
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(['/', '/icons/icon-192.png', '/icons/apple-touch-icon.png']).catch(() => {}))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.origin !== location.origin || u.pathname.startsWith('/api/')) return;
  if (r.mode === 'navigate') {
    e.respondWith(fetch(r).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put('/', copy)); return res; }).catch(() => caches.match('/')));
  } else if (u.pathname.startsWith('/icons/')) {
    e.respondWith(caches.match(r).then(hit => hit || fetch(r).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(r, copy)); return res; })));
  }
});
