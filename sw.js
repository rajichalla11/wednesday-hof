/* Wednesday 9–10 service worker: makes the site installable as an app. Pages always come from the network first (so updates
   show straight away) and fall back to the last copy when offline. Data calls (the sheet, the backend, /api) are left alone. */
const CACHE = 'w910-v3';
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(['/', '/icons/icon-192.png', '/icons/apple-touch-icon.png']).catch(() => {}))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.origin !== location.origin || u.pathname.startsWith('/api/')) return;
  if (r.mode === 'navigate') {
    // Open straight away from the saved copy, fetch the latest in the background, and tell the page if it changed
    // (it then offers "Tap to refresh"). No saved copy yet: wait for the network.
    e.respondWith(caches.open(CACHE).then(async c => {
      const hit = await c.match('/');
      const net = fetch(r).then(async res => {
        if (res.ok) { const fresh = res.clone(); const old = hit ? await hit.clone().text() : null; await c.put('/', res.clone());
          if (old !== null && old !== await fresh.text()) { const cs = await self.clients.matchAll({ type: 'window' }); cs.forEach(cl => cl.postMessage({ type: 'w910-update' })); } }
        return res;
      });
      if (hit) { e.waitUntil(net.catch(() => {})); return hit; }
      return net.catch(() => c.match('/'));
    }));
  } else if (u.pathname.startsWith('/icons/')) {
    e.respondWith(caches.match(r).then(hit => hit || fetch(r).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(r, copy)); return res; })));
  }
});

/* Push notifications (sent by the backend through /api/push): show them, and open the right page when tapped. */
self.addEventListener('push', e => {
  let d = {}; try { d = e.data ? e.data.json() : {}; } catch (_) { d = { title: 'Wednesday 9–10', body: e.data ? e.data.text() : '' }; }
  e.waitUntil(self.registration.showNotification(d.title || 'Wednesday 9–10', { body: d.body || '', icon: '/icons/icon-192.png', badge: '/icons/favicon-32.png', tag: d.tag || undefined, data: { url: d.url || '/' } }));
});
self.addEventListener('notificationclick', e => {
  e.notification.close(); const url = (e.notification.data && e.notification.data.url) || '/';
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(cs => { const c = cs.find(x => x.url.startsWith(self.location.origin)); if (c) { c.navigate(url).catch(() => {}); return c.focus(); } return self.clients.openWindow(url); }));
});
