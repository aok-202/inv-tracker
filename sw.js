/* Offline support. Keeps a copy of this app's own files so it opens without a connection.
   It only ever fetches files from this same site, and it never sees your Excel file or numbers.
   spx-data.js is fetched fresh from this site when online (the daily GitHub job updates it),
   with the saved copy used offline. When you change any other app file, change VERSION. */
const VERSION = '2026-11-07.1';
const CACHE = 'investment-tracker-' + VERSION;
const FILES = ['./', 'index.html', 'styles.css', 'spx-data.js', 'calc.js', 'xlsx.js', 'app.js', 'manifest.webmanifest',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png', 'icons/apple-touch-icon.png'];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k.startsWith('investment-tracker-') && k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
const FRESH = ['spx-data.js'];
async function freshFirst(req) {
  const cache = await caches.open(CACHE);
  try {
    const stop = new AbortController();
    const timer = setTimeout(() => stop.abort(), 3000);
    const res = await fetch(req, { cache: 'no-cache', signal: stop.signal });
    clearTimeout(timer);
    if (res.ok) { await cache.put(req, res.clone()); return res; }
  } catch (e) { /* offline or slow: use the saved copy */ }
  return (await cache.match(req, { ignoreSearch: true })) || Response.error();
}
self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (FRESH.some(f => url.pathname.endsWith('/' + f))) { event.respondWith(freshFirst(req)); return; }
  event.respondWith(
    caches.match(req, { ignoreSearch: true }).then(hit => hit || fetch(req).then(res => {
      if (res.ok && res.type === 'basic') { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return res;
    }).catch(() => (req.mode === 'navigate' ? caches.match('index.html') : Response.error())))
  );
});
