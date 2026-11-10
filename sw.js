/* Offline support. Keeps a copy of this app's own files so it opens without a connection.
   It only ever fetches files from this same site, and it never sees your Excel file or numbers.
   spx-data.js is fetched fresh from this site when online (the daily GitHub job updates it),
   with the saved copy used offline. When you change any other app file, change VERSION. */
const VERSION = '2026-11-10.1';
const CACHE = 'investment-tracker-' + VERSION;
/* The app can't run without these. */
const FILES = ['./', 'index.html', 'styles.css', 'spx-data.js', 'calc.js', 'xlsx.js', 'app.js'];
/* Nice to have: a missing icon must never stop the app from working offline or updating. */
const EXTRAS = ['manifest.webmanifest', 'apple-touch-icon.png', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png'];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache =>
    cache.addAll(FILES).then(() => Promise.all(EXTRAS.map(f => cache.add(f).catch(() => null))))
  ).then(() => self.skipWaiting()));
});
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k.startsWith('investment-tracker-') && k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
const FRESH = ['spx-data.js'];
/* The S&P prices: fresh from this site when it answers within 3 seconds, the saved copy otherwise.
   A slow answer still finishes in the background and is saved, so the next look has it. */
function saveFresh(req) {
  return fetch(req, { cache: 'no-cache' }).then(async res => {
    if (res.ok) await (await caches.open(CACHE)).put(req, res.clone());
    return res;
  });
}
async function freshFirst(req, net) {
  const quick = await Promise.race([net.catch(() => null), new Promise(done => setTimeout(() => done(null), 3000))]);
  if (quick && quick.ok) return quick;
  return (await caches.match(req, { ignoreSearch: true })) || net.catch(() => Response.error());
}
/* The page can't download anything itself (its security rules forbid it), so it asks here for
   the latest S&P prices: the same file from this same site, which also refreshes the saved copy. */
self.addEventListener('message', event => {
  const msg = event.data || {}, port = event.ports && event.ports[0];
  if (msg.type !== 'spx-refresh' || !port) return;
  const url = new URL('spx-data.js', self.registration.scope).href;
  event.waitUntil(saveFresh(new Request(url))
    .then(async res => { if (!res.ok) throw new Error('HTTP ' + res.status); port.postMessage({ ok: true, text: await res.text() }); })
    .catch(err => port.postMessage({ ok: false, error: String((err && err.message) || err) })));
});
self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (FRESH.some(f => url.pathname.endsWith('/' + f))) {
    const net = saveFresh(req);
    event.waitUntil(net.catch(() => null));
    event.respondWith(freshFirst(req, net));
    return;
  }
  event.respondWith(
    caches.match(req, { ignoreSearch: true }).then(hit => hit || fetch(req).then(res => {
      if (res.ok && res.type === 'basic') { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return res;
    }).catch(() => (req.mode === 'navigate' ? caches.match('index.html') : Response.error())))
  );
});
