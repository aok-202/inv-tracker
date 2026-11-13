/* Offline support. Keeps a copy of this app's own files so it opens without a connection.
   It only ever fetches files from this same site, and it never sees your Excel file or numbers.
   spx-data.js and bench-data.js are fetched fresh from this site when online (the daily job
   updates them), with the saved copies used offline. When you change any other app file, change VERSION. */
const VERSION = '2026-11-13.1';
const CACHE = 'investment-tracker-' + VERSION;
/* The app can't run without these. */
const FILES = ['./', 'index.html', 'styles.css', 'spx-data.js', 'bench-data.js', 'calc.js', 'xlsx.js', 'app.js'];
/* Nice to have: a missing icon must never stop the app from working offline or updating. */
const EXTRAS = ['manifest.webmanifest', 'apple-touch-icon.png', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png'];

/* These carry the version they were built as. All of them must match before a new version takes over. */
const STAMPED = ['index.html', 'styles.css', 'calc.js', 'app.js'];
const fresh = f => new Request(f, { cache: 'no-cache' });
self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    try {
      await cache.addAll(FILES.map(fresh));
      /* If the site is half-way through an update (some files new, some old), stop here. The version
         already on this device keeps working, and this is tried again the next time the app opens. */
      for (const f of STAMPED) {
        const res = await cache.match(f);
        if (!res || !(await res.text()).includes('build ' + VERSION)) throw new Error('waiting for the rest of version ' + VERSION + ': ' + f);
      }
    } catch (err) {
      await caches.delete(CACHE);
      throw err;
    }
    await Promise.all(EXTRAS.map(f => cache.add(fresh(f)).catch(() => null)));
    await self.skipWaiting();
  })());
});
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k.startsWith('investment-tracker-') && k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
const FRESH = ['spx-data.js', 'bench-data.js'];
/* The S&P prices: fresh from this site when it answers within a second and a half, the saved copy otherwise.
   A slow answer still finishes in the background and is saved, so the next look has it. */
function saveFresh(req) {
  return fetch(req, { cache: 'no-cache' }).then(async res => {
    if (res.ok) await (await caches.open(CACHE)).put(req, res.clone());
    return res;
  });
}
async function freshFirst(req, net) {
  const quick = await Promise.race([net.catch(() => null), new Promise(done => setTimeout(() => done(null), 1500))]);
  if (quick && quick.ok) return quick;
  return (await caches.match(req, { ignoreSearch: true })) || net.catch(() => Response.error());
}
/* The page can't download anything itself (its security rules forbid it), so it asks here for the
   latest prices: the two data files from this same site, which also refreshes the saved copies. */
self.addEventListener('message', event => {
  const msg = event.data || {}, port = event.ports && event.ports[0];
  if ((msg.type !== 'data-refresh' && msg.type !== 'spx-refresh') || !port) return;
  const get = async name => {
    const res = await saveFresh(new Request(new URL(name, self.registration.scope).href));
    if (!res.ok) throw new Error('HTTP ' + res.status);
    return res.text();
  };
  event.waitUntil(get('spx-data.js')
    .then(async text => port.postMessage({ ok: true, text, bench: msg.type === 'data-refresh' ? await get('bench-data.js').catch(() => null) : null }))
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
