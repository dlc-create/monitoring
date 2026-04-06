

const CACHE_NAME = 'pcc-cache-v3';
const CACHE_URLS  = [
  './',
  './index.html',
  './pcc-app.js',
  './pcc-charts.js',
  './logo.png',
  'https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js',
];

/* ── INSTALL: cache the app shell ── */
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(CACHE_URLS).catch(err => {
        console.warn('[SW] Some assets failed to cache:', err);
      });
    })
  );
  self.skipWaiting();
});

/* ── ACTIVATE: clean up old caches ── */
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

/* ── FETCH: network-first for API, cache-first for assets ── */
self.addEventListener('fetch', event => {
  const url = event.request.url;

  /* Always go network-first for Google APIs */
  if (
    url.includes('googleapis.com') ||
    url.includes('accounts.google.com') ||
    url.includes('oauth2.googleapis.com')
  ) {
    event.respondWith(
      fetch(event.request).catch(() => {
        /* If offline and it's an API call, return a friendly offline JSON */
        return new Response(
          JSON.stringify({ error: { message: 'You are offline. Please reconnect to sync data.' } }),
          { headers: { 'Content-Type': 'application/json' } }
        );
      })
    );
    return;
  }

  /* Cache-first for everything else (app shell, fonts, Chart.js) */
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        /* Only cache valid responses */
        if (!response || response.status !== 200 || response.type === 'opaque') {
          return response;
        }
        const clone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        return response;
      }).catch(() => {
        /* Offline fallback for navigation requests */
        if (event.request.mode === 'navigate') {
          return caches.match('./index.html');
        }
      });
    })
  );
});
