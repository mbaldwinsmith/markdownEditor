const CACHE_NAME = 'markdown-editor-cache-v3';
const OFFLINE_ASSETS = [
  './',
  './index.html',
  './styles.css',
  './app.js',
  './manifest.json',
  './privacy.html',
  './terms.html',
  './vendor/marked.min.js',
  './vendor/turndown.min.js',
  './vendor/turndown-plugin-gfm.min.js',
  './icons/icon-192.svg',
  './icons/icon-512.svg'
];

const OFFLINE_ASSET_URLS = new Set(
  OFFLINE_ASSETS.map((asset) => new URL(asset, self.location.origin).href)
);

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(OFFLINE_ASSETS))
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))
      )
    )
  );
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') {
    return;
  }

  const requestUrl = new URL(event.request.url);
  if (requestUrl.protocol !== 'http:' && requestUrl.protocol !== 'https:') {
    return;
  }

  if (requestUrl.origin !== self.location.origin) {
    return;
  }

  if (!OFFLINE_ASSET_URLS.has(requestUrl.href)) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse;
      }
      return fetch(event.request)
        .then((response) => {
          if (!response || response.status !== 200) {
            return response;
          }

          const responseClone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseClone));
          return response;
        })
        .catch(() => caches.match('./index.html'));
    })
  );
});
