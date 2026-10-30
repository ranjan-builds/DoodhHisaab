const CACHE_NAME = 'doodh-hisaab-shell-v1';
const CACHE_PREFIX = 'doodh-hisaab-shell-';
const APP_ROOT = new URL('./', self.location);
const APP_SHELL = [
  new URL('index.html', APP_ROOT).href,
  new URL('favicons/site.webmanifest', APP_ROOT).href,
  new URL('favicons/favicon-96x96.png', APP_ROOT).href,
  new URL('favicons/favicon.svg', APP_ROOT).href,
  new URL('favicons/favicon.ico', APP_ROOT).href,
  new URL('favicons/apple-touch-icon.png', APP_ROOT).href,
  new URL('favicons/icon-192.png', APP_ROOT).href,
  new URL('favicons/icon-512.png', APP_ROOT).href
];
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(key => key.startsWith(CACHE_PREFIX) && key !== CACHE_NAME).map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const request = event.request;
  const url = new URL(request.url);

  if (request.method !== 'GET') return;

  if (url.origin !== self.location.origin) return;

  if (request.mode === 'navigate') {
    event.respondWith(
      (async () => {
        try {
          const response = await fetch(request);
          if (response.ok) {
            await (await caches.open(CACHE_NAME)).put(request, response.clone());
          }
          return response;
        } catch {
          const cachedPage = await caches.match(request);
          return cachedPage || caches.match(new URL('index.html', APP_ROOT).href);
        }
      })()
    );
    return;
  }

  if (APP_SHELL.includes(request.url)) {
    event.respondWith(
      (async () => {
        const cached = await caches.match(request);
        if (cached) return cached;

        const response = await fetch(request);
        if (response.ok) {
          await (await caches.open(CACHE_NAME)).put(request, response.clone());
        }
        return response;
      })()
    );
  }
});
