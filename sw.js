const CACHE_NAME = 'cleopatra-church-v2';

self.addEventListener('install', (e) => {
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  // Pass-through to network for instant, fresh data without intercepting hangs
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
