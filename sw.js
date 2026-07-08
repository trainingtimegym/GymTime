const CACHE_NAME = 'training-time-v1';
const ARCHIVOS_CACHE = [
  './',
  './index.html',
  './config.js',
  './logo.png',
  './manifest.json'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ARCHIVOS_CACHE))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((nombres) =>
      Promise.all(
        nombres.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n))
      )
    )
  );
  self.clients.claim();
});

// Estrategia: red primero (para tener datos frescos), y si falla, cache como respaldo.
// Los pedidos a la API de Google Apps Script NUNCA se cachean, siempre van a la red.
self.addEventListener('fetch', (event) => {
  const url = event.request.url;
  if (url.includes('script.google.com')) {
    return; // dejamos pasar directo, sin intervenir
  }
  event.respondWith(
    fetch(event.request)
      .then((res) => {
        const resClonada = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, resClonada));
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});
