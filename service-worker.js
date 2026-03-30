const CACHE_NAME = 'datamatrix-offline-v1';

// All the files needed for the Add-in to run completely standalone
const ASSETS_TO_CACHE = [
  '/',
  '/src/taskpane/taskpane.html',
  '/src/taskpane/taskpane.css',
  '/src/taskpane/taskpane.js',
  '/assets/bwip-js.min.js',
  '/assets/icon-16.png',
  '/assets/icon-32.png',
  '/assets/icon-80.png',
  // Office JS fallback cache (though Office JS usually caches itself)
  'https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        return cache.addAll(ASSETS_TO_CACHE);
      })
  );
  self.skipWaiting(); // Force the waiting service worker to become the active one
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.filter(cacheName => cacheName !== CACHE_NAME)
          .map(cacheName => caches.delete(cacheName))
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  // Try finding the request in the cache first, otherwise fetch from the network
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        return response || fetch(event.request);
      })
  );
});
