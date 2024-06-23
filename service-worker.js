// Versionsnummer für den Cache
const CACHE_NAME = 'graph-visibility-controller-v1';

// Dateien, die gecacht werden sollen
const FILES_TO_CACHE = [
  '/',
  '/home.html',
  '/main.js',
  'https://appsforoffice.microsoft.com/lib/1/hosted/office.js'
];

// Installationsereignis des Service Workers
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(FILES_TO_CACHE);
    })
  );
});

// Abrufereignis des Service Workers
self.addEventListener('fetch', (event) => {
  event.respondWith(
    caches.match(event.request).then((response) => {
      // Gebe die gecachte Version zurück oder hole die Ressource aus dem Netzwerk
      return response || fetch(event.request);
    })
  );
});
