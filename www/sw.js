// Service worker de la version web installable (PC, navigateur) : garde
// l'interface en cache pour qu'elle s'ouvre même sans réseau. Les appels au
// serveur Apps Script (autre domaine, POST) ne passent jamais par ce cache.
var CACHE = 'brulages-fri-v1';
var SHELL = [
  './',
  'index.html',
  'config.js',
  'manifest.json',
  'assets/logo-fri.png',
  'assets/sponsors.jpg',
  'assets/icons/icon-192.png',
  'assets/icons/icon-512.png',
];

self.addEventListener('install', function (event) {
  event.waitUntil(caches.open(CACHE).then(function (cache) { return cache.addAll(SHELL); }));
  self.skipWaiting();
});

self.addEventListener('activate', function (event) {
  event.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

// Réseau d'abord (pour toujours avoir la dernière version publiée), cache
// en secours si hors-ligne.
self.addEventListener('fetch', function (event) {
  var req = event.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  event.respondWith(
    fetch(req).then(function (resp) {
      if (resp && resp.ok) {
        var copy = resp.clone();
        caches.open(CACHE).then(function (cache) { cache.put(req, copy); });
      }
      return resp;
    }).catch(function () {
      return caches.match(req).then(function (hit) { return hit || caches.match('index.html'); });
    })
  );
});
