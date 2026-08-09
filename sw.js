// DEWVEE Service Worker — network-first for API, cache-first for shell
const CACHE = 'dewvee-v17';
const SHELL = ['./index.html', './manifest.json', './dewvee-icon.svg', './dewvee-logo.svg'];

self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(c) { return c.addAll(SHELL); })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE; })
            .map(function(k)    { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(e) {
  var url = new URL(e.request.url);

  // Google Apps Script API — network-first, fall back to last cache
  if (url.hostname === 'script.google.com') {
    e.respondWith(
      fetch(e.request).then(function(resp) {
        if (resp.status === 200) {
          var clone = resp.clone();
          caches.open(CACHE).then(function(c) { c.put(e.request, clone); });
        }
        return resp;
      }).catch(function() {
        return caches.match(e.request).then(function(cached) {
          if (cached) return cached;
          // Return a synthetic offline response so the app can show a banner
          return new Response(JSON.stringify({offline: true}),
            { headers: { 'Content-Type': 'application/json' } });
        });
      })
    );
    return;
  }

  // Shell assets — cache-first, update in background
  if (url.origin === self.location.origin) {
    e.respondWith(
      caches.match(e.request).then(function(cached) {
        var network = fetch(e.request).then(function(resp) {
          if (resp.status === 200) {
            var clone = resp.clone();
            caches.open(CACHE).then(function(c) { c.put(e.request, clone); });
          }
          return resp;
        });
        return cached || network;
      })
    );
  }
});
