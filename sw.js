const CACHE = 'dewvee-v1';
const SHELL = ['./index.html', './Dashboard.html', './dewvee-icon.svg', './manifest.json'];

self.addEventListener('install', function(e){
  e.waitUntil(
    caches.open(CACHE).then(function(c){ return c.addAll(SHELL); })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(e){
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(k){ return k!==CACHE; }).map(function(k){ return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(e){
  // only intercept same-origin requests
  if(!e.request.url.startsWith(self.location.origin)) return;
  e.respondWith(
    caches.match(e.request).then(function(cached){
      var network = fetch(e.request).then(function(res){
        if(res.status === 200){
          var clone = res.clone();
          caches.open(CACHE).then(function(c){ c.put(e.request, clone); });
        }
        return res;
      });
      return cached || network;
    })
  );
});
