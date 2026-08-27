const CACHE = 'clinica-eb-v2';
const ASSETS = ['./index.html', './manifest.json', './logo.png', './config.js'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

/* Network-first: si hay conexión, siempre trae la versión más nueva del
   servidor (y la deja guardada). Si no hay conexión, usa la última copia
   guardada en el dispositivo. Así cada actualización se ve al instante
   sin tener que borrar caché a mano. */
self.addEventListener('fetch', e => {
  e.respondWith(
    fetch(e.request).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(e.request, copy)).catch(()=>{});
      return res;
    }).catch(() => caches.match(e.request))
  );
});
