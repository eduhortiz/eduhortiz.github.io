// Guarda solo los archivos de la app para que abra rápido y sin internet.
// Nunca guarda respuestas del Apps Script (datos de personal, stock, usuarios).
const CACHE = 'epp-v2';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'js/dc-runtime.js', 'js/react.production.min.js', 'js/react-dom.production.min.js', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return; // API y terceros: siempre directo a la red
  // Red primero (para recibir actualizaciones); si no hay internet, usa la copia guardada
  e.respondWith(fetch(e.request).then(r => {
    if (r.ok) { const cp = r.clone(); caches.open(CACHE).then(c => c.put(e.request, cp)); }
    return r;
  }).catch(() => caches.match(e.request).then(r => r || caches.match('index.html'))));
});
