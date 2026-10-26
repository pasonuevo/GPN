// Guarda la app de gasoil en el teléfono para que abra sin internet.
// Al cambiar la app, cambiar VERSION para que el teléfono baje la copia nueva.
const VERSION = 'gasoil-2026-10-26-1';
const FILES = ['./', './index.html', './manifest.json', './apple-touch-icon.png', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k.startsWith('gasoil-') && k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

// La app primero de la red (para recibir cambios) y si no hay señal, la copia guardada.
// Lo que va al servidor de la hoja no pasa por aquí.
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  e.respondWith(
    Promise.race([fetch(e.request), new Promise((_, no) => setTimeout(no, 5000))]).then(r => {
      if (r.ok) { const copia = r.clone(); caches.open(VERSION).then(c => c.put(e.request, copia)); }
      return r;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }).then(hit => hit || caches.match('./index.html')))
  );
});
