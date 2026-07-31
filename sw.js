// Service Worker - Lector CUD
// Cachea solo el "cascarón" de la app (HTML/CSS/JS/íconos) para que abra al instante.
// Las llamadas a Gemini o a tu backend SIEMPRE van a la red (nunca se cachean).

const CACHE_NAME = 'lector-cud-shell-v1';
const SHELL_FILES = [
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(()=>{})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = event.request.url;
  // Nunca cachear llamadas a APIs externas (Gemini, backend propio, fuentes de Google)
  if (url.includes('generativelanguage.googleapis.com') || url.includes('script.google.com') || url.includes('fonts.g')) {
    return; // dejar pasar directo a la red
  }
  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
