// Minimal service worker — exists only to satisfy Chrome/Android's PWA
// installability requirement (a fetch handler + manifest). It intentionally
// does NOT cache or intercept anything: this is a live ERP, and stale cached
// pages/data would be worse than no offline support at all. Every request
// just passes straight through to the network.

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', (event) => {
  event.respondWith(fetch(event.request));
});
