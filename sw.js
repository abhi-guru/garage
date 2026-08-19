const CACHE = 'garage-v1';
const SHELL = [
  '/garage/',
  '/garage/index.html',
  '/garage/manifest.json',
  '/garage/icon-192.png',
  '/garage/icon-512.png'
];

// Install: cache the app shell
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(SHELL))
  );
  self.skipWaiting();
});

// Activate: clean up old caches
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch: shell from cache, API calls always from network
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Always go network-first for the Google Apps Script API
  if (url.hostname.includes('script.google.com') ||
      url.hostname.includes('generativelanguage.googleapis.com') ||
      url.hostname.includes('api.qrserver.com')) {
    return; // browser handles normally
  }

  // Cache-first for app shell assets
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request))
  );
});
