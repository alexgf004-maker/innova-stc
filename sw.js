const CACHE = 'innova-stc-v2';
const PRECACHE = [
  '/innova-stc/',
  '/innova-stc/index.html',
  '/innova-stc/login.html',
  '/innova-stc/css/base.css',
  '/innova-stc/css/layout.css',
  '/innova-stc/css/components.css',
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(PRECACHE)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  // Network first for Firebase and API calls
  if (e.request.url.includes('firebase') || e.request.url.includes('googleapis') || e.request.url.includes('gstatic')) {
    return;
  }
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
