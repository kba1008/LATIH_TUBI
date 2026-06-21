const CACHE_NAME = 'latihai-v6-ms-tts';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  'https://cdn.tailwindcss.com',
  'https://unpkg.com/lucide@latest'
];

// Pemasangan (Install) Service Worker
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('Cache dibuka untuk Latih AI (Tema Baby Blue V4)');
        return cache.addAll(ASSETS);
      })
  );
});

// Menangkap permintaan dan kembalikan cache jika ada (Fetch)
self.addEventListener('fetch', event => {
  // Abaikan request ke API Groq dan Google Apps Script untuk memastikan data sentiasa terkini
  if (event.request.url.includes('api.groq.com') || event.request.url.includes('script.google.com')) {
    return;
  }
  
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        if (response) {
          return response; // Pulangkan dari cache
        }
        return fetch(event.request); // Ambil dari internet jika tiada dalam cache
      })
  );
});

// Pembersihan cache lama (Activate)
self.addEventListener('activate', event => {
  const cacheWhitelist = [CACHE_NAME];
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheWhitelist.indexOf(cacheName) === -1) {
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
});
