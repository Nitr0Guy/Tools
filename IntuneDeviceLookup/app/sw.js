/* ═══════════════════════════════════════════════════════
   Service Worker — cache-first for static, network-first for API
   ═══════════════════════════════════════════════════════ */

const CACHE_NAME = 'intune-lookup-v3';
const STATIC_ASSETS = [
    '/',
    '/index.html',
    '/css/style.css',
    '/js/app.js',
    '/js/auth.js',
    '/js/graph.js',
    '/js/intune.js',
    '/js/scanner.js',
    '/js/msal-browser.min.js',
    '/manifest.json',
    '/icons/icon-192.png',
    '/icons/icon-512.png',
];

// Install — cache static assets
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS))
    );
    self.skipWaiting();
});

// Activate — clean old caches
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
        )
    );
    self.clients.claim();
});

// Fetch — network-first for Graph API, cache-first for static
self.addEventListener('fetch', (event) => {
    const url = new URL(event.request.url);

    // Graph API and auth calls: always network
    if (url.hostname === 'graph.microsoft.com' ||
        url.hostname === 'login.microsoftonline.com' ||
        url.hostname === 'cdn.jsdelivr.net' ||
        url.hostname.endsWith('.cdn.jsdelivr.net') ||
        url.hostname.endsWith('.msauth.net')) {
        event.respondWith(fetch(event.request));
        return;
    }

    // Static assets: cache-first
    event.respondWith(
        caches.match(event.request).then(cached => {
            return cached || fetch(event.request).then(response => {
                // Cache successful GET responses
                if (event.request.method === 'GET' && response.status === 200) {
                    const clone = response.clone();
                    caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
                }
                return response;
            });
        })
    );
});
