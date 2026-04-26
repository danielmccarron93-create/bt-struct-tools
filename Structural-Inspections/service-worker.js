/**
 * BT Inspection Report — service worker (Phase 8)
 *
 * Strategy:
 *   - App-shell files: pre-cached on install (SHELL_ASSETS).
 *   - Third-party assets (CDN: Dexie, PDF.js, jsPDF, Google Fonts): cached on
 *     first successful fetch so the app can boot offline on subsequent loads.
 *   - Navigation requests fall back to cached index.html when offline.
 */

const VERSION = 'bt-inspect-v2.1.0-phase11-multitype';
const SHELL_CACHE   = `shell-${VERSION}`;
const RUNTIME_CACHE = `runtime-${VERSION}`;

const SHELL_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './css/styles.css',
  // App entry + core
  './js/app.js',
  './js/router.js',
  './js/db.js',
  './js/state.js',
  // Views
  './js/views/home.js',
  './js/views/projects.js',
  './js/views/project-detail.js',
  './js/views/drawing-viewer.js',
  './js/views/inspections.js',
  './js/views/inspection-detail.js',
  './js/views/new-inspection.js',
  './js/views/settings.js',
  // Components
  './js/components/header.js',
  './js/components/modal.js',
  './js/components/toast.js',
  './js/components/item-panel.js',
  './js/components/comment-library-picker.js',
  './js/components/import-project.js',
  // Libraries (in-repo)
  './js/lib/pdf.js',
  './js/lib/btproject.js',
  './js/lib/photos.js',
  './js/lib/report.js',
  './js/lib/pdf-fonts.js',
  './js/lib/share.js',
  './js/lib/ai-expand.js',
  './js/lib/anthropic.js',
  './js/lib/annotate.js',
  './js/lib/qr.js',
  './js/lib/export-bundle.js',
  // Assets
  './assets/logo.svg',
  './assets/logo-white.svg',
  './assets/favicon.svg',
  './assets/icon-maskable.svg',
  // PDF fonts — pre-cached so reports embed Libre Franklin even offline
  './assets/fonts/LibreFranklin-Regular.ttf',
  './assets/fonts/LibreFranklin-Medium.ttf',
  './assets/fonts/LibreFranklin-Bold.ttf',
  './assets/fonts/LibreFranklin-Italic.ttf'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(SHELL_CACHE)
      .then((cache) => cache.addAll(SHELL_ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((k) => k !== SHELL_CACHE && k !== RUNTIME_CACHE)
          .map((k) => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);

  // Strategy 1: Navigations — network-first (to pick up new HTML) with
  // offline fallback to cached index.html.
  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request)
        .then((r) => r)
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  // Strategy 2: Same-origin assets — cache-first, runtime-cache on miss.
  if (url.origin === self.location.origin) {
    event.respondWith(cacheFirst(request, SHELL_CACHE));
    return;
  }

  // Strategy 3: Cross-origin whitelisted CDNs — cache-first in runtime cache.
  if (isCachableCdn(url)) {
    event.respondWith(cacheFirst(request, RUNTIME_CACHE));
    return;
  }

  // Everything else — go to network, no cache.
});

function isCachableCdn(url) {
  return (
    url.hostname === 'unpkg.com' ||
    url.hostname === 'cdnjs.cloudflare.com' ||
    url.hostname === 'fonts.googleapis.com' ||
    url.hostname === 'fonts.gstatic.com'
  );
}

async function cacheFirst(request, cacheName) {
  const cache = await caches.open(cacheName);
  const cached = await cache.match(request);
  if (cached) return cached;
  try {
    const response = await fetch(request);
    if (response && response.ok) {
      // Clone before caching — response bodies are one-shot streams.
      cache.put(request, response.clone());
    }
    return response;
  } catch (err) {
    // Absolute last resort: fall through and let the browser show its error.
    return Response.error();
  }
}
