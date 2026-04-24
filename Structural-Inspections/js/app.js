/**
 * BT Inspection Report — app entry
 * Bootstraps DB, router, service worker, and mounts the header / initial view.
 */

import { initDb } from './db.js';
import { initRouter } from './router.js';
import { mountHeader } from './components/header.js';
import { appState } from './state.js';

const BOOT_MIN_MS = 350; // Avoid boot flash for fast loads
const bootStart = performance.now();

async function boot() {
  // 1. IndexedDB (Dexie)
  try {
    await initDb();
  } catch (err) {
    console.error('Database init failed:', err);
    // Non-fatal for Phase 1; continue so UI still renders
  }

  // 2. Service worker (PWA / offline)
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    try {
      await navigator.serviceWorker.register('service-worker.js');
    } catch (err) {
      console.warn('Service worker registration failed:', err);
    }
  }

  // 3. Mount persistent header
  mountHeader(document.getElementById('app-header'));

  // 4. Router — handles view mounting based on hash
  initRouter(document.getElementById('view-outlet'), document.getElementById('app-nav'));

  // 5. Hide boot screen once everything is ready (but not too fast)
  const elapsed = performance.now() - bootStart;
  const delay = Math.max(0, BOOT_MIN_MS - elapsed);
  setTimeout(() => {
    const boot = document.getElementById('boot');
    if (boot) {
      boot.classList.add('is-hidden');
      setTimeout(() => boot.remove(), 400);
    }
  }, delay);

  // 6. Expose for debugging in dev
  if (location.hostname === 'localhost' || location.hostname === '127.0.0.1') {
    window.__bt = { state: appState };
  }
}

boot();
