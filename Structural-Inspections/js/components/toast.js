/**
 * Ephemeral toast notification (bottom of screen).
 * Non-blocking, auto-dismisses.
 */

let host = null;

function ensureHost() {
  if (host) return host;
  host = document.createElement('div');
  host.className = 'toast-host';
  host.setAttribute('aria-live', 'polite');
  host.setAttribute('aria-atomic', 'true');
  document.body.appendChild(host);
  return host;
}

export function toast(message, { kind = 'info', duration = 3200 } = {}) {
  const node = document.createElement('div');
  node.className = `toast toast--${kind}`;
  node.textContent = message;
  ensureHost().appendChild(node);
  // trigger enter animation next frame
  requestAnimationFrame(() => node.classList.add('is-visible'));
  setTimeout(() => {
    node.classList.remove('is-visible');
    setTimeout(() => node.remove(), 200);
  }, duration);
}
