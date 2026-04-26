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

/**
 * Show a toast.
 *
 * `duration: 0` means "stay visible until .dismiss() is called" — useful for
 * long-running ops that want to swap themselves for a final success/error toast
 * without piling up.
 *
 * Returns a handle with .dismiss() so callers can clear long-lived toasts.
 */
export function toast(message, { kind = 'info', duration = 3200 } = {}) {
  const node = document.createElement('div');
  node.className = `toast toast--${kind}`;
  node.textContent = message;
  ensureHost().appendChild(node);
  // trigger enter animation next frame
  requestAnimationFrame(() => node.classList.add('is-visible'));

  let dismissed = false;
  const dismiss = () => {
    if (dismissed) return;
    dismissed = true;
    node.classList.remove('is-visible');
    setTimeout(() => node.remove(), 200);
  };

  let timer = null;
  if (duration > 0) timer = setTimeout(dismiss, duration);

  return {
    dismiss: () => { if (timer) clearTimeout(timer); dismiss(); }
  };
}
