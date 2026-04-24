/**
 * Accessible modal dialog.
 *
 * Usage:
 *
 *   const result = await openModal({
 *     title: 'New project',
 *     bodyHtml: '<div class="form-field">...</div>',
 *     primaryLabel: 'Create',
 *     onConfirm: async (dialogEl) => {
 *       const name = dialogEl.querySelector('[name=name]').value;
 *       if (!name) throw new Error('Name required');
 *       return { name };
 *     }
 *   });
 *
 * openModal returns the value returned by onConfirm, or null if the user
 * cancelled / closed the dialog.
 */

let activeDialog = null;

export function openModal(opts) {
  if (activeDialog) closeModal();

  return new Promise((resolve) => {
    const scrim = document.createElement('div');
    scrim.className = 'modal-scrim';
    scrim.setAttribute('role', 'presentation');

    const dialog = document.createElement('div');
    dialog.className = 'modal-dialog';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-labelledby', 'modal-title');

    dialog.innerHTML = `
      <header class="modal-dialog__header">
        <h2 id="modal-title">${escapeHtml(opts.title || '')}</h2>
        <button class="modal-dialog__close" aria-label="Close" type="button">&times;</button>
      </header>
      <div class="modal-dialog__body">${opts.bodyHtml || ''}</div>
      <footer class="modal-dialog__footer">
        <div class="modal-dialog__error" role="alert" aria-live="polite"></div>
        <div class="cluster cluster--end">
          <button class="btn btn--secondary" type="button" data-role="cancel">
            ${escapeHtml(opts.cancelLabel || 'Cancel')}
          </button>
          <button class="btn btn--primary" type="button" data-role="confirm">
            ${escapeHtml(opts.primaryLabel || 'Save')}
          </button>
        </div>
      </footer>
    `;

    scrim.appendChild(dialog);
    document.body.appendChild(scrim);
    activeDialog = scrim;

    const closeBtn   = dialog.querySelector('.modal-dialog__close');
    const cancelBtn  = dialog.querySelector('[data-role=cancel]');
    const confirmBtn = dialog.querySelector('[data-role=confirm]');
    const errorEl    = dialog.querySelector('.modal-dialog__error');

    // Remember the element that was focused when we opened so we can restore
    // it on close — important for keyboard + screen-reader users.
    const returnFocusTo = document.activeElement && document.activeElement !== document.body
      ? document.activeElement
      : null;

    function cleanup(result) {
      document.removeEventListener('keydown', onKeydown);
      scrim.classList.add('is-leaving');
      setTimeout(() => {
        scrim.remove();
        if (activeDialog === scrim) activeDialog = null;
        // Return focus so keyboard users don't end up at the top of the page.
        try { returnFocusTo?.focus?.({ preventScroll: true }); } catch {}
      }, 150);
      resolve(result);
    }

    // Focus trap — collect tabbable elements at Tab-time so dynamically
    // added content (e.g. photo cards in the item panel) is included.
    const TABBABLE = 'a[href], button:not([disabled]), input:not([disabled]):not([type=hidden]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';

    function trapTab(e) {
      if (e.key !== 'Tab') return;
      const focusables = Array.from(dialog.querySelectorAll(TABBABLE))
        .filter((el) => el.offsetParent !== null || el === document.activeElement);
      if (focusables.length === 0) return;
      const first = focusables[0];
      const last  = focusables[focusables.length - 1];
      if (e.shiftKey && document.activeElement === first) {
        e.preventDefault();
        last.focus();
      } else if (!e.shiftKey && document.activeElement === last) {
        e.preventDefault();
        first.focus();
      }
    }

    function onKeydown(e) {
      if (e.key === 'Escape') { cleanup(null); return; }
      trapTab(e);
    }

    async function onConfirm() {
      errorEl.textContent = '';
      confirmBtn.disabled = true;
      cancelBtn.disabled = true;
      try {
        const result = opts.onConfirm ? await opts.onConfirm(dialog) : true;
        cleanup(result);
      } catch (err) {
        console.warn('Modal confirm error:', err);
        errorEl.textContent = err.message || String(err);
      } finally {
        confirmBtn.disabled = false;
        cancelBtn.disabled = false;
      }
    }

    closeBtn.addEventListener('click', () => cleanup(null));
    cancelBtn.addEventListener('click', () => cleanup(null));
    confirmBtn.addEventListener('click', onConfirm);
    scrim.addEventListener('click', (e) => {
      if (e.target === scrim) cleanup(null);
    });
    document.addEventListener('keydown', onKeydown);

    // Enter-to-submit from any input (but not textarea)
    dialog.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && e.target.tagName === 'INPUT') {
        e.preventDefault();
        onConfirm();
      }
    });

    // Focus first input for quick entry
    setTimeout(() => {
      const firstInput = dialog.querySelector('input, textarea, select, button');
      if (firstInput) firstInput.focus();
    }, 50);

    // Optional: after-mount hook for complex bodies
    if (typeof opts.onMount === 'function') {
      opts.onMount(dialog);
    }
  });
}

export function closeModal() {
  if (activeDialog) {
    activeDialog.remove();
    activeDialog = null;
  }
}

/**
 * Confirmation dialog shortcut.
 */
export function confirmDialog({ title, message, confirmLabel = 'Confirm', danger = false }) {
  return openModal({
    title,
    bodyHtml: `<p>${escapeHtml(message || '')}</p>`,
    primaryLabel: confirmLabel,
    onConfirm: () => true,
    onMount: (dialog) => {
      if (danger) {
        const btn = dialog.querySelector('[data-role=confirm]');
        btn.classList.remove('btn--primary');
        btn.classList.add('btn--danger');
      }
    }
  });
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
