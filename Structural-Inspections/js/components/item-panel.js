/**
 * Item edit panel (Phase 4).
 *
 * Opens a modal for a single inspection item (pin):
 *   - Number badge (read-only, auto-assigned)
 *   - Severity   — observation | defect
 *   - Status     — open | closed   (for rectification tracking)
 *   - Comment    — free text (library-driven comments land in Phase 5)
 *   - Photos     — grid of thumbnails with captions + delete
 *                  "Take photo" (camera) and "Choose photo" (library) add flow
 *   - Delete item — destroys item + its photos
 *
 * Resolves to:
 *   { item }           — item was saved (returns the updated record)
 *   { deleted: true }  — item was deleted
 *   null               — user cancelled
 *
 * Photo semantics: photos are added / removed immediately (not staged) because
 * Blobs are heavy and stage-on-confirm is awkward. Caption edits are saved
 * when the user blurs the caption input. Comment/severity/status edits are
 * staged and saved together on confirm.
 */

import {
  getItem, updateItem, deleteItem,
  addItemPhoto, updateItemPhoto, deleteItemPhoto, listPhotosForItem,
  getInspection, getDrawing, getUserProfile
} from '../db.js';
import { openModal } from '../components/modal.js';
import { toast } from '../components/toast.js';
import {
  openPhotoCapture, openPhotoLibrary, compressImage, createUrlRegistry
} from '../lib/photos.js';
import { openCommentLibraryPicker } from './comment-library-picker.js';
import { openAnnotateModal } from '../lib/annotate.js';
import { expandComment } from '../lib/ai-expand.js';

export async function openItemPanel({ itemId }) {
  const item = await getItem(itemId);
  if (!item) return null;

  // Need the inspection's type key to filter the library picker.
  const inspection = await getInspection(item.inspectionId);
  const inspectionTypeKey = inspection ? inspection.inspectionTypeKey : null;

  // Tracked across the modal lifetime — updated when a library entry is picked.
  const libraryState = {
    libraryKey: item.libraryKey || null,
    asClause:   item.asClause   || ''
  };

  const result = { deleted: false, saved: null };

  const bodyHtml = `
    <div class="item-panel">
      <div class="item-panel__header">
        <span class="item-panel__num">Item ${item.itemNumber}</span>
        <div class="item-panel__status-toggle">
          <label>
            <input type="radio" name="status" value="open" ${item.status === 'open' ? 'checked' : ''}/>
            <span>Open</span>
          </label>
          <label>
            <input type="radio" name="status" value="closed" ${item.status === 'closed' ? 'checked' : ''}/>
            <span>Closed</span>
          </label>
        </div>
      </div>

      <div class="form-field">
        <label>Severity</label>
        <div class="radio-row">
          <label class="radio-chip radio-chip--observation">
            <input type="radio" name="severity" value="observation" ${item.severity === 'observation' ? 'checked' : ''}/>
            <span>Observation</span>
          </label>
          <label class="radio-chip radio-chip--defect">
            <input type="radio" name="severity" value="defect" ${item.severity === 'defect' ? 'checked' : ''}/>
            <span>Defect</span>
          </label>
          <label class="radio-chip radio-chip--holdpoint">
            <input type="radio" name="severity" value="holdpoint" ${item.severity === 'holdpoint' ? 'checked' : ''}/>
            <span>Hold point</span>
          </label>
        </div>
        <div class="form-field__help muted">
          Hold point = <strong>do not proceed</strong> until rectified. It prints a red banner on the report cover.
        </div>
      </div>

      <div class="form-field">
        <label for="ip-grid">Location / grid reference</label>
        <input id="ip-grid" name="gridRef" type="text"
          value="${escapeAttr(item.gridRef || '')}"
          placeholder="e.g. 3/B, GL-1, North edge of pour"
          autocomplete="off"/>
        <div class="form-field__help muted">
          Optional — printed in the item schedule so the builder can find the defect on site.
        </div>
      </div>

      <div class="form-field">
        <div class="form-field__label-row">
          <label for="ip-comment">Comment</label>
          <div class="cluster" style="gap:6px;">
            <button type="button" class="btn btn--ghost btn--sm" id="ip-btn-expand" title="Expand short-form note with Claude">
              <span aria-hidden="true">\u2728</span> Expand
            </button>
            <button type="button" class="btn btn--ghost btn--sm" id="ip-btn-library">
              <span aria-hidden="true">\ud83d\udcda</span> Use library
            </button>
          </div>
        </div>
        <textarea id="ip-comment" name="comment" rows="4"
          placeholder="Describe what was observed and any action required. Tap \u2728 Expand to polish a short note.">${escapeHtml(item.comment || '')}</textarea>
        <div class="form-field__help muted" id="ip-dictation-hint">
          <span class="dictation-hint__icon" aria-hidden="true">\ud83c\udf99</span>
          <span class="dictation-hint__ios">Tap the microphone on your keyboard to dictate.</span>
          <button type="button" class="dictation-hint__web btn btn--ghost btn--sm is-hidden" id="ip-dictate-btn">Start dictating</button>
        </div>
        <div class="form-field__help">
          <span class="library-clause-badge ${item.asClause ? '' : 'is-hidden'}" id="ip-clause-badge">
            <span class="muted">AS clause:</span>
            <strong id="ip-clause-text">${escapeHtml(item.asClause || '')}</strong>
            <button type="button" class="library-clause-badge__clear" id="ip-clause-clear" aria-label="Clear AS clause">&times;</button>
          </span>
        </div>
      </div>

      <div class="form-field">
        <label>Photos</label>
        <div class="photo-grid" id="ip-photo-grid" role="list">
          <div class="photo-grid__empty muted" id="ip-photo-empty">No photos yet.</div>
        </div>
        <div class="cluster" style="margin-top: var(--space-2);">
          <button type="button" class="btn btn--secondary btn--sm" id="ip-btn-camera">
            <span aria-hidden="true">📷</span> Take photo
          </button>
          <button type="button" class="btn btn--secondary btn--sm" id="ip-btn-photo-library">
            <span aria-hidden="true">🖼</span> Choose photo
          </button>
          <span class="muted" id="ip-photo-progress" style="font-size: 13px;"></span>
        </div>
      </div>

      <div class="item-panel__footer-actions">
        <button type="button" class="btn btn--ghost btn--sm" id="ip-btn-delete"
                style="color: var(--color-danger);">
          Delete item
        </button>
      </div>

      <div class="item-panel__confirm-delete is-hidden" id="ip-confirm-delete" role="alert">
        <p><strong>Delete item ${item.itemNumber}?</strong> This removes the pin and its photos. Item numbers are not re-issued.</p>
        <div class="cluster cluster--end">
          <button type="button" class="btn btn--secondary btn--sm" id="ip-confirm-cancel">Keep it</button>
          <button type="button" class="btn btn--danger btn--sm" id="ip-confirm-yes">Delete</button>
        </div>
      </div>
    </div>
  `;

  // We'll manage a registry so all object URLs get revoked when the modal closes.
  const urlRegistry = createUrlRegistry();

  await openModal({
    title: `Item ${item.itemNumber}`,
    primaryLabel: 'Save',
    bodyHtml,
    onMount: (dialog) => {
      dialog.classList.add('modal-dialog--wide');

      const grid        = dialog.querySelector('#ip-photo-grid');
      const empty       = dialog.querySelector('#ip-photo-empty');
      const progressEl  = dialog.querySelector('#ip-photo-progress');
      const btnCam      = dialog.querySelector('#ip-btn-camera');
      const btnLib      = dialog.querySelector('#ip-btn-photo-library');
      const btnDelete   = dialog.querySelector('#ip-btn-delete');
      const confirmBox  = dialog.querySelector('#ip-confirm-delete');
      const confirmYes  = dialog.querySelector('#ip-confirm-yes');
      const confirmNo   = dialog.querySelector('#ip-confirm-cancel');

      // --- Render photo grid from DB ---
      async function refreshPhotos() {
        urlRegistry.revokeAll();
        const photos = await listPhotosForItem(itemId);

        // Preserve the empty sentinel by clearing only photo cards
        grid.querySelectorAll('.photo-card').forEach((el) => el.remove());

        if (photos.length === 0) {
          empty.classList.remove('is-hidden');
          return;
        }
        empty.classList.add('is-hidden');

        for (const p of photos) {
          const url = urlRegistry.urlFor(p.blob);
          const card = document.createElement('div');
          card.className = 'photo-card';
          card.setAttribute('role', 'listitem');
          card.dataset.photoId = p.id;
          const annoCount = (p.annotations || []).length;
          const annoLabel = annoCount
            ? `\u270f ${annoCount} annotation${annoCount === 1 ? '' : 's'}`
            : '\u270f Annotate';
          card.innerHTML = `
            <div class="photo-card__thumb">
              <img alt="${escapeAttr(p.caption || 'Photo ' + p.id)}" src="${url}"/>
              <button type="button" class="photo-card__delete" aria-label="Delete photo"
                      data-photo-delete="${p.id}">&times;</button>
            </div>
            <div class="photo-card__row">
              <button type="button" class="btn btn--ghost btn--sm photo-card__annotate"
                      data-photo-annotate="${p.id}">${annoLabel}</button>
            </div>
            <input type="text" class="photo-card__caption" value="${escapeAttr(p.caption || '')}"
                   placeholder="Caption (optional)" data-photo-caption="${p.id}"/>
          `;
          grid.appendChild(card);
        }
      }

      // --- Delegation: delete button, annotate button, caption blur ---
      grid.addEventListener('click', async (e) => {
        const delBtn = e.target.closest('[data-photo-delete]');
        if (delBtn) {
          const pid = Number(delBtn.dataset.photoDelete);
          const card = delBtn.closest('.photo-card');
          card.classList.add('photo-card--deleting');
          delBtn.textContent = '…';
          try {
            await deleteItemPhoto(pid);
            await refreshPhotos();
          } catch (err) {
            toast('Couldn\u2019t delete photo', { kind: 'error' });
            console.warn(err);
            card.classList.remove('photo-card--deleting');
            delBtn.textContent = '\u00d7';
          }
          return;
        }
        const annoBtn = e.target.closest('[data-photo-annotate]');
        if (annoBtn) {
          const pid = Number(annoBtn.dataset.photoAnnotate);
          const photos = await listPhotosForItem(itemId);
          const photo = photos.find((p) => p.id === pid);
          if (!photo) return;
          const result = await openAnnotateModal({
            blob: photo.blob,
            annotations: photo.annotations || []
          });
          if (!result) return;
          await updateItemPhoto(pid, { annotations: result.annotations });
          await refreshPhotos();
          toast('Annotations saved', { kind: 'success' });
        }
      });

      grid.addEventListener('blur', async (e) => {
        const input = e.target.closest('[data-photo-caption]');
        if (!input) return;
        const pid = Number(input.dataset.photoCaption);
        try {
          await updateItemPhoto(pid, { caption: input.value });
        } catch (err) {
          console.warn('Caption save failed:', err);
        }
      }, true);

      // --- Add photo flow (camera / library) ---
      async function addPhoto(openFn) {
        progressEl.textContent = 'Opening camera…';
        const file = await openFn();
        if (!file) {
          progressEl.textContent = '';
          return;
        }
        progressEl.textContent = 'Compressing…';
        try {
          const { blob, width, height } = await compressImage(file);
          progressEl.textContent = 'Saving…';
          await addItemPhoto(itemId, { blob, width, height, caption: '' });
          await refreshPhotos();
          progressEl.textContent = '';
        } catch (err) {
          console.error('Photo add failed:', err);
          toast(`Couldn\u2019t add photo: ${err.message}`, { kind: 'error' });
          progressEl.textContent = '';
        }
      }
      btnCam.addEventListener('click', () => addPhoto(openPhotoCapture));
      btnLib.addEventListener('click', () => addPhoto(openPhotoLibrary));

      // --- Inline delete-item confirm ---
      btnDelete.addEventListener('click', () => {
        confirmBox.classList.remove('is-hidden');
        btnDelete.classList.add('is-hidden');
      });
      confirmNo.addEventListener('click', () => {
        confirmBox.classList.add('is-hidden');
        btnDelete.classList.remove('is-hidden');
      });
      confirmYes.addEventListener('click', async () => {
        confirmYes.disabled = true;
        confirmNo.disabled = true;
        try {
          await deleteItem(itemId);
          result.deleted = true;
          // Click the header close to dismiss modal (resolves openModal as null,
          // then the outer caller checks `result.deleted`).
          dialog.querySelector('.modal-dialog__close').click();
        } catch (err) {
          console.error(err);
          toast('Couldn\u2019t delete item', { kind: 'error' });
          confirmYes.disabled = false;
          confirmNo.disabled = false;
        }
      });

      // --- Library picker wiring ---
      const btnLibrary   = dialog.querySelector('#ip-btn-library');
      const clauseBadge  = dialog.querySelector('#ip-clause-badge');
      const clauseText   = dialog.querySelector('#ip-clause-text');
      const clauseClear  = dialog.querySelector('#ip-clause-clear');
      const commentEl    = dialog.querySelector('[name=comment]');

      btnLibrary.addEventListener('click', async () => {
        if (!inspectionTypeKey) {
          toast('No inspection type on this item', { kind: 'error' });
          return;
        }
        const currentSeverity = dialog.querySelector('[name=severity]:checked')?.value || null;
        const entry = await openCommentLibraryPicker({
          inspectionTypeKey,
          severity: currentSeverity
        });
        if (!entry) return;
        commentEl.value = entry.text;
        libraryState.libraryKey = entry.key || null;
        libraryState.asClause   = entry.asClause || '';
        clauseText.textContent = libraryState.asClause;
        clauseBadge.classList.toggle('is-hidden', !libraryState.asClause);
      });

      clauseClear.addEventListener('click', () => {
        libraryState.libraryKey = null;
        libraryState.asClause   = '';
        clauseText.textContent = '';
        clauseBadge.classList.add('is-hidden');
      });

      // --- Dictation (Web Speech API) wiring — graceful on unsupported browsers ---
      const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
      const dictateBtn = dialog.querySelector('#ip-dictate-btn');
      const iosHint    = dialog.querySelector('.dictation-hint__ios');
      if (SR && dictateBtn) {
        // Supported — show the button, hide the iOS-style hint (user can still
        // use keyboard mic too on mobile Safari, but desktop benefits from
        // an explicit button).
        dictateBtn.classList.remove('is-hidden');
        let rec = null;
        let listening = false;
        dictateBtn.addEventListener('click', () => {
          if (listening) { rec?.stop?.(); return; }
          rec = new SR();
          rec.continuous = true;
          rec.interimResults = false;
          rec.lang = 'en-AU';
          rec.onstart = () => {
            listening = true;
            dictateBtn.textContent = '\u25a0 Stop dictating';
            dictateBtn.classList.add('is-listening');
          };
          rec.onresult = (e) => {
            let transcript = '';
            for (let i = e.resultIndex; i < e.results.length; i += 1) {
              if (e.results[i].isFinal) transcript += e.results[i][0].transcript + ' ';
            }
            if (transcript) {
              const cur = commentEl.value;
              commentEl.value = cur ? cur.replace(/\s+$/, '') + ' ' + transcript.trim() : transcript.trim();
              commentEl.dispatchEvent(new Event('input', { bubbles: true }));
            }
          };
          rec.onerror = (e) => {
            console.warn('Speech recognition error:', e.error);
            if (e.error === 'not-allowed') {
              toast('Microphone permission denied — allow mic access to dictate.', { kind: 'error' });
            }
          };
          rec.onend = () => {
            listening = false;
            dictateBtn.textContent = 'Start dictating';
            dictateBtn.classList.remove('is-listening');
          };
          try { rec.start(); } catch (err) {
            toast('Couldn\u2019t start dictation', { kind: 'error' });
          }
        });
      }

      // --- AI Expand wiring ---
      const btnExpand = dialog.querySelector('#ip-btn-expand');
      btnExpand.addEventListener('click', async () => {
        const note = commentEl.value.trim();
        if (!note) {
          toast('Type or dictate a short note first, then tap Expand.', { kind: 'info' });
          commentEl.focus();
          return;
        }
        btnExpand.disabled = true;
        const prevLabel = btnExpand.innerHTML;
        btnExpand.innerHTML = '<span aria-hidden="true">\u2728</span> Expanding\u2026';
        try {
          const [profile, drawing] = await Promise.all([
            getUserProfile(),
            item.drawingId ? getDrawing(item.drawingId) : Promise.resolve(null)
          ]);
          const gridRef = dialog.querySelector('[name=gridRef]')?.value?.trim() || '';
          const severity = dialog.querySelector('[name=severity]:checked')?.value || 'observation';
          const result = await expandComment({
            apiKey: profile.anthropicKey,
            note,
            context: {
              inspectionTypeName: inspection?.inspectionTypeName || '',
              drawingSheet:       drawing?.sheetNumber || '',
              drawingRevision:    drawing?.revision || '',
              gridRef,
              severity
            }
          });
          // Replace comment with expanded version; stash original via alert for
          // undo — a compact two-line preview is too modal for a touch UI.
          commentEl.value = result.text;
          commentEl.dispatchEvent(new Event('input', { bubbles: true }));
          if (result.asClause) {
            libraryState.asClause = result.asClause;
            clauseText.textContent = result.asClause;
            clauseBadge.classList.remove('is-hidden');
          }
          toast('Expanded. Review and edit before saving.', { kind: 'success' });
        } catch (err) {
          toast(err.message || 'Expansion failed', { kind: 'error' });
        } finally {
          btnExpand.disabled = false;
          btnExpand.innerHTML = prevLabel;
        }
      });

      // Paint initial photos
      refreshPhotos();
    },
    onConfirm: async (dialog) => {
      const comment  = dialog.querySelector('[name=comment]').value;
      const gridRef  = dialog.querySelector('[name=gridRef]').value;
      const severity = dialog.querySelector('[name=severity]:checked')?.value || 'observation';
      const status   = dialog.querySelector('[name=status]:checked')?.value   || 'open';
      const updated = await updateItem(itemId, {
        comment, gridRef, severity, status,
        libraryKey: libraryState.libraryKey,
        asClause:   libraryState.asClause
      });
      result.saved = updated;
      return updated;
    }
  });

  // Revoke URLs on close (succeed or cancel)
  urlRegistry.revokeAll();

  if (result.deleted) return { deleted: true };
  if (result.saved)   return { item: result.saved };
  return null;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
function escapeAttr(s) { return escapeHtml(s).replace(/"/g, '&quot;'); }
