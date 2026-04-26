/**
 * Import a Cowork-produced .btproject bundle into the app.
 *
 * Flow:
 *   1. Open a hidden file input — user picks a .btproject (zip) file.
 *   2. Read + validate the bundle via lib/btproject.js.
 *   3. Show a preview modal: project metadata, drawing count, plan length,
 *      warnings. User can cancel or commit.
 *   4. On commit, run the atomic multi-table insert in db.js.
 *   5. Toast success and navigate to the new project.
 *
 * Errors are surfaced via toast — the user always knows what went wrong.
 */

import { readBundle, summariseBundle } from '../lib/btproject.js';
import { createProjectFromBundle } from '../db.js';
import { openModal } from './modal.js';
import { toast } from './toast.js';
import { go } from '../router.js';

/**
 * Trigger the import flow. Returns the new project's id (or null if cancelled).
 * Optional: pass a Blob to skip the file picker (used in tests + paste-from-share).
 */
export async function openImportProject({ blob: providedBlob } = {}) {
  let blob = providedBlob || null;

  // Step 1 — file picker (if no blob provided)
  if (!blob) {
    blob = await pickBundleFile();
    if (!blob) return null;     // user cancelled
  }

  // Step 2 — load + validate the bundle. Show a transient "reading…" toast
  // because large bundles (200+ MB) can take a couple of seconds.
  const readingToast = toast('Reading bundle…', { kind: 'info', duration: 0 });
  let parsed;
  try {
    parsed = await readBundle(blob);
  } catch (err) {
    readingToast?.dismiss?.();
    toast(err.message || 'Couldn\u2019t read bundle', { kind: 'error', duration: 6000 });
    return null;
  }
  readingToast?.dismiss?.();

  // Step 3 — preview modal
  const summary = summariseBundle(parsed.projectMap, parsed.pdfBlobs);
  const warnings = parsed.projectMap.warnings || [];
  const result = await openModal({
    title: 'Import project',
    primaryLabel: 'Import',
    cancelLabel:  'Cancel',
    bodyHtml: renderPreviewBody(parsed, summary, warnings),
    onConfirm: async () => parsed                  // pass through
  });

  if (!result) return null;

  // Step 4 — commit
  const committingToast = toast('Importing…', { kind: 'info', duration: 0 });
  let imported;
  try {
    imported = await createProjectFromBundle(parsed);
  } catch (err) {
    console.error('Import failed:', err);
    committingToast?.dismiss?.();
    toast(err.message || 'Import failed', { kind: 'error', duration: 8000 });
    return null;
  }
  committingToast?.dismiss?.();

  // Step 5 — success
  toast(
    `Imported ${imported.project.name} — ${imported.drawingCount} drawing${imported.drawingCount === 1 ? '' : 's'}`,
    { kind: 'success', duration: 4000 }
  );
  go('project', imported.id);
  return imported.id;
}

/* --------------------------------------------------------------------------
   Helpers
   -------------------------------------------------------------------------- */

/**
 * Open a hidden file input and resolve with the selected Blob, or null on
 * cancel. Accepts .btproject, .zip and any zip MIME type.
 */
function pickBundleFile() {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.btproject,.zip,application/zip,application/octet-stream';
    input.style.display = 'none';
    document.body.appendChild(input);

    let resolved = false;
    const cleanup = () => {
      try { document.body.removeChild(input); } catch {}
    };

    input.addEventListener('change', () => {
      resolved = true;
      const f = input.files && input.files[0] ? input.files[0] : null;
      cleanup();
      resolve(f);
    });

    // A user cancel doesn't fire 'change' on iOS — fall back to focus events
    // on the body that fire after the picker dismisses.
    const onFocus = () => {
      // Defer one tick so 'change' has a chance to fire first
      setTimeout(() => {
        if (!resolved) {
          window.removeEventListener('focus', onFocus);
          cleanup();
          resolve(null);
        }
      }, 350);
    };
    window.addEventListener('focus', onFocus);

    input.click();
  });
}

function renderPreviewBody(parsed, summary, warnings) {
  const p  = parsed.projectMap.project || {};
  const dr = parsed.projectMap.drawings || [];
  const ip = parsed.projectMap.inspectionPlan || [];

  // Tag distribution — show the top-N elements so the engineer can confirm
  // the classification looks sensible.
  const elementCounts = {};
  for (const d of dr) {
    for (const tag of (d.tags || [])) {
      if (!tag.startsWith('element:')) continue;
      const el = tag.slice('element:'.length);
      elementCounts[el] = (elementCounts[el] || 0) + 1;
    }
  }
  const topElements = Object.entries(elementCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  return `
    <div class="stack">
      <div class="import-preview__header">
        <div class="import-preview__job">${escapeHtml(p.jobNumber || '')}</div>
        <div class="import-preview__name">${escapeHtml(p.name || '(unnamed)')}</div>
        ${p.siteAddress ? `<div class="muted">${escapeHtml(p.siteAddress)}</div>` : ''}
      </div>

      <dl class="import-preview__meta">
        ${p.client      ? `<div><dt>Client</dt><dd>${escapeHtml(p.client)}</dd></div>` : ''}
        ${p.issueStatus ? `<div><dt>Issue</dt><dd>${escapeHtml(p.issueStatus)}</dd></div>` : ''}
        ${p.discipline  ? `<div><dt>Discipline</dt><dd>${escapeHtml(p.discipline)}</dd></div>` : ''}
        ${p.engineerOfRecord ? `<div><dt>Engineer of record</dt><dd>${escapeHtml(p.engineerOfRecord)}</dd></div>` : ''}
      </dl>

      <div class="import-preview__counts">
        <div><strong>${dr.length}</strong> drawing${dr.length === 1 ? '' : 's'} across <strong>${parsed.pdfBlobs.length}</strong> PDF${parsed.pdfBlobs.length === 1 ? '' : 's'}</div>
        <div><strong>${ip.length}</strong> inspection${ip.length === 1 ? '' : 's'} proposed</div>
      </div>

      ${topElements.length ? `
        <div>
          <div class="muted import-preview__section-title">Drawings classified by element</div>
          <ul class="import-preview__chips" role="list">
            ${topElements.map(([el, n]) =>
              `<li class="chip">${escapeHtml(el)} <span class="chip__count">${n}</span></li>`
            ).join('')}
          </ul>
        </div>
      ` : ''}

      ${ip.length ? `
        <div>
          <div class="muted import-preview__section-title">Inspection plan</div>
          <ol class="import-preview__plan">
            ${ip.slice(0, 8).map((entry) => `
              <li>
                <span class="import-preview__plan-title">${escapeHtml(entry.title || entry.type || '')}</span>
                ${entry.holdPoint ? `<span class="badge badge--hold">hold point</span>` : ''}
                ${entry.level ? `<span class="muted import-preview__plan-meta">${escapeHtml(entry.level)}</span>` : ''}
              </li>
            `).join('')}
            ${ip.length > 8 ? `<li class="muted">…and ${ip.length - 8} more</li>` : ''}
          </ol>
        </div>
      ` : ''}

      ${warnings.length ? `
        <div class="import-preview__warnings" role="alert">
          <div class="import-preview__warnings-title">${warnings.length} warning${warnings.length === 1 ? '' : 's'} from Cowork</div>
          <ul>
            ${warnings.slice(0, 6).map((w) => `<li>${escapeHtml(w)}</li>`).join('')}
            ${warnings.length > 6 ? `<li class="muted">…and ${warnings.length - 6} more</li>` : ''}
          </ul>
        </div>
      ` : ''}

      <p class="muted small">Click <strong>Import</strong> to commit. Drawings, inspection plan, and general notes will all be saved to this device.</p>
    </div>
  `;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
