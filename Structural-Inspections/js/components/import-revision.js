/**
 * Import a Cowork-produced revision-{ts}.json diff into the app.
 *
 * Flow:
 *   1. File picker — user selects a revision-*.json
 *   2. Parse + summarise (preview modal)
 *   3. On confirm, apply via db.applyRevisionDiff(projectId, diff)
 *   4. Toast + refresh project view
 */

import { applyRevisionDiff } from '../db.js';
import { openModal } from './modal.js';
import { toast } from './toast.js';
import { go } from '../router.js';

export async function openImportRevision(projectId) {
  const blob = await pickRevisionFile();
  if (!blob) return null;

  let diff;
  try {
    const text = await blob.text();
    diff = JSON.parse(text);
  } catch (err) {
    toast(`Couldn't parse revision file: ${err.message}`, { kind: 'error', duration: 6000 });
    return null;
  }

  if (diff.schemaVersion !== 1) {
    toast(`Unsupported revision schema v${diff.schemaVersion}`, { kind: 'error' });
    return null;
  }

  const summary = renderSummary(diff);
  const result = await openModal({
    title: 'Import revision',
    primaryLabel: 'Apply diff',
    cancelLabel:  'Cancel',
    bodyHtml: summary,
    onConfirm: () => true
  });
  if (!result) return null;

  try {
    const r = await applyRevisionDiff(projectId, diff);
    toast(
      `Revision applied — ${r.stampedDrawings} drawing${r.stampedDrawings === 1 ? '' : 's'} stamped, ${r.affectedPlanEntries} plan entr${r.affectedPlanEntries === 1 ? 'y' : 'ies'} flagged for review`,
      { kind: 'success', duration: 5000 }
    );
    go('project', projectId);
    return diff;
  } catch (err) {
    toast(err.message || 'Apply failed', { kind: 'error', duration: 8000 });
    return null;
  }
}


/* ─────────── helpers ─────────── */

function pickRevisionFile() {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json,application/json';
    input.style.display = 'none';
    document.body.appendChild(input);

    let resolved = false;
    const cleanup = () => { try { document.body.removeChild(input); } catch {} };

    input.addEventListener('change', () => {
      const f = input.files?.[0];
      if (!f) {
        if (!resolved) { resolved = true; cleanup(); resolve(null); }
        return;
      }
      if (!resolved) { resolved = true; cleanup(); resolve(f); }
    });
    input.addEventListener('cancel', () => {
      if (!resolved) { resolved = true; cleanup(); resolve(null); }
    });

    input.click();
    // Some browsers don't fire 'cancel' — fallback after a focus return
    window.addEventListener('focus', () => {
      setTimeout(() => {
        if (!resolved) { resolved = true; cleanup(); resolve(null); }
      }, 500);
    }, { once: true });
  });
}

function renderSummary(diff) {
  return `
    <div class="rev-summary">
      <p class="muted small">
        <strong>${escapeHtml(diff.newPdfFilename || 'New PDF')}</strong> compared against
        <strong>${escapeHtml(diff.againstIssue || 'current project')}</strong>.
      </p>

      <div class="rev-summary__grid">
        <div class="rev-summary__stat">
          <span class="rev-summary__num">${(diff.added || []).length}</span>
          <span class="rev-summary__lbl">Added</span>
        </div>
        <div class="rev-summary__stat">
          <span class="rev-summary__num rev-summary__num--accent">${(diff.changed || []).length}</span>
          <span class="rev-summary__lbl">Changed</span>
        </div>
        <div class="rev-summary__stat">
          <span class="rev-summary__num">${(diff.removed || []).length}</span>
          <span class="rev-summary__lbl">Removed</span>
        </div>
        <div class="rev-summary__stat">
          <span class="rev-summary__num">${(diff.affectedInspectionPlanIndices || []).length}</span>
          <span class="rev-summary__lbl">Plan items affected</span>
        </div>
      </div>

      ${(diff.changed || []).length ? `
        <div class="rev-summary__section">
          <div class="rev-summary__title">Changed sheets</div>
          <ul>
            ${diff.changed.slice(0, 12).map((c) => `
              <li><code>${escapeHtml(c.sheetNumber)}</code>
                  <span class="muted small">${escapeHtml(c.oldRevision || '—')} → ${escapeHtml(c.newRevision)}</span>
                  ${c.description ? `<div class="muted small">${escapeHtml(c.description)}</div>` : ''}
              </li>
            `).join('')}
            ${diff.changed.length > 12 ? `<li class="muted small">…and ${diff.changed.length - 12} more</li>` : ''}
          </ul>
        </div>
      ` : ''}

      ${(diff.added || []).length ? `
        <div class="rev-summary__section">
          <div class="rev-summary__title">Added sheets</div>
          <ul>
            ${diff.added.slice(0, 12).map((a) => `
              <li><code>${escapeHtml(a.sheetNumber)}</code>
                  <span class="muted small">rev ${escapeHtml(a.revision || '—')}</span>
                  ${a.description ? `<div class="muted small">${escapeHtml(a.description)}</div>` : ''}
              </li>
            `).join('')}
            ${diff.added.length > 12 ? `<li class="muted small">…and ${diff.added.length - 12} more</li>` : ''}
          </ul>
        </div>
      ` : ''}

      ${(diff.removed || []).length ? `
        <div class="rev-summary__section">
          <div class="rev-summary__title">Removed sheets</div>
          <ul>
            ${diff.removed.slice(0, 12).map((r) => `
              <li><code>${escapeHtml(r.sheetNumber)}</code>
                  ${r.description ? `<div class="muted small">${escapeHtml(r.description)}</div>` : ''}
              </li>
            `).join('')}
            ${diff.removed.length > 12 ? `<li class="muted small">…and ${diff.removed.length - 12} more</li>` : ''}
          </ul>
        </div>
      ` : ''}

      <div class="muted small" style="margin-top: 12px;">
        Applying this diff will stamp the affected drawings and plan entries with a "revision changed" badge.
        It does NOT replace the source PDFs — for that, generate a new .btproject bundle and re-import.
      </div>
    </div>
  `;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
