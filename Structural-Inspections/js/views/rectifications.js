/**
 * Outstanding Rectifications view — global across all projects.
 *
 * Per BUILD-PLAN.md Phase C1: the single most important screen for the senior
 * engineer between visits. Shows every open defect / hold-point item, sorted
 * oldest first. One-tap close-out with optional rectification notes + photo.
 *
 * Data via db.listOutstandingRectifications() — already enriched with project +
 * inspection + first-photo context.
 */

import {
  listOutstandingRectifications, listProjects,
  closeItem, reopenItem, addItemPhoto
} from '../db.js';
import { go } from '../router.js';
import { toast } from '../components/toast.js';
import { openModal } from '../components/modal.js';

let urlRegistry = null;   // local URL.createObjectURL cache for photo blobs

export async function render(root) {
  // Photo URL cache — release when leaving view (router replaces innerHTML)
  if (urlRegistry) urlRegistry.release();
  urlRegistry = createUrlRegistry();

  const [openRects, projects] = await Promise.all([
    listOutstandingRectifications(),
    listProjects()
  ]);

  // Build project filter dropdown
  const projectsWithRects = new Map();
  for (const r of openRects) {
    if (!projectsWithRects.has(r.project.id)) projectsWithRects.set(r.project.id, r.project);
  }

  const totalCount = openRects.length;
  const defectCount    = openRects.filter((r) => severityKey(r.item.severity) === 'defect').length;
  const holdPointCount = openRects.filter((r) => severityKey(r.item.severity) === 'holdpoint').length;
  const oldest = openRects.length > 0 ? openRects[0].ageDays : 0;

  root.innerHTML = `
    <div class="section-heading">
      <h1>Outstanding Rectifications</h1>
    </div>

    <div class="rect-stats">
      <div class="rect-stat">
        <span class="rect-stat__num ${totalCount > 0 ? 'rect-stat__num--accent' : ''}">${totalCount}</span>
        <span class="rect-stat__lbl">Open</span>
      </div>
      <div class="rect-stat">
        <span class="rect-stat__num">${defectCount}</span>
        <span class="rect-stat__lbl">Defects</span>
      </div>
      <div class="rect-stat">
        <span class="rect-stat__num ${holdPointCount > 0 ? 'rect-stat__num--accent' : ''}">${holdPointCount}</span>
        <span class="rect-stat__lbl">Hold points</span>
      </div>
      <div class="rect-stat">
        <span class="rect-stat__num">${oldest}<span class="rect-stat__sub">d</span></span>
        <span class="rect-stat__lbl">Oldest open</span>
      </div>
    </div>

    ${projectsWithRects.size > 1 ? `
      <div class="rect-filter">
        <label for="rect-project-filter" class="muted small">Filter by project:</label>
        <select id="rect-project-filter">
          <option value="">All projects (${totalCount})</option>
          ${[...projectsWithRects.values()].map((p) => `
            <option value="${p.id}">${escapeHtml(p.jobNumber || '')} ${escapeHtml(p.name)}</option>
          `).join('')}
        </select>
      </div>
    ` : ''}

    ${openRects.length === 0
      ? renderEmpty()
      : renderList(openRects)}
  `;

  // Filter dropdown
  const filterSel = root.querySelector('#rect-project-filter');
  if (filterSel) {
    filterSel.addEventListener('change', () => {
      const pid = filterSel.value;
      root.querySelectorAll('[data-item-id]').forEach((card) => {
        if (!pid || card.dataset.projectId === pid) {
          card.classList.remove('is-hidden');
        } else {
          card.classList.add('is-hidden');
        }
      });
    });
  }

  // Click on a rectification card → details + close-out modal
  root.querySelectorAll('[data-item-id]').forEach((card) => {
    card.addEventListener('click', async (e) => {
      // Ignore clicks on the close button
      if (e.target.closest('[data-action="close-rect"]')) return;
      const itemId = Number(card.dataset.itemId);
      const inspId = Number(card.dataset.inspectionId);
      // Just navigate to the inspection — engineer can review there
      go('inspection', inspId);
    });
  });

  // Close-out button on each card
  root.querySelectorAll('[data-action="close-rect"]').forEach((btn) => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const itemId = Number(btn.closest('[data-item-id]').dataset.itemId);
      const rect = openRects.find((r) => r.item.id === itemId);
      if (!rect) return;
      await openCloseOutModal(rect, () => location.reload());
    });
  });
}


/* ------------------------------------------------------------------------- */

function renderEmpty() {
  return `
    <div class="empty card">
      <h2>No outstanding rectifications</h2>
      <p class="muted">All defects and hold points across your projects are closed. When defects are raised on inspections they appear here until you close them out with a rectification note + photo.</p>
    </div>
  `;
}

function renderList(rects) {
  return `
    <ul class="rect-list stack" role="list">
      ${rects.map((r) => renderCard(r)).join('')}
    </ul>
  `;
}

function renderCard(r) {
  const { item, inspection, project, primaryDrawing, firstPhoto, ageDays } = r;
  const sev = severityKey(item.severity);
  const sevLabel = sev === 'holdpoint' ? 'Hold point' : (sev === 'defect' ? 'Defect' : 'Item');
  const previewRaw = (item.comment || '').trim();
  const previewText = previewRaw
    ? (previewRaw.length > 180 ? previewRaw.slice(0, 177) + '…' : previewRaw)
    : '(no comment)';
  const ageLabel = ageDays === 0 ? 'today' : (ageDays === 1 ? '1 day' : `${ageDays} days`);
  const ageBucket = ageDays >= 14 ? 'rect-card__age--old' : (ageDays >= 7 ? 'rect-card__age--mid' : '');

  const thumbHtml = firstPhoto
    ? `<img src="${urlRegistry.urlFor(firstPhoto.blob)}" alt="" class="rect-card__thumb-img" loading="lazy"/>`
    : `<span class="rect-card__thumb-placeholder" aria-hidden="true">—</span>`;

  return `
    <li class="rect-card rect-card--${sev}"
        data-item-id="${item.id}"
        data-inspection-id="${inspection.id}"
        data-project-id="${project.id}"
        tabindex="0" role="link"
        onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}">
      <div class="rect-card__num rect-card__num--${sev}">${item.itemNumber}</div>
      <div class="rect-card__thumb">${thumbHtml}</div>
      <div class="rect-card__body">
        <div class="rect-card__head">
          <span class="rect-card__sev rect-card__sev--${sev}">${sevLabel}</span>
          <span class="rect-card__age ${ageBucket}" title="Open ${ageLabel}">${ageLabel}</span>
        </div>
        <div class="rect-card__project muted">
          ${escapeHtml(project.jobNumber || '')} ${project.jobNumber ? '·' : ''} ${escapeHtml(project.name)}
        </div>
        <div class="rect-card__inspection">
          ${escapeHtml(inspection.inspectionTypeName || inspection.types?.[0] || '')}
          ${primaryDrawing?.sheetNumber ? ` <span class="muted">· ${escapeHtml(primaryDrawing.sheetNumber)}</span>` : ''}
        </div>
        <div class="rect-card__preview">${escapeHtml(previewText)}</div>
        ${item.gridRef ? `<div class="rect-card__grid muted small">Grid: ${escapeHtml(item.gridRef)}</div>` : ''}
        ${item.asClause ? `<div class="rect-card__clause muted small">${escapeHtml(item.asClause)}</div>` : ''}
      </div>
      <div class="rect-card__actions">
        <button class="btn btn--secondary btn--sm" data-action="close-rect" title="Mark this rectification as closed">
          ✓ Close
        </button>
      </div>
    </li>
  `;
}


/* ----- Close-out modal ----- */

async function openCloseOutModal(rect, onClosed) {
  const { item, inspection, project, primaryDrawing } = rect;
  const sevLabel = severityKey(item.severity) === 'holdpoint' ? 'hold point' : 'defect';

  const bodyHtml = `
    <div class="form">
      <p class="muted">
        You're closing out <strong>item #${item.itemNumber}</strong> (${sevLabel}) on
        <em>${escapeHtml(inspection.inspectionTypeName || '')}</em>
        for <strong>${escapeHtml(project.name)}</strong>.
      </p>

      <div class="form-field">
        <label class="muted small">Original comment:</label>
        <div class="muted" style="padding: 6px 10px; background: rgba(0,0,0,0.04); border-radius: 4px; font-size: 13px;">
          ${escapeHtml(item.comment || '(no comment)')}
        </div>
      </div>

      <div class="form-field">
        <label for="rect-notes">Rectification notes (optional)</label>
        <textarea id="rect-notes" name="closedNotes" rows="3" placeholder="What was done to rectify? Who confirmed? Any limitations?"></textarea>
      </div>

      <div class="form-field">
        <label for="rect-photo" class="muted small">Rectification photo (optional but recommended):</label>
        <input id="rect-photo" type="file" accept="image/*" capture="environment">
      </div>
    </div>
  `;

  const result = await openModal({
    title: `Close rectification #${item.itemNumber}`,
    primaryLabel: 'Close out',
    cancelLabel:  'Cancel',
    bodyHtml,
    onConfirm: async (modalRoot) => {
      const notes = modalRoot.querySelector('#rect-notes').value.trim();
      const photoInput = modalRoot.querySelector('#rect-photo');
      const photoFile = photoInput?.files?.[0] || null;

      let photoId = null;
      if (photoFile) {
        try {
          const photo = await addItemPhoto(item.id, { blob: photoFile, source: 'rectification' });
          photoId = photo?.id || null;
        } catch (err) {
          console.error('Failed to attach rectification photo', err);
          toast('Photo attach failed (item still closed)', { kind: 'warn' });
        }
      }

      await closeItem(item.id, { notes, photoId });
      return true;
    }
  });

  if (result) {
    toast('Rectification closed', { kind: 'success' });
    if (onClosed) onClosed();
  }
}


/* ----- Helpers ----- */

function severityKey(s) {
  const v = String(s || '').toLowerCase().replace(/[\s-]/g, '');
  if (v === 'holdpoint') return 'holdpoint';
  if (v === 'defect')    return 'defect';
  return 'observation';
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}

function createUrlRegistry() {
  const cache = new Map();
  return {
    urlFor(blob) {
      if (!blob) return '';
      let url = cache.get(blob);
      if (!url) {
        url = URL.createObjectURL(blob);
        cache.set(blob, url);
      }
      return url;
    },
    release() {
      for (const url of cache.values()) {
        try { URL.revokeObjectURL(url); } catch {}
      }
      cache.clear();
    }
  };
}
