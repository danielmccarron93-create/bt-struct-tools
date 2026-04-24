/**
 * Comment library picker (Phase 5).
 *
 * Shows library entries for an inspection type — grouped by category,
 * with an optional severity filter (observation / defect / all) and a
 * free-text search box. Resolves to the selected entry, or null if the
 * user cancels.
 *
 * Usage:
 *   const entry = await openCommentLibraryPicker({
 *     inspectionTypeKey: 'slab-prepour-ground',
 *     severity: 'defect',         // or null to show both
 *     title: 'Pick a comment'     // optional
 *   });
 *   if (entry) {
 *     // entry = { key, text, asClause, severity, category, inspectionTypeKey }
 *   }
 */

import {
  listCommentsForInspectionType,
  listGeneralCommentsForInspectionType
} from '../db.js';
import { openModal } from '../components/modal.js';

export async function openCommentLibraryPicker({
  inspectionTypeKey,
  severity = null,
  title = 'Pick a comment from the library'
} = {}) {
  const entries = await listCommentsForInspectionType(inspectionTypeKey, { severity });

  if (entries.length === 0) {
    return openModal({
      title: 'No matching library entries',
      bodyHtml: `
        <p>No comments in the library for this inspection type${severity ? ` and severity "${severity}"` : ''}.</p>
        <p class="muted">Edit the seed list in <code>js/db.js</code> or add entries via Settings (coming in Phase 8).</p>
      `,
      primaryLabel: 'OK',
      cancelLabel: 'Close',
      onConfirm: () => null
    }).then(() => null);
  }

  // Group by category
  const grouped = new Map();
  for (const e of entries) {
    const cat = e.category || 'Other';
    if (!grouped.has(cat)) grouped.set(cat, []);
    grouped.get(cat).push(e);
  }

  const bodyHtml = `
    <div class="library-picker">
      <div class="library-picker__toolbar">
        <input type="search" id="lp-search" class="library-picker__search"
               placeholder="Search comments…" autocomplete="off"/>
        <div class="library-picker__sev-filter" role="group" aria-label="Severity filter">
          <label class="radio-chip"><input type="radio" name="sev" value="" ${!severity ? 'checked' : ''}/><span>All</span></label>
          <label class="radio-chip radio-chip--observation"><input type="radio" name="sev" value="observation" ${severity === 'observation' ? 'checked' : ''}/><span>Observation</span></label>
          <label class="radio-chip radio-chip--defect"><input type="radio" name="sev" value="defect" ${severity === 'defect' ? 'checked' : ''}/><span>Defect</span></label>
          <label class="radio-chip radio-chip--holdpoint"><input type="radio" name="sev" value="holdpoint" ${severity === 'holdpoint' ? 'checked' : ''}/><span>Hold point</span></label>
        </div>
      </div>
      <div class="library-picker__list" id="lp-list" role="listbox">
        ${renderGrouped(grouped)}
      </div>
    </div>
  `;

  let chosen = null;

  await openModal({
    title,
    primaryLabel: 'Use selected',
    bodyHtml,
    onMount: (dialog) => {
      dialog.classList.add('modal-dialog--wide');
      const search = dialog.querySelector('#lp-search');
      const list   = dialog.querySelector('#lp-list');
      const sevRadios = dialog.querySelectorAll('[name=sev]');

      function applyFilters() {
        const term = search.value.trim().toLowerCase();
        const sev = Array.from(sevRadios).find((r) => r.checked)?.value || '';
        list.querySelectorAll('.library-picker__row').forEach((row) => {
          const matchSev  = !sev || row.dataset.sev === sev;
          const matchText = !term || row.textContent.toLowerCase().includes(term);
          row.classList.toggle('is-hidden', !(matchSev && matchText));
        });
        // Hide empty category headings
        list.querySelectorAll('.library-picker__category').forEach((cat) => {
          const next = cat.nextElementSibling;
          const anyVisible = Array.from(list.querySelectorAll(`.library-picker__row[data-category="${cat.dataset.category}"]`))
            .some((r) => !r.classList.contains('is-hidden'));
          cat.classList.toggle('is-hidden', !anyVisible);
        });
      }

      search.addEventListener('input', applyFilters);
      sevRadios.forEach((r) => r.addEventListener('change', applyFilters));

      list.addEventListener('click', (e) => {
        const row = e.target.closest('.library-picker__row');
        if (!row) return;
        // Toggle selection
        list.querySelectorAll('.library-picker__row.is-selected')
          .forEach((r) => r.classList.remove('is-selected'));
        row.classList.add('is-selected');
        chosen = entries.find((en) => en.key === row.dataset.key) || null;
      });

      // Double-click / Enter confirms immediately
      list.addEventListener('dblclick', (e) => {
        const row = e.target.closest('.library-picker__row');
        if (!row) return;
        chosen = entries.find((en) => en.key === row.dataset.key) || null;
        if (chosen) dialog.querySelector('[data-role=confirm]').click();
      });

      // Focus search box for quick keyboard filtering
      setTimeout(() => search.focus(), 60);
    },
    onConfirm: () => {
      if (!chosen) throw new Error('Pick a comment from the list first');
      return chosen;
    }
  });

  return chosen;
}

function renderGrouped(grouped) {
  let out = '';
  // Keep category order stable (Map preserves insertion order)
  for (const [cat, items] of grouped) {
    out += `<div class="library-picker__category" data-category="${escapeAttr(cat)}">${escapeHtml(cat)}</div>`;
    for (const e of items) {
      out += `
        <div class="library-picker__row" role="option" tabindex="0"
             data-key="${escapeAttr(e.key)}"
             data-sev="${escapeAttr(e.severity || '')}"
             data-category="${escapeAttr(cat)}">
          <div class="library-picker__row-main">
            <div class="library-picker__text">${escapeHtml(e.text)}</div>
            ${e.asClause ? `<div class="library-picker__clause muted">${escapeHtml(e.asClause)}</div>` : ''}
          </div>
          <div class="library-picker__badges">
            <span class="pill pill--${escapeAttr(e.severity || 'observation')}">${escapeHtml(e.severity || 'observation')}</span>
          </div>
        </div>
      `;
    }
  }
  return out;
}

/* --------------------------------------------------------------------------
   General-comments picker
   --------------------------------------------------------------------------
   A simpler variant for the "General comments" section on the inspection
   detail screen — flat list, no severity filter, multi-select so the
   engineer can tick several at once.

   Resolves to an array of { text, inspectionTypeKey } entries (may be empty).
   -------------------------------------------------------------------------- */
export async function openGeneralCommentPicker({ inspectionTypeKey, excludeTexts = [] } = {}) {
  const entries = await listGeneralCommentsForInspectionType(inspectionTypeKey);
  // Exclude anything that's already been added to the inspection so we don't
  // double up. Matching on exact text — good enough for seed data.
  const excludeSet = new Set(excludeTexts);
  const choices = entries.filter((e) => !excludeSet.has(e.text));

  if (choices.length === 0) {
    return openModal({
      title: 'Nothing new to add',
      bodyHtml: `<p>All library general-comments for this inspection type are already added. You can still add a custom comment.</p>`,
      primaryLabel: 'OK',
      cancelLabel: 'Close',
      onConfirm: () => []
    }).then(() => []);
  }

  const bodyHtml = `
    <div class="library-picker">
      <div class="library-picker__list" id="gp-list" role="listbox" aria-multiselectable="true">
        ${choices.map((e, idx) => `
          <label class="library-picker__row library-picker__row--checkable">
            <input type="checkbox" data-idx="${idx}"/>
            <div class="library-picker__row-main">
              <div class="library-picker__text">${escapeHtml(e.text)}</div>
              ${e.inspectionTypeKey && e.inspectionTypeKey !== '_any_'
                ? `<div class="library-picker__clause muted">Specific to this inspection type</div>`
                : `<div class="library-picker__clause muted">Universal</div>`}
            </div>
          </label>
        `).join('')}
      </div>
    </div>
  `;

  const picked = [];

  await openModal({
    title: 'Add from general-comments library',
    primaryLabel: 'Add selected',
    bodyHtml,
    onMount: (dialog) => {
      dialog.classList.add('modal-dialog--wide');
    },
    onConfirm: (dialog) => {
      const checked = dialog.querySelectorAll('#gp-list input[type=checkbox]:checked');
      if (checked.length === 0) throw new Error('Tick at least one comment, or cancel to skip');
      checked.forEach((cb) => {
        const idx = Number(cb.dataset.idx);
        if (choices[idx]) picked.push(choices[idx]);
      });
      return true;
    }
  });

  return picked;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
function escapeAttr(s) { return escapeHtml(s).replace(/"/g, '&quot;'); }
