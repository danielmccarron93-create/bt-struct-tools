/**
 * Pre-Inspection Brief modal (Phase C2 — v2.2)
 *
 * Shown when an engineer taps a pending inspection plan entry, BEFORE creating
 * the inspection. Surfaces the rich project-map data so the engineer arrives on
 * site sharp:
 *
 *   - Rationale / why this inspection
 *   - Cover values + concrete grade for the relevant element (from generalNotes)
 *   - Site-readiness check (per-entry override, else sensible default per type)
 *   - Certified-by-others exclusions for this project
 *   - Expected checklist with AS clause refs where available
 *   - Drawings to take to site (already resolved by the plan)
 *   - Hold point status + post-pour record requirements
 *
 * Two outcomes:
 *   - "Start inspection" → calls startInspectionFromPlanEntry
 *   - "Site not ready / cancel" → no inspection created
 */

import { startInspectionFromPlanEntry } from '../db.js';
import { openModal } from './modal.js';
import { toast } from './toast.js';
import { go } from '../router.js';
import { mergeWithProjectChecks } from '../lib/bt-standard-checks.js';

/**
 * Open the pre-inspection brief for a given plan entry.
 *
 * @param {Object} project        — full project record (with projectMap, inspectionPlan, etc.)
 * @param {number} planIndex      — index into project.inspectionPlan
 * @returns {Promise<Object|null>} the created inspection record, or null if cancelled
 */
export async function openPreInspectionBrief(project, planIndex) {
  const entry = project?.inspectionPlan?.[planIndex];
  if (!entry) return null;

  const generalNotes = project.projectMap?.generalNotes || {};
  const excludedFromBT = project.excludedFromBT || project.projectMap?.project?.excludedFromBT || [];

  const ELEMENT = inferCoverElementFromType(entry.type);
  const cover = ELEMENT ? generalNotes.concreteCover?.[ELEMENT] : null;
  const grade = ELEMENT ? generalNotes.concreteStrengths?.[ELEMENT] : null;

  const drawings = Array.isArray(entry.drawingRefs) ? entry.drawingRefs : [];

  // Merge BT standard checks with project-specific checks (deduped, project-first).
  // Returns { siteReadiness, onSiteChecks, pourDayRecords, certifications, standardComments }
  // each item tagged source: 'project' | 'standard' so the UI can show provenance.
  const merged = mergeWithProjectChecks(entry.type, entry);
  const siteReadiness = merged.siteReadiness;
  const checklist     = merged.onSiteChecks;
  const pourDayRecords = merged.pourDayRecords;
  const certifications = merged.certifications;

  const bodyHtml = `
    <div class="pre-brief">
      ${entry.holdPoint ? `<div class="pre-brief__hold-banner"><strong>Hold point</strong> — work must stop until BT (or geotech RPEQ) signs off.</div>` : ''}

      <div class="pre-brief__meta">
        ${entry.level     ? `<span class="pill pill--level">${escapeHtml(entry.level)}</span>` : ''}
        ${entry.stage     ? `<span class="pill">${escapeHtml(entry.stage)}</span>` : ''}
        ${entry.building && entry.building !== 'Main' ? `<span class="pill">${escapeHtml(entry.building)}</span>` : ''}
      </div>

      ${entry.rationale ? `
        <div class="pre-brief__section">
          <div class="pre-brief__section-title">Why</div>
          <p>${escapeHtml(entry.rationale)}</p>
        </div>
      ` : ''}

      ${cover || grade ? `
        <div class="pre-brief__section">
          <div class="pre-brief__section-title">Spec for this element ${ELEMENT ? `(<code>${escapeHtml(ELEMENT)}</code>)` : ''}</div>
          <table class="pre-brief__spec-table">
            <thead><tr><th></th><th>Bottom</th><th>Top</th><th>Sides</th><th>Grade</th></tr></thead>
            <tbody>
              <tr>
                <td><strong>Cover (mm)</strong></td>
                <td>${cover?.bottom ?? '—'}</td>
                <td>${cover?.top    ?? '—'}</td>
                <td>${cover?.sides  ?? '—'}</td>
                <td>${grade ? 'N' + grade : '—'}</td>
              </tr>
            </tbody>
          </table>
        </div>
      ` : ''}

      ${siteReadiness.length ? `
        <div class="pre-brief__section">
          <div class="pre-brief__section-title">Site readiness check (before you drive out)</div>
          <ul class="pre-brief__check">
            ${siteReadiness.map((q, i) => `
              <li>
                <label>
                  <input type="checkbox" data-readiness-idx="${i}">
                  ${escapeHtml(q.text)}
                  ${q.source === 'project' ? `<span class="pre-brief__src pre-brief__src--project" title="Project-specific (from Cowork)">project</span>` : ''}
                </label>
              </li>
            `).join('')}
          </ul>
          <div class="muted small">If any are 'no', consider rescheduling — you'll find them not done on site.</div>
        </div>
      ` : ''}

      ${checklist.length ? `
        <div class="pre-brief__section">
          <div class="pre-brief__section-title">What you'll be checking on site (${checklist.length} items)</div>
          <ul class="pre-brief__checklist">
            ${checklist.map((c) => renderChecklistItem(c)).join('')}
          </ul>
          <div class="muted small">Items marked <em class="pre-brief__src pre-brief__src--project">project</em> are tailored from Cowork; the rest are BT standard for this inspection type.</div>
        </div>
      ` : ''}

      ${pourDayRecords.length ? `
        <div class="pre-brief__section">
          <div class="pre-brief__section-title">Pour-day records to capture</div>
          <ul class="pre-brief__checklist">
            ${pourDayRecords.map((r) => `
              <li>
                ${escapeHtml(r.text)}
                ${r.captureType ? `<span class="pre-brief__capture pre-brief__capture--${escapeHtml(r.captureType)}">${escapeHtml(r.captureType)}</span>` : ''}
                ${r.source === 'project' ? `<span class="pre-brief__src pre-brief__src--project">project</span>` : ''}
              </li>
            `).join('')}
          </ul>
        </div>
      ` : ''}

      ${certifications.length ? `
        <div class="pre-brief__section">
          <div class="pre-brief__section-title">Certifications expected from contractor</div>
          <ul class="muted small">
            ${certifications.map((c) => `
              <li>
                ${escapeHtml(c.text)}
                ${c.providedBy ? `<span class="muted small"> — ${escapeHtml(c.providedBy)}</span>` : ''}
                ${c.formType ? ` <code>${escapeHtml(c.formType)}</code>` : ''}
                ${c.noteRef ? ` <code>${escapeHtml(c.noteRef)}</code>` : ''}
              </li>
            `).join('')}
          </ul>
          <div class="muted small">Track these in the Form 12 register so nothing is missed at handover.</div>
        </div>
      ` : ''}

      ${drawings.length ? `
        <div class="pre-brief__section">
          <div class="pre-brief__section-title">Drawings to take (${drawings.length})</div>
          <ul class="pre-brief__drawings">
            ${drawings.map((d) => `
              <li><code>${escapeHtml(d.sheetNumber || '')}</code> ${escapeHtml(d.reason || '')}</li>
            `).join('')}
          </ul>
          <div class="muted small">All ${drawings.length} drawing${drawings.length === 1 ? '' : 's'} will be attached to the inspection.</div>
        </div>
      ` : ''}

      ${excludedFromBT.length ? `
        <div class="pre-brief__section pre-brief__section--excluded">
          <div class="pre-brief__section-title">Not in BT scope on this project</div>
          <ul class="muted small">
            ${excludedFromBT.slice(0, 5).map((e) => `
              <li><strong>${escapeHtml(e.element)}</strong>${e.responsibility ? ` — ${escapeHtml(e.responsibility)}` : ''}${e.noteRef ? ` <code>${escapeHtml(e.noteRef)}</code>` : ''}</li>
            `).join('')}
            ${excludedFromBT.length > 5 ? `<li class="muted small">…and ${excludedFromBT.length - 5} more</li>` : ''}
          </ul>
          <div class="muted small">Don't waste time inspecting these — the contractor's other RPEQs certify them.</div>
        </div>
      ` : ''}

      ${entry.scopeNote ? `
        <div class="pre-brief__section pre-brief__section--note">
          <div class="muted small"><em>${escapeHtml(entry.scopeNote)}</em></div>
        </div>
      ` : ''}
    </div>
  `;

  const result = await openModal({
    title: entry.title || entry.type,
    primaryLabel: 'Start inspection',
    cancelLabel:  'Site not ready / cancel',
    bodyHtml,
    onConfirm: () => true
  });

  if (!result) return null;

  // Create the inspection
  try {
    const inspection = await startInspectionFromPlanEntry(project.id, planIndex);
    toast(`Started: ${inspection.inspectionTypeName || entry.title}`, { kind: 'success' });
    go('inspection', inspection.id);
    return inspection;
  } catch (err) {
    toast(err.message || 'Couldn’t start inspection', { kind: 'error', duration: 6000 });
    return null;
  }
}


/* ─────────── helpers ─────────── */

function renderChecklistItem(c) {
  if (typeof c === 'string') {
    return `<li>${escapeHtml(c)}</li>`;
  }
  if (c && typeof c === 'object' && c.text) {
    const refs = [];
    if (c.asClauseRef) refs.push(`<code>${escapeHtml(c.asClauseRef)}</code>`);
    if (c.noteRef)     refs.push(`<code>${escapeHtml(c.noteRef)}</code>`);
    const refHtml = refs.length ? ` <span class="muted small">${refs.join(' · ')}</span>` : '';
    const critHtml = c.critical ? `<span class="pre-brief__crit" title="Critical check">!</span>` : '';
    const srcHtml  = c.source === 'project' ? `<span class="pre-brief__src pre-brief__src--project">project</span>` : '';
    const liClass  = c.critical ? 'pre-brief__item--critical' : '';
    return `<li class="${liClass}">${critHtml}${escapeHtml(c.text)}${refHtml}${srcHtml}</li>`;
  }
  return '';
}

/**
 * Map a plan entry's inspection-type key → the cover-schedule element key
 * to look up cover/grade from generalNotes. Best-effort — returns null if
 * no good mapping (e.g. for steel inspections, no concrete cover applies).
 */
function inferCoverElementFromType(type) {
  switch (type) {
    case 'pad-footing-prepour':
    case 'strip-footing-prepour':       return 'footings';
    case 'raft-footing-prepour':        return 'footings';
    case 'slab-prepour-ground':         return 'slabOnGroundInternal';   // try internal first; UI shows a fallback
    case 'slab-prepour-suspended':      return 'suspendedSlab';
    case 'beam-prepour':                return 'suspendedSlab';
    case 'column-prepour':              return 'columns';
    case 'concrete-wall-prepour':       return 'walls';
    case 'retaining-wall-prepour':      return 'walls';
    case 'stair-prepour':               return 'stairs';
    case 'post-tension-strand':         return 'suspendedSlabPT';
    case 'pile-cfa-install':
    case 'pile-bored-install':          return 'boredPiers';
    default:                            return null;   // steel, blockwork, timber → no concrete cover
  }
}

// Note: per-type site-readiness defaults moved to js/lib/bt-standard-checks.js
// (BT_STANDARD_CHECKS[type].siteReadiness) — single source of truth, also
// consumed by the inspection-detail view.

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
