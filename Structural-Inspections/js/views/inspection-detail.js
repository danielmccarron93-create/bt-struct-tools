/**
 * Inspection detail view.
 *
 * Shows metadata for a single inspection, the primary drawing, the items
 * (pinned observations / defects) recorded against it, and the extent of
 * inspection.
 *
 * Actions here:
 *   - Edit metadata (type, drawing, date, inspector, attendees, weather, notes)
 *   - Mark complete / mark draft (status toggle)
 *   - Delete inspection (cascades items + photos + highlights)
 *   - Mark up the plan (drawing viewer in markup mode)
 *   - Tap an item row to re-open its edit panel
 */

import {
  getInspection, updateInspection, deleteInspection,
  getProject, getDrawing, getDrawingBlob,
  listInspectionTypes, listDrawingsForProject,
  listItemsForInspection, listPhotosForItem,
  listHighlightsForInspection,
  setInspectionGeneralComments,
  saveReport, listReportsForInspection, getReport, deleteReport
} from '../db.js';
import { openModal, confirmDialog } from '../components/modal.js';
import { toast } from '../components/toast.js';
import { go } from '../router.js';
import { openItemPanel } from '../components/item-panel.js';
import { openGeneralCommentPicker } from '../components/comment-library-picker.js';
import { createUrlRegistry } from '../lib/photos.js';
import { renderThumbnail as renderPdfThumb } from '../lib/pdf.js';
import { generateReport } from '../lib/report.js';
import { shareReport, openPdfPreview, canShareFiles } from '../lib/share.js';
import { buildExportBundle } from '../lib/export-bundle.js';
import { getStandardChecksForType } from '../lib/bt-standard-checks.js';

export async function render(root, params) {
  const id = Number(params[0]);
  if (!id) { go('inspections'); return; }

  const inspection = await getInspection(id);
  if (!inspection) {
    root.innerHTML = `
      <div class="card"><h2>Inspection not found</h2>
        <p class="muted"><a href="#/inspections">Back to inspections</a></p></div>`;
    return;
  }

  const [project, drawing, items, highlights, reports] = await Promise.all([
    getProject(inspection.projectId),
    inspection.primaryDrawingId ? getDrawing(inspection.primaryDrawingId) : Promise.resolve(null),
    listItemsForInspection(inspection.id),
    listHighlightsForInspection(inspection.id),
    listReportsForInspection(inspection.id)
  ]);

  // Photos for item thumbnails — fetch in parallel.
  const firstPhotos = await Promise.all(items.map(async (it) => {
    const ps = await listPhotosForItem(it.id);
    return ps[0] || null;
  }));

  // URL registry for photo thumbnails; revoked on next render pass.
  const urlRegistry = createUrlRegistry();

  root.innerHTML = `
    <nav class="breadcrumb">
      <a href="#/inspections">&lsaquo; Inspections</a>
      ${project ? `· <a href="#/project/${project.id}">${escapeHtml(project.jobNumber)} ${escapeHtml(project.name)}</a>` : ''}
    </nav>

    <header class="inspection-header">
      <div class="inspection-header__category muted">${escapeHtml(inspection.inspectionTypeCategory || '')}</div>
      <h1>${escapeHtml(inspection.inspectionTypeName || '')}</h1>
      <div class="inspection-header__meta">
        <span class="badge ${inspection.status === 'complete' ? 'badge--complete' : 'badge--draft'}">
          ${escapeHtml(inspection.status || 'draft')}
        </span>
        <span class="muted">${formatDate(inspection.date)}</span>
        ${inspection.inspectorName ? `<span class="muted">&middot; ${escapeHtml(inspection.inspectorName)}</span>` : ''}
      </div>
    </header>

    ${renderPlanContextCard(inspection)}

    <section class="project-section">
      <div class="section-heading"><h2>Plan</h2></div>
      ${drawing ? renderPlanCard(drawing) : `
        <div class="empty card">
          <p class="muted">Primary drawing not found — may have been deleted.</p>
        </div>
      `}
    </section>

    <section class="project-section">
      <div class="section-heading"><h2>Details</h2></div>
      <dl class="detail-grid">
        ${detailRow('Project', project ? `${escapeHtml(project.jobNumber)} — ${escapeHtml(project.name)}` : '—')}
        ${detailRow('Site', project?.siteAddress)}
        ${detailRow('Client', project?.client)}
        ${detailRow('Inspector', inspection.inspectorName)}
        ${detailRow('Attendees', inspection.attendees)}
        ${detailRow('Weather', inspection.weather)}
        ${detailRow('Notes', inspection.notes, true)}
      </dl>
    </section>

    <section class="project-section">
      <div class="section-heading">
        <h2>Items <span class="badge badge--count">${items.length}</span></h2>
        <div class="cluster">
          ${drawing ? `<button class="btn btn--primary btn--sm" id="btn-markup-items">
            <span aria-hidden="true">+</span> Mark up plan
          </button>` : ''}
        </div>
      </div>
      ${items.length === 0
        ? renderEmptyItems(!!drawing)
        : renderItemsList(items, firstPhotos, urlRegistry)}
    </section>

    <section class="project-section">
      <div class="section-heading">
        <h2>Extent of inspection <span class="badge badge--count">${highlights.length}</span></h2>
      </div>
      ${renderExtentSummary(highlights.length, !!drawing)}
    </section>

    <section class="project-section">
      <div class="section-heading">
        <h2>General comments <span class="badge badge--count">${(inspection.generalComments || []).length}</span></h2>
        <div class="cluster">
          <button class="btn btn--secondary btn--sm" id="btn-add-general-library">
            <span aria-hidden="true">📚</span> Add from library
          </button>
          <button class="btn btn--secondary btn--sm" id="btn-add-general-custom">
            <span aria-hidden="true">+</span> Add custom
          </button>
        </div>
      </div>
      ${renderGeneralComments(inspection.generalComments || [])}
    </section>

    <section class="project-section">
      <div class="section-heading">
        <h2>Reports <span class="badge badge--count">${reports.length}</span></h2>
        <div class="cluster">
          <button class="btn btn--primary btn--sm" id="btn-generate-report">
            <span aria-hidden="true">📄</span> Generate report
          </button>
          <button class="btn btn--secondary btn--sm" id="btn-export-bundle"
            title="Build a ZIP in the OneDrive folder layout — Report.pdf + JSON + Photos/">
            <span aria-hidden="true">📦</span> Export bundle
          </button>
        </div>
      </div>
      ${renderReportsList(reports)}
    </section>

    <section class="project-section">
      <div class="section-heading"><h2>Actions</h2></div>
      <div class="cluster">
        <button class="btn btn--secondary" id="btn-edit-inspection">Edit details</button>
        <button class="btn btn--secondary" id="btn-toggle-status">
          ${inspection.status === 'complete' ? 'Mark draft' : 'Mark complete'}
        </button>
        <button class="btn btn--ghost" id="btn-delete-inspection" style="color: var(--color-danger);">
          Delete inspection
        </button>
      </div>
    </section>
  `;

  root.querySelector('#btn-edit-inspection').addEventListener('click', () => onEdit(inspection));
  root.querySelector('#btn-toggle-status').addEventListener('click', () => onToggleStatus(inspection));
  root.querySelector('#btn-delete-inspection').addEventListener('click', () => onDelete(inspection));
  root.querySelector('#btn-generate-report').addEventListener('click', () => onGenerateReport(inspection));
  root.querySelector('#btn-export-bundle').addEventListener('click', () => onExportBundle(inspection));

  const openPlan = root.querySelector('#btn-open-plan');
  if (openPlan) {
    openPlan.addEventListener('click', () => go('markup', inspection.id));
  }

  const markupItems = root.querySelector('#btn-markup-items');
  if (markupItems) {
    markupItems.addEventListener('click', () => go('markup', inspection.id));
  }

  // Kick off the plan-card thumbnail in the background — don't block the first paint.
  if (drawing) {
    renderPlanThumbAsync(drawing).catch((err) =>
      console.warn('Plan thumbnail failed:', err));
  }

  // Delegate: clicking an item row opens its edit panel
  const itemList = root.querySelector('.item-list');
  if (itemList) {
    itemList.addEventListener('click', async (e) => {
      const row = e.target.closest('[data-item-id]');
      if (!row) return;
      const itemId = Number(row.dataset.itemId);
      const result = await openItemPanel({ itemId });
      if (result) {
        // Re-render so thumbnails / counts / status reflect any change
        go('inspection', inspection.id);
      }
    });
  }

  // General comments — library + custom + delete
  root.querySelector('#btn-add-general-library').addEventListener('click', () =>
    onAddGeneralFromLibrary(inspection));
  root.querySelector('#btn-add-general-custom').addEventListener('click', () =>
    onAddGeneralCustom(inspection));

  const genList = root.querySelector('.general-comment-list');
  if (genList) {
    genList.addEventListener('click', async (e) => {
      const delBtn = e.target.closest('[data-gen-delete]');
      if (!delBtn) return;
      const idx = Number(delBtn.dataset.genDelete);
      const next = (inspection.generalComments || []).slice();
      next.splice(idx, 1);
      await setInspectionGeneralComments(inspection.id, next);
      toast('General comment removed', { kind: 'info' });
      go('inspection', inspection.id);
    });
  }

  // Reports list — Open / Share / Delete via delegation
  const reportList = root.querySelector('.report-list');
  if (reportList) {
    reportList.addEventListener('click', (e) => {
      const row = e.target.closest('[data-report-id]');
      if (!row) return;
      const id = Number(row.dataset.reportId);

      if (e.target.closest('[data-report-open]'))   { onOpenReport(id); return; }
      if (e.target.closest('[data-report-share]'))  { onShareReport(id, inspection, project); return; }
      if (e.target.closest('[data-report-delete]')) { onDeleteReport(id, inspection); return; }

      // Clicking the row itself (outside any action) → open
      onOpenReport(id);
    });
  }
}

async function onAddGeneralFromLibrary(inspection) {
  const existingTexts = (inspection.generalComments || []).map((c) => c.text);
  const picks = await openGeneralCommentPicker({
    inspectionTypeKey: inspection.inspectionTypeKey,
    excludeTexts: existingTexts
  });
  if (!picks || picks.length === 0) return;
  const now = new Date().toISOString();
  const additions = picks.map((p) => ({
    text:       p.text,
    libraryKey: null,           // generalComments table doesn't carry keys in the v1 seed
    addedAt:    now
  }));
  const next = (inspection.generalComments || []).concat(additions);
  await setInspectionGeneralComments(inspection.id, next);
  toast(`Added ${picks.length} general comment${picks.length === 1 ? '' : 's'}`, { kind: 'success' });
  go('inspection', inspection.id);
}

async function onAddGeneralCustom(inspection) {
  const result = await openModal({
    title: 'Add custom general comment',
    primaryLabel: 'Add',
    bodyHtml: `
      <div class="form">
        <div class="form-field">
          <label for="gc-text">Comment</label>
          <textarea id="gc-text" name="text" rows="4"
            placeholder="e.g. Access to the south-east corner was limited by formwork."></textarea>
        </div>
      </div>
    `,
    onConfirm: (dialog) => {
      const text = dialog.querySelector('[name=text]').value.trim();
      if (!text) throw new Error('Enter some text first');
      return text;
    }
  });
  if (!result) return;
  const next = (inspection.generalComments || []).concat([{
    text:       result,
    libraryKey: null,
    addedAt:    new Date().toISOString()
  }]);
  await setInspectionGeneralComments(inspection.id, next);
  toast('General comment added', { kind: 'success' });
  go('inspection', inspection.id);
}

/* --------------------------------------------------------------------------
   Rendering helpers
   -------------------------------------------------------------------------- */

/**
 * Plan-context card — only renders when this inspection was started from a
 * project plan entry (fromPlanIndex set during startInspectionFromPlanEntry).
 * Shows the senior-engineer rationale and the expectedChecklist so the
 * engineer arrives on site already knowing what they're checking.
 */
function renderPlanContextCard(inspection) {
  // Pull BT standard checks for this inspection type — these surface even
  // when the inspection wasn't started from a plan entry (e.g. ad-hoc).
  const typeKey = inspection.inspectionTypeKey || (Array.isArray(inspection.types) && inspection.types[0]) || null;
  const standard = typeKey ? getStandardChecksForType(typeKey) : null;

  const projectChecklist = Array.isArray(inspection.expectedChecklist) ? inspection.expectedChecklist : [];

  const hasPlanContent = inspection.fromPlanIndex != null && (inspection.rationale || projectChecklist.length || inspection.holdPoint);
  const hasStandardContent = standard && (
    standard.onSiteChecks.length || standard.pourDayRecords.length || standard.certifications.length
  );
  if (!hasPlanContent && !hasStandardContent) return '';

  const meta = [
    inspection.level     ? escapeHtml(inspection.level) : '',
    inspection.building && inspection.building !== 'Main' ? escapeHtml(inspection.building) : '',
    inspection.stage     ? escapeHtml(inspection.stage) : ''
  ].filter(Boolean).join(' · ');

  // De-dupe project + standard on-site checks (project text wins)
  const projectChecksTexts = new Set(projectChecklist.map((c) => {
    const t = (typeof c === 'string') ? c : (c && c.text) || '';
    return t.toLowerCase();
  }));
  const standardOnly = (standard?.onSiteChecks || []).filter(
    (c) => !projectChecksTexts.has((c.text || '').toLowerCase())
  );

  return `
    <section class="project-section">
      <details class="plan-context-card card" open>
        <summary>
          <span class="plan-context-card__title">${hasPlanContent ? 'From the project plan' : 'BT standard checks for this inspection'}</span>
          ${inspection.holdPoint ? '<span class="badge badge--hold">hold point</span>' : ''}
          ${meta ? `<span class="muted small">${meta}</span>` : ''}
        </summary>
        <div class="plan-context-card__body">
          ${inspection.rationale ? `
            <p class="plan-context-card__rationale">${escapeHtml(inspection.rationale)}</p>
          ` : ''}

          ${projectChecklist.length ? `
            <div class="plan-context-card__checklist">
              <div class="plan-context-card__section-title muted small">Project-specific checks (from drawings + General Notes)</div>
              <ul>
                ${projectChecklist.map((c) => renderChecklistRow(c)).join('')}
              </ul>
            </div>
          ` : ''}

          ${standardOnly.length ? `
            <div class="plan-context-card__checklist">
              <div class="plan-context-card__section-title muted small">BT standard checks for ${escapeHtml(standard.name)}</div>
              <ul>
                ${standardOnly.map((c) => renderChecklistRow(c)).join('')}
              </ul>
            </div>
          ` : ''}

          ${standard?.pourDayRecords?.length ? `
            <div class="plan-context-card__checklist">
              <div class="plan-context-card__section-title muted small">Pour-day records to capture</div>
              <ul>
                ${standard.pourDayRecords.map((r) => `
                  <li>${escapeHtml(r.text)}${r.captureType ? ` <span class="muted small">[${escapeHtml(r.captureType)}]</span>` : ''}</li>
                `).join('')}
              </ul>
            </div>
          ` : ''}

          ${standard?.certifications?.length ? `
            <div class="plan-context-card__checklist">
              <div class="plan-context-card__section-title muted small">Certifications expected from contractor</div>
              <ul>
                ${standard.certifications.map((c) => `
                  <li>
                    ${escapeHtml(c.text)}
                    ${c.providedBy ? ` <span class="muted small">— ${escapeHtml(c.providedBy)}</span>` : ''}
                    ${c.formType ? ` <code>${escapeHtml(c.formType)}</code>` : ''}
                    ${c.noteRef ? ` <code>${escapeHtml(c.noteRef)}</code>` : ''}
                  </li>
                `).join('')}
              </ul>
            </div>
          ` : ''}

          <div class="muted small">Drop pins on the plan to record observations / defects against any of these.</div>
        </div>
      </details>
    </section>
  `;
}

function renderChecklistRow(c) {
  if (typeof c === 'string') return `<li>${escapeHtml(c)}</li>`;
  if (!c || typeof c !== 'object') return '';
  const refs = [];
  if (c.asClauseRef) refs.push(`<code>${escapeHtml(c.asClauseRef)}</code>`);
  if (c.noteRef)     refs.push(`<code>${escapeHtml(c.noteRef)}</code>`);
  const refHtml = refs.length ? ` <span class="muted small">${refs.join(' · ')}</span>` : '';
  const critHtml = c.critical ? `<span class="check-crit" title="Critical check">!</span>` : '';
  return `<li class="${c.critical ? 'check-row--critical' : ''}">${critHtml}${escapeHtml(c.text || '')}${refHtml}</li>`;
}

function renderPlanCard(drawing) {
  return `
    <div class="plan-card">
      <div class="plan-card__thumb" id="plan-thumb">
        <div class="drawing-card__thumb-placeholder">PDF</div>
      </div>
      <div class="plan-card__body">
        <div class="plan-card__sheet">
          ${escapeHtml(drawing.sheetNumber || '(no sheet)')}
          ${drawing.revision ? `<span class="drawing-card__rev">Rev ${escapeHtml(drawing.revision)}</span>` : ''}
        </div>
        <div class="plan-card__desc">${escapeHtml(drawing.description || drawing.filename || '—')}</div>
        <div class="plan-card__meta muted">
          ${drawing.calibration
            ? `<span class="drawing-card__cal drawing-card__cal--yes">Calibrated</span>`
            : `<span class="drawing-card__cal drawing-card__cal--no">Not calibrated — calibrate before marking up</span>`}
        </div>
      </div>
      <div class="plan-card__action">
        <button class="btn btn--primary" id="btn-open-plan">Open plan</button>
      </div>
    </div>
  `;
}

function detailRow(label, value, wide = false) {
  const v = value && String(value).trim() ? escapeHtml(value) : '<span class="muted">—</span>';
  return `
    <div class="detail-row ${wide ? 'detail-row--wide' : ''}">
      <dt>${escapeHtml(label)}</dt>
      <dd>${v}</dd>
    </div>
  `;
}

function renderEmptyItems(hasDrawing) {
  if (!hasDrawing) {
    return `
      <div class="empty card">
        <p class="muted">No primary drawing attached — re-attach one before marking up.</p>
      </div>
    `;
  }
  return `
    <div class="empty card">
      <p class="muted">No items pinned yet. Open the plan in markup mode, tap to drop numbered pins, and attach photos + comments to each one. Each item becomes a row in the final report's rectification legend.</p>
    </div>
  `;
}

function renderItemsList(items, firstPhotos, urlRegistry) {
  return `
    <ul class="item-list" role="list">
      ${items.map((it, idx) => {
        const thumb = firstPhotos[idx];
        const thumbHtml = thumb
          ? `<img src="${urlRegistry.urlFor(thumb.blob)}" alt="" />`
          : `<span class="item-card__thumb-placeholder" aria-hidden="true">—</span>`;
        const preview = (it.comment || '').trim();
        const previewHtml = preview
          ? escapeHtml(preview.length > 140 ? preview.slice(0, 137) + '…' : preview)
          : '<span class="muted">No comment yet</span>';
        const sev = it.severity || 'observation';
        const sevLabel = sev === 'holdpoint' ? 'hold point' : sev;
        return `
          <li class="item-card" data-item-id="${it.id}" tabindex="0" role="button"
              onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}">
            <div class="item-card__num item-card__num--${escapeAttr(sev)}">
              ${it.itemNumber}
            </div>
            <div class="item-card__thumb">${thumbHtml}</div>
            <div class="item-card__body">
              <div class="item-card__preview">${previewHtml}</div>
              <div class="item-card__meta muted">
                <span class="pill pill--${escapeAttr(sev)}">${escapeHtml(sevLabel)}</span>
                &middot;
                <span class="pill pill--status-${escapeAttr(it.status || 'open')}">${escapeHtml(it.status || 'open')}</span>
                ${it.gridRef ? ' &middot; <span class="muted">' + escapeHtml(it.gridRef) + '</span>' : ''}
              </div>
            </div>
          </li>
        `;
      }).join('')}
    </ul>
  `;
}

function renderReportsList(reports) {
  if (!reports || reports.length === 0) {
    return `
      <div class="empty card">
        <p class="muted">No reports generated yet. Click <strong>Generate report</strong> to build a 3-page PDF (cover, marked plan, item schedule) saved here for later sharing.</p>
      </div>
    `;
  }
  const canShare = canShareFiles();
  const shareLabel = canShare ? 'Share' : 'Email';
  return `
    <ul class="report-list" role="list">
      ${reports.map((r) => `
        <li class="report-row" data-report-id="${r.id}" tabindex="0" role="button"
            onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}">
          <div class="report-row__icon" aria-hidden="true">📄</div>
          <div class="report-row__main">
            <div class="report-row__filename">${escapeHtml(r.filename)}</div>
            <div class="report-row__meta muted">
              ${formatDateTime(r.generatedAt)}
              ${r.pageCount ? ` &middot; ${r.pageCount} page${r.pageCount === 1 ? '' : 's'}` : ''}
              &middot; ${formatBytes(r.sizeBytes || 0)}
              ${r.generatedBy ? ' &middot; ' + escapeHtml(r.generatedBy) : ''}
            </div>
          </div>
          <div class="report-row__actions">
            <button class="icon-btn" data-report-open aria-label="Open report" title="Open report">
              <svg viewBox="0 0 24 24"><path d="M15 3h6v6M10 14L21 3M21 14v5a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5"/></svg>
            </button>
            <button class="btn btn--secondary btn--sm" data-report-share>
              <span aria-hidden="true">↗</span> ${escapeHtml(shareLabel)}
            </button>
            <button class="icon-btn" data-report-delete aria-label="Delete report" style="color: var(--color-danger);">
              <svg viewBox="0 0 24 24"><path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/></svg>
            </button>
          </div>
        </li>
      `).join('')}
    </ul>
  `;
}

function formatDateTime(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.valueOf())) return String(iso);
  return d.toLocaleString('en-AU', {
    day: '2-digit', month: 'short', year: 'numeric',
    hour: '2-digit', minute: '2-digit'
  });
}

function formatBytes(bytes) {
  if (!bytes) return '0 KB';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function renderGeneralComments(list) {
  if (!list || list.length === 0) {
    return `
      <div class="empty card">
        <p class="muted">No general comments added. Library picks are seeded for this inspection type — use <strong>Add from library</strong> or write a custom one.</p>
      </div>
    `;
  }
  return `
    <ul class="general-comment-list" role="list">
      ${list.map((c, idx) => `
        <li class="general-comment-row">
          <div class="general-comment-row__text">${escapeHtml(c.text)}</div>
          <button type="button" class="icon-btn general-comment-row__del"
                  aria-label="Remove general comment" data-gen-delete="${idx}">
            <svg viewBox="0 0 24 24"><path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/></svg>
          </button>
        </li>
      `).join('')}
    </ul>
  `;
}

function renderExtentSummary(count, hasDrawing) {
  if (!hasDrawing) {
    return `<div class="empty card">
      <p class="muted">Attach a primary drawing to record extent of inspection.</p>
    </div>`;
  }
  if (count === 0) {
    return `<div class="empty card">
      <p class="muted">No area marked yet. Open the plan and use <strong>Extent</strong> to drag a yellow highlight over the area you physically inspected — this goes on the marked plan in the report.</p>
    </div>`;
  }
  return `<div class="card">
    <p>${count} area${count === 1 ? '' : 's'} marked. Open the plan to review, add, or remove highlights.</p>
  </div>`;
}

/* --------------------------------------------------------------------------
   Handlers
   -------------------------------------------------------------------------- */

async function onEdit(inspection) {
  const [types, drawings] = await Promise.all([
    listInspectionTypes(),
    listDrawingsForProject(inspection.projectId)
  ]);

  const grouped = new Map();
  for (const t of types) {
    if (!grouped.has(t.category)) grouped.set(t.category, []);
    grouped.get(t.category).push(t);
  }

  const bodyHtml = `
    <div class="form">
      <div class="form-field">
        <label for="e-type">Inspection type <span class="req">*</span></label>
        <select id="e-type" name="inspectionTypeKey" required>
          ${Array.from(grouped.entries()).map(([cat, items]) => `
            <optgroup label="${escapeHtml(cat)}">
              ${items.map((t) => `
                <option value="${escapeAttr(t.key)}" ${t.key === inspection.inspectionTypeKey ? 'selected' : ''}>
                  ${escapeHtml(t.name)}
                </option>
              `).join('')}
            </optgroup>
          `).join('')}
        </select>
      </div>

      <div class="form-field">
        <label for="e-drawing">Primary drawing <span class="req">*</span></label>
        <select id="e-drawing" name="primaryDrawingId" required>
          ${drawings.map((d) => {
            const label = d.sheetNumber
              ? `${d.sheetNumber}${d.revision ? ` (Rev ${d.revision})` : ''}${d.description ? ' — ' + d.description : ''}`
              : (d.filename || 'Untitled drawing');
            return `<option value="${d.id}" ${d.id === inspection.primaryDrawingId ? 'selected' : ''}>${escapeHtml(label)}</option>`;
          }).join('')}
        </select>
      </div>

      <div class="form-field">
        <label for="e-date">Date <span class="req">*</span></label>
        <input id="e-date" name="date" type="date" required value="${escapeAttr(inspection.date || '')}"/>
      </div>

      <div class="form-field">
        <label for="e-inspector">Inspector</label>
        <input id="e-inspector" name="inspectorName" type="text" value="${escapeAttr(inspection.inspectorName || '')}"/>
      </div>

      <div class="form-field">
        <label for="e-attendees">Attendees</label>
        <input id="e-attendees" name="attendees" type="text" value="${escapeAttr(inspection.attendees || '')}"/>
      </div>

      <div class="form-field">
        <label for="e-weather">Weather</label>
        <input id="e-weather" name="weather" type="text" value="${escapeAttr(inspection.weather || '')}"/>
      </div>

      <div class="form-field">
        <label for="e-notes">Notes</label>
        <textarea id="e-notes" name="notes" rows="3">${escapeHtml(inspection.notes || '')}</textarea>
      </div>
    </div>
  `;

  const result = await openModal({
    title: 'Edit inspection',
    primaryLabel: 'Save changes',
    bodyHtml,
    onConfirm: async (dialog) => {
      return updateInspection(inspection.id, {
        inspectionTypeKey: dialog.querySelector('[name=inspectionTypeKey]').value,
        primaryDrawingId:  dialog.querySelector('[name=primaryDrawingId]').value,
        date:              dialog.querySelector('[name=date]').value,
        inspectorName:     dialog.querySelector('[name=inspectorName]').value,
        attendees:         dialog.querySelector('[name=attendees]').value,
        weather:           dialog.querySelector('[name=weather]').value,
        notes:             dialog.querySelector('[name=notes]').value
      });
    }
  });

  if (result) {
    toast('Inspection updated', { kind: 'success' });
    go('inspection', inspection.id);
  }
}

async function onToggleStatus(inspection) {
  const next = inspection.status === 'complete' ? 'draft' : 'complete';
  await updateInspection(inspection.id, { status: next });
  toast(`Inspection marked ${next}`, { kind: 'success' });
  go('inspection', inspection.id);
}

async function onGenerateReport(inspection) {
  const btn = document.getElementById('btn-generate-report');
  const originalLabel = btn ? btn.innerHTML : '';
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span aria-hidden="true">⏳</span> Generating…';
  }
  // Immediate user feedback — a landscape-plan render + photo re-compress can
  // take a few seconds. Without this the button state change alone is too
  // subtle on desktop and invisible when the button is off-screen on mobile.
  toast('Generating report — this usually takes a few seconds.', { kind: 'info' });
  try {
    const { blob, filename, pageCount } = await generateReport(inspection.id, {
      onProgress: (msg) => {
        if (btn) btn.innerHTML = `<span aria-hidden="true">⏳</span> ${escapeHtml(msg)}`;
      }
    });

    // Persist under the inspection so it shows up in the Reports section.
    await saveReport(inspection.id, {
      filename,
      pdfBlob:     blob,
      sizeBytes:   blob.size,
      pageCount,
      generatedBy: inspection.inspectorName || ''
    });

    // Open a preview tab so the engineer can eyeball it. Do NOT auto-download
    // — the Reports list row has its own Share / Open actions.
    openPdfPreview(blob);
    // Revoke the preview URL after the user has had time to look. openPdfPreview
    // hands the URL to the browser synchronously, so a generous timeout is fine.
    // (openPdfPreview doesn't return the URL directly; we skip the revoke here
    // and rely on tab close + GC. Minor memory cost acceptable for v1.)

    toast('Report generated and saved', { kind: 'success' });
    go('inspection', inspection.id);
  } catch (err) {
    console.error('Report generation failed:', err);
    toast(`Couldn\u2019t generate report: ${err.message || err}`, { kind: 'error' });
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = originalLabel;
    }
  }
}

async function onExportBundle(inspection) {
  const btn = document.getElementById('btn-export-bundle');
  const original = btn ? btn.innerHTML : '';
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span aria-hidden="true">\u23f3</span> Exporting\u2026';
  }
  toast('Building export bundle\u2026', { kind: 'info' });
  try {
    const { blob, filename } = await buildExportBundle(inspection.id, {
      onProgress: (msg) => {
        if (btn) btn.innerHTML = `<span aria-hidden="true">\u23f3</span> ${escapeHtml(msg)}`;
      }
    });
    // Trigger a download — the engineer drops the zip into OneDrive.
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 0);
    toast('Export bundle downloaded \u2014 drop it into OneDrive / SharePoint.', { kind: 'success' });
  } catch (err) {
    console.error('Export bundle failed:', err);
    toast(`Couldn\u2019t export: ${err.message || err}`, { kind: 'error' });
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = original;
    }
  }
  // Re-render so the new persisted report appears in the list.
  go('inspection', inspection.id);
}

async function onOpenReport(reportId) {
  const r = await getReport(reportId);
  if (!r) { toast('Report not found', { kind: 'error' }); return; }
  openPdfPreview(r.pdfBlob);
}

async function onShareReport(reportId, inspection, project) {
  const r = await getReport(reportId);
  if (!r) { toast('Report not found', { kind: 'error' }); return; }
  const result = await shareReport(r.pdfBlob, r.filename, { inspection, project });
  if (result.method === 'share') {
    toast('Shared', { kind: 'success' });
  } else if (result.method === 'download+mailto') {
    toast('Report downloaded — attach it to the email draft', { kind: 'info' });
  } else if (result.method === 'error') {
    toast('Couldn\u2019t share the report', { kind: 'error' });
  }
  // 'cancelled' is silent — user picked cancel on the share sheet
}

async function onDeleteReport(reportId, inspection) {
  const ok = await confirmDialog({
    title: 'Delete this report version?',
    message: 'This removes the saved PDF. Other versions of the report (if any) are kept. You can regenerate at any time.',
    confirmLabel: 'Delete',
    danger: true
  });
  if (!ok) return;
  await deleteReport(reportId);
  toast('Report deleted', { kind: 'info' });
  go('inspection', inspection.id);
}

async function onDelete(inspection) {
  const ok = await confirmDialog({
    title: 'Delete inspection?',
    message: 'This will permanently remove the inspection and any items / photos recorded against it. This cannot be undone.',
    confirmLabel: 'Delete',
    danger: true
  });
  if (!ok) return;
  await deleteInspection(inspection.id);
  toast('Inspection deleted', { kind: 'info' });
  go('project', inspection.projectId);
}

/* --------------------------------------------------------------------------
   Utilities
   -------------------------------------------------------------------------- */

function formatDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  return d.toLocaleDateString('en-AU', { day: '2-digit', month: 'short', year: 'numeric' });
}

async function renderPlanThumbAsync(drawing) {
  const holder = document.getElementById('plan-thumb');
  if (!holder) return;
  const blob = await getDrawingBlob(drawing);
  if (!blob) return;
  const dataUrl = await renderPdfThumb(blob, 360, drawing.pageNumber || 1);
  // Only paint if the holder is still mounted (user may have navigated away).
  if (document.getElementById('plan-thumb') === holder) {
    holder.innerHTML = `<img alt="" src="${dataUrl}"/>`;
  }
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
function escapeAttr(s) { return escapeHtml(s).replace(/"/g, '&quot;'); }
