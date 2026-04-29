/**
 * Project detail view (Phase 11 rebuild).
 *
 * For projects imported from a Cowork .btproject bundle this renders:
 *   - Warnings banner (if any)
 *   - Extended project header (issue status, discipline, engineer of record)
 *   - General Notes summary card (concrete cover, bearing, geotech, codes)
 *   - Inspection plan (ordered list, status badges, tap to start)
 *   - Drawings grouped by level → kind, with filter chips
 *   - Inspections (the actual records)
 *
 * For legacy projects (no projectMap) the layout falls back to the original
 * flat drawings + inspections structure so existing data keeps working.
 */

import {
  getProject, updateProject, deleteProject,
  listDrawingsForProject, addDrawingsFromSource, updateDrawing, deleteDrawing,
  getDrawing, getDrawingBlob,
  listInspectionsForProject,
  startInspectionFromPlanEntry, skipPlanEntry, resetPlanEntry,
  getProjectProgress, markForm12Issued, unmarkForm12Issued
} from '../db.js';
import { openModal, confirmDialog } from '../components/modal.js';
import { toast } from '../components/toast.js';
import { go } from '../router.js';
import {
  renderThumbnail,
  extractPageMetadata,
  extractSheetNumberFromFilename,
  extractRevisionFromFilename
} from '../lib/pdf.js';
import { openNewInspection } from './new-inspection.js';
import { openPreInspectionBrief } from '../components/pre-inspection-brief.js';
import { openImportRevision } from '../components/import-revision.js';
import { acknowledgeRevisionChange } from '../db.js';

export async function render(root, params) {
  const id = Number(params[0]);
  if (!id) {
    go('projects');
    return;
  }

  const project = await getProject(id);
  if (!project) {
    root.innerHTML = `
      <div class="card">
        <h2>Project not found</h2>
        <p class="muted">It may have been deleted. <a href="#/projects">Back to projects</a>.</p>
      </div>`;
    return;
  }

  const [drawings, inspections, progress] = await Promise.all([
    listDrawingsForProject(id),
    listInspectionsForProject(id),
    getProjectProgress(id).catch(() => null)   // graceful: legacy projects without a plan still render
  ]);

  const hasProjectMap = !!project.projectMap;
  const generalNotes  = project.projectMap?.generalNotes || null;
  const warnings      = project.projectMap?.warnings || [];
  const plan          = Array.isArray(project.inspectionPlan) ? project.inspectionPlan : [];
  const excludedFromBT = project.excludedFromBT || project.projectMap?.project?.excludedFromBT || [];

  const pendingRevisions = Array.isArray(project.pendingRevisions) ? project.pendingRevisions : [];

  root.innerHTML = `
    <nav class="breadcrumb"><a href="#/projects">&lsaquo; Projects</a></nav>

    ${warnings.length ? renderWarningsBanner(warnings) : ''}
    ${pendingRevisions.length ? renderPendingRevisionsBanner(pendingRevisions, plan) : ''}
    ${renderProjectHeader(project)}
    ${plan.length && progress ? renderProgressStrip(progress, project) : ''}
    ${hasProjectMap && generalNotes ? renderGeneralNotesCard(generalNotes) : ''}
    ${excludedFromBT.length ? renderExcludedFromBT(excludedFromBT) : ''}
    ${plan.length ? renderInspectionPlan(plan, inspections) : ''}

    <section class="project-section">
      <div class="section-heading">
        <h2>Drawings <span class="badge badge--count">${drawings.length}</span></h2>
        <div class="cluster">
          <label class="btn btn--secondary btn--sm" for="input-drawings">
            <span aria-hidden="true">+</span> Upload PDF
          </label>
          <input id="input-drawings" type="file" accept="application/pdf" multiple class="sr-only"/>
        </div>
      </div>
      <div class="upload-progress is-hidden" id="upload-progress" role="status" aria-live="polite"></div>

      ${drawings.length === 0
        ? renderEmptyDrawings()
        : (hasProjectMap ? renderGroupedDrawings(drawings) : renderFlatDrawings(drawings))}
    </section>

    <section class="project-section">
      <div class="section-heading">
        <h2>Inspections <span class="badge badge--count">${inspections.length}</span></h2>
        <div class="cluster">
          <button class="btn btn--primary btn--sm" id="btn-new-inspection">
            <span aria-hidden="true">+</span> New inspection
          </button>
        </div>
      </div>
      ${inspections.length === 0
        ? renderEmptyInspections(drawings.length === 0)
        : renderInspectionList(inspections)}
    </section>
  `;

  /* ---- Wire handlers ---- */
  root.querySelector('#btn-edit-project').addEventListener('click', () => onEditProject(project));
  root.querySelector('#btn-delete-project').addEventListener('click', () => onDeleteProject(project));
  const importRevBtn = root.querySelector('#btn-import-revision');
  if (importRevBtn) importRevBtn.addEventListener('click', () => openImportRevision(project.id));

  // Acknowledge revision-changed badge on a plan entry
  root.querySelectorAll('[data-plan-action="ack-revision"]').forEach((btn) => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const idx = Number(btn.closest('[data-plan-index]').dataset.planIndex);
      await acknowledgeRevisionChange(project.id, idx);
      toast('Revision acknowledged', { kind: 'info' });
      go('project', project.id);
    });
  });
  root.querySelector('#input-drawings').addEventListener('change', (e) => onUploadDrawings(e, project.id));
  root.querySelector('#btn-new-inspection').addEventListener('click', () => onNewInspection(project.id));

  // Plan entry click — tap-to-start (legacy openNewInspection from "+ New" button still works for ad-hoc)
  root.querySelectorAll('[data-plan-index]').forEach((entry) => {
    entry.addEventListener('click', (e) => {
      // Ignore taps on action buttons inside the entry.
      if (e.target.closest('[data-plan-action]')) return;
      const idx = Number(entry.dataset.planIndex);
      const status = entry.dataset.planStatus;
      if (status === 'in-progress' || status === 'done') {
        // Jump to the linked inspection if there is one.
        const inspId = Number(entry.dataset.linkedInspectionId);
        if (inspId) go('inspection', inspId);
      } else {
        onStartPlanEntry(project.id, idx);
      }
    });
  });
  root.querySelectorAll('[data-plan-action="skip"]').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      onSkipPlanEntry(project.id, Number(btn.closest('[data-plan-index]').dataset.planIndex));
    });
  });
  root.querySelectorAll('[data-plan-action="reset"]').forEach((btn) => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      await resetPlanEntry(project.id, Number(btn.closest('[data-plan-index]').dataset.planIndex));
      go('project', project.id);
    });
  });

  // Form 12 issue / unissue
  const issueBtn = root.querySelector('[data-action="form12-issue"]');
  if (issueBtn) {
    issueBtn.addEventListener('click', async () => {
      const ok = await confirmDialog({
        title: 'Issue Form 12 / RPEQ certification?',
        message: 'This marks the project as Form 12 issued (status only — PDF generation comes in v2.3). You can undo this from the same place.',
        confirmLabel: 'Mark issued',
        cancelLabel:  'Cancel'
      });
      if (!ok) return;
      await markForm12Issued(project.id);
      toast('Form 12 marked issued', { kind: 'success' });
      go('project', project.id);
    });
  }
  const unissueBtn = root.querySelector('[data-action="form12-unissue"]');
  if (unissueBtn) {
    unissueBtn.addEventListener('click', async () => {
      const ok = await confirmDialog({
        title: 'Undo Form 12 issuance?',
        message: 'This clears the issued timestamp. The project goes back to "ready to issue" if criteria still met.',
        confirmLabel: 'Undo',
        cancelLabel:  'Cancel'
      });
      if (!ok) return;
      await unmarkForm12Issued(project.id);
      toast('Form 12 issuance cleared', { kind: 'info' });
      go('project', project.id);
    });
  }

  // Drawing filter chips
  const chipBar = root.querySelector('#dwg-filter-chips');
  if (chipBar) {
    chipBar.addEventListener('click', (e) => {
      const chip = e.target.closest('[data-filter]');
      if (!chip) return;
      chipBar.querySelectorAll('[data-filter]').forEach((c) => c.classList.remove('is-active'));
      chip.classList.add('is-active');
      applyDrawingFilter(root, chip.dataset.filter);
    });
  }

  // Inspection list delegation
  const inspList = root.querySelector('.inspection-list');
  if (inspList) {
    inspList.addEventListener('click', (e) => {
      const card = e.target.closest('[data-inspection-id]');
      if (card) go('inspection', card.dataset.inspectionId);
    });
  }
  root.querySelectorAll('[data-action="trigger-upload"]').forEach((btn) => {
    btn.addEventListener('click', () => document.getElementById('input-drawings').click());
  });

  // Drawing list delegation (grouped or flat)
  const drawingContainer = root.querySelector('[data-drawings-root]');
  if (drawingContainer) {
    drawingContainer.addEventListener('click', (e) => {
      const card = e.target.closest('[data-drawing-id]');
      if (!card) return;
      if (e.target.closest('[data-action=edit]')) {
        e.preventDefault(); e.stopPropagation();
        onEditDrawing(Number(card.dataset.drawingId), project.id);
        return;
      }
      if (e.target.closest('[data-action=delete]')) {
        e.preventDefault(); e.stopPropagation();
        onDeleteDrawing(Number(card.dataset.drawingId), project.id);
        return;
      }
      go('drawing', card.dataset.drawingId);
    });
  }

  // Kick off thumbnails in background
  renderThumbnailsSequentially(drawings);
}

/* --------------------------------------------------------------------------
   Section renderers
   -------------------------------------------------------------------------- */

function renderWarningsBanner(warnings) {
  return `
    <div class="warnings-banner" role="alert">
      <details>
        <summary>
          <span class="warnings-banner__count">${warnings.length}</span>
          ${warnings.length === 1 ? 'warning' : 'warnings'} from project setup
        </summary>
        <ul>
          ${warnings.map((w) => `<li>${escapeHtml(w)}</li>`).join('')}
        </ul>
      </details>
    </div>
  `;
}

/**
 * Pending revisions banner — shown after applyRevisionDiff stamps the project.
 * Lists each revision import and how many plan entries still need acknowledgement.
 */
function renderPendingRevisionsBanner(pendingRevisions, plan) {
  if (!pendingRevisions.length) return '';
  const unacknowledged = (plan || []).filter(
    (e) => e.revisionChangedAt && !e.revisionAcknowledgedAt
  ).length;
  return `
    <div class="rev-banner" role="alert">
      <details ${unacknowledged > 0 ? 'open' : ''}>
        <summary>
          <strong>Revision${pendingRevisions.length === 1 ? '' : 's'} imported</strong>
          ${unacknowledged > 0
            ? ` — <span class="rev-banner__pending">${unacknowledged} plan entr${unacknowledged === 1 ? 'y' : 'ies'} need review</span>`
            : ' — all changes acknowledged'}
        </summary>
        <ul>
          ${pendingRevisions.map((r) => `
            <li>
              <strong>${escapeHtml(r.newPdf || 'revision')}</strong>
              <span class="muted small"> · imported ${formatRelativeDate(r.importedAt)}</span>
              <div class="muted small">${escapeHtml(r.summary || '')}</div>
            </li>
          `).join('')}
        </ul>
      </details>
    </div>
  `;
}

function formatRelativeDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return d.toLocaleString('en-AU', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
}

function renderProjectHeader(p) {
  const issue       = p.issueStatus || '';
  const discipline  = p.discipline || '';
  const eor         = p.engineerOfRecord || '';
  const builder     = p.builder || '';
  return `
    <header class="project-header">
      <div class="project-header__top">
        <div class="project-header__job">${escapeHtml(p.jobNumber || '—')}</div>
        ${issue ? `<span class="badge badge--issue">${escapeHtml(issue)}</span>` : ''}
        ${discipline ? `<span class="badge badge--discipline">${escapeHtml(discipline)}</span>` : ''}
      </div>
      <h1>${escapeHtml(p.name)}</h1>
      <dl class="project-header__meta">
        ${p.client      ? `<div><dt>Client</dt><dd>${escapeHtml(p.client)}</dd></div>` : ''}
        ${p.siteAddress ? `<div><dt>Site</dt><dd>${escapeHtml(p.siteAddress)}</dd></div>` : ''}
        ${eor           ? `<div><dt>Engineer of record</dt><dd>${escapeHtml(eor)}</dd></div>` : ''}
        ${builder       ? `<div><dt>Builder</dt><dd>${escapeHtml(builder)}</dd></div>` : ''}
        ${p.notes       ? `<div><dt>Notes</dt><dd>${escapeHtml(p.notes)}</dd></div>` : ''}
      </dl>
      <div class="cluster">
        <button class="btn btn--secondary btn--sm" id="btn-edit-project">Edit project</button>
        <button class="btn btn--ghost btn--sm" id="btn-import-revision" title="Import a revision diff JSON produced by tools/diff_revision.py">Import revision</button>
        <button class="btn btn--ghost btn--sm" id="btn-delete-project" style="color: var(--color-danger);">Delete project</button>
      </div>
    </header>
  `;
}

/**
 * Form 12 / RPEQ progress strip — top of project view.
 * Shows: total inspections, % complete, hold-points outstanding,
 * rectifications outstanding, "Form 12 ready" status, "Issue Form 12" button.
 */
function renderProgressStrip(progress, project) {
  const { plan, rectifications, form12Ready, form12IssuedAt } = progress;
  const pct = plan.total > 0
    ? Math.round((plan.complete + plan.skipped) / plan.total * 100)
    : 0;
  const issued = !!form12IssuedAt;

  let statusBadge, action;
  if (issued) {
    const dt = new Date(form12IssuedAt);
    const dateLabel = dt.toLocaleDateString('en-AU', { day: '2-digit', month: 'short', year: 'numeric' });
    statusBadge = `<span class="badge badge--complete">Form 12 issued · ${dateLabel}</span>`;
    action = `<button class="btn btn--ghost btn--sm" data-action="form12-unissue">Undo</button>`;
  } else if (form12Ready) {
    statusBadge = `<span class="badge badge--ready">Form 12 ready to issue</span>`;
    action = `<button class="btn btn--primary btn--sm" data-action="form12-issue">Issue Form 12</button>`;
  } else {
    const blockers = [];
    if (plan.complete + plan.skipped < plan.total) {
      const left = plan.total - plan.complete - plan.skipped;
      blockers.push(`${left} inspection${left === 1 ? '' : 's'} pending`);
    }
    if (rectifications.outstanding > 0) {
      blockers.push(`${rectifications.outstanding} rectification${rectifications.outstanding === 1 ? '' : 's'} outstanding`);
    }
    statusBadge = `<span class="muted small">Form 12 ready when: ${blockers.join(' · ')}</span>`;
    action = '';
  }

  return `
    <section class="form12-strip">
      <div class="form12-strip__bar">
        <div class="form12-strip__bar-track">
          <div class="form12-strip__bar-fill" style="width: ${pct}%;"></div>
        </div>
        <div class="form12-strip__pct">${pct}%</div>
      </div>
      <div class="form12-strip__stats">
        <div class="form12-stat">
          <span class="form12-stat__num">${plan.complete}<span class="form12-stat__sub">/${plan.total}</span></span>
          <span class="form12-stat__lbl">Complete</span>
        </div>
        <div class="form12-stat">
          <span class="form12-stat__num ${plan.holdsOutstanding > 0 ? 'form12-stat__num--accent' : ''}">${plan.holdsOutstanding}</span>
          <span class="form12-stat__lbl">Hold points open</span>
        </div>
        <div class="form12-stat">
          <span class="form12-stat__num ${rectifications.outstanding > 0 ? 'form12-stat__num--accent' : ''}">${rectifications.outstanding}</span>
          <span class="form12-stat__lbl">Rectifications</span>
        </div>
        <div class="form12-stat">
          <span class="form12-stat__num">${plan.skipped}</span>
          <span class="form12-stat__lbl">Skipped</span>
        </div>
      </div>
      <div class="form12-strip__action">
        ${statusBadge}
        ${action}
      </div>
    </section>
  `;
}

/**
 * Excluded from BT — shows the certified-by-others summary so engineer is
 * aware of what they're NOT inspecting (precast lifting, steel temp props,
 * stud framing, roof safety systems).
 */
function renderExcludedFromBT(excluded) {
  return `
    <section class="excluded-from-bt card">
      <details>
        <summary>
          <strong>Not in BT scope (${excluded.length})</strong>
          <span class="muted small">— certified by other parties; do not inspect</span>
        </summary>
        <ul class="excluded-list">
          ${excluded.map((e) => `
            <li>
              <strong>${escapeHtml(e.element)}</strong>
              ${e.responsibility ? `<div class="muted small">Responsibility: ${escapeHtml(e.responsibility)}</div>` : ''}
              ${e.scope ? `<div class="muted small">Scope: ${escapeHtml(e.scope)}</div>` : ''}
              ${e.noteRef ? `<div class="muted small">Ref: <code>${escapeHtml(e.noteRef)}</code></div>` : ''}
            </li>
          `).join('')}
        </ul>
      </details>
    </section>
  `;
}

function renderGeneralNotesCard(notes) {
  if (!notes || Object.keys(notes).length === 0) return '';

  const codes  = notes.designCodes && notes.designCodes.length
    ? notes.designCodes.join(', ') : null;
  const exposure = notes.exposureClass || null;
  const cover    = notes.concreteCover || null;
  const bearing  = notes.bearingCapacity || null;
  const geotech  = notes.geotechReport || null;
  const certBy   = notes.certifiedByOthers || [];
  const special  = notes.specialNotes || [];

  const coverRows = cover
    ? Object.entries(cover).map(([k, v]) => {
        const parts = [];
        if (v.bottom != null) parts.push(`bot ${v.bottom}`);
        if (v.top    != null) parts.push(`top ${v.top}`);
        if (v.sides  != null) parts.push(`side ${v.sides}`);
        return `<dt>${escapeHtml(k)}</dt><dd>${parts.join(' · ')} mm</dd>`;
      }).join('')
    : '';

  const bearingRows = bearing
    ? [
        bearing.padFootings   != null ? `pad ${bearing.padFootings}` : null,
        bearing.stripFootings != null ? `strip ${bearing.stripFootings}` : null,
        bearing.boredPiers    != null ? `pier ${bearing.boredPiers}` : null,
        bearing.shaftAdhesion != null ? `shaft adh. ${bearing.shaftAdhesion}` : null
      ].filter(Boolean).join(' · ') + ' kPa'
    : null;

  return `
    <details class="general-notes-card card" open>
      <summary>
        <span class="general-notes-card__title">General Notes</span>
        <span class="muted small">extracted from drawings</span>
      </summary>
      <div class="general-notes-card__body">
        ${codes ? `
          <div class="gn-row">
            <div class="gn-row__label">Codes</div>
            <div class="gn-row__value">${escapeHtml(codes)}</div>
          </div>` : ''}
        ${exposure ? `
          <div class="gn-row">
            <div class="gn-row__label">Exposure class</div>
            <div class="gn-row__value">${escapeHtml(exposure)}</div>
          </div>` : ''}
        ${bearingRows ? `
          <div class="gn-row">
            <div class="gn-row__label">Bearing</div>
            <div class="gn-row__value">${escapeHtml(bearingRows)}</div>
          </div>` : ''}
        ${cover && coverRows ? `
          <div class="gn-row">
            <div class="gn-row__label">Concrete cover</div>
            <div class="gn-row__value">
              <dl class="gn-cover">${coverRows}</dl>
            </div>
          </div>` : ''}
        ${geotech && (geotech.consultant || geotech.reportNumber) ? `
          <div class="gn-row">
            <div class="gn-row__label">Geotech</div>
            <div class="gn-row__value">
              ${escapeHtml(geotech.consultant || '')}${geotech.reportNumber ? ' · #' + escapeHtml(geotech.reportNumber) : ''}${geotech.date ? ' · ' + escapeHtml(geotech.date) : ''}
            </div>
          </div>` : ''}
        ${certBy.length ? `
          <div class="gn-row">
            <div class="gn-row__label">Certified by others</div>
            <div class="gn-row__value">
              <ul class="gn-cert">
                ${certBy.map((c) => `<li><strong>${escapeHtml(c.element || '')}</strong> — ${escapeHtml(c.responsibility || '')}${c.noteRef ? ' <span class="muted">(' + escapeHtml(c.noteRef) + ')</span>' : ''}</li>`).join('')}
              </ul>
            </div>
          </div>` : ''}
        ${special.length ? `
          <div class="gn-row">
            <div class="gn-row__label">Special notes</div>
            <div class="gn-row__value">
              <ul>${special.map((s) => `<li>${escapeHtml(s)}</li>`).join('')}</ul>
            </div>
          </div>` : ''}
      </div>
    </details>
  `;
}

function renderInspectionPlan(plan, inspections) {
  // Build a quick lookup for inspections by id
  const inspById = new Map(inspections.map((i) => [i.id, i]));

  return `
    <section class="project-section">
      <div class="section-heading">
        <h2>Inspection plan
          <span class="badge badge--count">${plan.length}</span>
        </h2>
        <div class="muted small plan-progress">
          ${planProgressLabel(plan)}
        </div>
      </div>
      <ol class="plan-list" role="list">
        ${plan.map((entry, i) => renderPlanEntry(entry, i, inspById)).join('')}
      </ol>
    </section>
  `;
}

function planProgressLabel(plan) {
  const total = plan.length;
  const done = plan.filter((p) => p.status === 'done').length;
  const inProg = plan.filter((p) => p.status === 'in-progress').length;
  const skipped = plan.filter((p) => p.status === 'skipped').length;
  const pending = total - done - inProg - skipped;
  const parts = [];
  if (done > 0)     parts.push(`${done} done`);
  if (inProg > 0)   parts.push(`${inProg} in progress`);
  if (skipped > 0)  parts.push(`${skipped} skipped`);
  if (pending > 0)  parts.push(`${pending} pending`);
  return parts.join(' · ');
}

function renderPlanEntry(entry, index, inspById) {
  const status = entry.status || 'pending';
  const linkedInspId = entry.completedInspectionId || '';
  const linkedInsp = linkedInspId ? inspById.get(Number(linkedInspId)) : null;

  // Status auto-promotes from 'in-progress' → 'done' when its inspection is complete.
  const effectiveStatus = (linkedInsp && linkedInsp.status === 'complete') ? 'done' : status;

  const statusBadge = {
    'pending':     '<span class="badge badge--pending">pending</span>',
    'in-progress': '<span class="badge badge--in-progress">in progress</span>',
    'done':        '<span class="badge badge--done">done</span>',
    'skipped':     '<span class="badge badge--skipped">skipped</span>'
  }[effectiveStatus] || '';

  const refs = (entry.drawingRefs || []).map((r) => r.sheetNumber).filter(Boolean);
  const refsLabel = refs.length === 0 ? '' :
    refs.length <= 3 ? refs.join(' · ') : `${refs.slice(0, 3).join(' · ')} +${refs.length - 3}`;

  // Revision-changed badge — set when applyRevisionDiff stamps this entry
  // and not yet acknowledged by the engineer.
  const revPending = !!entry.revisionChangedAt && !entry.revisionAcknowledgedAt;

  return `
    <li class="plan-entry plan-entry--${effectiveStatus} ${revPending ? 'plan-entry--rev-pending' : ''}"
        data-plan-index="${index}"
        data-plan-status="${effectiveStatus}"
        data-linked-inspection-id="${linkedInspId}"
        tabindex="0" role="link"
        onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}">
      <div class="plan-entry__seq">${entry.sequence ?? index + 1}</div>
      <div class="plan-entry__main">
        <div class="plan-entry__title-row">
          <span class="plan-entry__title">${escapeHtml(entry.title || entry.type || '')}</span>
          ${entry.holdPoint ? `<span class="badge badge--hold">hold point</span>` : ''}
          ${statusBadge}
          ${revPending ? `<span class="badge badge--rev-pending" title="A drawing referenced by this inspection has been revised — review and acknowledge.">revision changed</span>` : ''}
        </div>
        <div class="plan-entry__meta muted small">
          ${entry.level ? escapeHtml(entry.level) : ''}${entry.building && entry.building !== 'Main' ? ' · ' + escapeHtml(entry.building) : ''}
          ${refsLabel ? ' · ' + escapeHtml(refsLabel) : ''}
          ${entry.stage ? ' · ' + escapeHtml(entry.stage) : ''}
        </div>
        ${entry.rationale ? `<div class="plan-entry__rationale muted small">${escapeHtml(entry.rationale)}</div>` : ''}
        ${entry.skippedReason ? `<div class="plan-entry__skipped muted small">Skipped: ${escapeHtml(entry.skippedReason)}</div>` : ''}
      </div>
      <div class="plan-entry__actions">
        ${revPending
          ? `<button class="icon-btn icon-btn--accent" data-plan-action="ack-revision" aria-label="Acknowledge revision" title="Mark this revision change as reviewed">✓</button>`
          : ''}
        ${effectiveStatus === 'pending'
          ? `<button class="icon-btn" data-plan-action="skip" aria-label="Skip" title="Skip this inspection">⊘</button>`
          : ''}
        ${effectiveStatus === 'skipped' || effectiveStatus === 'in-progress'
          ? `<button class="icon-btn" data-plan-action="reset" aria-label="Reset" title="Reset to pending">↺</button>`
          : ''}
      </div>
    </li>
  `;
}

function renderEmptyDrawings() {
  return `
    <div class="empty card">
      <p class="muted">No drawings uploaded yet. Drop your structural drawing PDFs here (multiple at once is fine) — each one will become its own sheet you can open on site.</p>
      <button class="btn btn--primary" data-action="trigger-upload">Upload drawings</button>
    </div>
  `;
}

function renderFlatDrawings(drawings) {
  return `
    <div data-drawings-root>
      <ul class="drawing-list" role="list">
        ${drawings.map(renderDrawingCard).join('')}
      </ul>
    </div>
  `;
}

/**
 * Grouped layout for Cowork-imported projects: filter chips at top, then
 * drawings grouped by level with kind labels.
 */
function renderGroupedDrawings(drawings) {
  // Build filter facets
  const kindCounts = {};
  for (const d of drawings) {
    const kindTag = (d.tags || []).find((t) => t.startsWith('kind:'));
    const k = kindTag ? kindTag.slice(5) : 'other';
    kindCounts[k] = (kindCounts[k] || 0) + 1;
  }
  const kindOrder = ['plan', 'section', 'elevation', 'detail', 'typical-detail', 'schedule', 'cover', 'notes', 'other'];
  const presentKinds = kindOrder.filter((k) => kindCounts[k]);

  // Group by level
  const levelOrder = ['ground', 'l1', 'l2', 'l3', 'l4', 'l5', 'lower-roof', 'roof', 'basement', 'multi', 'na'];
  const groups = new Map();
  for (const d of drawings) {
    const levelTag = (d.tags || []).find((t) => t.startsWith('level:'));
    const lvl = levelTag ? levelTag.slice(6) : 'na';
    if (!groups.has(lvl)) groups.set(lvl, []);
    groups.get(lvl).push(d);
  }
  const orderedLevels = levelOrder.filter((l) => groups.has(l))
    .concat(Array.from(groups.keys()).filter((l) => !levelOrder.includes(l)));

  return `
    <div data-drawings-root>
      <div class="dwg-filter-chips" id="dwg-filter-chips" role="tablist" aria-label="Filter drawings by kind">
        <button class="chip is-active" data-filter="all" type="button">All <span class="chip__count">${drawings.length}</span></button>
        ${presentKinds.map((k) => `
          <button class="chip" data-filter="${escapeAttr('kind:' + k)}" type="button">
            ${escapeHtml(prettyKind(k))} <span class="chip__count">${kindCounts[k]}</span>
          </button>`).join('')}
      </div>

      ${orderedLevels.map((lvl) => `
        <div class="dwg-group" data-level="${escapeAttr(lvl)}">
          <h3 class="dwg-group__title">${escapeHtml(prettyLevel(lvl))} <span class="muted small">${groups.get(lvl).length}</span></h3>
          <ul class="drawing-list" role="list">
            ${groups.get(lvl).map(renderDrawingCard).join('')}
          </ul>
        </div>
      `).join('')}
    </div>
  `;
}

function applyDrawingFilter(root, filter) {
  const cards = root.querySelectorAll('.drawing-card');
  let visibleCount = 0;
  cards.forEach((card) => {
    if (filter === 'all') {
      card.classList.remove('is-hidden');
      visibleCount += 1;
      return;
    }
    const tags = (card.dataset.tags || '').split(' ').filter(Boolean);
    if (tags.includes(filter)) {
      card.classList.remove('is-hidden');
      visibleCount += 1;
    } else {
      card.classList.add('is-hidden');
    }
  });
  // Hide empty groups
  root.querySelectorAll('.dwg-group').forEach((g) => {
    const visible = g.querySelectorAll('.drawing-card:not(.is-hidden)').length;
    g.classList.toggle('is-hidden', visible === 0);
  });
}

function renderDrawingCard(d) {
  const sheet = d.sheetNumber ? d.sheetNumber : `Page ${d.pageNumber || 1}`;
  const desc  = d.description || d.filename || '';
  const rev   = d.revision ? ` (Rev ${escapeHtml(d.revision)})` : '';
  const titleLine = `${escapeHtml(sheet)}${desc ? ' — ' + escapeHtml(desc) : ''}${rev}`;

  // Pull element tags onto a chip strip
  const elementTags = (d.tags || []).filter((t) => t.startsWith('element:'))
    .slice(0, 3)
    .map((t) => t.slice(8));

  return `
    <li class="drawing-card" data-drawing-id="${d.id}" data-tags="${escapeAttr((d.tags || []).join(' '))}"
        tabindex="0" role="link"
        onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}">
      <div class="drawing-card__thumb" id="thumb-${d.id}">
        <div class="drawing-card__thumb-placeholder">PDF</div>
      </div>
      <div class="drawing-card__body">
        <div class="drawing-card__title">${titleLine}</div>
        ${elementTags.length ? `
          <div class="drawing-card__chips">
            ${elementTags.map((t) => `<span class="chip chip--xs">${escapeHtml(t)}</span>`).join('')}
          </div>` : ''}
        <div class="drawing-card__meta muted">
          ${d.calibration
            ? `<span class="drawing-card__cal drawing-card__cal--yes">Calibrated</span>`
            : `<span class="drawing-card__cal drawing-card__cal--no">Not calibrated</span>`}
          &middot; ${formatDate(d.uploadedAt)}
        </div>
      </div>
      <div class="drawing-card__actions">
        <button class="icon-btn" data-action="edit" aria-label="Edit drawing">
          <svg viewBox="0 0 24 24"><path d="M12 20h9M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>
        </button>
        <button class="icon-btn" data-action="delete" aria-label="Delete drawing">
          <svg viewBox="0 0 24 24"><path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/></svg>
        </button>
      </div>
    </li>
  `;
}

function renderEmptyInspections(noDrawings) {
  if (noDrawings) {
    return `
      <div class="empty card">
        <p class="muted">Upload at least one structural drawing above before starting an inspection.</p>
      </div>
    `;
  }
  return `
    <div class="empty card">
      <p class="muted">Tap a plan entry above to start an inspection from the program, or use <strong>+ New inspection</strong> for ad-hoc work.</p>
    </div>
  `;
}

function renderInspectionList(inspections) {
  return `
    <ul class="inspection-list stack" role="list">
      ${inspections.map((i) => `
        <li class="inspection-card" data-inspection-id="${i.id}" tabindex="0" role="link"
            onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}">
          <div class="inspection-card__main">
            <div class="inspection-card__type">${escapeHtml(i.inspectionTypeName || '')}</div>
            <div class="inspection-card__meta muted">
              ${formatDate(i.date)}${i.inspectorName ? ' &middot; ' + escapeHtml(i.inspectorName) : ''}
            </div>
          </div>
          <div class="inspection-card__status">
            <span class="badge ${i.status === 'complete' ? 'badge--complete' : 'badge--draft'}">
              ${escapeHtml(i.status || 'draft')}
            </span>
          </div>
        </li>
      `).join('')}
    </ul>
  `;
}

/* --------------------------------------------------------------------------
   Handlers
   -------------------------------------------------------------------------- */

async function onStartPlanEntry(projectId, planIndex) {
  const project = await getProject(projectId);
  if (!project) return;
  // Open the rich pre-inspection brief modal (Phase C2 — v2.2). It handles the
  // confirmation flow + creates the inspection on confirm + navigates.
  await openPreInspectionBrief(project, planIndex);
}

async function onSkipPlanEntry(projectId, planIndex) {
  const project = await getProject(projectId);
  const entry = project?.inspectionPlan?.[planIndex];
  if (!entry) return;

  let reason = '';
  await openModal({
    title: `Skip: ${entry.title}?`,
    primaryLabel: 'Skip',
    bodyHtml: `
      <p class="muted">Skipped entries stay in the plan with the reason recorded — useful for audit (e.g. "builder cast without notification").</p>
      <div class="form-field">
        <label for="skip-reason">Reason</label>
        <textarea id="skip-reason" name="reason" rows="3" placeholder="e.g. Manufacturer-certified, no BT inspection required."></textarea>
      </div>
    `,
    onConfirm: (dialog) => {
      reason = dialog.querySelector('[name=reason]').value.trim();
      return true;
    }
  });
  if (!reason) return;        // dialog cancelled or empty

  await skipPlanEntry(projectId, planIndex, reason);
  toast('Plan entry skipped', { kind: 'info' });
  go('project', projectId);
}

async function onNewInspection(projectId) {
  const inspection = await openNewInspection({ projectId });
  if (inspection) {
    toast('Inspection started', { kind: 'success' });
    go('inspection', inspection.id);
  }
}

async function renderThumbnailsSequentially(drawings) {
  for (const d of drawings) {
    try {
      const blob = await getDrawingBlob(d);
      if (!blob) continue;
      const url = await renderThumbnail(blob, 280, d.pageNumber || 1);
      const holder = document.getElementById(`thumb-${d.id}`);
      if (holder) holder.innerHTML = `<img alt="" src="${url}"/>`;
    } catch (err) {
      console.warn('Thumbnail render failed for drawing', d.id, err);
    }
  }
}

async function onEditProject(project) {
  const result = await openModal({
    title: 'Edit project',
    primaryLabel: 'Save changes',
    bodyHtml: `
      <div class="form">
        <div class="form-field">
          <label for="p-job">Job number <span class="req">*</span></label>
          <input id="p-job" name="jobNumber" type="text" value="${escapeAttr(project.jobNumber)}" required/>
        </div>
        <div class="form-field">
          <label for="p-name">Project name <span class="req">*</span></label>
          <input id="p-name" name="name" type="text" value="${escapeAttr(project.name)}" required/>
        </div>
        <div class="form-field">
          <label for="p-client">Client</label>
          <input id="p-client" name="client" type="text" value="${escapeAttr(project.client || '')}"/>
        </div>
        <div class="form-field">
          <label for="p-addr">Site address</label>
          <input id="p-addr" name="siteAddress" type="text" value="${escapeAttr(project.siteAddress || '')}"/>
        </div>
        <div class="form-field">
          <label for="p-builder">Builder</label>
          <input id="p-builder" name="builder" type="text" value="${escapeAttr(project.builder || '')}"/>
        </div>
        <div class="form-field">
          <label for="p-notes">Notes</label>
          <textarea id="p-notes" name="notes" rows="3">${escapeHtml(project.notes || '')}</textarea>
        </div>
      </div>
    `,
    onConfirm: async (dialog) => {
      const patch = {
        jobNumber:   dialog.querySelector('[name=jobNumber]').value,
        name:        dialog.querySelector('[name=name]').value,
        client:      dialog.querySelector('[name=client]').value,
        siteAddress: dialog.querySelector('[name=siteAddress]').value,
        builder:     dialog.querySelector('[name=builder]').value,
        notes:       dialog.querySelector('[name=notes]').value
      };
      if (!patch.jobNumber.trim()) throw new Error('Job number is required');
      if (!patch.name.trim())      throw new Error('Project name is required');
      return updateProject(project.id, patch);
    }
  });

  if (result) {
    toast('Project updated', { kind: 'success' });
    go('project', project.id);
  }
}

async function onDeleteProject(project) {
  const ok = await confirmDialog({
    title: `Delete project ${project.jobNumber}?`,
    message: `This will permanently remove the project, all uploaded drawings, and any inspections recorded against it. This cannot be undone.`,
    confirmLabel: 'Delete project',
    danger: true
  });
  if (!ok) return;
  await deleteProject(project.id);
  toast('Project deleted', { kind: 'info' });
  go('projects');
}

async function onUploadDrawings(event, projectId) {
  const files = Array.from(event.target.files || []);
  event.target.value = '';
  if (files.length === 0) return;

  const progress = document.getElementById('upload-progress');
  progress.classList.remove('is-hidden');

  let createdDrawings = 0;
  for (const file of files) {
    progress.textContent = `Reading ${file.name}…`;
    try {
      if (file.type !== 'application/pdf' && !/\.pdf$/i.test(file.name)) {
        throw new Error('Not a PDF');
      }
      progress.textContent = `Scanning pages of ${file.name}…`;
      const pages = await extractPageMetadata(file);
      if (pages.length === 0) throw new Error('Could not read any pages');

      const filenameSheet = extractSheetNumberFromFilename(file.name);
      const filenameRev   = extractRevisionFromFilename(file.name);

      const pageRecords = pages.map((p) => ({
        pageNumber:  p.pageNumber,
        pageWidth:   p.width,
        pageHeight:  p.height,
        sheetNumber: p.sheetNumber || (pages.length === 1 ? filenameSheet : ''),
        revision:    p.revision    || (pages.length === 1 ? filenameRev   : ''),
        description: p.description || ''
      }));

      progress.textContent = `Saving ${pages.length} drawing${pages.length === 1 ? '' : 's'} from ${file.name}…`;
      const created = await addDrawingsFromSource(projectId, {
        pdfBlob:   file,
        filename:  file.name,
        pageCount: pages.length
      }, pageRecords);

      createdDrawings += created.length;
    } catch (err) {
      console.error('Upload failed for', file.name, err);
      toast(`Failed to upload ${file.name}: ${err.message}`, { kind: 'error' });
    }
  }

  progress.classList.add('is-hidden');
  if (createdDrawings > 0) {
    toast(`Uploaded ${createdDrawings} drawing${createdDrawings === 1 ? '' : 's'}`, { kind: 'success' });
  }
  go('project', projectId);
}

async function onEditDrawing(drawingId, projectId) {
  const d = await getDrawing(drawingId);
  if (!d) return;

  const result = await openModal({
    title: 'Edit drawing',
    primaryLabel: 'Save',
    bodyHtml: `
      <div class="form">
        <div class="form-field">
          <label for="d-sheet">Sheet number</label>
          <input id="d-sheet" name="sheetNumber" type="text" value="${escapeAttr(d.sheetNumber || '')}" placeholder="e.g. S-01"/>
        </div>
        <div class="form-field">
          <label for="d-rev">Revision</label>
          <input id="d-rev" name="revision" type="text" value="${escapeAttr(d.revision || '')}" placeholder="e.g. B"/>
        </div>
        <div class="form-field">
          <label for="d-desc">Description</label>
          <input id="d-desc" name="description" type="text" value="${escapeAttr(d.description || '')}" placeholder="e.g. Ground floor slab layout"/>
        </div>
        <div class="form-field form-field--readonly">
          <label>File</label>
          <div class="muted">${escapeHtml(d.filename || '—')}</div>
        </div>
      </div>
    `,
    onConfirm: async (dialog) => {
      return updateDrawing(drawingId, {
        sheetNumber: dialog.querySelector('[name=sheetNumber]').value,
        revision:    dialog.querySelector('[name=revision]').value,
        description: dialog.querySelector('[name=description]').value
      });
    }
  });

  if (result) {
    toast('Drawing updated', { kind: 'success' });
    go('project', projectId);
  }
}

async function onDeleteDrawing(drawingId, projectId) {
  const ok = await confirmDialog({
    title: 'Delete drawing?',
    message: 'The drawing and its calibration will be removed from this project. This cannot be undone.',
    confirmLabel: 'Delete',
    danger: true
  });
  if (!ok) return;
  await deleteDrawing(drawingId);
  toast('Drawing deleted', { kind: 'info' });
  go('project', projectId);
}

/* --------------------------------------------------------------------------
   Utilities
   -------------------------------------------------------------------------- */

function prettyLevel(lvl) {
  return ({
    'ground':     'Ground floor',
    'l1':         'Level 1',
    'l2':         'Level 2',
    'l3':         'Level 3',
    'l4':         'Level 4',
    'l5':         'Level 5',
    'roof':       'Roof',
    'lower-roof': 'Lower roof',
    'basement':   'Basement',
    'multi':      'Multi-level',
    'na':         'General / details'
  }[lvl] || lvl);
}

function prettyKind(k) {
  return ({
    'plan':           'Plans',
    'section':        'Sections',
    'elevation':      'Elevations',
    'detail':         'Details',
    'typical-detail': 'Typical details',
    'schedule':       'Schedules',
    'cover':          'Cover',
    'notes':          'Notes',
    'other':          'Other'
  }[k] || k);
}

function formatDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  return d.toLocaleDateString('en-AU', { day: '2-digit', month: 'short', year: 'numeric' });
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
function escapeAttr(s) { return escapeHtml(s).replace(/"/g, '&quot;'); }
