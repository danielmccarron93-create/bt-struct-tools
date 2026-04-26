/**
 * Home screen — landing after boot.
 *
 * Shows quick-action tiles, plus compact lists of the three most recent
 * projects and the three most recent inspections so you can get back into
 * whatever you were working on with one tap.
 */

import {
  db, listProjects, listInspections, listNextUpFromPlans,
  startInspectionFromPlanEntry
} from '../db.js';
import { confirmDialog } from '../components/modal.js';
import { toast } from '../components/toast.js';
import { go } from '../router.js';

export async function render(root) {
  const [projectCount, inspectionCount, draftCount,
         recentProjects, recentInspections, nextUp] = await Promise.all([
    db.projects.count(),
    db.inspections.count(),
    db.inspections.where('status').equals('draft').count().catch(() => 0),
    listProjects().then((arr) => arr.slice(0, 3)),
    listInspections().then((arr) => arr.slice(0, 3)),
    listNextUpFromPlans(3)
  ]);

  // Need project metadata for the inspection rows (job number + project name)
  const projectById = new Map(
    (await listProjects()).map((p) => [p.id, p])
  );

  root.innerHTML = `
    <section class="hero">
      <h1>Site inspection, finished on site.</h1>
      <p>Open a project, set up today's inspection, and leave site with the report ready to email.</p>
      <div class="hero__meta">Bligh Tanner &middot; Structural</div>
    </section>

    <div class="section-heading">
      <h2>Quick actions</h2>
    </div>

    <div class="tile-grid stack-lg">
      <a class="tile" href="#/projects">
        <span class="tile__icon" aria-hidden="true">
          <svg viewBox="0 0 24 24"><path d="M3 6a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/></svg>
        </span>
        <span class="tile__title">Projects</span>
        <span class="tile__meta">${projectCount === 0 ? 'No projects yet — create one to get started.' : `${projectCount} project${projectCount === 1 ? '' : 's'}.`}</span>
      </a>
      <a class="tile" href="#/inspections">
        <span class="tile__icon" aria-hidden="true">
          <svg viewBox="0 0 24 24"><path d="M8 2h8l4 4v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6Z"/><path d="M8 10h8M8 14h8M8 18h5"/></svg>
        </span>
        <span class="tile__title">Inspections</span>
        <span class="tile__meta">
          ${inspectionCount === 0
            ? 'None yet — start one from a project or the Inspections tab.'
            : `${inspectionCount} total${draftCount ? ` &middot; ${draftCount} draft${draftCount === 1 ? '' : 's'}` : ''}.`}
        </span>
      </a>
      <a class="tile" href="#/settings">
        <span class="tile__icon" aria-hidden="true">
          <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="m19.4 15-.6 1 1.4 2.2-2 2-2.2-1.4-1 .6-.6 2.6h-2.8l-.6-2.6-1-.6-2.2 1.4-2-2L5.8 16l-.6-1-2.6-.6v-2.8l2.6-.6.6-1-1.4-2.2 2-2 2.2 1.4 1-.6.6-2.6h2.8l.6 2.6 1 .6 2.2-1.4 2 2L18.2 9l.6 1 2.6.6v2.8Z"/></svg>
        </span>
        <span class="tile__title">Settings</span>
        <span class="tile__meta">Inspector profile, RPEQ, diagnostics.</span>
      </a>
    </div>

    ${nextUp.length > 0 ? `
      <div class="section-heading" style="margin-top: var(--space-6);">
        <h2>Next up</h2>
        <span class="muted small">from your project plans</span>
      </div>
      <ul class="next-up-list stack" role="list" id="next-up">
        ${nextUp.map(({ project, entry, planIndex }) => {
          const refs = (entry.drawingRefs || []).map((r) => r.sheetNumber).filter(Boolean);
          const refsLabel = refs.length === 0 ? '' :
            refs.length <= 3 ? refs.join(' · ') : `${refs.slice(0, 3).join(' · ')} +${refs.length - 3}`;
          const status = entry.status || 'pending';
          return `
            <li class="next-up-card next-up-card--${status}"
                data-project-id="${project.id}" data-plan-index="${planIndex}"
                tabindex="0" role="link"
                onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}">
              <div class="next-up-card__seq">${entry.sequence ?? planIndex + 1}</div>
              <div class="next-up-card__main">
                <div class="next-up-card__project muted small">
                  ${escapeHtml(project.jobNumber || '')}${project.jobNumber ? ' · ' : ''}${escapeHtml(project.name)}
                </div>
                <div class="next-up-card__title">${escapeHtml(entry.title || entry.type || '')}</div>
                <div class="next-up-card__meta muted small">
                  ${entry.level ? escapeHtml(entry.level) : ''}${refsLabel ? ' · ' + escapeHtml(refsLabel) : ''}
                  ${entry.holdPoint ? ' · <span class="badge badge--hold">hold point</span>' : ''}
                  ${status === 'in-progress' ? ' · <span class="badge badge--in-progress">in progress</span>' : ''}
                </div>
              </div>
              <div class="next-up-card__cta">
                ${status === 'in-progress' ? 'Resume' : 'Start'}
                <span aria-hidden="true">›</span>
              </div>
            </li>
          `;
        }).join('')}
      </ul>
    ` : ''}

    ${recentInspections.length > 0 ? `
      <div class="section-heading" style="margin-top: var(--space-6);">
        <h2>Recent inspections</h2>
        <a href="#/inspections" class="btn btn--ghost btn--sm">See all &rsaquo;</a>
      </div>
      <ul class="inspection-list stack" role="list" id="recent-inspections">
        ${recentInspections.map((i) => {
          const p = projectById.get(i.projectId);
          return `
            <li class="inspection-card" data-inspection-id="${i.id}" tabindex="0" role="link"
                onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}">
              <div class="inspection-card__main">
                <div class="inspection-card__project muted">
                  ${p ? `${escapeHtml(p.jobNumber)} &middot; ${escapeHtml(p.name)}` : 'Unknown project'}
                </div>
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
          `;
        }).join('')}
      </ul>
    ` : ''}

    ${recentProjects.length > 0 ? `
      <div class="section-heading" style="margin-top: var(--space-6);">
        <h2>Recent projects</h2>
        <a href="#/projects" class="btn btn--ghost btn--sm">See all &rsaquo;</a>
      </div>
      <ul class="project-list stack" role="list" id="recent-projects">
        ${recentProjects.map((p) => `
          <li class="project-card" data-project-id="${p.id}" tabindex="0" role="link"
              onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}">
            <div class="project-card__main">
              <div class="project-card__job">${escapeHtml(p.jobNumber)}</div>
              <div class="project-card__title">${escapeHtml(p.name)}</div>
              <div class="project-card__meta muted">
                ${p.client ? escapeHtml(p.client) : '—'}${p.siteAddress ? ' &middot; ' + escapeHtml(p.siteAddress) : ''}
              </div>
            </div>
            <div class="project-card__aside muted">
              <div class="project-card__aside-label">Updated</div>
              <div>${formatDate(p.updatedAt)}</div>
            </div>
          </li>
        `).join('')}
      </ul>
    ` : ''}
  `;

  const recentP = root.querySelector('#recent-projects');
  if (recentP) {
    recentP.addEventListener('click', (e) => {
      const card = e.target.closest('[data-project-id]');
      if (card) go('project', card.dataset.projectId);
    });
  }
  const recentI = root.querySelector('#recent-inspections');
  if (recentI) {
    recentI.addEventListener('click', (e) => {
      const card = e.target.closest('[data-inspection-id]');
      if (card) go('inspection', card.dataset.inspectionId);
    });
  }
  const nextUpEl = root.querySelector('#next-up');
  if (nextUpEl) {
    nextUpEl.addEventListener('click', async (e) => {
      const card = e.target.closest('[data-project-id][data-plan-index]');
      if (!card) return;
      const projectId = Number(card.dataset.projectId);
      const planIdx   = Number(card.dataset.planIndex);
      const entry     = nextUp.find((n) => n.project.id === projectId && n.planIndex === planIdx)?.entry;
      if (!entry) return;

      // In-progress → jump straight to the linked inspection.
      if (entry.completedInspectionId && (entry.status || 'pending') === 'in-progress') {
        go('inspection', entry.completedInspectionId);
        return;
      }

      // Otherwise confirm + start.
      const ok = await confirmDialog({
        title: `Start: ${entry.title || entry.type}?`,
        message: entry.rationale || 'Pre-fills the inspection from the project plan.',
        confirmLabel: 'Start inspection'
      });
      if (!ok) return;

      try {
        const inspection = await startInspectionFromPlanEntry(projectId, planIdx);
        toast(`Started: ${inspection.inspectionTypeName}`, { kind: 'success' });
        go('inspection', inspection.id);
      } catch (err) {
        toast(err.message || 'Couldn\u2019t start inspection', { kind: 'error', duration: 6000 });
      }
    });
  }
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
