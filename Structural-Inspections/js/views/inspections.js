/**
 * Inspections list view — global across all projects.
 */

import { listInspections, listProjects } from '../db.js';
import { toast } from '../components/toast.js';
import { go } from '../router.js';
import { openNewInspection } from './new-inspection.js';

export async function render(root) {
  const [inspections, projects] = await Promise.all([
    listInspections(),
    listProjects()
  ]);

  const projectMap = new Map(projects.map((p) => [p.id, p]));

  root.innerHTML = `
    <div class="section-heading">
      <h1>Inspections</h1>
      <div class="cluster">
        <button class="btn btn--primary" id="btn-new-inspection">
          <span aria-hidden="true">+</span> New inspection
        </button>
      </div>
    </div>

    ${inspections.length === 0 ? renderEmpty() : renderList(inspections, projectMap)}
  `;

  root.querySelector('#btn-new-inspection').addEventListener('click', async () => {
    const inspection = await openNewInspection();
    if (inspection) {
      toast('Inspection started', { kind: 'success' });
      go('inspection', inspection.id);
    }
  });

  const emptyBtn = root.querySelector('#btn-new-inspection-empty');
  if (emptyBtn) {
    emptyBtn.addEventListener('click', async () => {
      const inspection = await openNewInspection();
      if (inspection) {
        toast('Inspection started', { kind: 'success' });
        go('inspection', inspection.id);
      }
    });
  }

  // Delegate: open an inspection
  const list = root.querySelector('.inspection-list');
  if (list) {
    list.addEventListener('click', (e) => {
      const card = e.target.closest('[data-inspection-id]');
      if (card) go('inspection', card.dataset.inspectionId);
    });
  }
}

function renderEmpty() {
  return `
    <div class="empty card">
      <h2>No inspections yet</h2>
      <p>Start an inspection for one of your projects. You'll pick the inspection type, the primary drawing you're marking up, and the date — then you're ready to walk the plan.</p>
      <button class="btn btn--primary" id="btn-new-inspection-empty">+ Start your first inspection</button>
    </div>
  `;
}

function renderList(inspections, projectMap) {
  return `
    <ul class="inspection-list stack" role="list">
      ${inspections.map((i) => {
        const project = projectMap.get(i.projectId);
        return `
          <li class="inspection-card" data-inspection-id="${i.id}" tabindex="0" role="link"
              onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}">
            <div class="inspection-card__main">
              <div class="inspection-card__project muted">
                ${project ? `${escapeHtml(project.jobNumber)} &middot; ${escapeHtml(project.name)}` : 'Unknown project'}
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
  `;
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
