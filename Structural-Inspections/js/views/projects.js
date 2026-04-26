/**
 * Projects list view.
 *
 * Shows all projects with their job number, client, and updated date,
 * and lets the user create a new project.
 */

import { listProjects, createProject } from '../db.js';
import { openModal } from '../components/modal.js';
import { toast } from '../components/toast.js';
import { openImportProject } from '../components/import-project.js';
import { go } from '../router.js';

export async function render(root) {
  const projects = await listProjects();

  root.innerHTML = `
    <div class="section-heading">
      <h1>Projects</h1>
      <div class="cluster">
        <button class="btn btn--secondary" id="btn-import-project" title="Import a Cowork-built project bundle">
          <span aria-hidden="true">↧</span> Import
        </button>
        <button class="btn btn--primary" id="btn-new-project">
          <span aria-hidden="true">+</span> New project
        </button>
      </div>
    </div>

    ${projects.length === 0 ? renderEmpty() : renderList(projects)}
  `;

  root.querySelector('#btn-new-project').addEventListener('click', onNewProject);
  root.querySelector('#btn-import-project').addEventListener('click', () => openImportProject());

  const emptyBtn = root.querySelector('#btn-new-project-empty');
  if (emptyBtn) emptyBtn.addEventListener('click', onNewProject);

  const emptyImportBtn = root.querySelector('#btn-import-project-empty');
  if (emptyImportBtn) emptyImportBtn.addEventListener('click', () => openImportProject());

  // Delegate: clicking a project card opens it
  root.addEventListener('click', (e) => {
    const card = e.target.closest('[data-project-id]');
    if (card && !e.defaultPrevented) {
      const id = card.dataset.projectId;
      go('project', id);
    }
  });
}

function renderEmpty() {
  return `
    <div class="empty card">
      <h2>No projects yet</h2>
      <p>Two ways to start:</p>
      <ul class="muted" style="margin: 0 0 var(--space-4);">
        <li><strong>Import</strong> — pick a <code>.btproject</code> bundle Cowork built from your drawings (recommended).</li>
        <li><strong>Create from scratch</strong> — set up project metadata manually, then upload PDFs page-by-page.</li>
      </ul>
      <div class="cluster">
        <button class="btn btn--primary" id="btn-import-project-empty">↧ Import bundle</button>
        <button class="btn btn--secondary" id="btn-new-project-empty">+ Create from scratch</button>
      </div>
    </div>
  `;
}

function renderList(projects) {
  return `
    <ul class="project-list stack" role="list">
      ${projects.map((p) => `
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
  `;
}

async function onNewProject() {
  const result = await openModal({
    title: 'New project',
    primaryLabel: 'Create project',
    bodyHtml: `
      <div class="form">
        <div class="form-field">
          <label for="p-job">Job number <span class="req">*</span></label>
          <input id="p-job" name="jobNumber" type="text" autocomplete="off" required placeholder="e.g. 23123"/>
        </div>
        <div class="form-field">
          <label for="p-name">Project name <span class="req">*</span></label>
          <input id="p-name" name="name" type="text" autocomplete="off" required placeholder="e.g. 123 Smith St Alterations"/>
        </div>
        <div class="form-field">
          <label for="p-client">Client</label>
          <input id="p-client" name="client" type="text" autocomplete="off" placeholder="Builder or principal"/>
        </div>
        <div class="form-field">
          <label for="p-addr">Site address</label>
          <input id="p-addr" name="siteAddress" type="text" autocomplete="off" placeholder="Street, suburb, state"/>
        </div>
        <div class="form-field">
          <label for="p-notes">Notes</label>
          <textarea id="p-notes" name="notes" rows="3" placeholder="Scope, drawing set reference, anything worth remembering"></textarea>
        </div>
      </div>
    `,
    onConfirm: async (dialog) => {
      const data = {
        jobNumber:   dialog.querySelector('[name=jobNumber]').value,
        name:        dialog.querySelector('[name=name]').value,
        client:      dialog.querySelector('[name=client]').value,
        siteAddress: dialog.querySelector('[name=siteAddress]').value,
        notes:       dialog.querySelector('[name=notes]').value
      };
      return createProject(data);
    }
  });

  if (result) {
    toast(`Project ${result.jobNumber} created`, { kind: 'success' });
    go('project', result.id);
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
