/**
 * Project detail view.
 *
 * Shows a single project's metadata, its drawings, and (later) its
 * inspections. Primary actions here: edit project, upload drawings,
 * open a drawing, delete a drawing, delete the project.
 */

import {
  getProject, updateProject, deleteProject,
  listDrawingsForProject, addDrawingsFromSource, updateDrawing, deleteDrawing,
  getDrawing, getDrawingBlob,
  listInspectionsForProject
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

  const [drawings, inspections] = await Promise.all([
    listDrawingsForProject(id),
    listInspectionsForProject(id)
  ]);

  root.innerHTML = `
    <nav class="breadcrumb"><a href="#/projects">&lsaquo; Projects</a></nav>

    <header class="project-header">
      <div class="project-header__job">${escapeHtml(project.jobNumber)}</div>
      <h1>${escapeHtml(project.name)}</h1>
      <dl class="project-header__meta">
        ${project.client      ? `<div><dt>Client</dt><dd>${escapeHtml(project.client)}</dd></div>` : ''}
        ${project.siteAddress ? `<div><dt>Site</dt><dd>${escapeHtml(project.siteAddress)}</dd></div>` : ''}
        ${project.notes       ? `<div><dt>Notes</dt><dd>${escapeHtml(project.notes)}</dd></div>` : ''}
      </dl>
      <div class="cluster">
        <button class="btn btn--secondary btn--sm" id="btn-edit-project">Edit project</button>
        <button class="btn btn--ghost btn--sm" id="btn-delete-project" style="color: var(--color-danger);">Delete project</button>
      </div>
    </header>

    <section class="project-section">
      <div class="section-heading">
        <h2>Drawings <span class="badge badge--count">${drawings.length}</span></h2>
        <div class="cluster">
          <label class="btn btn--primary btn--sm" for="input-drawings">
            <span aria-hidden="true">+</span> Upload PDF
          </label>
          <input id="input-drawings" type="file" accept="application/pdf" multiple class="sr-only"/>
        </div>
      </div>

      <div class="upload-progress is-hidden" id="upload-progress" role="status" aria-live="polite"></div>

      ${drawings.length === 0 ? renderEmptyDrawings() : renderDrawingList(drawings)}
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

  root.querySelector('#btn-edit-project').addEventListener('click', () => onEditProject(project));
  root.querySelector('#btn-delete-project').addEventListener('click', () => onDeleteProject(project));
  root.querySelector('#input-drawings').addEventListener('change', (e) => onUploadDrawings(e, project.id));
  root.querySelector('#btn-new-inspection').addEventListener('click', () => onNewInspection(project.id));

  // Inspection list delegation
  const inspList = root.querySelector('.inspection-list');
  if (inspList) {
    inspList.addEventListener('click', (e) => {
      const card = e.target.closest('[data-inspection-id]');
      if (card) go('inspection', card.dataset.inspectionId);
    });
  }

  // Also wire the "upload drawings" button in the empty state (if present)
  root.querySelectorAll('[data-action="trigger-upload"]').forEach((btn) => {
    btn.addEventListener('click', () => document.getElementById('input-drawings').click());
  });

  // Delegate on drawing list
  const drawingList = root.querySelector('.drawing-list');
  if (drawingList) {
    drawingList.addEventListener('click', (e) => {
      const card = e.target.closest('[data-drawing-id]');
      if (!card) return;

      if (e.target.closest('[data-action=edit]')) {
        e.preventDefault();
        e.stopPropagation();
        onEditDrawing(Number(card.dataset.drawingId), project.id);
        return;
      }
      if (e.target.closest('[data-action=delete]')) {
        e.preventDefault();
        e.stopPropagation();
        onDeleteDrawing(Number(card.dataset.drawingId), project.id);
        return;
      }
      // Open drawing viewer
      go('drawing', card.dataset.drawingId);
    });
  }

  // Kick off thumbnails in the background (don't block first paint).
  renderThumbnailsSequentially(drawings);
}

/* --------------------------------------------------------------------------
   Rendering helpers
   -------------------------------------------------------------------------- */

function renderEmptyDrawings() {
  return `
    <div class="empty card">
      <p class="muted">No drawings uploaded yet. Drop your structural drawing PDFs here (multiple at once is fine) — each one will become its own sheet you can open on site.</p>
      <button class="btn btn--primary" data-action="trigger-upload">Upload drawings</button>
    </div>
  `;
}

function renderDrawingList(drawings) {
  return `
    <ul class="drawing-list" role="list">
      ${drawings.map((d) => `
        <li class="drawing-card" data-drawing-id="${d.id}" tabindex="0" role="link"
            onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}">
          <div class="drawing-card__thumb" id="thumb-${d.id}">
            <div class="drawing-card__thumb-placeholder">PDF</div>
          </div>
          <div class="drawing-card__body">
            <div class="drawing-card__sheet">
              ${escapeHtml(d.sheetNumber || `Page ${d.pageNumber || 1}`)}
              ${d.revision ? `<span class="drawing-card__rev">Rev ${escapeHtml(d.revision)}</span>` : ''}
            </div>
            <div class="drawing-card__desc">${escapeHtml(d.description || d.filename || '—')}</div>
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
      `).join('')}
    </ul>
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
      <p class="muted">No inspections yet for this project. Start an inspection — pick the type (e.g. Slab on Ground Pre-Pour), the drawing you're marking up, and the date.</p>
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
      if (holder) {
        holder.innerHTML = `<img alt="" src="${url}"/>`;
      }
    } catch (err) {
      console.warn('Thumbnail render failed for drawing', d.id, err);
    }
  }
}

/* --------------------------------------------------------------------------
   Handlers — project
   -------------------------------------------------------------------------- */

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

/* --------------------------------------------------------------------------
   Handlers — drawings
   -------------------------------------------------------------------------- */

async function onUploadDrawings(event, projectId) {
  const files = Array.from(event.target.files || []);
  event.target.value = ''; // reset so same file can be reselected later
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

      // Scan every page of the PDF for size + title-block metadata.
      progress.textContent = `Scanning pages of ${file.name}…`;
      const pages = await extractPageMetadata(file);
      if (pages.length === 0) throw new Error('Could not read any pages');

      // Filename-level fallbacks — only used if per-page extraction comes up empty.
      const filenameSheet = extractSheetNumberFromFilename(file.name);
      const filenameRev   = extractRevisionFromFilename(file.name);

      // Build per-page drawing records.
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

  // Re-render this view
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
