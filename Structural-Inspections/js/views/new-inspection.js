/**
 * "New inspection" modal flow.
 *
 * Exports `openNewInspection({ projectId })` which opens the modal and
 * returns the newly-created inspection record (or null if cancelled).
 *
 * If projectId is supplied, the project select is pre-selected and locked
 * (the user is already inside a project). Otherwise they pick from all
 * available projects.
 */

import {
  listProjects, listInspectionTypes, listDrawingsForProject,
  createInspection, getUserProfile
} from '../db.js';
import { openModal } from '../components/modal.js';

export async function openNewInspection({ projectId = null } = {}) {
  const [projects, types] = await Promise.all([
    listProjects(),
    listInspectionTypes()
  ]);

  if (projects.length === 0) {
    return openModal({
      title: 'No projects yet',
      bodyHtml: `<p>Create a project first and upload at least one structural drawing against it — then you can start an inspection against that drawing.</p>`,
      primaryLabel: 'OK',
      cancelLabel: 'Close',
      onConfirm: () => true
    }).then(() => null);
  }

  // If we have a preselected project with no drawings, bail early with a
  // friendly message rather than a disabled form.
  let initialDrawings = [];
  if (projectId) {
    initialDrawings = await listDrawingsForProject(projectId);
    if (initialDrawings.length === 0) {
      await openModal({
        title: 'Upload a drawing first',
        bodyHtml: `<p>This project has no drawings yet. Upload the structural PDF you'll be inspecting against, then start the inspection.</p>`,
        primaryLabel: 'OK',
        cancelLabel: 'Close',
        onConfirm: () => true
      });
      return null;
    }
  }

  const today = new Date().toISOString().slice(0, 10);

  // Group inspection types by category for optgroup display
  const grouped = new Map();
  for (const t of types) {
    if (!grouped.has(t.category)) grouped.set(t.category, []);
    grouped.get(t.category).push(t);
  }

  const bodyHtml = `
    <div class="form">
      <div class="form-field">
        <label for="ni-project">Project <span class="req">*</span></label>
        <select id="ni-project" name="projectId" required ${projectId ? 'disabled' : ''}>
          ${projectId ? '' : '<option value="">Choose project…</option>'}
          ${projects.map((p) => `
            <option value="${p.id}" ${Number(projectId) === p.id ? 'selected' : ''}>
              ${escapeHtml(p.jobNumber)} — ${escapeHtml(p.name)}
            </option>
          `).join('')}
        </select>
      </div>

      <div class="form-field">
        <label for="ni-type">Inspection type <span class="req">*</span></label>
        <select id="ni-type" name="inspectionTypeKey" required>
          <option value="">Choose type…</option>
          ${Array.from(grouped.entries()).map(([cat, items]) => `
            <optgroup label="${escapeHtml(cat)}">
              ${items.map((t) => `<option value="${escapeAttr(t.key)}">${escapeHtml(t.name)}</option>`).join('')}
            </optgroup>
          `).join('')}
        </select>
      </div>

      <div class="form-field">
        <label for="ni-drawing">Primary drawing <span class="req">*</span></label>
        <select id="ni-drawing" name="primaryDrawingId" required ${initialDrawings.length === 0 ? 'disabled' : ''}>
          ${projectId
            ? `<option value="">Choose drawing…</option>${initialDrawings.map(drawingOption).join('')}`
            : '<option value="">Select project first</option>'}
        </select>
        <div class="form-field__help muted">The plan you'll mark up. You can reference other drawings in your comments.</div>
      </div>

      <div class="form-field">
        <label for="ni-date">Date <span class="req">*</span></label>
        <input id="ni-date" name="date" type="date" required value="${today}"/>
      </div>

      <div class="form-field">
        <label for="ni-inspector">Inspector</label>
        <input id="ni-inspector" name="inspectorName" type="text" placeholder="Your name (for the report)"/>
      </div>

      <div class="form-field">
        <label for="ni-attendees">Attendees</label>
        <input id="ni-attendees" name="attendees" type="text" placeholder="e.g. Site supervisor, concreter"/>
      </div>

      <div class="form-field">
        <label for="ni-weather">Weather</label>
        <input id="ni-weather" name="weather" type="text" placeholder="e.g. Fine, 24°C"/>
      </div>

      <div class="form-field">
        <label for="ni-notes">Notes</label>
        <textarea id="ni-notes" name="notes" rows="2" placeholder="Any general context"></textarea>
      </div>
    </div>
  `;

  return openModal({
    title: 'New inspection',
    primaryLabel: 'Start inspection',
    bodyHtml,
    onMount: (dialog) => {
      // If project wasn't preselected, wire up the cascade:
      // changing project → reloads drawing options.
      const projectSelect = dialog.querySelector('[name=projectId]');
      const drawingSelect = dialog.querySelector('[name=primaryDrawingId]');

      if (!projectId) {
        projectSelect.addEventListener('change', async (e) => {
          const pid = Number(e.target.value);
          if (!pid) {
            drawingSelect.innerHTML = '<option value="">Select project first</option>';
            drawingSelect.disabled = true;
            return;
          }
          const drawings = await listDrawingsForProject(pid);
          if (drawings.length === 0) {
            drawingSelect.innerHTML = '<option value="">No drawings in this project</option>';
            drawingSelect.disabled = true;
          } else {
            drawingSelect.innerHTML =
              `<option value="">Choose drawing…</option>` +
              drawings.map(drawingOption).join('');
            drawingSelect.disabled = false;
          }
        });
      }

      // Pre-fill inspector name from the saved profile (Phase 8). Falls back
      // to localStorage for anyone who had a name before the profile was added.
      (async () => {
        let name = '';
        try {
          const profile = await getUserProfile();
          name = (profile?.name || '').trim();
        } catch {}
        if (!name) {
          try { name = localStorage.getItem('bt.lastInspectorName') || ''; } catch {}
        }
        const field = dialog.querySelector('[name=inspectorName]');
        if (field && name && !field.value) field.value = name;
      })();
    },
    onConfirm: async (dialog) => {
      const data = {
        projectId:          Number(dialog.querySelector('[name=projectId]').value),
        inspectionTypeKey:  dialog.querySelector('[name=inspectionTypeKey]').value,
        primaryDrawingId:   Number(dialog.querySelector('[name=primaryDrawingId]').value),
        date:               dialog.querySelector('[name=date]').value,
        inspectorName:      dialog.querySelector('[name=inspectorName]').value,
        attendees:          dialog.querySelector('[name=attendees]').value,
        weather:            dialog.querySelector('[name=weather]').value,
        notes:              dialog.querySelector('[name=notes]').value
      };
      const inspection = await createInspection(data);
      // Remember inspector for next time
      try { localStorage.setItem('bt.lastInspectorName', data.inspectorName); } catch {}
      return inspection;
    }
  });
}

/**
 * Truncate any loose free-text description that made it past extraction. Keeps
 * the option line a single glanceable row even if a drawing's description is
 * long or noisy.
 */
function drawingOption(d) {
  const desc = (d.description || '').trim();
  const descShort = desc.length > 48 ? desc.slice(0, 46) + '…' : desc;
  const sheet = d.sheetNumber || `Page ${d.pageNumber || 1}`;
  const label = `${sheet}${d.revision ? ` (Rev ${d.revision})` : ''}${descShort ? ' — ' + descShort : ''}`;
  return `<option value="${d.id}">${escapeHtml(label)}</option>`;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
function escapeAttr(s) { return escapeHtml(s).replace(/"/g, '&quot;'); }
