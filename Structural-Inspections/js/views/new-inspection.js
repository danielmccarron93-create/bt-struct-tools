/**
 * "New inspection" modal flow (Phase 11 rebuild).
 *
 * What changed vs. the v2.0 version:
 *
 *   - Inspection type is now a search-as-you-type combobox supporting
 *     MULTI-SELECT — one site visit can cover several inspection types
 *     (e.g. "blockwork wall reinforcement" + "concrete columns" both
 *     happening on the same level on the same morning).
 *   - The drawings to take to site are auto-selected based on the chosen
 *     types' affinity tags (drawings.applicableInspectionTypes — populated
 *     at import time by Cowork).
 *   - The user can tick / untick to override the smart pick.
 *   - The inspector field is the engineer dropdown from Phase 11.
 *
 * Shape of the result returned to the caller is the same as before
 * (the inspection record). Schema-wise we add `types[]` and `drawingIds[]`
 * onto the record — the v5 Dexie schema already accommodates them.
 */

import {
  listProjects, listInspectionTypes, listDrawingsForProject,
  createInspection, getUserProfile, addEngineer, db, getProject
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

  // If we have a preselected project with no drawings, bail early.
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

  // Pull the project's plan (if it has one) so we can flag the "in plan"
  // inspection types in the search list — those float to the top.
  let planTypeKeys = new Set();
  if (projectId) {
    const project = await getProject(projectId);
    if (project?.inspectionPlan?.length) {
      for (const entry of project.inspectionPlan) {
        if (entry.type) planTypeKeys.add(entry.type);
      }
    }
  }

  // Group types by category for visual structure inside the picker.
  const grouped = new Map();
  for (const t of types) {
    if (!grouped.has(t.category)) grouped.set(t.category, []);
    grouped.get(t.category).push(t);
  }

  const bodyHtml = `
    <div class="form ni-form">
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
        <label>Inspection types <span class="req">*</span></label>
        <input id="ni-type-search" type="search" placeholder="Search types — e.g. blockwork, slab, footing"
               autocomplete="off" class="ni-type-search"/>
        <div class="ni-type-picker" id="ni-type-list" role="listbox" aria-multiselectable="true">
          ${renderTypeList(grouped, planTypeKeys, '')}
        </div>
        <div class="ni-type-summary muted small" id="ni-type-summary">
          0 selected — pick one or more.
        </div>
      </div>

      <div class="form-field">
        <label>Drawings to take to site <span class="req">*</span></label>
        <div class="ni-drawing-list" id="ni-drawing-list" role="listbox" aria-multiselectable="true">
          ${initialDrawings.length === 0
            ? `<p class="muted small">${projectId ? 'No drawings in this project — upload some first.' : 'Choose a project to see its drawings.'}</p>`
            : initialDrawings.map((d) => renderDrawingCheckbox(d, false)).join('')}
        </div>
        <div class="form-field__help muted small" id="ni-drawing-hint">
          Pick the drawings you'll mark up on site. We pre-tick the relevant ones once you choose your inspection type(s).
        </div>
      </div>

      <div class="form-field">
        <label for="ni-date">Date <span class="req">*</span></label>
        <input id="ni-date" name="date" type="date" required value="${today}"/>
      </div>

      <div class="form-field">
        <label for="ni-inspector">Inspector</label>
        <div class="input-cluster">
          <select id="ni-inspector" name="inspectorName">
            <!-- options populated in onMount -->
          </select>
          <button class="btn btn--ghost btn--sm" type="button" id="ni-inspector-add" title="Add new engineer">+ Add</button>
        </div>
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

  // Local state shared by the modal's event handlers — `selectedTypes` and
  // `selectedDrawings` are the source of truth for what gets submitted.
  const state = {
    drawings:          initialDrawings.slice(),    // refreshed when project changes
    selectedTypes:     new Set(),                  // type keys
    selectedDrawings:  new Set(),                  // drawing ids
    autoSelected:      new Set(),                  // ids that are ticked automatically
                                                   // — distinguishes "user kept" from "user added"
    typeFilter:        ''
  };

  return openModal({
    title: 'New inspection',
    primaryLabel: 'Start inspection',
    bodyHtml,
    onMount: (dialog) => {
      const projectSelect = dialog.querySelector('[name=projectId]');
      const typeListEl    = dialog.querySelector('#ni-type-list');
      const typeSearchEl  = dialog.querySelector('#ni-type-search');
      const typeSummaryEl = dialog.querySelector('#ni-type-summary');
      const drawingListEl = dialog.querySelector('#ni-drawing-list');

      // --- Type list interactions ----------------------------------------
      function refreshTypeList() {
        typeListEl.innerHTML = renderTypeList(grouped, planTypeKeys, state.typeFilter, state.selectedTypes);
      }
      function refreshTypeSummary() {
        const n = state.selectedTypes.size;
        if (n === 0) {
          typeSummaryEl.textContent = '0 selected — pick one or more.';
        } else {
          const labels = Array.from(state.selectedTypes)
            .map((k) => types.find((t) => t.key === k)?.name || k)
            .filter(Boolean);
          typeSummaryEl.textContent = `${n} selected: ${labels.join(', ')}`;
        }
      }

      typeSearchEl.addEventListener('input', () => {
        state.typeFilter = typeSearchEl.value.trim().toLowerCase();
        refreshTypeList();
      });
      typeListEl.addEventListener('click', (e) => {
        const item = e.target.closest('[data-type-key]');
        if (!item) return;
        const key = item.dataset.typeKey;
        if (state.selectedTypes.has(key)) state.selectedTypes.delete(key);
        else state.selectedTypes.add(key);
        refreshTypeList();
        refreshTypeSummary();
        recomputeAutoSelectedDrawings();
        renderDrawingChecks();
      });

      // --- Drawing interactions ------------------------------------------
      function recomputeAutoSelectedDrawings() {
        // For every drawing whose applicableInspectionTypes overlaps the
        // selected types, mark it as auto-selected (and, if the user hasn't
        // explicitly unticked it, add to selectedDrawings).
        const newAuto = new Set();
        for (const d of state.drawings) {
          const applies = (d.applicableInspectionTypes || [])
            .some((t) => state.selectedTypes.has(t));
          if (applies) newAuto.add(d.id);
        }
        // Drawings auto-selected before but no longer auto-applicable: drop them
        // from selectedDrawings UNLESS the user has manually re-ticked them
        // (which we don't currently track distinct from auto). Simple rule for
        // v1: dropped auto-selections also drop from selected.
        for (const id of state.autoSelected) {
          if (!newAuto.has(id)) state.selectedDrawings.delete(id);
        }
        // Add any new auto-selections to the selected set.
        for (const id of newAuto) state.selectedDrawings.add(id);
        state.autoSelected = newAuto;
      }
      function renderDrawingChecks() {
        if (state.drawings.length === 0) {
          drawingListEl.innerHTML = `<p class="muted small">${projectSelect.value ? 'No drawings in this project — upload some first.' : 'Choose a project to see its drawings.'}</p>`;
          return;
        }
        drawingListEl.innerHTML = state.drawings
          .map((d) => renderDrawingCheckbox(d, state.selectedDrawings.has(d.id), state.autoSelected.has(d.id)))
          .join('');
        updateDrawingHint();
      }
      function updateDrawingHint() {
        const hint = dialog.querySelector('#ni-drawing-hint');
        if (!hint) return;
        const total = state.drawings.length;
        const selected = state.selectedDrawings.size;
        const auto = state.autoSelected.size;
        if (state.selectedTypes.size === 0) {
          hint.textContent = `${total} drawings available — pick inspection types to auto-select the relevant ones.`;
        } else if (auto === 0) {
          hint.textContent = `No drawings tagged for these inspection types — tick the ones you want manually.`;
        } else {
          hint.textContent = `${selected} selected (${auto} auto-picked from your inspection types).`;
        }
      }

      drawingListEl.addEventListener('click', (e) => {
        const row = e.target.closest('[data-drawing-id]');
        if (!row) return;
        const id = Number(row.dataset.drawingId);
        if (state.selectedDrawings.has(id)) state.selectedDrawings.delete(id);
        else state.selectedDrawings.add(id);
        // If user toggled an auto-pick, the autoSelected set diverges — track:
        // we leave autoSelected as-is so the visual cue ("auto") still shows.
        renderDrawingChecks();
      });

      // --- Project change → refresh drawings + re-evaluate auto-pick -----
      projectSelect.addEventListener('change', async (e) => {
        const pid = Number(e.target.value);
        if (!pid) {
          state.drawings = [];
          state.selectedDrawings = new Set();
          state.autoSelected = new Set();
          renderDrawingChecks();
          return;
        }
        const drawings = await listDrawingsForProject(pid);
        state.drawings = drawings;
        state.selectedDrawings = new Set();
        state.autoSelected = new Set();
        recomputeAutoSelectedDrawings();
        renderDrawingChecks();

        // Refresh the type list with this project's plan badges.
        const project = await getProject(pid);
        planTypeKeys = new Set(
          (project?.inspectionPlan || []).map((e) => e.type).filter(Boolean)
        );
        refreshTypeList();
      });

      // First render
      refreshTypeSummary();
      // Initial state when project is preselected — recompute drawings.
      if (projectId) {
        recomputeAutoSelectedDrawings();
        renderDrawingChecks();
      }

      // --- Inspector dropdown --------------------------------------------
      const populateInspectorDropdown = async (preselect) => {
        const select = dialog.querySelector('[name=inspectorName]');
        if (!select) return;
        const profile = await getUserProfile().catch(() => null);
        const primary = (profile?.name || '').trim();
        const others = Array.isArray(profile?.engineers) ? profile.engineers : [];
        const lastUsed = (() => {
          try { return localStorage.getItem('bt.lastInspectorName') || ''; } catch { return ''; }
        })();
        const names = [];
        const pushUnique = (n) => {
          if (n && !names.some((x) => x.toLowerCase() === n.toLowerCase())) names.push(n);
        };
        pushUnique(primary);
        for (const e of others) pushUnique(e.name);
        pushUnique(lastUsed);
        const value = preselect || primary || lastUsed || '';
        select.innerHTML = `
          ${names.map((n) =>
            `<option value="${escapeAttr(n)}" ${n === value ? 'selected' : ''}>${escapeHtml(n)}</option>`
          ).join('')}
          ${names.length === 0 ? '<option value="">(none — tap + Add)</option>' : ''}
        `;
      };
      populateInspectorDropdown();

      const addBtn = dialog.querySelector('#ni-inspector-add');
      if (addBtn) {
        addBtn.addEventListener('click', async () => {
          const name = window.prompt('Engineer name:');
          if (!name || !name.trim()) return;
          const rpeq = window.prompt('RPEQ number (optional):') || '';
          try {
            await addEngineer({ name, rpeq });
            await populateInspectorDropdown(name.trim());
          } catch (err) {
            alert(err.message || 'Couldn\u2019t add engineer');
          }
        });
      }
    },

    onConfirm: async (dialog) => {
      const pid = Number(dialog.querySelector('[name=projectId]').value);
      if (!pid) throw new Error('Pick a project');

      const typeKeys = Array.from(state.selectedTypes);
      if (typeKeys.length === 0) throw new Error('Pick at least one inspection type');

      const drawingIds = Array.from(state.selectedDrawings);
      if (drawingIds.length === 0) throw new Error('Pick at least one drawing to take on site');

      // Sort drawing IDs by sheet number / page so the lowest-page drawing is
      // the visual primary the markup viewer opens to.
      const drawingsById = new Map(state.drawings.map((d) => [d.id, d]));
      drawingIds.sort((a, b) => {
        const da = drawingsById.get(a), dbw = drawingsById.get(b);
        const sa = (da?.sheetNumber || '').trim();
        const sb = (dbw?.sheetNumber || '').trim();
        if (sa && sb) {
          const cmp = sa.localeCompare(sb, undefined, { numeric: true });
          if (cmp !== 0) return cmp;
        }
        return (da?.pageNumber || 0) - (dbw?.pageNumber || 0);
      });

      const data = {
        projectId:          pid,
        inspectionTypeKey:  typeKeys[0],          // primary type — drives display name + boilerplate
        primaryDrawingId:   drawingIds[0],        // first sheet by sort
        date:               dialog.querySelector('[name=date]').value,
        inspectorName:      dialog.querySelector('[name=inspectorName]').value,
        attendees:          dialog.querySelector('[name=attendees]').value,
        weather:            dialog.querySelector('[name=weather]').value,
        notes:              dialog.querySelector('[name=notes]').value
      };
      const inspection = await createInspection(data);

      // Add multi-type / multi-drawing fields. Schema accommodates them; the
      // markup viewer still opens primaryDrawingId, but inspection-detail and
      // future per-type sections can read from `types[]` and `drawingIds[]`.
      await db.inspections.update(inspection.id, {
        types:      typeKeys,
        drawingIds: drawingIds
      });

      // Remember inspector for next time
      try { localStorage.setItem('bt.lastInspectorName', data.inspectorName); } catch {}
      return { ...inspection, types: typeKeys, drawingIds };
    }
  });
}

/* --------------------------------------------------------------------------
   Renderers
   -------------------------------------------------------------------------- */

function renderTypeList(grouped, planTypeKeys, filter, selectedTypes) {
  // Filter on type.name + keywords (when present) — cheap substring match.
  const matches = (t) => {
    if (!filter) return true;
    const hay = `${t.name} ${t.category} ${(t.keywords || []).join(' ')}`.toLowerCase();
    return hay.includes(filter);
  };

  const inPlan = [];
  const others = [];
  for (const [cat, items] of grouped) {
    for (const t of items) {
      if (!matches(t)) continue;
      if (planTypeKeys.has(t.key)) inPlan.push(t);
      else others.push({ ...t, _category: cat });
    }
  }

  const renderItem = (t, opts = {}) => {
    const checked = selectedTypes && selectedTypes.has(t.key);
    return `
      <button type="button" class="ni-type-item ${checked ? 'is-selected' : ''}"
              data-type-key="${escapeAttr(t.key)}" role="option" aria-selected="${checked ? 'true' : 'false'}">
        <span class="ni-type-item__check" aria-hidden="true">${checked ? '✓' : ''}</span>
        <span class="ni-type-item__name">${escapeHtml(t.name)}</span>
        ${opts.inPlan ? '<span class="badge badge--phase">in plan</span>' : ''}
        <span class="ni-type-item__cat muted small">${escapeHtml(opts.cat || t.category)}</span>
      </button>
    `;
  };

  const planSection = inPlan.length === 0 ? '' : `
    <div class="ni-type-group">
      <div class="ni-type-group__title">From this project's plan</div>
      ${inPlan.map((t) => renderItem(t, { inPlan: true })).join('')}
    </div>
  `;

  // Re-group "others" by category for display.
  const otherByCat = new Map();
  for (const t of others) {
    if (!otherByCat.has(t._category)) otherByCat.set(t._category, []);
    otherByCat.get(t._category).push(t);
  }
  const otherSection = Array.from(otherByCat.entries()).map(([cat, ts]) => `
    <div class="ni-type-group">
      <div class="ni-type-group__title">${escapeHtml(cat)}</div>
      ${ts.map((t) => renderItem(t, { cat })).join('')}
    </div>
  `).join('');

  if (!planSection && !otherSection) {
    return `<p class="muted small">No inspection types match "${escapeHtml(filter)}".</p>`;
  }
  return planSection + otherSection;
}

function renderDrawingCheckbox(d, isSelected, isAuto) {
  const sheet = d.sheetNumber || `Page ${d.pageNumber || 1}`;
  const desc  = (d.description || '').trim();
  const rev   = d.revision ? ` (Rev ${escapeHtml(d.revision)})` : '';
  const titleLine = `${escapeHtml(sheet)}${desc ? ' — ' + escapeHtml(desc) : ''}${rev}`;
  const elementTags = (d.tags || [])
    .filter((t) => t.startsWith('element:'))
    .slice(0, 3)
    .map((t) => t.slice(8));

  return `
    <button type="button"
            class="ni-drawing-item ${isSelected ? 'is-selected' : ''} ${isAuto ? 'is-auto' : ''}"
            data-drawing-id="${d.id}" role="option" aria-selected="${isSelected ? 'true' : 'false'}">
      <span class="ni-drawing-item__check" aria-hidden="true">${isSelected ? '✓' : ''}</span>
      <span class="ni-drawing-item__main">
        <span class="ni-drawing-item__title">${titleLine}</span>
        ${elementTags.length ? `
          <span class="ni-drawing-item__chips">
            ${elementTags.map((t) => `<span class="chip chip--xs">${escapeHtml(t)}</span>`).join('')}
          </span>` : ''}
      </span>
      ${isAuto ? '<span class="badge badge--auto">auto</span>' : ''}
    </button>
  `;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
function escapeAttr(s) { return escapeHtml(s).replace(/"/g, '&quot;'); }
