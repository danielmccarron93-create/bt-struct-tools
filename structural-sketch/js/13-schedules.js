// ── 3D PHASE 8: AUTO-GENERATED SCHEDULES ─────────────────
// ══════════════════════════════════════════════════════════

const schedOverlay = document.getElementById('schedule-overlay');
const schedBody = document.getElementById('schedule-body');
const schedSummary = document.getElementById('schedule-summary');
let activeScheduleTab = 'padfooting';

document.getElementById('btn-schedules').addEventListener('click', () => {
    openSchedules();
});

// Tab switching
for (const tab of document.querySelectorAll('.schedule-tab')) {
    tab.addEventListener('click', () => {
        activeScheduleTab = tab.dataset.tab;
        document.querySelectorAll('.schedule-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        renderSchedule();
    });
}

function openSchedules() {
    schedOverlay.classList.remove('hidden');
    renderSchedule();
    // Auto-focus the first primary editable field (PF1 size for pad footings)
    setTimeout(() => {
        const firstPrimary = document.querySelector('#schedule-tbody .dim-value[data-primary="1"]');
        if (firstPrimary) {
            firstPrimary.click();
        } else {
            const firstDim = document.querySelector('#schedule-tbody .dim-value');
            if (firstDim) firstDim.click();
        }
    }, 50);
}

document.getElementById('schedule-close').addEventListener('click', () => {
    schedOverlay.classList.add('hidden');
});
schedOverlay.addEventListener('click', (e) => {
    if (e.target === schedOverlay) schedOverlay.classList.add('hidden');
});

// ── NEW SCHEDULE EDITOR: Type-based system ────────────────

function getScheduleColWidths(category) {
    // Returns column width percentages: [type, colour, ...fields, count, delete]
    if (category === 'padfooting') {
        // mark, colour, dimensions, depth, reinf, setdown + count + delete
        return ['7%', '4%', '24%', '18%', '18%', '10%', '6%', '5%'];
    } else if (category === 'stripfooting') {
        // type, colour, width, depth, reo, setdown + count + delete
        return ['7%', '5%', '16%', '16%', '18%', '14%', '8%', '6%'];
    } else if (category === 'beam' || category === 'column') {
        // type, colour, sectionType, size, grade, description + count + delete
        return ['6%', '4%', '12%', '22%', '8%', '20%', '6%', '5%'];
    } else if (category === 'wall') {
        // type, colour, wallType, thickness, description + count + delete
        return ['7%', '5%', '18%', '16%', '22%', '8%', '6%'];
    } else if (category === 'bracingWall') {
        // type, colour, bracingType, capacity, minLength, description + count + delete
        return ['7%', '5%', '14%', '12%', '12%', '18%', '8%', '6%'];
    }
    return ['10%', '6%', '24%', '28%', '12%', '8%'];
}

function renderSchedule() {
    const category = activeScheduleTab; // 'padfooting', 'stripfooting', 'beam', 'column', 'wall'
    const types = project.scheduleTypes[category] || {};

    const schedTable = document.getElementById('schedule-table');
    const schedThead = document.getElementById('schedule-thead');
    const schedTbody = document.getElementById('schedule-tbody');

    // Set fixed column widths via colgroup
    const colWidths = getScheduleColWidths(category);
    let existingColgroup = schedTable.querySelector('colgroup');
    if (existingColgroup) existingColgroup.remove();
    const colgroup = document.createElement('colgroup');
    for (const w of colWidths) {
        const col = document.createElement('col');
        col.style.width = w;
        colgroup.appendChild(col);
    }
    schedTable.insertBefore(colgroup, schedThead);

    // Build header based on category
    const headers = getScheduleHeaders(category);
    let headerHtml = '<tr>';
    for (const h of headers) {
        headerHtml += `<th>${h}</th>`;
    }
    headerHtml += '<th>Count</th><th></th></tr>';
    schedThead.innerHTML = headerHtml;

    // Build rows for each type
    let bodyHtml = '';
    for (const [typeRef, typeData] of Object.entries(types)) {
        const count = countElementsOfType(category, typeRef);
        bodyHtml += renderScheduleRow(category, typeRef, typeData, count);
    }
    schedTbody.innerHTML = bodyHtml;

    // Summary
    const totalCount = Object.keys(types).reduce((sum, t) => sum + countElementsOfType(category, t), 0);
    schedSummary.innerHTML = `<span><strong>${totalCount}</strong> ${category} items</span>`;

    // Setup add type button
    const addBtn = document.getElementById('schedule-add-type');
    addBtn.onclick = () => addScheduleType(category);
}

function getScheduleHeaders(category) {
    if (category === 'padfooting') {
        return ['Mark', 'Colour', 'Dimensions', 'Depth', 'Reinf.', 'Setdown'];
    } else if (category === 'stripfooting') {
        return ['Type', 'Colour', 'Width (mm)', 'Depth (mm)', 'Reinforcement', 'Setdown (mm)', 'TOP'];
    } else if (category === 'beam') {
        return ['Type', 'Colour', 'Section', 'Size', 'Grade', 'Description'];
    } else if (category === 'column') {
        return ['Type', 'Colour', 'Section', 'Size', 'Grade', 'Description'];
    } else if (category === 'wall') {
        return ['Type', 'Colour', 'Wall Type', 'Thickness (mm)', 'Description'];
    } else if (category === 'bracingWall') {
        return ['Type', 'Colour', 'Bracing Type', 'Capacity (kN/m)', 'Min Length (mm)', 'Description'];
    }
    return ['Type', 'Colour'];
}

function countElementsOfType(category, typeRef) {
    let count = 0;
    for (const el of project.elements) {
        if (category === 'padfooting' && el.type === 'footing' && (el.typeRef === typeRef || el.mark === typeRef || (!el.typeRef && typeRef === 'PF1'))) {
            count++;
        } else if (category === 'stripfooting' && el.type === 'stripFooting' && (el.typeRef === typeRef || el.tag === typeRef || (!el.typeRef && typeRef === 'SF1'))) {
            count++;
        } else if (category === 'beam' && el.type === 'line' && el.layer === 'S-BEAM' && (el.typeRef === typeRef || el.tag === typeRef || (!el.typeRef && typeRef === 'SB1'))) {
            count++;
        } else if (category === 'column' && el.type === 'column' && (el.typeRef === typeRef || el.tag === typeRef || (!el.typeRef && typeRef === 'SC1'))) {
            count++;
        } else if (category === 'wall' && el.type === 'wall' && (el.typeRef === typeRef || el.tag === typeRef || (!el.typeRef && typeRef === 'BW1'))) {
            count++;
        } else if (category === 'bracingWall' && el.type === 'bracingWall' && (el.typeRef === typeRef || el.tag === typeRef || (!el.typeRef && typeRef === 'BR1'))) {
            count++;
        }
    }
    return count;
}

function renderScheduleRow(category, typeRef, typeData, count) {
    let html = '<tr>';

    // Type name
    html += `<td class="type-col" data-cat="${category}" data-type="${typeRef}">${typeRef}</td>`;

    // Colour swatch
    html += `<td><div class="color-swatch" style="background:${typeData.color};" data-cat="${category}" data-type="${typeRef}" onclick="openColorPicker(event)"></div></td>`;

    // Helper: show value or empty placeholder
    const v = (val) => (val === '' || val === undefined || val === null) ? '' : val;
    const emptyClass = (val) => (val === '' || val === undefined || val === null) ? ' empty-cell' : '';

    // Category-specific fields
    if (category === 'padfooting') {
        const isRect = !!typeData.rect;
        const w = v(typeData.width);
        const l = v(typeData.length);
        const wEmpty = w === '' ? ' empty' : '';
        const lEmpty = l === '' ? ' empty' : '';
        // Shape toggle icon: linked chain = square, broken chain = rectangular
        const toggleSvg = isRect
            ? '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M18 6L6 18"/><path d="M8 6h-2a2 2 0 0 0-2 2v2"/><path d="M16 18h2a2 2 0 0 0 2-2v-2"/></svg>'
            : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M9 17H7a4 4 0 0 1 0-8h2"/><path d="M15 7h2a4 4 0 0 1 0 8h-2"/><line x1="8" y1="12" x2="16" y2="12"/></svg>';
        const toggleTitle = isRect ? 'Rectangular — click to switch to Square' : 'Square — click to switch to Rectangular';

        html += `<td class="pf-size-cell" data-cat="${category}" data-type="${typeRef}">`;
        html += `<div class="size-cell-wrap">`;
        html += `<button class="shape-toggle${isRect ? ' rect' : ''}" title="${toggleTitle}" onclick="togglePfShape('${typeRef}')">${toggleSvg}</button>`;
        // Width value + spinner
        html += `<div class="dim-group">`;
        html += `<span class="dim-value${wEmpty}" data-field="width" data-type="${typeRef}" data-primary="1" onclick="editPfDim(this)">${w || '—'}</span>`;
        html += `<div class="dim-spinner"><button onclick="spinPfDim('${typeRef}','width',100)" title="+100">▲</button><button onclick="spinPfDim('${typeRef}','width',-100)" title="−100">▼</button></div>`;
        html += `</div>`;
        if (isRect) {
            html += `<span class="dim-sep">×</span>`;
            html += `<div class="dim-group">`;
            html += `<span class="dim-value${lEmpty}" data-field="length" data-type="${typeRef}" onclick="editPfDim(this)">${l || '—'}</span>`;
            html += `<div class="dim-spinner"><button onclick="spinPfDim('${typeRef}','length',100)" title="+100">▲</button><button onclick="spinPfDim('${typeRef}','length',-100)" title="−100">▼</button></div>`;
            html += `</div>`;
        } else {
            html += `<span class="sq-label">SQ</span>`;
        }
        html += `</div></td>`;
        // Depth with spinner
        html += `<td class="pf-dim-cell" data-cat="${category}" data-type="${typeRef}">`;
        html += `<div class="size-cell-wrap">`;
        html += `<div class="dim-group">`;
        html += `<span class="dim-value${v(typeData.depth) === '' ? ' empty' : ''}" data-field="depth" data-type="${typeRef}" data-primary="1" onclick="editPfDim(this)">${v(typeData.depth) || '—'}</span>`;
        html += `<div class="dim-spinner"><button onclick="spinPfDim('${typeRef}','depth',100)" title="+100">▲</button><button onclick="spinPfDim('${typeRef}','depth',-100)" title="−100">▼</button></div>`;
        html += `</div></div></td>`;
        html += `<td class="edit-cell${emptyClass(typeData.reo)}" onclick="editScheduleCell(event)">${v(typeData.reo)}</td>`;
        // Setdown with spinner
        html += `<td class="pf-dim-cell" data-cat="${category}" data-type="${typeRef}">`;
        html += `<div class="size-cell-wrap">`;
        html += `<div class="dim-group">`;
        html += `<span class="dim-value${v(typeData.setdown) === '' ? ' empty' : ''}" data-field="setdown" data-type="${typeRef}" onclick="editPfDim(this)">${v(typeData.setdown) || '—'}</span>`;
        html += `<div class="dim-spinner"><button onclick="spinPfDim('${typeRef}','setdown',100)" title="+100">▲</button><button onclick="spinPfDim('${typeRef}','setdown',-100)" title="−100">▼</button></div>`;
        html += `</div></div></td>`;
    } else if (category === 'stripfooting') {
        // Width with spinner
        html += `<td class="pf-dim-cell" data-cat="${category}" data-type="${typeRef}">`;
        html += `<div class="size-cell-wrap"><div class="dim-group">`;
        html += `<span class="dim-value${v(typeData.width) === '' ? ' empty' : ''}" data-field="width" data-type="${typeRef}" data-cat="stripfooting" onclick="editPfDim(this)">${v(typeData.width) || '—'}</span>`;
        html += `<div class="dim-spinner"><button onclick="spinPfDim('${typeRef}','width',100,'stripfooting')" title="+100">▲</button><button onclick="spinPfDim('${typeRef}','width',-100,'stripfooting')" title="−100">▼</button></div>`;
        html += `</div></div></td>`;
        // Depth with spinner
        html += `<td class="pf-dim-cell" data-cat="${category}" data-type="${typeRef}">`;
        html += `<div class="size-cell-wrap"><div class="dim-group">`;
        html += `<span class="dim-value${v(typeData.depth) === '' ? ' empty' : ''}" data-field="depth" data-type="${typeRef}" data-cat="stripfooting" onclick="editPfDim(this)">${v(typeData.depth) || '—'}</span>`;
        html += `<div class="dim-spinner"><button onclick="spinPfDim('${typeRef}','depth',100,'stripfooting')" title="+100">▲</button><button onclick="spinPfDim('${typeRef}','depth',-100,'stripfooting')" title="−100">▼</button></div>`;
        html += `</div></div></td>`;
        html += `<td class="edit-cell${emptyClass(typeData.reo)}" onclick="editScheduleCell(event)">${v(typeData.reo)}</td>`;
        // Setdown with spinner
        html += `<td class="pf-dim-cell" data-cat="${category}" data-type="${typeRef}">`;
        html += `<div class="size-cell-wrap"><div class="dim-group">`;
        html += `<span class="dim-value${v(typeData.setdown) === '' ? ' empty' : ''}" data-field="setdown" data-type="${typeRef}" data-cat="stripfooting" onclick="editPfDim(this)">${v(typeData.setdown) || '—'}</span>`;
        html += `<div class="dim-spinner"><button onclick="spinPfDim('${typeRef}','setdown',100,'stripfooting')" title="+100">▲</button><button onclick="spinPfDim('${typeRef}','setdown',-100,'stripfooting')" title="−100">▼</button></div>`;
        html += `</div></div></td>`;
        // TOP cell (editable text)
        html += `<td class="edit-cell${emptyClass(typeData.top)}" data-primary="1" onclick="editScheduleCell(event)">${v(typeData.top)}</td>`;
    } else if (category === 'beam' || category === 'column') {
        // Section type dropdown
        const curST = v(typeData.sectionType);
        html += `<td><select class="sched-select" data-cat="${category}" data-type="${typeRef}" data-field="sectionType" onchange="onScheduleDropdown(this)">`;
        html += `<option value="">—</option>`;
        for (const st of Object.keys(STEEL_SECTIONS)) {
            html += `<option value="${st}"${curST === st ? ' selected' : ''}>${st}</option>`;
        }
        html += `</select></td>`;
        // Size dropdown (populated based on section type)
        const curSize = v(typeData.size);
        const sizes = curST && STEEL_SECTIONS[curST] ? STEEL_SECTIONS[curST] : [];
        html += `<td><select class="sched-select sched-size-select" data-cat="${category}" data-type="${typeRef}" data-field="size" onchange="onScheduleDropdown(this)">`;
        html += `<option value="">—</option>`;
        for (const sz of sizes) {
            const displayName = formatSectionName(curST, sz);
            html += `<option value="${displayName}"${curSize === displayName ? ' selected' : ''}>${displayName}</option>`;
        }
        html += `</select></td>`;
        // Grade dropdown
        const curGrade = v(typeData.grade) || '300';
        html += `<td><select class="sched-select" data-cat="${category}" data-type="${typeRef}" data-field="grade" onchange="onScheduleDropdown(this)">`;
        for (const g of STEEL_GRADES) {
            html += `<option value="${g}"${curGrade === g ? ' selected' : ''}>${g}</option>`;
        }
        html += `</select></td>`;
        // Description (free text)
        html += `<td class="edit-cell${emptyClass(typeData.description)}" onclick="editScheduleCell(event)">${v(typeData.description)}</td>`;
    } else if (category === 'wall') {
        // Wall type dropdown
        const curWT = v(typeData.wallType);
        html += `<td><select class="sched-select" data-cat="${category}" data-type="${typeRef}" data-field="wallType" onchange="onScheduleDropdown(this)">`;
        html += `<option value="">—</option>`;
        for (const wt of Object.keys(WALL_TYPES)) {
            html += `<option value="${wt}"${curWT === wt ? ' selected' : ''}>${wt}</option>`;
        }
        html += `</select></td>`;
        // Thickness with spinner
        html += `<td class="pf-dim-cell" data-cat="${category}" data-type="${typeRef}">`;
        html += `<div class="size-cell-wrap"><div class="dim-group">`;
        html += `<span class="dim-value${v(typeData.thickness) === '' ? ' empty' : ''}" data-field="thickness" data-type="${typeRef}" data-cat="wall" onclick="editPfDim(this)">${v(typeData.thickness) || '—'}</span>`;
        html += `<div class="dim-spinner"><button onclick="spinPfDim('${typeRef}','thickness',10,'wall')" title="+10">▲</button><button onclick="spinPfDim('${typeRef}','thickness',-10,'wall')" title="−10">▼</button></div>`;
        html += `</div></div></td>`;
        // Description (free text)
        html += `<td class="edit-cell${emptyClass(typeData.description)}" onclick="editScheduleCell(event)">${v(typeData.description)}</td>`;
    } else if (category === 'bracingWall') {
        // Bracing type dropdown
        const bracingTypes = typeof BRACING_TYPES !== 'undefined' ? Object.keys(BRACING_TYPES) : ['g','h-A','h-B','nominal-1','nominal-2'];
        const curBT = v(typeData.bracingType);
        html += `<td><select class="sched-select" data-cat="${category}" data-type="${typeRef}" data-field="bracingType" onchange="onScheduleDropdown(this)">`;
        html += `<option value="">—</option>`;
        for (const bt of bracingTypes) {
            const btData = typeof BRACING_TYPES !== 'undefined' ? BRACING_TYPES[bt] : null;
            const label = btData ? `${bt} (${btData.capacity || '?'} kN/m)` : bt;
            html += `<option value="${bt}"${curBT === bt ? ' selected' : ''}>${label}</option>`;
        }
        html += `</select></td>`;
        // Capacity (read-only from bracing type)
        const capVal = typeof BRACING_TYPES !== 'undefined' && BRACING_TYPES[curBT] ? BRACING_TYPES[curBT].capacity : '—';
        html += `<td style="text-align:center;font-weight:600;">${capVal !== null ? capVal : 'Eng.'}</td>`;
        // Min length with spinner
        html += `<td class="pf-dim-cell" data-cat="${category}" data-type="${typeRef}">`;
        html += `<div class="size-cell-wrap"><div class="dim-group">`;
        html += `<span class="dim-value${v(typeData.minLength) === '' ? ' empty' : ''}" data-field="minLength" data-type="${typeRef}" data-cat="bracingWall" onclick="editPfDim(this)">${v(typeData.minLength) || '—'}</span>`;
        html += `<div class="dim-spinner"><button onclick="spinPfDim('${typeRef}','minLength',50,'bracingWall')" title="+50">▲</button><button onclick="spinPfDim('${typeRef}','minLength',-50,'bracingWall')" title="−50">▼</button></div>`;
        html += `</div></div></td>`;
        // Description (free text)
        html += `<td class="edit-cell${emptyClass(typeData.description)}" onclick="editScheduleCell(event)">${v(typeData.description)}</td>`;
    }

    // Count
    html += `<td style="text-align:center;"><strong>${count}</strong></td>`;

    // Delete button
    html += `<td><button class="delete-btn" onclick="deleteScheduleType('${category}', '${typeRef}')">Delete</button></td>`;

    html += '</tr>';
    return html;
}

function editScheduleCell(e) {
    const cell = e.target.closest('.edit-cell');
    if (!cell || cell.querySelector('input')) return; // already editing

    const oldVal = cell.textContent;
    const input = document.createElement('input');
    input.type = 'text';
    input.value = oldVal;
    input.style.cssText = 'width:100%;padding:2px 4px;font-size:11px;border:1px solid var(--accent);box-sizing:border-box;';

    cell.innerHTML = '';
    cell.appendChild(input);
    input.focus();
    input.select();

    let saved = false;
    function saveEdit(moveTo) {
        if (saved) return;
        saved = true;

        const newVal = input.value.trim();
        cell.textContent = newVal;
        if (newVal === '') cell.classList.add('empty-cell');
        else cell.classList.remove('empty-cell');

        // Extract context from parent row
        const row = cell.closest('tr');
        const typeRef = row.querySelector('.type-col').textContent;
        const category = activeScheduleTab;
        const typeData = project.scheduleTypes[category][typeRef];

        // Find which field was edited based on position
        const cellIndex = Array.from(row.querySelectorAll('td')).indexOf(cell);

        // Map cell index to field name
        if (cellIndex === 0 || cellIndex === 1) return; // Type/Color, skip

        const fieldIndex = cellIndex - 2;
        const fields = getScheduleFields(category);
        const fieldName = fields[fieldIndex];

        if (fieldName && typeData) {
            if (newVal === '') {
                typeData[fieldName] = '';
            } else {
                typeData[fieldName] = isNaN(newVal) ? newVal : parseFloat(newVal);
            }
            engine.requestRender();
        }

        // Tab navigation: move to next/prev editable cell
        if (moveTo) {
            setTimeout(() => {
                const allCells = Array.from(document.querySelectorAll('#schedule-tbody .edit-cell'));
                const idx = allCells.indexOf(cell);
                const nextIdx = moveTo === 'next' ? idx + 1 : idx - 1;
                if (nextIdx >= 0 && nextIdx < allCells.length) {
                    allCells[nextIdx].click();
                }
            }, 10);
        }
    }

    input.addEventListener('blur', () => saveEdit(null));
    input.addEventListener('keydown', (ev) => {
        if (ev.key === 'Tab') {
            ev.preventDefault();
            saveEdit(ev.shiftKey ? 'prev' : 'next');
        } else if (ev.key === 'Enter') {
            ev.preventDefault();
            saveEdit('next');
        } else if (ev.key === 'Escape') {
            saved = true;
            cell.textContent = oldVal;
            if (oldVal === '') cell.classList.add('empty-cell');
        }
    });
}

function getScheduleFields(category) {
    if (category === 'padfooting') {
        // size & depth & setdown handled by dim-value cells; only 'reo' uses generic edit-cell
        return ['_size', '_depth', 'reo', '_setdown'];
    } else if (category === 'stripfooting') {
        // width, depth & setdown handled by dim-value spinners; 'reo' and 'top' use generic edit-cell
        return ['_width', '_depth', 'reo', '_setdown', 'top'];
    } else if (category === 'beam' || category === 'column') {
        // sectionType, size, grade are dropdowns; only 'description' uses generic edit-cell
        return ['_sectionType', '_size', '_grade', 'description'];
    } else if (category === 'wall') {
        // wallType is dropdown, thickness is spinner; only 'description' uses generic edit-cell
        return ['_wallType', '_thickness', 'description'];
    } else if (category === 'bracingWall') {
        // bracingType is dropdown, capacity is read-only, minLength is spinner; only 'description' uses generic edit-cell
        return ['_bracingType', '_capacity', '_minLength', 'description'];
    }
    return [];
}

// ── PAD FOOTING: Custom dimension editing ─────────────────

/** Toggle a pad footing between square and rectangular */
window.togglePfShape = function(typeRef) {
    const td = project.scheduleTypes.padfooting[typeRef];
    if (!td) return;
    td.rect = !td.rect;
    // When switching to square, clear length so it mirrors width
    if (!td.rect) td.length = '';
    renderSchedule();
    engine.requestRender();
};

/** Spin a pad footing dimension up/down by delta (±100) */
window.spinPfDim = function(typeRef, field, delta, cat) {
    const category = cat || 'padfooting';
    const td = project.scheduleTypes[category][typeRef];
    if (!td) return;
    let val = parseFloat(td[field]) || 0;
    val = Math.max(0, val + delta);
    td[field] = val;
    // If square mode pad footing, keep length in sync with width
    if (category === 'padfooting' && !td.rect && field === 'width') td.length = '';
    renderSchedule();
    engine.requestRender();
};

/** Dropdown change handler for schedule selects */
window.onScheduleDropdown = function(select) {
    const cat = select.dataset.cat;
    const typeRef = select.dataset.type;
    const field = select.dataset.field;
    const td = project.scheduleTypes[cat][typeRef];
    if (!td) return;

    td[field] = select.value;

    // When section type changes, reset the size and re-render to update size dropdown
    if (field === 'sectionType') {
        td.size = '';
        renderSchedule();
    }

    // When wall type changes, auto-fill thickness
    if (field === 'wallType' && WALL_TYPES[select.value]) {
        td.thickness = WALL_TYPES[select.value].thickness;
        renderSchedule();
    }

    // When bracing type changes, auto-fill capacity and minLength
    if (field === 'bracingType' && typeof BRACING_TYPES !== 'undefined' && BRACING_TYPES[select.value]) {
        td.capacity = BRACING_TYPES[select.value].capacity;
        // Set default min lengths per type
        if (select.value === 'h-A') td.minLength = 600;
        else if (select.value === 'g' || select.value === 'h-B') td.minLength = 900;
        else if (select.value.startsWith('nominal')) td.minLength = 450;
        else td.minLength = td.minLength || 900;
        td.description = BRACING_TYPES[select.value].desc || '';
        renderSchedule();
        if (typeof updateBracingSummaryPanel === 'function') updateBracingSummaryPanel();
    }

    engine.requestRender();
};

/** Inline-edit a dimension value (works for padfooting, stripfooting, wall) */
window.editPfDim = function(span) {
    if (span.querySelector('input')) return;
    const field = span.dataset.field;
    const typeRef = span.dataset.type;
    const category = span.dataset.cat || 'padfooting';
    const td = project.scheduleTypes[category][typeRef];
    if (!td) return;

    const oldVal = td[field] !== '' && td[field] !== undefined ? String(td[field]) : '';
    const input = document.createElement('input');
    input.type = 'text';
    input.value = oldVal;
    input.style.cssText = 'width:52px;padding:2px 4px;font-size:11px;border:1px solid var(--accent);border-radius:2px;font-family:var(--font-mono);box-sizing:border-box;outline:none;text-align:center;';

    span.textContent = '';
    span.appendChild(input);
    input.focus();
    input.select();

    let saved = false;
    function save(moveTo) {
        if (saved) return;
        saved = true;
        const newVal = input.value.trim();
        if (newVal === '') {
            td[field] = '';
        } else {
            td[field] = isNaN(newVal) ? newVal : parseFloat(newVal);
        }
        // If square mode pad footing and editing width, sync length
        if (category === 'padfooting' && !td.rect && field === 'width') td.length = '';
        engine.requestRender();

        if (moveTo) {
            // Navigate through editable fields
            // For pad footings: Tab/Enter only moves through primary fields (Size → Depth → next row Size → ...)
            // For other types: navigate through all dim-values
            setTimeout(() => {
                const isPrimary = span.dataset.primary === '1';
                const selector = isPrimary
                    ? '#schedule-tbody .dim-value[data-primary="1"]'
                    : '#schedule-tbody .dim-value';
                const allEditable = Array.from(document.querySelectorAll(selector));
                const currentIdx = allEditable.indexOf(span);
                let targetIdx = moveTo === 'next' ? currentIdx + 1 : moveTo === 'prev' ? currentIdx - 1 : -1;
                if (moveTo === 'next-row') {
                    // Jump to first primary field of next row
                    const row = span.closest('tr');
                    const nextRow = row.nextElementSibling;
                    if (nextRow) {
                        const firstDim = isPrimary
                            ? nextRow.querySelector('.dim-value[data-primary="1"]')
                            : nextRow.querySelector('.dim-value');
                        if (firstDim) firstDim.click();
                    }
                    return;
                }
                if (targetIdx >= 0 && targetIdx < allEditable.length) {
                    allEditable[targetIdx].click();
                }
            }, 10);
        } else {
            renderSchedule();
        }
    }

    input.addEventListener('blur', () => save(null));
    input.addEventListener('keydown', (ev) => {
        if (ev.key === 'Tab') {
            ev.preventDefault();
            save(ev.shiftKey ? 'prev' : 'next');
        } else if (ev.key === 'Enter') {
            ev.preventDefault();
            // Enter: for primary fields, jump to next primary (Size→Depth→next row Size)
            // For non-primary, jump to next dim-value
            const isPrimary = span.dataset.primary === '1';
            if (isPrimary) {
                // Check if this is the last primary field in the row
                const row = span.closest('tr');
                const primariesInRow = Array.from(row.querySelectorAll('.dim-value[data-primary="1"]'));
                const isLastPrimary = primariesInRow.indexOf(span) === primariesInRow.length - 1;
                save(isLastPrimary ? 'next-row' : 'next');
            } else {
                const row = span.closest('tr');
                const dimsInRow = Array.from(row.querySelectorAll('.dim-value'));
                const isLast = dimsInRow.indexOf(span) === dimsInRow.length - 1;
                save(isLast ? 'next-row' : 'next');
            }
        } else if (ev.key === 'Escape') {
            saved = true;
            renderSchedule();
        }
    });
};

function openColorPicker(e) {
    const swatch = e.target;
    const cat = swatch.dataset.cat;
    const typeRef = swatch.dataset.type;
    const input = document.createElement('input');
    input.type = 'color';
    input.value = project.scheduleTypes[cat][typeRef].color;
    input.onchange = () => {
        project.scheduleTypes[cat][typeRef].color = input.value;
        swatch.style.background = input.value;
        engine.requestRender();
    };
    input.click();
}

function addScheduleType(category) {
    const types = project.scheduleTypes[category];
    // Determine prefix for this category
    const prefix = category === 'padfooting' ? 'PF' : category === 'stripfooting' ? 'SF' :
                   category === 'beam' ? 'SB' : category === 'column' ? 'SC' : 'BW';
    // Find next available number
    let nextNum = 1;
    while (types[prefix + nextNum]) {
        nextNum++;
    }
    const newTypeRef = prefix + nextNum;

    // Get next color from palette
    const colorIndex = Object.keys(types).length % SCHEDULE_COLORS.length;
    const defaultColor = SCHEDULE_COLORS[colorIndex];

    // Create new type with defaults
    let newType = { color: defaultColor };
    if (category === 'padfooting') {
        newType = { width: '', length: '', depth: '', reo: '', setdown: 200, rect: false, color: defaultColor };
    } else if (category === 'stripfooting') {
        newType = { width: 300, depth: 500, reo: '', setdown: 200, color: defaultColor };
    } else if (category === 'beam') {
        newType = { sectionType: '', size: '', description: '', grade: '300', color: defaultColor };
    } else if (category === 'column') {
        newType = { sectionType: '', size: '', description: '', grade: '300', color: defaultColor };
    } else if (category === 'wall') {
        newType = { wallType: '', thickness: '', description: '', color: defaultColor };
    }

    project.scheduleTypes[category][newTypeRef] = newType;
    renderSchedule();
}

window.deleteScheduleType = function(category, typeRef) {
    if (!confirm(`Delete type ${typeRef}? Elements using this type will revert to the default.`)) return;
    delete project.scheduleTypes[category][typeRef];
    renderSchedule();
    engine.requestRender();
};

window.editScheduleCell = editScheduleCell;
window.openColorPicker = openColorPicker;

// ── CSV Export ───────────────────────────────────────────

document.getElementById('schedule-export-csv').addEventListener('click', () => {
    let csv = '';
    const rows = schedBody.querySelectorAll('tr');
    for (const row of rows) {
        const cells = row.querySelectorAll('th, td');
        csv += Array.from(cells).map(c => '"' + c.textContent.replace(/"/g, '""') + '"').join(',') + '\n';
    }

    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = activeScheduleTab + '-schedule.csv';
    a.click();
    URL.revokeObjectURL(url);
});

// ══════════════════════════════════════════════════════════
// ── LEVELS & RL EDITOR ─────────────────────────────────────
// ══════════════════════════════════════════════════════════

const levelsRLOverlay = document.getElementById('levels-rl-overlay');
const groundRLInput = document.getElementById('ground-rl-input');
const tofToggleBtn = document.getElementById('btn-tof-toggle');

function openLevelsRLEditor() {
    levelsRLOverlay.classList.remove('hidden');
    groundRLInput.value = levelSystem.groundRL || '';
    tofToggleBtn.classList.toggle('active', levelSystem.showTOFTags);
    renderLevelsRLTable();
    drawRLSectionDiagram();
}

document.getElementById('levels-rl-close').addEventListener('click', () => {
    levelsRLOverlay.classList.add('hidden');
});
levelsRLOverlay.addEventListener('click', (e) => {
    if (e.target === levelsRLOverlay) levelsRLOverlay.classList.add('hidden');
});

// Ground RL input
groundRLInput.addEventListener('change', () => {
    const val = parseFloat(groundRLInput.value);
    levelSystem.groundRL = isNaN(val) ? 0 : val;
    renderLevelsRLTable();
    drawRLSectionDiagram();
    engine.requestRender();
});

// TOF toggle
tofToggleBtn.addEventListener('click', () => {
    levelSystem.showTOFTags = !levelSystem.showTOFTags;
    tofToggleBtn.classList.toggle('active', levelSystem.showTOFTags);
    engine.requestRender();
});

function renderLevelsRLTable() {
    const tbody = document.getElementById('levels-rl-tbody');
    let html = '';
    for (let i = 0; i < levelSystem.levels.length; i++) {
        const lv = levelSystem.levels[i];
        const rl = levelSystem.groundRL + (lv.elevation / 1000);
        const isGF = i === 0;
        html += '<tr>';
        html += `<td style="font-weight:700;color:var(--accent);">${lv.id}</td>`;
        html += `<td class="level-name-cell">${lv.name}</td>`;
        html += `<td><input type="number" value="${lv.height}" data-idx="${i}" data-field="height" step="100" onchange="updateLevelField(this)" ${lv.height === 0 && i === levelSystem.levels.length - 1 ? 'style="color:var(--text-disabled);"' : ''}></td>`;
        html += `<td style="font-family:var(--font-mono);color:var(--text-secondary);">${lv.elevation}</td>`;
        html += `<td class="rl-cell">RL ${rl.toFixed(3)}</td>`;
        html += `<td style="font-family:var(--font-mono);font-size:11px;color:var(--text-tertiary);">${isGF ? '200 (default)' : '—'}</td>`;
        html += '</tr>';
    }
    tbody.innerHTML = html;
}

window.updateLevelField = function(input) {
    const idx = parseInt(input.dataset.idx);
    const field = input.dataset.field;
    const val = parseFloat(input.value);
    if (isNaN(val) || idx < 0 || idx >= levelSystem.levels.length) return;
    levelSystem.levels[idx][field] = val;
    recalcElevations();
    renderLevelsRLTable();
    drawRLSectionDiagram();
    buildLevelTabs();
    engine.requestRender();
};

function drawRLSectionDiagram() {
    const canvas = document.getElementById('rl-section-canvas');
    const ctx = canvas.getContext('2d');
    const W = canvas.width, H = canvas.height;
    ctx.clearRect(0, 0, W, H);

    const levels = levelSystem.levels;
    if (levels.length === 0) return;

    const topLevel = levels[levels.length - 1];
    const maxElev = topLevel.elevation + (topLevel.height > 0 ? topLevel.height : 1000);
    const minElev = -500; // show a bit below GF for footings

    const margin = { top: 12, bottom: 20, left: 80, right: 30 };
    const plotH = H - margin.top - margin.bottom;
    const plotW = W - margin.left - margin.right;

    function yFromElev(elev) {
        return margin.top + plotH * (1 - (elev - minElev) / (maxElev - minElev));
    }

    // Draw level lines
    for (const lv of levels) {
        const y = yFromElev(lv.elevation);
        const rl = levelSystem.groundRL + (lv.elevation / 1000);

        // Dashed line
        ctx.strokeStyle = '#ccc';
        ctx.lineWidth = 1;
        ctx.setLineDash([4, 3]);
        ctx.beginPath();
        ctx.moveTo(margin.left, y);
        ctx.lineTo(W - margin.right, y);
        ctx.stroke();
        ctx.setLineDash([]);

        // Level name + RL
        ctx.font = '600 11px "Segoe UI", Arial, sans-serif';
        ctx.fillStyle = '#374151';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(lv.name, margin.left - 6, y);

        ctx.font = '500 9px "Consolas", monospace';
        ctx.fillStyle = '#2B7CD0';
        ctx.fillText('RL ' + rl.toFixed(3), margin.left - 6, y + 12);
    }

    // Draw storey height arrows between levels
    for (let i = 0; i < levels.length - 1; i++) {
        const y1 = yFromElev(levels[i].elevation);
        const y2 = yFromElev(levels[i + 1].elevation);
        const xMid = margin.left + plotW * 0.5;

        // Vertical line
        ctx.strokeStyle = '#aaa';
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(xMid, y1 - 1);
        ctx.lineTo(xMid, y2 + 1);
        ctx.stroke();

        // Arrow tips
        ctx.fillStyle = '#aaa';
        ctx.beginPath(); ctx.moveTo(xMid, y1 - 1); ctx.lineTo(xMid - 3, y1 - 6); ctx.lineTo(xMid + 3, y1 - 6); ctx.closePath(); ctx.fill();
        ctx.beginPath(); ctx.moveTo(xMid, y2 + 1); ctx.lineTo(xMid - 3, y2 + 6); ctx.lineTo(xMid + 3, y2 + 6); ctx.closePath(); ctx.fill();

        // Height text
        ctx.font = '500 10px "Consolas", monospace';
        ctx.fillStyle = '#555';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(levels[i].height + ' mm', xMid + 40, (y1 + y2) / 2);
    }

    // Draw schematic building outline
    const bx1 = margin.left + plotW * 0.15;
    const bx2 = margin.left + plotW * 0.85;
    for (let i = 0; i < levels.length; i++) {
        const y = yFromElev(levels[i].elevation);
        // Slab line (thicker)
        ctx.strokeStyle = '#6b7280';
        ctx.lineWidth = 3;
        ctx.beginPath();
        ctx.moveTo(bx1, y);
        ctx.lineTo(bx2, y);
        ctx.stroke();
    }
    // Vertical walls
    if (levels.length > 1) {
        const yTop = yFromElev(levels[levels.length - 1].elevation);
        const yBot = yFromElev(levels[0].elevation);
        ctx.strokeStyle = '#9ca3af';
        ctx.lineWidth = 1.5;
        ctx.beginPath();
        ctx.moveTo(bx1, yTop); ctx.lineTo(bx1, yBot);
        ctx.moveTo(bx2, yTop); ctx.lineTo(bx2, yBot);
        ctx.stroke();
    }

    // Default TOF line if GF RL is set
    if (levelSystem.groundRL !== 0) {
        const tofElev = -200; // default setdown
        const y = yFromElev(tofElev);
        ctx.strokeStyle = '#E68A00';
        ctx.lineWidth = 1;
        ctx.setLineDash([3, 2]);
        ctx.beginPath();
        ctx.moveTo(bx1 - 10, y);
        ctx.lineTo(bx2 + 10, y);
        ctx.stroke();
        ctx.setLineDash([]);
        ctx.font = '600 9px "Segoe UI", Arial, sans-serif';
        ctx.fillStyle = '#E68A00';
        ctx.textAlign = 'left';
        ctx.fillText('TOF (default)', bx2 + 14, y);
        ctx.font = '500 8px "Consolas", monospace';
        ctx.fillText('RL ' + (levelSystem.groundRL - 0.2).toFixed(3), bx2 + 14, y + 10);
    }
}

// ══════════════════════════════════════════════════════════
// ── TOF TAG RENDERING ON PLAN ────────────────────────────
// ══════════════════════════════════════════════════════════

engine.onRender((ctx, eng) => {
    if (!levelSystem.showTOFTags || levelSystem.groundRL === 0) return;
    const zoom = eng.viewport.zoom;

    for (const el of project.getVisibleElements()) {
        if (el.type !== 'footing' && el.type !== 'stripFooting') continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const tof = getFootingTOF(el);
        const hasOverride = el.tofOverride !== undefined && el.tofOverride !== null && el.tofOverride !== '';

        let pos;
        if (el.type === 'footing') {
            pos = eng.coords.realToScreen(el.x, el.y);
        } else {
            // Strip footing — use midpoint
            const mx = (el.x1 + el.x2) / 2, my = (el.y1 + el.y2) / 2;
            pos = eng.coords.realToScreen(mx, my);
        }

        // Draw TOF tag
        const tagText = 'TOF ' + tof.toFixed(3);
        const fontSize = Math.max(7, 2.5 * zoom);
        ctx.font = `600 ${fontSize}px "Consolas", monospace`;
        const textW = ctx.measureText(tagText).width;
        const padX = 3, padY = 2;

        // Get footing half-width for positioning
        let halfW = 10;
        if (el.type === 'footing') {
            const typeRef = el.typeRef || el.mark || 'PF1';
            const td = project.scheduleTypes.padfooting[typeRef] || {};
            const fw = td.width || el.footingWidth || 1000;
            halfW = (fw / 2 / CONFIG.drawingScale) * zoom;
        }

        const tx = pos.x + halfW + 4;
        const ty = pos.y - fontSize / 2 - padY;

        // Background
        ctx.fillStyle = hasOverride ? 'rgba(230,138,0,0.12)' : 'rgba(255,255,255,0.92)';
        ctx.strokeStyle = hasOverride ? '#E68A00' : '#aaa';
        ctx.lineWidth = 0.8;
        const rr = 2;
        ctx.beginPath();
        if (ctx.roundRect) ctx.roundRect(tx - padX, ty - padY, textW + padX * 2, fontSize + padY * 2, rr);
        else ctx.rect(tx - padX, ty - padY, textW + padX * 2, fontSize + padY * 2);
        ctx.fill();
        ctx.stroke();

        // Text
        ctx.fillStyle = hasOverride ? '#B86E00' : '#555';
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';
        ctx.fillText(tagText, tx, ty);
    }
});

// ══════════════════════════════════════════════════════════
