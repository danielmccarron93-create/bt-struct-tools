// ── 3D PHASE 7: PROPERTIES PANEL ─────────────────────────
// ══════════════════════════════════════════════════════════

const propsPanel = document.getElementById('props-panel');
const propsType = document.getElementById('props-type');
const propsBody = document.getElementById('props-body');

// CRITICAL: Stop clicks inside the properties panel from propagating
// to the canvas container, which would deselect the element and close the panel.
propsPanel.addEventListener('mousedown', (e) => { e.stopPropagation(); });
propsPanel.addEventListener('dblclick', (e) => { e.stopPropagation(); });
propsPanel.addEventListener('click', (e) => { e.stopPropagation(); });

let _lastPropsElement = null;

function updatePropsPanel() {
    const el = selectedElement;

    if (!el) {
        propsPanel.classList.add('hidden');
        _lastPropsElement = null;
        return;
    }

    propsPanel.classList.remove('hidden');
    if (el === _lastPropsElement) return; // no change
    _lastPropsElement = el;

    // Type label
    const typeNames = {
        line: 'Line', polyline: 'Polyline', column: 'Column',
        text: 'Text', dimension: 'Dimension', leader: 'Leader',
        cloud: 'Cloud', notesbox: 'Notes Panel', table: 'Table',
        slabcallout: 'Slab Callout', footing: 'Pad Footing',
        stripFooting: 'Strip Footing', wall: 'Wall',
        borehole: 'Borehole', rlmarker: 'RL Marker'
    };
    let displayType = typeNames[el.type] || el.type;
    if (el.type === 'line' && el.layer === 'S-BEAM') displayType = 'Beam';
    else if (el.type === 'line' && el.layer === 'S-WALL') displayType = 'Wall';
    propsType.textContent = displayType;

    // Build property rows
    let html = '';

    // ── Common properties ──
    const layerName = project.layers[el.layer]?.name || el.layer;
    html += propRow('Layer', layerName);

    const lv = levelSystem.levels.find(l => l.id === el.level);
    html += propRow('Level', lv ? lv.name : (el.level || '—'));

    // ── Type-specific properties ──
    if (el.type === 'line') {
        const len = Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2));
        html += propDivider();
        html += propRow('Length', fmtLen(len));
        if (el.layer === 'S-BEAM' && (el.tag || el.typeRef)) {
            const bTypeRef = el.typeRef || el.tag;
            const bSchedData = project.scheduleTypes.beam[bTypeRef] || {};
            const bColor = bSchedData.color || '#93C5FD';
            html += `<div class="prop-row"><span class="prop-label">Type</span><span class="prop-type-badge" onclick="showTypeReassignmentPicker(selectedElement,'beam',150,80)" title="Click to change type"><span class="color-dot" style="background:${bColor}"></span>${bTypeRef}</span></div>`;
            html += propDivider();
            // Editable section type
            const secDisplay = bSchedData.sectionType || '<em style="color:#999">click to set</em>';
            html += `<div class="prop-row"><span class="prop-label">Section</span><span class="prop-value prop-click-edit" onclick="propEditSteelSection('${el.id}','beam')">${secDisplay}</span></div>`;
            // Editable size
            const sizeDisplay = bSchedData.size ? formatSectionName(bSchedData.sectionType, bSchedData.size) : '<em style="color:#999">click to set</em>';
            html += `<div class="prop-row"><span class="prop-label">Size</span><span class="prop-value prop-click-edit" onclick="propEditSteelSize('${el.id}','beam')" style="font-weight:600;">${sizeDisplay}</span></div>`;
            // Editable grade
            html += `<div class="prop-row"><span class="prop-label">Grade</span><span class="prop-value prop-click-edit" onclick="propEditSteelGrade('${el.id}','beam')">${bSchedData.grade || '300'}</span></div>`;
            // Editable description
            const descDisplay = bSchedData.description || '<em style="color:#999">click to add</em>';
            html += `<div class="prop-row"><span class="prop-label">Desc.</span><span class="prop-value prop-click-edit" onclick="propEditDescription('${el.id}','beam')">${descDisplay}</span></div>`;
            // ── Design Check (AS 4100) ──
            if (typeof buildDesignCheckHTML === 'function') {
                html += buildDesignCheckHTML(el);
            }
        }
        html += propDivider();
        html += propRow('X1', fmtCoord(el.x1));
        html += propRow('Y1', fmtCoord(el.y1));
        html += propRow('X2', fmtCoord(el.x2));
        html += propRow('Y2', fmtCoord(el.y2));
    }

    if (el.type === 'column') {
        html += propDivider();
        const cTypeRef = el.typeRef || el.tag;
        const cSchedData = project.scheduleTypes.column[cTypeRef] || {};
        const cColor = cSchedData.color || '#93C5FD';
        html += `<div class="prop-row"><span class="prop-label">Type</span><span class="prop-type-badge" onclick="showTypeReassignmentPicker(selectedElement,'column',150,80)" title="Click to change type"><span class="color-dot" style="background:${cColor}"></span>${cTypeRef || '—'}</span></div>`;
        html += propDivider();
        // Editable section type
        const cSecDisplay = cSchedData.sectionType || '<em style="color:#999">click to set</em>';
        html += `<div class="prop-row"><span class="prop-label">Section</span><span class="prop-value prop-click-edit" onclick="propEditSteelSection('${el.id}','column')">${cSecDisplay}</span></div>`;
        // Editable size
        const cSizeDisplay = cSchedData.size ? formatSectionName(cSchedData.sectionType, cSchedData.size) : '<em style="color:#999">click to set</em>';
        html += `<div class="prop-row"><span class="prop-label">Size</span><span class="prop-value prop-click-edit" onclick="propEditSteelSize('${el.id}','column')" style="font-weight:600;">${cSizeDisplay}</span></div>`;
        // Editable grade
        html += `<div class="prop-row"><span class="prop-label">Grade</span><span class="prop-value prop-click-edit" onclick="propEditSteelGrade('${el.id}','column')">${cSchedData.grade || '300'}</span></div>`;
        // Editable description
        const cDescDisplay = cSchedData.description || '<em style="color:#999">click to add</em>';
        html += `<div class="prop-row"><span class="prop-label">Desc.</span><span class="prop-value prop-click-edit" onclick="propEditDescription('${el.id}','column')">${cDescDisplay}</span></div>`;
        html += propDivider();
        html += propRow('Extends', el.extends || 'below');
        html += propDivider();
        html += propRow('X', fmtCoord(el.x));
        html += propRow('Y', fmtCoord(el.y));
    }

    if (el.type === 'text') {
        html += propDivider();
        html += propRowEditable('Text', el.text, 'text');
        html += propRow('Font Size', (el.fontSize || 3.5).toFixed(1) + ' mm');
    }

    if (el.type === 'polyline') {
        const pts = el.points || [];
        let perimeter = 0;
        for (let i = 0; i < pts.length - 1; i++) {
            perimeter += Math.sqrt(Math.pow(pts[i+1].x - pts[i].x, 2) + Math.pow(pts[i+1].y - pts[i].y, 2));
        }
        if (el.closed && pts.length > 2) {
            perimeter += Math.sqrt(Math.pow(pts[0].x - pts[pts.length-1].x, 2) + Math.pow(pts[0].y - pts[pts.length-1].y, 2));
        }
        html += propDivider();
        html += propRow('Vertices', pts.length);
        html += propRow('Closed', el.closed ? 'Yes' : 'No');
        html += propRow('Perimeter', fmtLen(perimeter));
        if (el.hatch && el.hatch !== 'none') {
            html += propRow('Hatch', el.hatch);
            html += propRow('Fill Color', el.fillColor || '#CCC');
        }
        // Area for closed shapes
        if (el.closed && pts.length >= 3) {
            let area = 0;
            for (let i = 0; i < pts.length; i++) {
                const j = (i + 1) % pts.length;
                area += pts[i].x * pts[j].y - pts[j].x * pts[i].y;
            }
            area = Math.abs(area) / 2;
            html += propRow('Area', (area / 1e6).toFixed(2) + ' m²');
        }
    }

    if (el.type === 'dimension') {
        const len = Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2));
        const isH = Math.abs(el.x2 - el.x1) >= Math.abs(el.y2 - el.y1);
        const measured = isH ? Math.abs(el.x2 - el.x1) : Math.abs(el.y2 - el.y1);
        html += propDivider();
        html += propRow('Measured', fmtLen(measured));
        html += propRow('Direction', isH ? 'Horizontal' : 'Vertical');
    }

    if (el.type === 'leader') {
        html += propDivider();
        if (el.text) html += propRowEditable('Label', el.text, 'text');
        const len = Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2));
        html += propRow('Length', fmtLen(len));
    }

    if (el.type === 'cloud') {
        html += propDivider();
        html += propRow('Vertices', (el.points || []).length);
    }

    if (el.type === 'notesbox') {
        html += propDivider();
        html += propRowEditable('Heading', el.heading || 'NOTES', 'heading');
        html += propRowEditable('Body', el.body || '', 'body');
        html += propRow('Font Size', (el.fontSize || 3.5).toFixed(1) + ' mm');
    }

    if (el.type === 'table') {
        html += propDivider();
        html += propRow('Template', el.template || 'custom');
        html += propRow('Rows', (el.rows || []).length);
        html += propRow('Columns', (el.rows && el.rows[0] ? el.rows[0].length : 0));
    }

    if (el.type === 'slabcallout') {
        html += propDivider();
        html += propRowEditable('Thickness', el.thickness || '250', 'thickness');
    }

    // ── Pad Footing properties (EDITABLE) ──
    if (el.type === 'footing') {
        const typeRef = el.typeRef || el.mark || 'PF1';
        const typeData = project.scheduleTypes.padfooting[typeRef] || {};
        const w = typeData.width || el.footingWidth || 1000;
        const l = typeData.rect ? (typeData.length || w) : w;
        const depthVal = typeData.depth || el.footingDepth || 300;
        const setdownVal = el.depthBelowFSL || 200;
        const tofVal = getFootingTOF(el);
        const tofStr = levelSystem.groundRL !== 0 ? tofVal.toFixed(3) : '—';
        const hasOverride = el.tofOverride !== undefined && el.tofOverride !== null && el.tofOverride !== '';
        const pfColor = typeData.color || '#93C5FD';

        html += propDivider();
        html += `<div class="prop-row"><span class="prop-label">Type</span><span class="prop-type-badge" onclick="showTypeReassignmentPicker(selectedElement,'padfooting',150,80)" title="Click to change type"><span class="color-dot" style="background:${pfColor}"></span>${typeRef}</span></div>`;
        html += propDivider();
        // Editable width
        html += `<div class="prop-row"><span class="prop-label">Width</span><span class="prop-value prop-click-edit" onclick="propEditPfDim('${el.id}','width')" style="font-weight:600;">${w}</span><span style="font-size:9px;color:var(--text-tertiary);">mm</span></div>`;
        // Square/Rect toggle + length
        if (typeData.rect) {
            html += `<div class="prop-row"><span class="prop-label">Length</span><span class="prop-value prop-click-edit" onclick="propEditPfDim('${el.id}','length')" style="font-weight:600;">${l}</span><span style="font-size:9px;color:var(--text-tertiary);">mm</span></div>`;
        }
        html += `<div class="prop-row"><span class="prop-label">Shape</span><span class="prop-value prop-click-edit" onclick="propTogglePfRect('${el.id}')" style="font-size:10px;color:var(--accent);">${typeData.rect ? '▣ RECT — click for SQ' : '▣ SQ — click for RECT'}</span></div>`;
        // Editable depth
        html += `<div class="prop-row"><span class="prop-label">Depth</span><span class="prop-value prop-click-edit" onclick="propEditPfDim('${el.id}','depth')" style="font-weight:600;">${depthVal}</span><span style="font-size:9px;color:var(--text-tertiary);">mm</span></div>`;
        html += propDivider();
        // Setdown with spinner
        html += `<div class="prop-row">
            <span class="prop-label">Setdown</span>
            <div style="display:flex;align-items:center;gap:2px;">
                <span class="prop-value pf-prop-setdown" data-elid="${el.id}" onclick="editFootingSetdown(this)" style="cursor:pointer;padding:2px 4px;border-radius:2px;min-width:40px;text-align:center;">${setdownVal}</span>
                <span style="font-size:9px;color:var(--text-tertiary);">mm</span>
                <div class="dim-spinner" style="margin-left:2px;">
                    <button onclick="spinFootingSetdown('${el.id}', -100)" title="−100mm (shallower)">▲</button>
                    <button onclick="spinFootingSetdown('${el.id}', 100)" title="+100mm (deeper)">▼</button>
                </div>
            </div>
        </div>`;
        // TOF with override capability
        html += `<div class="prop-row">
            <span class="prop-label">TOF RL</span>
            <div style="display:flex;align-items:center;gap:2px;">
                <span class="prop-value pf-prop-tof" data-elid="${el.id}" onclick="editFootingTOF(this)" style="cursor:pointer;padding:2px 4px;border-radius:2px;min-width:50px;text-align:center;${hasOverride ? 'color:var(--warning);font-weight:600;' : ''}">${tofStr}</span>
                ${hasOverride ? '<span style="font-size:8px;color:var(--warning);font-weight:700;" title="Custom TOF override">OVR</span>' : ''}
                ${hasOverride ? `<button onclick="clearFootingTOFOverride('${el.id}')" style="font-size:8px;padding:1px 4px;border:1px solid #ddd;border-radius:2px;background:#fff;cursor:pointer;color:var(--text-tertiary);" title="Reset to default">✕</button>` : ''}
            </div>
        </div>`;
        // Editable reo
        const pfReoDisplay = typeData.reo || '<em style="color:#999">click to add</em>';
        html += `<div class="prop-row"><span class="prop-label">Reo</span><span class="prop-value prop-click-edit" onclick="propEditPfReo('${el.id}')">${pfReoDisplay}</span></div>`;
        html += propDivider();
        html += propRow('X', fmtCoord(el.x));
        html += propRow('Y', fmtCoord(el.y));
    }

    // ── Strip Footing properties (EDITABLE) ──
    if (el.type === 'stripFooting') {
        const sfTypeRef = el.typeRef || el.tag || 'SF1';
        const sfTypeData = project.scheduleTypes.stripfooting[sfTypeRef] || {};
        const sfColor = sfTypeData.color || '#93C5FD';
        const sfSetdownVal = el.depthBelowFSL || 200;
        const sfTofVal = getFootingTOF(el);
        const sfTofStr = levelSystem.groundRL !== 0 ? sfTofVal.toFixed(3) : '—';
        const sfHasOverride = el.tofOverride !== undefined && el.tofOverride !== null && el.tofOverride !== '';
        // Check schedule-level TOP override
        const sfHasSchedTop = sfTypeData.top !== undefined && sfTypeData.top !== null && sfTypeData.top !== '';

        html += propDivider();
        html += `<div class="prop-row"><span class="prop-label">Type</span><span class="prop-type-badge" onclick="showTypeReassignmentPicker(selectedElement,'stripfooting',150,80)" title="Click to change type"><span class="color-dot" style="background:${sfColor}"></span>${sfTypeRef}</span></div>`;
        html += propDivider();
        // Editable width
        html += `<div class="prop-row"><span class="prop-label">Width</span><span class="prop-value prop-click-edit" onclick="propEditSfDim('${el.id}','width')" style="font-weight:600;">${sfTypeData.width || 300}</span><span style="font-size:9px;color:var(--text-tertiary);">mm</span></div>`;
        // Editable depth
        html += `<div class="prop-row"><span class="prop-label">Depth</span><span class="prop-value prop-click-edit" onclick="propEditSfDim('${el.id}','depth')" style="font-weight:600;">${sfTypeData.depth || 500}</span><span style="font-size:9px;color:var(--text-tertiary);">mm</span></div>`;
        html += propDivider();
        // Setdown with spinner
        html += `<div class="prop-row">
            <span class="prop-label">Setdown</span>
            <div style="display:flex;align-items:center;gap:2px;">
                <span class="prop-value pf-prop-setdown" data-elid="${el.id}" onclick="editFootingSetdown(this)" style="cursor:pointer;padding:2px 4px;border-radius:2px;min-width:40px;text-align:center;">${sfSetdownVal}</span>
                <span style="font-size:9px;color:var(--text-tertiary);">mm</span>
                <div class="dim-spinner" style="margin-left:2px;">
                    <button onclick="spinFootingSetdown('${el.id}', -100)" title="−100mm (shallower)">▲</button>
                    <button onclick="spinFootingSetdown('${el.id}', 100)" title="+100mm (deeper)">▼</button>
                </div>
            </div>
        </div>`;
        // TOP RL — editable per-element override (same as pad footing TOF)
        html += `<div class="prop-row">
            <span class="prop-label">TOP RL</span>
            <div style="display:flex;align-items:center;gap:2px;">
                <span class="prop-value pf-prop-tof" data-elid="${el.id}" onclick="editFootingTOF(this)" style="cursor:pointer;padding:2px 4px;border-radius:2px;min-width:50px;text-align:center;${sfHasOverride ? 'color:var(--warning);font-weight:600;' : ''}">${sfTofStr}</span>
                ${sfHasOverride ? '<span style="font-size:8px;color:var(--warning);font-weight:700;" title="Custom TOP override">OVR</span>' : ''}
                ${sfHasOverride ? `<button onclick="clearFootingTOFOverride('${el.id}')" style="font-size:8px;padding:1px 4px;border:1px solid #ddd;border-radius:2px;background:#fff;cursor:pointer;color:var(--text-tertiary);" title="Reset to default">✕</button>` : ''}
            </div>
        </div>`;
        // Schedule-level TOP (stored in schedule type, shared across all of this type)
        const sfTopDisplay = sfHasSchedTop ? sfTypeData.top : '<em style="color:#999">click to set</em>';
        html += `<div class="prop-row"><span class="prop-label">Sched TOP</span><span class="prop-value prop-click-edit" onclick="propEditSfTop('${el.id}')" style="font-size:10px;">${sfTopDisplay}</span></div>`;
        html += propDivider();
        // Editable reo
        const sfReoDisplay = sfTypeData.reo || '<em style="color:#999">click to add</em>';
        html += `<div class="prop-row"><span class="prop-label">Reo</span><span class="prop-value prop-click-edit" onclick="propEditSfReo('${el.id}')">${sfReoDisplay}</span></div>`;
        html += propDivider();
        // Length (computed, read-only)
        const sfLen = Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2));
        html += propRow('Length', fmtLen(sfLen));
        html += propDivider();
        html += propRow('X1', fmtCoord(el.x1));
        html += propRow('Y1', fmtCoord(el.y1));
        html += propRow('X2', fmtCoord(el.x2));
        html += propRow('Y2', fmtCoord(el.y2));
    }

    // ── Wall properties (EDITABLE) ──
    if (el.type === 'wall') {
        const wTypeRef = el.typeRef || el.tag || 'BW1';
        const wTypeData = project.scheduleTypes.wall[wTypeRef] || {};
        const wColor = wTypeData.color || '#93C5FD';
        html += propDivider();
        html += `<div class="prop-row"><span class="prop-label">Type</span><span class="prop-type-badge" onclick="showTypeReassignmentPicker(selectedElement,'wall',150,80)" title="Click to change type"><span class="color-dot" style="background:${wColor}"></span>${wTypeRef}</span></div>`;
        html += propDivider();
        // Editable wall type
        const wWTDisplay = wTypeData.wallType || '<em style="color:#999">click to set</em>';
        html += `<div class="prop-row"><span class="prop-label">Wall Type</span><span class="prop-value prop-click-edit" onclick="propEditWallType('${el.id}')">${wWTDisplay}</span></div>`;
        // Editable thickness
        const wThkDisplay = wTypeData.thickness || '<em style="color:#999">click to set</em>';
        html += `<div class="prop-row"><span class="prop-label">Thickness</span><span class="prop-value prop-click-edit" onclick="propEditWallThickness('${el.id}')" style="font-weight:600;">${wThkDisplay}</span>${wTypeData.thickness ? '<span style="font-size:9px;color:var(--text-tertiary);">mm</span>' : ''}</div>`;
        // Editable description
        const wDescDisplay = wTypeData.description || '<em style="color:#999">click to add</em>';
        html += `<div class="prop-row"><span class="prop-label">Desc.</span><span class="prop-value prop-click-edit" onclick="propEditDescription('${el.id}','wall')">${wDescDisplay}</span></div>`;
        // Length
        const wLen = Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2));
        html += propDivider();
        html += propRow('Length', fmtLen(wLen));
        html += propDivider();
        html += propRow('X1', fmtCoord(el.x1));
        html += propRow('Y1', fmtCoord(el.y1));
        html += propRow('X2', fmtCoord(el.x2));
        html += propRow('Y2', fmtCoord(el.y2));
    }

    // ── Bracing Wall properties ──
    if (el.type === 'bracingWall') {
        const bwTypeRef = el.typeRef || el.tag || 'BR1';
        const bwTypeData = project.scheduleTypes.bracingWall?.[bwTypeRef] || {};
        const bwColor = bwTypeData.color || '#000000';
        const bwBracingType = bwTypeData.bracingType || 'g';
        const bwCapPerM = typeof BRACING_TYPES !== 'undefined' ? (BRACING_TYPES[bwBracingType]?.capacity || 0) : 0;
        html += propDivider();
        html += `<div class="prop-row"><span class="prop-label">Type</span><span class="prop-type-badge" onclick="showTypeReassignmentPicker(selectedElement,'bracingWall',150,80)" title="Click to change type"><span class="color-dot" style="background:${bwColor}"></span>${bwTypeRef}</span></div>`;
        html += propDivider();
        html += propRow('Bracing Type', bwBracingType);
        html += propRow('Base Capacity', bwCapPerM !== null ? bwCapPerM + ' kN/m' : 'Engineered');
        html += propRow('Description', bwTypeData.description || '—');
        // Length and effective capacity
        const bwLen = Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2));
        const bwDir = typeof getBracingDirection !== 'undefined' ? getBracingDirection(el) : '—';
        html += propDivider();
        html += propRow('Length', fmtLen(bwLen));
        html += propRow('Direction', bwDir === 'X' ? '↕ X-dir' : bwDir === 'Y' ? '↔ Y-dir' : bwDir);
        // Effective capacity
        if (typeof calcBracingCapacity !== 'undefined') {
            const bwCap = calcBracingCapacity(el, bwTypeData,
                typeof bracingSettings !== 'undefined' ? bracingSettings.ceilingHeight : 2700,
                typeof bracingSettings !== 'undefined' ? bracingSettings.jointGroup : 'JD6');
            html += propRow('Eff. Capacity', `<strong style="color:${bwColor}">${bwCap.toFixed(1)} kN</strong>`);
        }
        html += propDivider();
        html += propRow('X1', fmtCoord(el.x1));
        html += propRow('Y1', fmtCoord(el.y1));
        html += propRow('X2', fmtCoord(el.x2));
        html += propRow('Y2', fmtCoord(el.y2));
    }

    // ── Borehole properties ──
    if (el.type === 'borehole') {
        const groundRL = el.groundRL || levelSystem.groundRL;
        const rockRL = el.rockRL !== undefined ? el.rockRL : (groundRL - el.depthToRock);
        html += propDivider();
        html += propRow('Tag', `<strong style="color:#8B4513">${el.tag}</strong>`);
        html += propRow('Depth to Rock', el.depthToRock.toFixed(2) + ' m');
        if (levelSystem.groundRL !== 0) {
            html += propRow('Ground RL', (el.groundRL || levelSystem.groundRL).toFixed(3));
            html += propRow('Rock RL', rockRL.toFixed(3));
        }
        html += propDivider();
        html += propRow('X', fmtCoord(el.x));
        html += propRow('Y', fmtCoord(el.y));
    }

    // ID
    html += propDivider();
    html += propRow('ID', el.id);

    propsBody.innerHTML = html;

    // Wire editable fields
    for (const input of propsBody.querySelectorAll('.prop-input')) {
        input.addEventListener('keydown', (ev) => {
            ev.stopPropagation();
            if (ev.key === 'Enter') {
                const field = input.dataset.field;
                const oldVal = el[field];
                const newVal = input.value.trim();
                if (newVal && newVal !== String(oldVal)) {
                    history.execute({
                        description: 'Edit property: ' + field,
                        execute() { el[field] = newVal; },
                        undo() { el[field] = oldVal; }
                    });
                    engine.requestRender();
                }
                input.blur();
            }
            if (ev.key === 'Escape') input.blur();
        });
        input.addEventListener('blur', () => {
            // Re-render to reset display
            _lastPropsElement = null;
            updatePropsPanel();
        });
    }
}

function propRow(label, value) {
    return `<div class="prop-row"><span class="prop-label">${label}</span><span class="prop-value">${value}</span></div>`;
}

function propRowEditable(label, value, field) {
    return `<div class="prop-row"><span class="prop-label">${label}</span><input class="prop-input" data-field="${field}" value="${escHtml(String(value))}"></div>`;
}

function propDivider() {
    return '<div class="prop-divider"></div>';
}

function fmtLen(mm) {
    if (mm >= 1000) return (mm / 1000).toFixed(mm >= 10000 ? 1 : 2) + ' m';
    return Math.round(mm) + ' mm';
}

function fmtCoord(mm) {
    return (mm / 1000).toFixed(3) + ' m';
}

function escHtml(str) {
    return str.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

// ── Footing Property Interactions ────────────────────────

/** Spin setdown on a specific footing by delta mm */
window.spinFootingSetdown = function(elId, delta) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    let val = (el.depthBelowFSL || 200) + delta;
    val = Math.max(0, Math.round(val));
    el.depthBelowFSL = val;
    // Clear any TOF override since user is adjusting setdown directly
    delete el.tofOverride;
    _lastPropsElement = null;
    updatePropsPanel();
    engine.requestRender();
};

/** Inline-edit setdown value */
window.editFootingSetdown = function(span) {
    if (span.querySelector('input')) return;
    const elId = span.dataset.elid;
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const oldVal = el.depthBelowFSL || 200;
    const input = document.createElement('input');
    input.type = 'text';
    input.value = oldVal;
    input.style.cssText = 'width:50px;padding:2px 4px;font-size:11px;border:1px solid var(--accent);border-radius:2px;font-family:var(--font-mono);text-align:center;outline:none;';
    span.textContent = '';
    span.appendChild(input);
    input.focus();
    input.select();
    let saved = false;
    function save() {
        if (saved) return;
        saved = true;
        const newVal = Math.max(0, Math.round(parseFloat(input.value) || 200));
        el.depthBelowFSL = newVal;
        delete el.tofOverride;
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    input.addEventListener('blur', save);
    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); save(); }
        if (ev.key === 'Escape') { saved = true; _lastPropsElement = null; updatePropsPanel(); }
    });
};

/** Inline-edit TOF RL value (creates an override) */
window.editFootingTOF = function(span) {
    if (span.querySelector('input') || levelSystem.groundRL === 0) return;
    const elId = span.dataset.elid;
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const currentTOF = getFootingTOF(el);
    const input = document.createElement('input');
    input.type = 'text';
    input.value = currentTOF.toFixed(3);
    input.style.cssText = 'width:60px;padding:2px 4px;font-size:11px;border:1px solid var(--warning);border-radius:2px;font-family:var(--font-mono);text-align:center;outline:none;';
    span.textContent = '';
    span.appendChild(input);
    input.focus();
    input.select();
    let saved = false;
    function save() {
        if (saved) return;
        saved = true;
        const newTOF = parseFloat(input.value);
        if (!isNaN(newTOF)) {
            el.tofOverride = newTOF;
            // Back-calculate the setdown implied by this TOF
            el.depthBelowFSL = tofToSetdown(newTOF, el.level || 'GF');
        }
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    input.addEventListener('blur', save);
    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); save(); }
        if (ev.key === 'Escape') { saved = true; _lastPropsElement = null; updatePropsPanel(); }
    });
};

/** Clear TOF override — revert to default (level RL - type setdown) */
window.clearFootingTOFOverride = function(elId) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    delete el.tofOverride;
    _lastPropsElement = null;
    updatePropsPanel();
    engine.requestRender();
};

// ── Schedule Sync: Find or Create Type ─────────────────
// Given a category and desired properties, find an existing schedule
// type that matches or create a new one. Returns the typeRef string.

function findOrCreateScheduleType(category, matchProps) {
    const types = project.scheduleTypes[category];

    // 1. Try to find an existing type with matching dimensions
    for (const [ref, data] of Object.entries(types)) {
        if (scheduleTypeMatches(category, data, matchProps)) {
            return ref;
        }
    }

    // 2. Try to reuse an empty/unused slot (e.g. PF1 with no dimensions set
    //    and no elements on the plan referencing it). This avoids skipping
    //    to PF6 when PF1–PF5 are all blank placeholders.
    for (const [ref, data] of Object.entries(types)) {
        if (isScheduleTypeEmpty(category, data) && countTypeUsage(category, ref) === 0) {
            // Fill this empty slot with the new properties
            if (category === 'padfooting') {
                data.width = matchProps.width || ''; data.length = matchProps.length || '';
                data.depth = matchProps.depth || ''; data.reo = matchProps.reo || '';
                data.setdown = matchProps.setdown || 200; data.rect = matchProps.rect || false;
            } else if (category === 'stripfooting') {
                data.width = matchProps.width || 300; data.depth = matchProps.depth || 500;
                data.reo = matchProps.reo || ''; data.setdown = matchProps.setdown || 200;
                data.top = matchProps.top || '';
            } else if (category === 'beam' || category === 'column') {
                data.sectionType = matchProps.sectionType || ''; data.size = matchProps.size || '';
                data.description = matchProps.description || ''; data.grade = matchProps.grade || '300';
            } else if (category === 'wall') {
                data.wallType = matchProps.wallType || ''; data.thickness = matchProps.thickness || '';
                data.description = matchProps.description || '';
            }
            return ref;
        }
    }

    // 3. No match and no empty slots — create a new type
    const prefix = category === 'padfooting' ? 'PF' : category === 'stripfooting' ? 'SF' :
                   category === 'beam' ? 'SB' : category === 'column' ? 'SC' : 'BW';
    let nextNum = 1;
    while (types[prefix + nextNum]) nextNum++;
    const newRef = prefix + nextNum;

    const colorIndex = Object.keys(types).length % SCHEDULE_COLORS.length;
    const defaultColor = SCHEDULE_COLORS[colorIndex];

    let newType;
    if (category === 'padfooting') {
        newType = { width: matchProps.width || '', length: matchProps.length || '', depth: matchProps.depth || '', reo: matchProps.reo || '', setdown: matchProps.setdown || 200, rect: matchProps.rect || false, color: defaultColor };
    } else if (category === 'stripfooting') {
        newType = { width: matchProps.width || 300, depth: matchProps.depth || 500, reo: matchProps.reo || '', setdown: matchProps.setdown || 200, top: matchProps.top || '', color: defaultColor };
    } else if (category === 'beam') {
        newType = { sectionType: matchProps.sectionType || '', size: matchProps.size || '', description: matchProps.description || '', grade: matchProps.grade || '300', color: defaultColor };
    } else if (category === 'column') {
        newType = { sectionType: matchProps.sectionType || '', size: matchProps.size || '', description: matchProps.description || '', grade: matchProps.grade || '300', color: defaultColor };
    } else if (category === 'wall') {
        newType = { wallType: matchProps.wallType || '', thickness: matchProps.thickness || '', description: matchProps.description || '', color: defaultColor };
    }
    types[newRef] = newType;
    return newRef;
}

// Check if a schedule type slot is empty (no meaningful data filled in)
function isScheduleTypeEmpty(category, data) {
    if (category === 'padfooting') {
        return !data.width && !data.depth;
    }
    if (category === 'stripfooting') {
        return !data.width && !data.depth;
    }
    if (category === 'beam' || category === 'column') {
        return !data.sectionType && !data.size;
    }
    if (category === 'wall') {
        return !data.wallType && !data.thickness;
    }
    return false;
}

function scheduleTypeMatches(category, data, match) {
    if (category === 'padfooting') {
        return String(data.width) === String(match.width) &&
               String(data.depth) === String(match.depth) &&
               Boolean(data.rect) === Boolean(match.rect) &&
               (!match.rect || String(data.length) === String(match.length));
    }
    if (category === 'stripfooting') {
        return String(data.width) === String(match.width) &&
               String(data.depth) === String(match.depth);
    }
    if (category === 'beam' || category === 'column') {
        return data.sectionType === match.sectionType &&
               data.size === match.size;
    }
    if (category === 'wall') {
        return data.wallType === match.wallType &&
               String(data.thickness) === String(match.thickness);
    }
    return false;
}

// Count how many elements reference a given schedule type
function countTypeUsage(category, typeRef) {
    return project.elements.filter(el => {
        if (category === 'padfooting') return el.type === 'footing' && (el.typeRef || el.mark) === typeRef;
        if (category === 'stripfooting') return el.type === 'stripFooting' && (el.typeRef || el.tag) === typeRef;
        if (category === 'beam') return el.type === 'line' && el.layer === 'S-BEAM' && (el.typeRef || el.tag) === typeRef;
        if (category === 'column') return el.type === 'column' && (el.typeRef || el.tag) === typeRef;
        if (category === 'wall') return el.type === 'wall' && (el.typeRef || el.tag) === typeRef;
        return false;
    }).length;
}

// When an empty type is shared by multiple elements and the user edits one,
// keep the edited element on its current type (fill it with new data) and
// reassign the OTHER elements to the next available empty slot.
function reassignOtherElements(category, currentTypeRef, keepElId) {
    const types = project.scheduleTypes[category];
    // Find all elements sharing currentTypeRef except the one being edited
    const others = project.elements.filter(el => {
        const elRef = (category === 'padfooting') ? (el.typeRef || el.mark) : (el.typeRef || el.tag);
        if (elRef !== currentTypeRef) return false;
        const elId = el.id;
        const keepId = parseInt(keepElId) || keepElId;
        if (elId === keepId) return false;
        if (category === 'padfooting') return el.type === 'footing';
        if (category === 'stripfooting') return el.type === 'stripFooting';
        if (category === 'beam') return el.type === 'line' && el.layer === 'S-BEAM';
        if (category === 'column') return el.type === 'column';
        if (category === 'wall') return el.type === 'wall';
        return false;
    });
    if (others.length === 0) return;
    // Find the next empty+unused slot (skip currentTypeRef since it's now filled)
    let targetRef = null;
    for (const [ref, data] of Object.entries(types)) {
        if (ref === currentTypeRef) continue;
        if (isScheduleTypeEmpty(category, data) && countTypeUsage(category, ref) === 0) {
            targetRef = ref;
            break;
        }
    }
    if (!targetRef) {
        // No empty slot — create a new empty one
        const prefix = category === 'padfooting' ? 'PF' : category === 'stripfooting' ? 'SF' :
                       category === 'beam' ? 'SB' : category === 'column' ? 'SC' : 'BW';
        let n = 1;
        while (types[prefix + n]) n++;
        targetRef = prefix + n;
        const colorIndex = Object.keys(types).length % SCHEDULE_COLORS.length;
        if (category === 'padfooting') {
            types[targetRef] = { width: '', length: '', depth: '', reo: '', setdown: 200, rect: false, color: SCHEDULE_COLORS[colorIndex] };
        } else if (category === 'stripfooting') {
            types[targetRef] = { width: '', depth: '', reo: '', setdown: 200, top: '', color: SCHEDULE_COLORS[colorIndex] };
        } else if (category === 'beam' || category === 'column') {
            types[targetRef] = { sectionType: '', size: '', description: '', grade: '300', color: SCHEDULE_COLORS[colorIndex] };
        } else if (category === 'wall') {
            types[targetRef] = { wallType: '', thickness: '', description: '', color: SCHEDULE_COLORS[colorIndex] };
        }
    }
    // Move all other elements to the target empty slot
    others.forEach(el => {
        if (category === 'padfooting') { el.mark = targetRef; el.typeRef = targetRef; }
        else { el.tag = targetRef; el.typeRef = targetRef; }
    });
}

// Update element's type reference and sync element properties from schedule
function applyTypeToElement(el, category, typeRef) {
    const typeData = project.scheduleTypes[category][typeRef] || {};
    if (category === 'padfooting') {
        el.mark = typeRef;
        el.typeRef = typeRef;
        if (typeData.width) el.footingWidth = parseInt(typeData.width) || el.footingWidth;
        if (typeData.depth) el.footingDepth = parseInt(typeData.depth) || el.footingDepth;
        if (typeData.reo) el.reinforcement = typeData.reo;
        if (typeData.setdown !== undefined) el.depthBelowFSL = typeData.setdown;
    } else if (category === 'stripfooting') {
        el.tag = typeRef;
        el.typeRef = typeRef;
        if (typeData.width) el.footingWidth = parseInt(typeData.width) || el.footingWidth;
        if (typeData.depth) el.footingDepth = parseInt(typeData.depth) || el.footingDepth;
    } else if (category === 'beam') {
        el.tag = typeRef;
        el.typeRef = typeRef;
        if (typeData.sectionType) el.memberCategory = typeData.sectionType;
        if (typeData.size) el.memberSize = typeData.size;
    } else if (category === 'column') {
        el.tag = typeRef;
        el.typeRef = typeRef;
        if (typeData.sectionType) el.memberCategory = typeData.sectionType;
        if (typeData.size) el.memberSize = typeData.size;
    } else if (category === 'wall') {
        el.tag = typeRef;
        el.typeRef = typeRef;
        if (typeData.wallType) el.wallType = typeData.wallType;
        if (typeData.thickness) el.thickness = parseInt(typeData.thickness) || el.thickness;
    }
}

// ── Editable Property: update schedule and element on property change ──

// For pad footings: user edits width/length/depth directly in properties panel
window.propEditPfDim = function(elId, field) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.mark || 'PF1';
    const typeData = project.scheduleTypes.padfooting[typeRef] || {};
    const currentVal = (field === 'width') ? (typeData.width || el.footingWidth || 1000) :
                       (field === 'length') ? (typeData.length || typeData.width || el.footingWidth || 1000) :
                       (field === 'depth') ? (typeData.depth || el.footingDepth || 300) : '';
    const span = event.target.closest('.prop-click-edit') || event.target;
    if (span.querySelector('input')) return;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'prop-inline-input';
    input.value = currentVal;
    span.textContent = '';
    span.appendChild(input);
    input.focus();
    input.select();
    let saved = false;
    function save() {
        if (saved) return;
        saved = true;
        const newVal = Math.round(parseFloat(input.value) || 0);
        if (newVal <= 0) { _lastPropsElement = null; updatePropsPanel(); return; }
        // Determine if this is the only element using this type
        const usage = countTypeUsage('padfooting', typeRef);
        const newProps = {
            width: field === 'width' ? newVal : (typeData.width || el.footingWidth || 1000),
            depth: field === 'depth' ? newVal : (typeData.depth || el.footingDepth || 300),
            rect: typeData.rect || false,
            length: typeData.length || ''
        };
        if (field === 'length') {
            newProps.length = newVal;
            newProps.rect = true;
        }
        if (usage <= 1) {
            // Only this element uses it — update the type in place
            if (field === 'width') typeData.width = newVal;
            else if (field === 'depth') typeData.depth = newVal;
            else if (field === 'length') { typeData.length = newVal; typeData.rect = true; }
            el.footingWidth = parseInt(typeData.width) || el.footingWidth;
            el.footingDepth = parseInt(typeData.depth) || el.footingDepth;
        } else if (isScheduleTypeEmpty('padfooting', typeData)) {
            // Multiple elements share an empty type — keep this element on
            // its current slot (fill it), move the others to empty slots
            if (field === 'width') typeData.width = newVal;
            else if (field === 'depth') typeData.depth = newVal;
            else if (field === 'length') { typeData.length = newVal; typeData.rect = true; }
            el.footingWidth = parseInt(typeData.width) || el.footingWidth;
            el.footingDepth = parseInt(typeData.depth) || el.footingDepth;
            reassignOtherElements('padfooting', typeRef, elId);
        } else {
            // Multiple elements use a non-empty type — find or create a matching type
            const newRef = findOrCreateScheduleType('padfooting', newProps);
            applyTypeToElement(el, 'padfooting', newRef);
        }
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    input.addEventListener('blur', save);
    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); save(); }
        if (ev.key === 'Escape') { saved = true; _lastPropsElement = null; updatePropsPanel(); }
    });
};

// For strip footings: edit width/depth
window.propEditSfDim = function(elId, field) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.tag || 'SF1';
    const typeData = project.scheduleTypes.stripfooting[typeRef] || {};
    const currentVal = (field === 'width') ? (typeData.width || 300) : (typeData.depth || 500);
    const span = event.target.closest('.prop-click-edit') || event.target;
    if (span.querySelector('input')) return;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'prop-inline-input';
    input.value = currentVal;
    span.textContent = '';
    span.appendChild(input);
    input.focus();
    input.select();
    let saved = false;
    function save() {
        if (saved) return;
        saved = true;
        const newVal = Math.round(parseFloat(input.value) || 0);
        if (newVal <= 0) { _lastPropsElement = null; updatePropsPanel(); return; }
        const usage = countTypeUsage('stripfooting', typeRef);
        const newProps = {
            width: field === 'width' ? newVal : (typeData.width || 300),
            depth: field === 'depth' ? newVal : (typeData.depth || 500)
        };
        if (usage <= 1) {
            typeData[field] = newVal;
            el.footingWidth = parseInt(typeData.width) || el.footingWidth;
            el.footingDepth = parseInt(typeData.depth) || el.footingDepth;
        } else if (isScheduleTypeEmpty('stripfooting', typeData)) {
            typeData[field] = newVal;
            el.footingWidth = parseInt(typeData.width) || el.footingWidth;
            el.footingDepth = parseInt(typeData.depth) || el.footingDepth;
            reassignOtherElements('stripfooting', typeRef, elId);
        } else {
            const newRef = findOrCreateScheduleType('stripfooting', newProps);
            applyTypeToElement(el, 'stripfooting', newRef);
        }
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    input.addEventListener('blur', save);
    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); save(); }
        if (ev.key === 'Escape') { saved = true; _lastPropsElement = null; updatePropsPanel(); }
    });
};

// For strip footings: edit reinforcement inline
window.propEditSfReo = function(elId) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.tag || 'SF1';
    const typeData = project.scheduleTypes.stripfooting[typeRef] || {};
    const span = event.target.closest('.prop-click-edit') || event.target;
    if (span.querySelector('input')) return;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'prop-inline-input';
    input.style.width = '120px';
    input.value = typeData.reo || '';
    input.placeholder = 'e.g. N12-200 EW';
    span.textContent = '';
    span.appendChild(input);
    input.focus();
    input.select();
    let saved = false;
    function save() {
        if (saved) return;
        saved = true;
        typeData.reo = input.value.trim();
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    input.addEventListener('blur', save);
    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); save(); }
        if (ev.key === 'Escape') { saved = true; _lastPropsElement = null; updatePropsPanel(); }
    });
};

// For strip footings: edit schedule-level TOP value
window.propEditSfTop = function(elId) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.tag || 'SF1';
    const typeData = project.scheduleTypes.stripfooting[typeRef] || {};
    const span = event.target.closest('.prop-click-edit') || event.target;
    if (span.querySelector('input')) return;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'prop-inline-input';
    input.style.width = '80px';
    input.value = typeData.top || '';
    input.placeholder = 'e.g. -0.600';
    span.textContent = '';
    span.appendChild(input);
    input.focus();
    input.select();
    let saved = false;
    function save() {
        if (saved) return;
        saved = true;
        const val = input.value.trim();
        typeData.top = val;
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    input.addEventListener('blur', save);
    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); save(); }
        if (ev.key === 'Escape') { saved = true; _lastPropsElement = null; updatePropsPanel(); }
    });
};

// For beams/columns: edit section type via dropdown
window.propEditSteelSection = function(elId, category) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.tag;
    const typeData = project.scheduleTypes[category][typeRef] || {};
    const span = event.target.closest('.prop-click-edit') || event.target;
    if (span.querySelector('select')) return;
    const select = document.createElement('select');
    select.className = 'prop-select-edit';
    select.innerHTML = '<option value="">— Select —</option>' +
        Object.keys(STEEL_SECTIONS).map(st =>
            `<option value="${st}"${typeData.sectionType === st ? ' selected' : ''}>${st}</option>`
        ).join('');
    span.textContent = '';
    span.appendChild(select);
    select.focus();
    function save() {
        const newST = select.value;
        if (newST === typeData.sectionType) { _lastPropsElement = null; updatePropsPanel(); return; }
        const usage = countTypeUsage(category, typeRef);
        if (usage <= 1) {
            typeData.sectionType = newST;
            typeData.size = ''; // reset size when section type changes
            if (newST) el.memberCategory = newST;
            el.memberSize = '';
        } else if (isScheduleTypeEmpty(category, typeData)) {
            typeData.sectionType = newST;
            typeData.size = '';
            if (newST) el.memberCategory = newST;
            el.memberSize = '';
            reassignOtherElements(category, typeRef, elId);
        } else {
            const newRef = findOrCreateScheduleType(category, { sectionType: newST, size: '' });
            applyTypeToElement(el, category, newRef);
        }
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    select.addEventListener('change', save);
    select.addEventListener('blur', () => { _lastPropsElement = null; updatePropsPanel(); });
};

// For beams/columns: edit size via dropdown (filtered by section type)
window.propEditSteelSize = function(elId, category) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.tag;
    const typeData = project.scheduleTypes[category][typeRef] || {};
    if (!typeData.sectionType) { alert('Select a section type first'); return; }
    const span = event.target.closest('.prop-click-edit') || event.target;
    if (span.querySelector('select')) return;
    const sizes = STEEL_SECTIONS[typeData.sectionType] || [];
    const select = document.createElement('select');
    select.className = 'prop-select-edit';
    select.innerHTML = '<option value="">— Select —</option>' +
        sizes.map(s => {
            const display = formatSectionName(typeData.sectionType, s);
            return `<option value="${s}"${typeData.size === s ? ' selected' : ''}>${display}</option>`;
        }).join('');
    span.textContent = '';
    span.appendChild(select);
    select.focus();
    function save() {
        const newSize = select.value;
        if (newSize === typeData.size) { _lastPropsElement = null; updatePropsPanel(); return; }
        const usage = countTypeUsage(category, typeRef);
        if (usage <= 1) {
            typeData.size = newSize;
            if (newSize) el.memberSize = newSize;
        } else if (isScheduleTypeEmpty(category, typeData)) {
            typeData.size = newSize;
            if (newSize) el.memberSize = newSize;
            reassignOtherElements(category, typeRef, elId);
        } else {
            const newRef = findOrCreateScheduleType(category, {
                sectionType: typeData.sectionType,
                size: newSize
            });
            applyTypeToElement(el, category, newRef);
        }
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    select.addEventListener('change', save);
    select.addEventListener('blur', () => { _lastPropsElement = null; updatePropsPanel(); });
};

// For beams/columns: edit grade via dropdown
window.propEditSteelGrade = function(elId, category) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.tag;
    const typeData = project.scheduleTypes[category][typeRef] || {};
    const span = event.target.closest('.prop-click-edit') || event.target;
    if (span.querySelector('select')) return;
    const select = document.createElement('select');
    select.className = 'prop-select-edit';
    select.innerHTML = STEEL_GRADES.map(g =>
        `<option value="${g}"${typeData.grade === g ? ' selected' : ''}>${g}</option>`
    ).join('');
    span.textContent = '';
    span.appendChild(select);
    select.focus();
    function save() {
        const newGrade = select.value;
        typeData.grade = newGrade;
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    select.addEventListener('change', save);
    select.addEventListener('blur', () => { _lastPropsElement = null; updatePropsPanel(); });
};

// For beams/columns: edit description inline
window.propEditDescription = function(elId, category) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.tag;
    const typeData = project.scheduleTypes[category][typeRef] || {};
    const span = event.target.closest('.prop-click-edit') || event.target;
    if (span.querySelector('input')) return;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'prop-inline-input';
    input.style.width = '120px';
    input.value = typeData.description || '';
    input.placeholder = 'e.g. BEARER';
    span.textContent = '';
    span.appendChild(input);
    input.focus();
    input.select();
    let saved = false;
    function save() {
        if (saved) return;
        saved = true;
        typeData.description = input.value.trim();
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    input.addEventListener('blur', save);
    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); save(); }
        if (ev.key === 'Escape') { saved = true; _lastPropsElement = null; updatePropsPanel(); }
    });
};

// For walls: edit wall type via dropdown
window.propEditWallType = function(elId) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.tag;
    const typeData = project.scheduleTypes.wall[typeRef] || {};
    const span = event.target.closest('.prop-click-edit') || event.target;
    if (span.querySelector('select')) return;
    const select = document.createElement('select');
    select.className = 'prop-select-edit';
    select.innerHTML = '<option value="">— Select —</option>' +
        Object.keys(WALL_TYPES).map(wt =>
            `<option value="${wt}"${typeData.wallType === wt ? ' selected' : ''}>${wt}</option>`
        ).join('');
    span.textContent = '';
    span.appendChild(select);
    select.focus();
    function save() {
        const newWT = select.value;
        if (newWT === typeData.wallType) { _lastPropsElement = null; updatePropsPanel(); return; }
        const autoThk = WALL_TYPES[newWT]?.thickness || '';
        const usage = countTypeUsage('wall', typeRef);
        if (usage <= 1) {
            typeData.wallType = newWT;
            typeData.thickness = autoThk;
            if (newWT) el.wallType = newWT;
            if (autoThk) el.thickness = autoThk;
        } else if (isScheduleTypeEmpty('wall', typeData)) {
            typeData.wallType = newWT;
            typeData.thickness = autoThk;
            if (newWT) el.wallType = newWT;
            if (autoThk) el.thickness = autoThk;
            reassignOtherElements('wall', typeRef, elId);
        } else {
            const newRef = findOrCreateScheduleType('wall', { wallType: newWT, thickness: autoThk });
            applyTypeToElement(el, 'wall', newRef);
        }
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    select.addEventListener('change', save);
    select.addEventListener('blur', () => { _lastPropsElement = null; updatePropsPanel(); });
};

// For walls: edit thickness
window.propEditWallThickness = function(elId) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.tag;
    const typeData = project.scheduleTypes.wall[typeRef] || {};
    const currentVal = typeData.thickness || el.thickness || '';
    const span = event.target.closest('.prop-click-edit') || event.target;
    if (span.querySelector('input')) return;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'prop-inline-input';
    input.value = currentVal;
    span.textContent = '';
    span.appendChild(input);
    input.focus();
    input.select();
    let saved = false;
    function save() {
        if (saved) return;
        saved = true;
        const newVal = Math.round(parseFloat(input.value) || 0);
        if (newVal <= 0) { _lastPropsElement = null; updatePropsPanel(); return; }
        const usage = countTypeUsage('wall', typeRef);
        if (usage <= 1) {
            typeData.thickness = newVal;
            el.thickness = newVal;
        } else if (isScheduleTypeEmpty('wall', typeData)) {
            typeData.thickness = newVal;
            el.thickness = newVal;
            reassignOtherElements('wall', typeRef, elId);
        } else {
            const newRef = findOrCreateScheduleType('wall', {
                wallType: typeData.wallType || '',
                thickness: newVal
            });
            applyTypeToElement(el, 'wall', newRef);
        }
        _lastPropsElement = null;
        updatePropsPanel();
        engine.requestRender();
    }
    input.addEventListener('blur', save);
    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); save(); }
        if (ev.key === 'Escape') { saved = true; _lastPropsElement = null; updatePropsPanel(); }
    });
};

// For pad footings: toggle between square and rectangular
window.propTogglePfRect = function(elId) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.mark || 'PF1';
    const typeData = project.scheduleTypes.padfooting[typeRef] || {};
    const usage = countTypeUsage('padfooting', typeRef);
    if (usage <= 1) {
        typeData.rect = !typeData.rect;
        if (!typeData.rect) typeData.length = '';
    } else if (isScheduleTypeEmpty('padfooting', typeData)) {
        typeData.rect = !typeData.rect;
        if (!typeData.rect) typeData.length = '';
        reassignOtherElements('padfooting', typeRef, elId);
    } else {
        const newProps = {
            width: typeData.width || el.footingWidth || 1000,
            depth: typeData.depth || el.footingDepth || 300,
            rect: !typeData.rect,
            length: typeData.rect ? '' : (typeData.width || el.footingWidth || 1000)
        };
        const newRef = findOrCreateScheduleType('padfooting', newProps);
        applyTypeToElement(el, 'padfooting', newRef);
    }
    _lastPropsElement = null;
    updatePropsPanel();
    engine.requestRender();
};

// For pad footings: edit reo
window.propEditPfReo = function(elId) {
    const el = project.elements.find(e => e.id === parseInt(elId) || e.id === elId);
    if (!el) return;
    const typeRef = el.typeRef || el.mark || 'PF1';
    const typeData = project.scheduleTypes.padfooting[typeRef] || {};
    const span = event.target.closest('.prop-click-edit') || event.target;
    if (span.querySelector('input')) return;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'prop-inline-input';
    input.style.width = '120px';
    input.value = typeData.reo || '';
    input.placeholder = 'e.g. N12-200 EW';
    span.textContent = '';
    span.appendChild(input);
    input.focus();
    input.select();
    let saved = false;
    function save() {
        if (saved) return;
        saved = true;
        typeData.reo = input.value.trim();
        _lastPropsElement = null;
        updatePropsPanel();
    }
    input.addEventListener('blur', save);
    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); save(); }
        if (ev.key === 'Escape') { saved = true; _lastPropsElement = null; updatePropsPanel(); }
    });
};

// ── Bulk Properties Panel (Multi-Select) ─────────────────
// Shows editable properties grouped by type when multiple elements are selected.
// Editing a field updates ALL selected elements of that type in one undo step.

let _lastBulkCount = 0;

function updateBulkPropsPanel() {
    const els = (typeof selectedElements !== 'undefined') ? selectedElements : [];

    // If 0 or 1 selected, hide bulk panel and let single panel handle it
    if (els.length <= 1) {
        propsPanel.classList.toggle('hidden', !selectedElement);
        _lastBulkCount = 0;
        _lastPropsElement = null; // force single panel refresh
        return;
    }

    // Multi-select: show bulk properties panel
    propsPanel.classList.remove('hidden');

    // Skip re-render if same count (perf optimisation for mousemove)
    if (els.length === _lastBulkCount && _lastPropsElement === 'BULK') return;
    _lastBulkCount = els.length;
    _lastPropsElement = 'BULK';

    propsType.textContent = els.length + ' Elements';

    // Group by element type
    const groups = {};
    for (const el of els) {
        let groupKey = el.type;
        if (el.type === 'line' && el.layer === 'S-BEAM') groupKey = 'beam';
        if (!groups[groupKey]) groups[groupKey] = [];
        groups[groupKey].push(el);
    }

    const typeLabels = {
        beam: 'Beams', wall: 'Walls', column: 'Columns',
        stripFooting: 'Strip Footings', footing: 'Pad Footings',
        line: 'Lines', polyline: 'Polylines', text: 'Text',
        dimension: 'Dimensions', leader: 'Leaders',
        cloud: 'Clouds', edge: 'Edges', slab: 'Slabs'
    };

    let html = '';
    html += `<div class="bulk-summary">${els.length} elements selected</div>`;

    for (const [groupKey, groupEls] of Object.entries(groups)) {
        const label = typeLabels[groupKey] || groupKey;
        html += `<div class="bulk-props-section">`;
        html += `<div class="bulk-props-type-header">${label}<span class="count-badge">${groupEls.length}</span></div>`;

        // ── Strip Footing bulk properties ──
        if (groupKey === 'stripFooting') {
            const widths = groupEls.map(el => {
                const td = project.scheduleTypes.stripfooting[el.typeRef || el.tag] || {};
                return td.width || 300;
            });
            const depths = groupEls.map(el => {
                const td = project.scheduleTypes.stripfooting[el.typeRef || el.tag] || {};
                return td.depth || 500;
            });
            const setdowns = groupEls.map(el => el.depthBelowFSL || 200);
            const typeRefs = [...new Set(groupEls.map(el => el.typeRef || el.tag || 'SF1'))];

            if (typeRefs.length === 1) {
                html += bulkPropRow('Type', typeRefs[0], null);
            } else {
                html += bulkPropRow('Type', 'Varies (' + typeRefs.join(', ') + ')', null);
            }
            html += bulkEditRow('Width', uniqOrVaries(widths), 'mm', 'bulk-sf-width');
            html += bulkEditRow('Depth', uniqOrVaries(depths), 'mm', 'bulk-sf-depth');
            html += bulkEditRow('Setdown', uniqOrVaries(setdowns), 'mm', 'bulk-sf-setdown');
            html += bulkPropRow('Total Length', fmtLen(groupEls.reduce((sum, el) =>
                sum + Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2)), 0)), null);
        }

        // ── Pad Footing bulk properties ──
        else if (groupKey === 'footing') {
            const widths = groupEls.map(el => {
                const td = project.scheduleTypes.padfooting[el.typeRef || el.mark] || {};
                return td.width || el.footingWidth || 1000;
            });
            const depths = groupEls.map(el => {
                const td = project.scheduleTypes.padfooting[el.typeRef || el.mark] || {};
                return td.depth || el.footingDepth || 300;
            });
            const setdowns = groupEls.map(el => el.depthBelowFSL || 200);
            const typeRefs = [...new Set(groupEls.map(el => el.typeRef || el.mark || 'PF1'))];

            if (typeRefs.length === 1) {
                html += bulkPropRow('Type', typeRefs[0], null);
            } else {
                html += bulkPropRow('Type', 'Varies (' + typeRefs.join(', ') + ')', null);
            }
            html += bulkEditRow('Width', uniqOrVaries(widths), 'mm', 'bulk-pf-width');
            html += bulkEditRow('Depth', uniqOrVaries(depths), 'mm', 'bulk-pf-depth');
            html += bulkEditRow('Setdown', uniqOrVaries(setdowns), 'mm', 'bulk-pf-setdown');
        }

        // ── Wall bulk properties ──
        else if (groupKey === 'wall') {
            const thicknesses = groupEls.map(el => {
                const td = project.scheduleTypes.wall[el.typeRef || el.tag] || {};
                return td.thickness || el.thickness || 0;
            });
            const wallTypes = groupEls.map(el => {
                const td = project.scheduleTypes.wall[el.typeRef || el.tag] || {};
                return td.wallType || '';
            });
            const typeRefs = [...new Set(groupEls.map(el => el.typeRef || el.tag || ''))];

            if (typeRefs.length === 1) {
                html += bulkPropRow('Type', typeRefs[0], null);
            } else {
                html += bulkPropRow('Type', 'Varies (' + typeRefs.join(', ') + ')', null);
            }
            html += bulkPropRow('Wall Type', uniqOrVaries(wallTypes) || '—', null);
            html += bulkEditRow('Thickness', uniqOrVaries(thicknesses), 'mm', 'bulk-wall-thickness');
            html += bulkPropRow('Total Length', fmtLen(groupEls.reduce((sum, el) =>
                sum + Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2)), 0)), null);
        }

        // ── Beam bulk properties ──
        else if (groupKey === 'beam') {
            const typeRefs = [...new Set(groupEls.map(el => el.typeRef || el.tag || ''))];
            const sizes = groupEls.map(el => {
                const td = project.scheduleTypes.beam[el.typeRef || el.tag] || {};
                return td.size || '';
            });
            if (typeRefs.length === 1) {
                html += bulkPropRow('Type', typeRefs[0], null);
            } else {
                html += bulkPropRow('Type', 'Varies (' + typeRefs.join(', ') + ')', null);
            }
            html += bulkPropRow('Size', uniqOrVaries(sizes) || '—', null);
            html += bulkPropRow('Total Length', fmtLen(groupEls.reduce((sum, el) =>
                sum + Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2)), 0)), null);
        }

        // ── Column bulk properties ──
        else if (groupKey === 'column') {
            const typeRefs = [...new Set(groupEls.map(el => el.typeRef || el.tag || ''))];
            const sizes = groupEls.map(el => {
                const td = project.scheduleTypes.column[el.typeRef || el.tag] || {};
                return td.size || '';
            });
            if (typeRefs.length === 1) {
                html += bulkPropRow('Type', typeRefs[0], null);
            } else {
                html += bulkPropRow('Type', 'Varies (' + typeRefs.join(', ') + ')', null);
            }
            html += bulkPropRow('Size', uniqOrVaries(sizes) || '—', null);
            html += bulkPropRow('Count', groupEls.length, null);
        }

        // ── Generic fallback (count only) ──
        else {
            html += bulkPropRow('Count', groupEls.length, null);
        }

        html += `</div>`;
    }

    propsBody.innerHTML = html;

    // ── Wire up bulk edit inputs ──
    // Strip footing bulk edits — postSync re-reads schedule type into each element's rendering props
    const sfSync = (els) => {
        for (const el of els) {
            const ref = el.typeRef || el.tag || 'SF1';
            const td = project.scheduleTypes.stripfooting[ref] || {};
            if (td.width) el.footingWidth = parseInt(td.width);
            if (td.depth) el.footingDepth = parseInt(td.depth);
        }
    };
    wireBulkInput('bulk-sf-width', groups.stripFooting || [], (el, val) => {
        const ref = el.typeRef || el.tag || 'SF1';
        if (!project.scheduleTypes.stripfooting[ref]) project.scheduleTypes.stripfooting[ref] = {};
        return { target: project.scheduleTypes.stripfooting[ref], field: 'width' };
    }, sfSync);
    wireBulkInput('bulk-sf-depth', groups.stripFooting || [], (el, val) => {
        const ref = el.typeRef || el.tag || 'SF1';
        if (!project.scheduleTypes.stripfooting[ref]) project.scheduleTypes.stripfooting[ref] = {};
        return { target: project.scheduleTypes.stripfooting[ref], field: 'depth' };
    }, sfSync);
    wireBulkInput('bulk-sf-setdown', groups.stripFooting || [], (el, val) => {
        return { target: el, field: 'depthBelowFSL' };
    });

    // Pad footing bulk edits — same pattern
    const pfSync = (els) => {
        for (const el of els) {
            const ref = el.typeRef || el.mark || 'PF1';
            const td = project.scheduleTypes.padfooting[ref] || {};
            if (td.width) el.footingWidth = parseInt(td.width);
            if (td.depth) el.footingDepth = parseInt(td.depth);
        }
    };
    wireBulkInput('bulk-pf-width', groups.footing || [], (el, val) => {
        const ref = el.typeRef || el.mark || 'PF1';
        if (!project.scheduleTypes.padfooting[ref]) project.scheduleTypes.padfooting[ref] = {};
        return { target: project.scheduleTypes.padfooting[ref], field: 'width' };
    }, pfSync);
    wireBulkInput('bulk-pf-depth', groups.footing || [], (el, val) => {
        const ref = el.typeRef || el.mark || 'PF1';
        if (!project.scheduleTypes.padfooting[ref]) project.scheduleTypes.padfooting[ref] = {};
        return { target: project.scheduleTypes.padfooting[ref], field: 'depth' };
    }, pfSync);
    wireBulkInput('bulk-pf-setdown', groups.footing || [], (el, val) => {
        return { target: el, field: 'depthBelowFSL' };
    });

    // Wall bulk edits — sync thickness to element for rendering
    const wallSync = (els) => {
        for (const el of els) {
            const ref = el.typeRef || el.tag || 'BW1';
            const td = project.scheduleTypes.wall[ref] || {};
            if (td.thickness) el.thickness = parseInt(td.thickness);
        }
    };
    wireBulkInput('bulk-wall-thickness', groups.wall || [], (el, val) => {
        const ref = el.typeRef || el.tag || 'BW1';
        if (!project.scheduleTypes.wall[ref]) project.scheduleTypes.wall[ref] = {};
        return { target: project.scheduleTypes.wall[ref], field: 'thickness' };
    }, wallSync);
}

/** Wire a bulk edit input: on Enter, update all elements in one undo step */
function wireBulkInput(inputId, elements, getTargetField, postSync) {
    const input = document.getElementById(inputId);
    if (!input || elements.length === 0) return;

    const commit = () => {
        const newVal = parseInt(input.value);
        if (isNaN(newVal) || newVal <= 0) return;

        // Capture old values for undo
        const changes = [];
        // Track unique schedule types to avoid duplicate updates
        const seenTargets = new Set();

        for (const el of elements) {
            const { target, field } = getTargetField(el, newVal);
            const key = (target === el ? 'el-' + el.id : JSON.stringify(target)) + ':' + field;
            if (seenTargets.has(key)) continue;
            seenTargets.add(key);
            const oldVal = target[field];
            if (oldVal === newVal) continue;
            changes.push({ target, field, oldVal, newVal });
        }

        if (changes.length === 0) return;

        history.execute({
            description: 'Bulk update ' + changes.length + ' propert' + (changes.length === 1 ? 'y' : 'ies'),
            execute() {
                for (const c of changes) c.target[c.field] = c.newVal;
                if (postSync) postSync(elements);
            },
            undo() {
                for (const c of changes) c.target[c.field] = c.oldVal;
                if (postSync) postSync(elements);
            }
        });

        _lastBulkCount = 0; // force refresh
        updateBulkPropsPanel();
        engine.requestRender();
    };

    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); commit(); }
        if (ev.key === 'Escape') { input.blur(); }
    });
    input.addEventListener('blur', () => {
        // Don't auto-commit on blur — only on Enter
    });
}

/** Helper: returns the single value if all are the same, or 'Varies' */
function uniqOrVaries(arr) {
    const unique = [...new Set(arr)];
    if (unique.length === 1) return unique[0];
    return 'Varies';
}

/** Read-only bulk property row */
function bulkPropRow(label, value) {
    return `<div class="bulk-prop-row"><span class="bulk-prop-label">${label}</span><span class="prop-value">${value}</span></div>`;
}

/** Editable bulk property row with input field */
function bulkEditRow(label, currentValue, unit, inputId) {
    const isVaries = currentValue === 'Varies';
    const displayVal = isVaries ? '' : currentValue;
    const placeholder = isVaries ? 'Varies' : '';
    return `<div class="bulk-prop-row">
        <span class="bulk-prop-label">${label}</span>
        <input type="number" class="bulk-prop-input" id="${inputId}"
               value="${displayVal}" placeholder="${placeholder}"
               title="Edit and press Enter to update all">
        ${unit ? `<span class="bulk-prop-unit">${unit}</span>` : ''}
    </div>`;
}

// Hook into render cycle: show bulk panel when multi-selected, single panel otherwise
const _origUpdatePropsPanel = updatePropsPanel;
updatePropsPanel = function() {
    const els = (typeof selectedElements !== 'undefined') ? selectedElements : [];
    if (els.length > 1) {
        updateBulkPropsPanel();
        return;
    }
    _lastBulkCount = 0;
    _origUpdatePropsPanel();
};

// Update properties panel on every render
engine.onRender(updatePropsPanel);

// ══════════════════════════════════════════════════════════
// ── SIMPLIFIED STRUCTURAL TOOLS ──────────────────────────
// ══════════════════════════════════════════════════════════

// ── Beam Tool (dedicated beam drawing) ───────────────────
// Draws a line on S-BEAM layer with solid line style

const beamBtn = document.getElementById('btn-beam');
const slabBtn = document.getElementById('btn-slab');

beamBtn.addEventListener('click', () => {
    document.getElementById('element-type').value = 'S-BEAM';
    memberCatSelect.value = 'UB';
    populateSizeDropdown();
    setActiveTool('line');
    // Update UI
    beamBtn.classList.add('active');
    slabBtn.classList.remove('active');
    document.getElementById('status-tool').textContent = 'Beam';
});

// ── Slab Tool (draw rectangle, auto-hatch concrete) ─────
// Draws a closed rectangle on S-SLAB layer with concrete hatch

slabBtn.addEventListener('click', () => {
    document.getElementById('element-type').value = 'S-SLAB';
    setActiveTool('slab');
    slabBtn.classList.add('active');
    beamBtn.classList.remove('active');
    document.getElementById('status-tool').textContent = 'Slab';
});

// B key = beam, shift+S = slab (S is already snap toggle)
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'b') { beamBtn.click(); }
});

// Patch: when a slab rectangle is committed, auto-apply concrete hatch + fill colour
const origHistExec3 = history.execute;
history.execute = function(cmd) {
    origHistExec3.call(this, cmd);

    // Check for newly created slab polylines and auto-hatch with fill colour
    const slabColorInput = document.getElementById('slab-color');
    const slabColor = slabColorInput ? slabColorInput.value : '#E8D8A0';
    for (const el of project.elements) {
        if (el.type === 'polyline' && el.layer === 'S-SLAB' && el.closed && !el.hatch) {
            el.hatch = 'concrete';
            el.fillColor = slabColor;
            el.fillOpacity = 0.2;  // 20% opacity for sand/yellow
        }
    }
};

// Update tool button highlighting when tools change
// When column tool is selected, default to SHS category
document.getElementById('btn-column').addEventListener('click', () => {
    memberCatSelect.value = 'SHS';
    populateSizeDropdown();
    // Default to 89x5.0 SHS for steel columns
    if (memberSizeSelect.querySelector('option[value="89x89x5SHS"]')) {
        memberSizeSelect.value = '89x89x5SHS';
    }
    beamBtn.classList.remove('active');
    slabBtn.classList.remove('active');
}, true);

// ══════════════════════════════════════════════════════════
