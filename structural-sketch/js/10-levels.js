// ── 3D PHASE 1: MULTI-LEVEL SYSTEM ───────────────────────
// ══════════════════════════════════════════════════════════

// ── Level Data Model ─────────────────────────────────────

const levelSystem = {
    levels: [
        { id: 'GF',  name: 'Ground Floor', height: 2700, elevation: 0, subSlabGap: 200 },
        { id: 'L1',  name: 'Level 1',      height: 2700, elevation: 2700 },
        { id: 'L2',  name: 'Level 2',      height: 2700, elevation: 5400 },
        { id: 'RF',  name: 'Roof',          height: 0,    elevation: 8100 },
    ],
    activeLevelIndex: 0,
    groundRL: 0,          // RL of Ground Floor top-of-slab (datum) in metres, e.g. 13.29
    showTOFTags: false,   // Toggle TOF tags on plan
};

// Auto-detect GF slab thickness from drawn slabs
function getGFSlabThickness() {
    const gfSlabs = project.elements.filter(el =>
        el.level === 'GF' && el.type === 'polyline' && el.layer === 'S-SLAB'
    );
    if (gfSlabs.length > 0) return gfSlabs[0].slabThickness || 200;
    return 200; // fallback if no slab drawn yet
}

// ── RL & TOF Calculation Helpers ──────────────────────────

/** Get the RL (in metres) of a given level's top-of-slab */
function getLevelRL(levelId) {
    const lv = levelSystem.levels.find(l => l.id === levelId) || levelSystem.levels[0];
    return levelSystem.groundRL + (lv.elevation / 1000);
}

/** Get Top-of-Footing RL (metres) for a pad footing element.
 *  If the footing has a custom tofOverride (in metres), use that.
 *  Otherwise: Level RL - setdown (mm→m). */
function getFootingTOF(el) {
    if (el.tofOverride !== undefined && el.tofOverride !== null && el.tofOverride !== '') {
        return parseFloat(el.tofOverride);
    }
    const levelRL = getLevelRL(el.level || 'GF');
    const setdownM = (el.depthBelowFSL || 200) / 1000;
    return levelRL - setdownM;
}

/** Get the setdown in mm implied by a given TOF RL override */
function tofToSetdown(tofRL, levelId) {
    const levelRL = getLevelRL(levelId || 'GF');
    return Math.round((levelRL - tofRL) * 1000);
}

// Legacy helper — elevation relative to GF in mm (used by 3D, section views)
function getFootingTopElevation(el) {
    const depthBelowFSL = el.depthBelowFSL || 200;
    return -depthBelowFSL;
}

function getActiveLevel() {
    return levelSystem.levels[levelSystem.activeLevelIndex];
}

// Convenience getter — bracing code references levelSystem.currentLevel
Object.defineProperty(levelSystem, 'currentLevel', {
    get() { return levelSystem.levels[levelSystem.activeLevelIndex]?.id || 'GF'; },
    enumerable: false
});

function recalcElevations() {
    let elev = 0;
    for (let i = 0; i < levelSystem.levels.length; i++) {
        levelSystem.levels[i].elevation = elev;
        elev += levelSystem.levels[i].height;
    }
}

// ── Assign Level to New Elements ─────────────────────────

// Patch: every element creation should stamp the active level ID.
// We'll use a global hook that ProjectData checks.
// Override project.elements.push to auto-stamp level
const origPush = Array.prototype.push;
const elementsProxy = project.elements;

// Simpler approach: patch into each tool's element creation.
// We'll add a `_stampLevel` function called before each history.execute
function stampLevel(el) {
    if (!el.level) {
        el.level = getActiveLevel().id;
    }
    return el;
}

// Patch existing element creation by adding level to elements that don't have one
// (elements created before this phase won't have a level — assign to GF)
for (const el of project.elements) {
    if (!el.level) el.level = 'GF';
}

// Hook into all mousedown/commit handlers: when a new element is created,
// stamp it. We'll do this by overriding generateId to also set a global
// "next element level" flag.
const origGenerateId = generateId;
// Actually, simplest: patch the history.execute to stamp any new element
const origHistoryExecute = history.execute.bind(history);
history.execute = function(cmd) {
    // Wrap the execute to stamp level on any elements pushed
    const origExec = cmd.execute;
    cmd.execute = function() {
        origExec();
        // Find newly added elements without a level and stamp them
        for (const el of project.elements) {
            if (!el.level) el.level = getActiveLevel().id;
        }
    };
    origHistoryExecute(cmd);
};

// ── Element Filtering by Level ───────────────────────────

// Override getVisibleElements to filter by active level
const origGetVisible = project.getVisibleElements.bind(project);
project.getVisibleElements = function() {
    // V18: When on details sheet, return empty array — details sheet renders its own content
    if (typeof detailsSystem !== 'undefined' && detailsSystem.activeSheet === 'details') {
        return [];
    }
    const activeId = getActiveLevel().id;
    return origGetVisible().filter(el => el.level === activeId);
};

// Also need a method that gets elements from OTHER levels (for ghost)
project.getGhostElements = function() {
    const activeIdx = levelSystem.activeLevelIndex;
    if (activeIdx <= 0) return []; // no level below ground
    const belowId = levelSystem.levels[activeIdx - 1].id;
    return this.elements.filter(el => {
        const layer = this.layers[el.layer];
        return layer && layer.visible && el.level === belowId;
    });
};

// For getting ALL visible (unfiltered) — needed for save/export
project.getAllVisibleElements = origGetVisible;

// ── Ghost Underlay Rendering ─────────────────────────────

function drawGhostUnderlay(ctx, eng) {
    const ghostEls = project.getGhostElements();
    if (ghostEls.length === 0) return;

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    ctx.save();

    for (const el of ghostEls) {
        const layer = project.layers[el.layer];
        if (!layer) continue;

        // ── Ghost Walls ──────────────────────────────────
        if (el.type === 'wall') {
            const wdx = el.x2 - el.x1;
            const wdy = el.y2 - el.y1;
            const wlen = Math.sqrt(wdx * wdx + wdy * wdy);
            if (wlen < 0.1) continue;

            const nx = -wdy / wlen, ny = wdx / wlen;
            const halfT = el.thickness / 2;

            // Convert corner points to screen
            const s1a = coords.realToScreen(el.x1 + nx * halfT, el.y1 + ny * halfT);
            const s1b = coords.realToScreen(el.x1 - nx * halfT, el.y1 - ny * halfT);
            const s2a = coords.realToScreen(el.x2 + nx * halfT, el.y2 + ny * halfT);
            const s2b = coords.realToScreen(el.x2 - nx * halfT, el.y2 - ny * halfT);

            // Subtle grey fill for wall width
            ctx.globalAlpha = 0.12;
            ctx.fillStyle = '#B0B0B0';
            ctx.beginPath();
            ctx.moveTo(s1a.x, s1a.y);
            ctx.lineTo(s2a.x, s2a.y);
            ctx.lineTo(s2b.x, s2b.y);
            ctx.lineTo(s1b.x, s1b.y);
            ctx.closePath();
            ctx.fill();

            // Dashed grey outline
            ctx.globalAlpha = 0.3;
            ctx.strokeStyle = '#999999';
            ctx.lineWidth = Math.max(0.5, 0.3 * zoom);
            ctx.setLineDash([3 * zoom, 2 * zoom]);
            ctx.beginPath();
            ctx.moveTo(s1a.x, s1a.y);
            ctx.lineTo(s2a.x, s2a.y);
            ctx.lineTo(s2b.x, s2b.y);
            ctx.lineTo(s1b.x, s1b.y);
            ctx.closePath();
            ctx.stroke();
            ctx.setLineDash([]);

            // Ghost wall tag
            if (el.tag && zoom > 0.5) {
                const mx = (s1a.x + s2b.x) / 2, my = (s1a.y + s2b.y) / 2;
                const fontSize = Math.max(1, 1.8 * zoom);
                ctx.globalAlpha = 0.25;
                ctx.font = `${fontSize}px "Segoe UI", Arial, sans-serif`;
                ctx.fillStyle = '#666666';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.fillText(el.tag, mx, my);
            }
            continue;
        }

        // ── Ghost Strip Footings ─────────────────────────
        if (el.type === 'stripFooting') {
            const p1 = coords.realToScreen(el.x1, el.y1);
            const p2 = coords.realToScreen(el.x2, el.y2);
            const dx = p2.x - p1.x, dy = p2.y - p1.y;
            const len = Math.sqrt(dx * dx + dy * dy);
            if (len < 0.5) continue;
            const halfW = ((el.footingWidth || 600) / 2 / CONFIG.drawingScale) * zoom;
            const ux = dx / len, uy = dy / len;
            const nnx = -uy * halfW, nny = ux * halfW;

            // Subtle fill
            ctx.globalAlpha = 0.1;
            ctx.fillStyle = '#B0B0B0';
            ctx.beginPath();
            ctx.moveTo(p1.x + nnx, p1.y + nny);
            ctx.lineTo(p2.x + nnx, p2.y + nny);
            ctx.lineTo(p2.x - nnx, p2.y - nny);
            ctx.lineTo(p1.x - nnx, p1.y - nny);
            ctx.closePath();
            ctx.fill();

            // Dashed outline
            ctx.globalAlpha = 0.25;
            ctx.strokeStyle = '#999999';
            ctx.lineWidth = Math.max(0.5, 0.2 * zoom);
            ctx.setLineDash([3 * zoom, 2 * zoom]);
            ctx.stroke();
            ctx.setLineDash([]);
            continue;
        }

        // ── Ghost Pad Footings ───────────────────────────
        if (el.type === 'footing') {
            const pos = coords.realToScreen(el.x, el.y);
            const fW = el.footingWidth || el.width || 1000;
            const halfW = (fW / 2 / CONFIG.drawingScale) * zoom;

            ctx.globalAlpha = 0.1;
            ctx.fillStyle = '#B0B0B0';
            ctx.fillRect(pos.x - halfW, pos.y - halfW, halfW * 2, halfW * 2);

            ctx.globalAlpha = 0.25;
            ctx.strokeStyle = '#999999';
            ctx.lineWidth = Math.max(0.5, 0.2 * zoom);
            ctx.setLineDash([3 * zoom, 2 * zoom]);
            ctx.strokeRect(pos.x - halfW, pos.y - halfW, halfW * 2, halfW * 2);
            ctx.setLineDash([]);
            continue;
        }

        // ── Existing ghost types (lines, columns, polylines, text) ──
        ctx.globalAlpha = 0.15;
        ctx.strokeStyle = '#888888';
        ctx.fillStyle = '#888888';
        ctx.lineWidth = Math.max(0.5, layer.lineWeight * zoom * 0.5);
        ctx.setLineDash([2 * zoom, 2 * zoom]);

        if (el.type === 'line' || el.type === 'leader') {
            const p1 = coords.realToScreen(el.x1, el.y1);
            const p2 = coords.realToScreen(el.x2, el.y2);
            ctx.beginPath(); ctx.moveTo(p1.x, p1.y); ctx.lineTo(p2.x, p2.y); ctx.stroke();
        }
        if (el.type === 'column') {
            const cp = coords.realToScreen(el.x, el.y);
            const r = Math.max(3, (el.size / CONFIG.drawingScale) * zoom / 2);
            ctx.strokeRect(cp.x - r, cp.y - r, r * 2, r * 2);
        }
        if (el.type === 'polyline') {
            const pts = el.points;
            if (pts.length < 2) continue;
            ctx.beginPath();
            const p0 = coords.realToScreen(pts[0].x, pts[0].y);
            ctx.moveTo(p0.x, p0.y);
            for (let i = 1; i < pts.length; i++) {
                const p = coords.realToScreen(pts[i].x, pts[i].y);
                ctx.lineTo(p.x, p.y);
            }
            if (el.closed) ctx.closePath();
            ctx.stroke();
        }
        if (el.type === 'text') {
            const tp = coords.realToScreen(el.x, el.y);
            const fontSize = Math.max(7, (el.fontSize || 3.5) * zoom);
            ctx.font = `${fontSize}px "Architects Daughter", cursive`;
            ctx.textAlign = 'left';
            ctx.textBaseline = 'middle';
            ctx.fillText(el.text, tp.x, tp.y);
        }

        ctx.setLineDash([]);
    }

    ctx.restore();
}

// Register ghost renderer — draw BEFORE current level elements
// Insert early in the callback chain (after pdf, statusbar, grids, hatch)
engine._renderCallbacks.splice(4, 0, drawGhostUnderlay);

// ── Level Switcher UI ────────────────────────────────────

const levelTabsEl = document.getElementById('level-tabs');
const levelPrevBtn = document.getElementById('level-prev');
const levelNextBtn = document.getElementById('level-next');

function buildLevelTabs() {
    levelTabsEl.innerHTML = '';
    for (let i = 0; i < levelSystem.levels.length; i++) {
        const lv = levelSystem.levels[i];
        const tab = document.createElement('button');
        tab.className = 'level-tab' + (i === levelSystem.activeLevelIndex ? ' active' : '');

        const htText = lv.height > 0 ? (lv.height / 1000).toFixed(1) + 'm' : '';
        tab.innerHTML = lv.name + (htText ? '<span class="level-ht">' + htText + '</span>' : '');

        tab.addEventListener('click', () => {
            switchToLevel(i);
        });
        levelTabsEl.appendChild(tab);
    }

    levelPrevBtn.disabled = levelSystem.activeLevelIndex <= 0;
    levelNextBtn.disabled = levelSystem.activeLevelIndex >= levelSystem.levels.length - 1;
}

function switchToLevel(index) {
    if (index < 0 || index >= levelSystem.levels.length) return;
    levelSystem.activeLevelIndex = index;

    // Clear selection when switching levels
    selectedElement = null;
    selectedElements = [];

    buildLevelTabs();
    engine.requestRender();
    updateStatusBar();

    // Update drawing title to show level name
    const lv = getActiveLevel();
    document.getElementById('status-tool').textContent =
        document.getElementById('status-tool').textContent; // keep current tool
    console.log('[Level] Switched to: ' + lv.name + ' (elevation: ' + lv.elevation + 'mm)');
}

levelPrevBtn.addEventListener('click', () => {
    switchToLevel(levelSystem.activeLevelIndex - 1);
});

levelNextBtn.addEventListener('click', () => {
    switchToLevel(levelSystem.activeLevelIndex + 1);
});

// Page Up / Page Down for level switching
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.key === 'PageUp') {
        e.preventDefault();
        switchToLevel(levelSystem.activeLevelIndex + 1);
    }
    if (e.key === 'PageDown') {
        e.preventDefault();
        switchToLevel(levelSystem.activeLevelIndex - 1);
    }
});

// ── Level Manager Dialog ─────────────────────────────────

const levelManager = document.getElementById('level-manager');
const levelMgrBody = document.getElementById('level-mgr-body');

document.getElementById('level-manage-btn').addEventListener('click', () => {
    openLevelManager();
});

function openLevelManager() {
    buildLevelManagerTable();
    levelManager.classList.remove('hidden');
}

function buildLevelManagerTable() {
    levelMgrBody.innerHTML = '';
    for (let i = levelSystem.levels.length - 1; i >= 0; i--) {
        const lv = levelSystem.levels[i];
        const tr = document.createElement('tr');

        const tdName = document.createElement('td');
        const nameInput = document.createElement('input');
        nameInput.type = 'text';
        nameInput.value = lv.name;
        nameInput.dataset.idx = i;
        nameInput.dataset.field = 'name';
        tdName.appendChild(nameInput);

        const tdHeight = document.createElement('td');
        const htInput = document.createElement('input');
        htInput.type = 'number';
        htInput.value = lv.height;
        htInput.min = 0;
        htInput.step = 100;
        htInput.dataset.idx = i;
        htInput.dataset.field = 'height';
        tdHeight.appendChild(htInput);

        const tdElev = document.createElement('td');
        tdElev.className = 'level-elev';
        tdElev.textContent = (lv.elevation / 1000).toFixed(1) + ' m';

        const tdDel = document.createElement('td');
        if (levelSystem.levels.length > 1) {
            const delBtn = document.createElement('button');
            delBtn.className = 'del-btn';
            delBtn.textContent = '✕';
            delBtn.title = 'Delete level';
            delBtn.addEventListener('click', () => {
                levelSystem.levels.splice(i, 1);
                if (levelSystem.activeLevelIndex >= levelSystem.levels.length) {
                    levelSystem.activeLevelIndex = levelSystem.levels.length - 1;
                }
                recalcElevations();
                buildLevelManagerTable();
            });
            tdDel.appendChild(delBtn);
        }

        tr.appendChild(tdName);
        tr.appendChild(tdHeight);
        tr.appendChild(tdElev);
        tr.appendChild(tdDel);
        levelMgrBody.appendChild(tr);
    }
}

document.getElementById('level-add-btn').addEventListener('click', () => {
    const topLevel = levelSystem.levels[levelSystem.levels.length - 1];
    const newId = 'L' + (levelSystem.levels.length);
    levelSystem.levels.push({
        id: newId,
        name: 'Level ' + levelSystem.levels.length,
        height: 2700,
        elevation: topLevel.elevation + topLevel.height,
    });
    recalcElevations();
    buildLevelManagerTable();
});

document.getElementById('level-mgr-apply').addEventListener('click', () => {
    // Read values from the table inputs
    const inputs = levelMgrBody.querySelectorAll('input');
    for (const input of inputs) {
        const idx = parseInt(input.dataset.idx);
        const field = input.dataset.field;
        if (field === 'name') {
            levelSystem.levels[idx].name = input.value;
        } else if (field === 'height') {
            levelSystem.levels[idx].height = parseInt(input.value) || 0;
        }
    }
    recalcElevations();
    levelManager.classList.add('hidden');
    buildLevelTabs();
    engine.requestRender();
});

document.getElementById('level-mgr-cancel').addEventListener('click', () => {
    levelManager.classList.add('hidden');
});

// ── Level indicator on status bar ────────────────────────

// Add level name to the status bar
const statusBarEl = document.getElementById('status-bar');
const levelStatusItem = document.createElement('div');
levelStatusItem.className = 'status-item';
levelStatusItem.innerHTML = '<span class="status-label">Level</span><span class="status-value" id="status-level">Ground Floor</span>';
statusBarEl.insertBefore(levelStatusItem, statusBarEl.firstChild);

const statusLevelEl = document.getElementById('status-level');

// Patch updateStatusBar to include level
const origUpdateStatus = updateStatusBar;
updateStatusBar = function(e) {
    origUpdateStatus(e);
    if (statusLevelEl) {
        statusLevelEl.textContent = getActiveLevel().name;
    }
};

// ── Save/Load: include levels ────────────────────────────

// Patch saveProject to include level data
const origSaveProject = saveProject;
saveProject = function() {
    // Temporarily store levels in project for save
    const data = {
        version: 3,
        appName: 'StructuralSketch',
        savedAt: new Date().toISOString(),
        config: {
            drawingScale: CONFIG.drawingScale,
            gridVisible: CONFIG.gridVisible,
        },
        projectInfo: project.projectInfo,
        elements: project.elements, // ALL elements, not filtered
        layers: project.layers,
        activeLayerId: project.activeLayerId,
        structuralGrids: structuralGrids.map(g => ({
            id: g.id, axis: g.axis, position: g.position, label: g.label
        })),
        gridLabelState: gridLabelState,
        nextId: _nextId,
        colNextNum: _colNextNum,
        ftgNextNum: _ftgNextNum,
        // 3D-specific
        levels: levelSystem.levels,
        activeLevelIndex: levelSystem.activeLevelIndex,
        // Schedule types
        scheduleTypes: project.scheduleTypes,
        // Bracing settings
        bracingSettings: typeof bracingSettings !== 'undefined' ? bracingSettings : undefined,
    };

    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const name = (project.projectInfo.drawingNumber || 'drawing').replace(/[^a-zA-Z0-9_-]/g, '_');
    a.href = url;
    a.download = name + '.sdraw';
    a.click();
    URL.revokeObjectURL(url);
    console.log('[Save] Project saved as ' + name + '.sdraw');
};

// Patch loadProject to restore levels and scheduleTypes
const origLoadProject = loadProject;
loadProject = function(json) {
    origLoadProject(json);
    try {
        const data = JSON.parse(json);
        if (data.levels) {
            levelSystem.levels = data.levels;
            levelSystem.activeLevelIndex = data.activeLevelIndex || 0;
            recalcElevations();
        }
        if (data.groundRL !== undefined) levelSystem.groundRL = data.groundRL;
        if (data.showTOFTags !== undefined) levelSystem.showTOFTags = data.showTOFTags;
        if (data.scheduleTypes) {
            project.scheduleTypes = data.scheduleTypes;
            // Migrate old B-prefix beams to SB-prefix
            if (project.scheduleTypes.beam) {
                const beams = project.scheduleTypes.beam;
                for (const key of Object.keys(beams)) {
                    if (/^B\d+$/.test(key)) {
                        const newKey = 'S' + key;
                        beams[newKey] = beams[key];
                        // Migrate old 'category' field to 'sectionType'
                        if (beams[newKey].category && !beams[newKey].sectionType) {
                            beams[newKey].sectionType = beams[newKey].category;
                        }
                        if (!beams[newKey].description) beams[newKey].description = '';
                        if (!beams[newKey].grade) beams[newKey].grade = '300';
                        delete beams[key];
                    }
                }
            }
            // Migrate old 'category' field to 'sectionType' for columns
            if (project.scheduleTypes.column) {
                for (const td of Object.values(project.scheduleTypes.column)) {
                    if (td.category && !td.sectionType) {
                        td.sectionType = td.category;
                    }
                    if (!td.description) td.description = '';
                    if (!td.grade) td.grade = '300';
                }
            }
            // Add description field to walls if missing
            if (project.scheduleTypes.wall) {
                for (const td of Object.values(project.scheduleTypes.wall)) {
                    if (!td.description) td.description = '';
                }
            }
            // Also migrate element typeRefs and tags
            for (const el of project.elements) {
                if (el.type === 'line' && el.layer === 'S-BEAM') {
                    if (el.typeRef && /^B\d+$/.test(el.typeRef)) el.typeRef = 'S' + el.typeRef;
                    if (el.tag && /^B\d+/.test(el.tag)) el.tag = 'S' + el.tag;
                }
            }
        }
        // Restore bracing settings (merge with defaults for forward-compat)
        if (data.bracingSettings && typeof bracingSettings !== 'undefined') {
            Object.assign(bracingSettings, data.bracingSettings);
        }
    } catch (e) { /* ignore parse errors — base loader handles it */ }
    buildLevelTabs();
    updateStatusBar();
};

// ── Initialise Level UI ──────────────────────────────────

buildLevelTabs();
window._app.levelSystem = levelSystem;

// ══════════════════════════════════════════════════════════
// ── 3D PHASE 2: COLUMN VERTICAL EXTENSION ────────────────
// ══════════════════════════════════════════════════════════

// ── Column Direction Popup ───────────────────────────────

const colPopup = document.getElementById('col-popup');
let colPopupTarget = null;

function showColPopup(col, screenX, screenY) {
    colPopupTarget = col;

    // Position popup near the column
    colPopup.style.left = (screenX + 10) + 'px';
    colPopup.style.top = (screenY - 40) + 'px';
    colPopup.classList.remove('hidden');

    // Highlight current direction
    for (const btn of colPopup.querySelectorAll('.col-popup-btn')) {
        btn.classList.toggle('active', btn.dataset.dir === (col.extends || 'below'));
    }
}

function hideColPopup() {
    colPopup.classList.add('hidden');
    colPopupTarget = null;
}

// Click handlers for popup buttons
for (const btn of colPopup.querySelectorAll('.col-popup-btn')) {
    btn.addEventListener('click', (e) => {
        e.stopPropagation();
        if (!colPopupTarget) return;

        const el = colPopupTarget;
        const oldDir = el.extends || 'below';
        const newDir = btn.dataset.dir;

        if (oldDir !== newDir) {
            history.execute({
                description: 'Set column ' + el.tag + ' extends: ' + newDir,
                execute() { el.extends = newDir; },
                undo() { el.extends = oldDir; }
            });
        }

        hideColPopup();
        engine.requestRender();
    });
}

// Show popup on double-click on a column
container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'select') return;
    if (!selectedElement || selectedElement.type !== 'column') return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    showColPopup(selectedElement, sx, sy);
});

// Hide popup when clicking elsewhere
container.addEventListener('mousedown', (e) => {
    if (colPopupTarget && !colPopup.contains(e.target)) {
        hideColPopup();
    }
});

window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && colPopupTarget) hideColPopup();
});

// ── Inherited Columns from Adjacent Levels ───────────────

/**
 * Get columns from other levels that extend into the current level.
 * - Columns on the level above with extends='below' or 'both' → show on current level
 * - Columns on the level below with extends='above' or 'both' → show on current level
 */
function getInheritedColumns() {
    const activeIdx = levelSystem.activeLevelIndex;
    const inherited = [];

    // Columns from level ABOVE that extend down
    if (activeIdx < levelSystem.levels.length - 1) {
        const aboveId = levelSystem.levels[activeIdx + 1].id;
        for (const el of project.elements) {
            if (el.type !== 'column' || el.level !== aboveId) continue;
            if (el.extends === 'below' || el.extends === 'both') {
                inherited.push({ ...el, _inherited: true, _fromLevel: aboveId, _direction: 'from above' });
            }
        }
    }

    // Columns from level BELOW that extend up
    if (activeIdx > 0) {
        const belowId = levelSystem.levels[activeIdx - 1].id;
        for (const el of project.elements) {
            if (el.type !== 'column' || el.level !== belowId) continue;
            if (el.extends === 'above' || el.extends === 'both') {
                inherited.push({ ...el, _inherited: true, _fromLevel: belowId, _direction: 'from below' });
            }
        }
    }

    return inherited;
}

// ── Render Inherited Columns ─────────────────────────────

function drawInheritedColumns(ctx, eng) {
    const inherited = getInheritedColumns();
    if (inherited.length === 0) return;

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    ctx.save();
    ctx.globalAlpha = 0.35;

    for (const col of inherited) {
        const cp = coords.realToScreen(col.x, col.y);
        const halfSize = (col.size / CONFIG.drawingScale) * zoom / 2;
        const r = Math.max(3, halfSize);

        // Dashed outline (inherited style)
        ctx.strokeStyle = '#2B7CD0';
        ctx.lineWidth = Math.max(1, 0.35 * zoom);
        ctx.setLineDash([3 * zoom, 2 * zoom]);
        ctx.strokeRect(cp.x - r, cp.y - r, r * 2, r * 2);

        // X cross
        ctx.beginPath();
        ctx.moveTo(cp.x - r, cp.y - r); ctx.lineTo(cp.x + r, cp.y + r);
        ctx.moveTo(cp.x + r, cp.y - r); ctx.lineTo(cp.x - r, cp.y + r);
        ctx.stroke();
        ctx.setLineDash([]);

        // Tag + direction indicator — show size if assigned, else type ref
        if (col.tag && zoom > 0.3) {
            const colTypeRef = col.typeRef || col.tag;
            const colSchedData = project.scheduleTypes.column[colTypeRef];
            const shortTag = (colSchedData && colSchedData.size) ? colSchedData.size : colTypeRef;
            const arrow = col._direction === 'from above' ? '↓' : '↑';
            const fontSize = Math.max(6, 1.8 * zoom);
            ctx.save();
            ctx.translate(cp.x + r + 3, cp.y + r + 3);
            ctx.rotate(-Math.PI / 4);
            ctx.font = `${fontSize}px "Segoe UI", Arial, sans-serif`;
            ctx.fillStyle = '#2B7CD0';
            ctx.textAlign = 'left';
            ctx.textBaseline = 'middle';
            ctx.fillText(shortTag + arrow, 0, 0);
            ctx.restore();
        }
    }

    ctx.restore();
    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';
}

// Register inherited columns renderer — after ghost underlay, before main elements
engine._renderCallbacks.splice(5, 0, drawInheritedColumns);

// ── Column Direction Indicator in Main Rendering ──────────

// Patch: add a small direction arrow on column rendering.
// We'll add this as a post-render overlay on the existing column drawing.
function drawColumnDirectionIndicators(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;
    if (zoom < 0.5) return; // too small to show indicators

    for (const el of project.getVisibleElements()) {
        if (el.type !== 'column') continue;

        const cp = coords.realToScreen(el.x, el.y);
        const r = Math.max(3, (el.size / CONFIG.drawingScale) * zoom / 2);
        const dir = el.extends || 'below';

        // Small arrow indicator above/below the column square
        ctx.fillStyle = '#666';
        ctx.font = `${Math.max(8, 2.5 * zoom)}px "Segoe UI", Arial, sans-serif`;
        ctx.textAlign = 'center';

        if (dir === 'below' || dir === 'both') {
            ctx.textBaseline = 'top';
            ctx.fillText('↓', cp.x + r + 6, cp.y - 3);
        }
        if (dir === 'above' || dir === 'both') {
            ctx.textBaseline = 'bottom';
            ctx.fillText('↑', cp.x + r + 6, cp.y + 3);
        }
    }

    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';
}

// Register after main element rendering
engine.onRender(drawColumnDirectionIndicators);

// ── Existing Columns: Backfill extends Property ──────────

// Ensure all existing columns have an extends property
for (const el of project.elements) {
    if (el.type === 'column' && !el.extends) {
        // Determine based on level
        const levelIdx = levelSystem.levels.findIndex(l => l.id === el.level);
        el.extends = (levelIdx === 0) ? 'above' : 'below';
    }
}


// ══════════════════════════════════════════════════════════
