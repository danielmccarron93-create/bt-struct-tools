// ── 3D PHASE 9: CHAIN DIMS + AUTO COLS + BEAM SNAP ───────
// ══════════════════════════════════════════════════════════

// ── Chain Dimension Tool ─────────────────────────────────
// Click multiple points in sequence, right-click to finish.
// Renders as a continuous dimension string + overall dimension above.

const chainDimState = {
    points: [],      // sheet-mm points
    currentEnd: null,
};

document.getElementById('btn-chain-dim').addEventListener('click', () => setActiveTool('chaindim'));

window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.key === 'D' && !e.ctrlKey && !e.metaKey) setActiveTool('chaindim'); // Shift+D
});

function getChainDimPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    if (chainDimState.points.length > 0) {
        const last = chainDimState.points[chainDimState.points.length - 1];
        pos = applyOrtho(pos.x, pos.y, last.x, last.y);
    }
    return pos;
}

container.addEventListener('mousemove', (e) => {
    if (activeTool === 'chaindim') {
        chainDimState.currentEnd = getChainDimPos(e);
        engine.requestRender();
    }
});

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'chaindim') return;
    chainDimState.points.push(getChainDimPos(e));
    engine.requestRender();
});

container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'chaindim' && chainDimState.points.length >= 2) {
        e.preventDefault();
        commitChainDim();
    } else if (activeTool === 'chaindim') {
        e.preventDefault();
        chainDimState.points = [];
        engine.requestRender();
    }
});

window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && activeTool === 'chaindim' && chainDimState.points.length >= 2) {
        commitChainDim();
    }
});

function commitChainDim() {
    const pts = chainDimState.points;
    if (pts.length < 2) return;

    const realPts = pts.map(p => engine.coords.sheetToReal(p.x, p.y));

    const newChainDim = {
        id: generateId(),
        type: 'chaindim',
        layer: 'S-DIMS',
        points: realPts,
        offset: 5, // mm offset from measured line
    };

    history.execute({
        description: 'Chain dimension',
        execute() { project.elements.push(newChainDim); },
        undo() {
            const i = project.elements.indexOf(newChainDim);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    chainDimState.points = [];
    chainDimState.currentEnd = null;
    engine.requestRender();
}

// ── Chain Dimension Rendering ────────────────────────────

function drawChainDimensions(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    for (const el of project.getVisibleElements()) {
        if (el.type !== 'chaindim') continue;
        const pts = el.points;
        if (pts.length < 2) continue;

        const isSelected = isElementSelected(el);
        const color = isSelected ? '#2B7CD0' : '#000000';
        const offsetMM = (el.offset || 5);

        // Determine direction (horizontal or vertical chain)
        const totalDx = Math.abs(pts[pts.length-1].x - pts[0].x);
        const totalDy = Math.abs(pts[pts.length-1].y - pts[0].y);
        const isHoriz = totalDx >= totalDy;

        // Screen-space positions
        const screenPts = pts.map(p => coords.realToScreen(p.x, p.y));
        const offsetPx = offsetMM * zoom;
        const overallOffsetPx = offsetPx + 8 * zoom;

        ctx.strokeStyle = color;
        ctx.fillStyle = color;
        ctx.lineWidth = Math.max(0.5, 0.18 * zoom);
        ctx.setLineDash([]);

        // Individual segment dimensions
        for (let i = 0; i < screenPts.length - 1; i++) {
            const p1 = screenPts[i], p2 = screenPts[i + 1];
            let d1, d2;
            if (isHoriz) {
                d1 = { x: p1.x, y: p1.y - offsetPx };
                d2 = { x: p2.x, y: p2.y - offsetPx };
            } else {
                d1 = { x: p1.x - offsetPx, y: p1.y };
                d2 = { x: p2.x - offsetPx, y: p2.y };
            }

            // Extension lines
            ctx.beginPath();
            if (isHoriz) {
                ctx.moveTo(p1.x, p1.y - 1.5 * zoom); ctx.lineTo(d1.x, d1.y - 1.5 * zoom);
                ctx.moveTo(p2.x, p2.y - 1.5 * zoom); ctx.lineTo(d2.x, d2.y - 1.5 * zoom);
            } else {
                ctx.moveTo(p1.x - 1.5 * zoom, p1.y); ctx.lineTo(d1.x - 1.5 * zoom, d1.y);
                ctx.moveTo(p2.x - 1.5 * zoom, p2.y); ctx.lineTo(d2.x - 1.5 * zoom, d2.y);
            }
            ctx.stroke();

            // Dim line
            ctx.beginPath(); ctx.moveTo(d1.x, d1.y); ctx.lineTo(d2.x, d2.y); ctx.stroke();

            // Tick marks
            const tk = 2.5 * zoom;
            ctx.lineWidth = Math.max(1, 0.25 * zoom);
            for (const tp of [d1, d2]) {
                ctx.beginPath();
                ctx.moveTo(tp.x - tk, tp.y + tk);
                ctx.lineTo(tp.x + tk, tp.y - tk);
                ctx.stroke();
            }
            ctx.lineWidth = Math.max(0.5, 0.18 * zoom);

            // Segment measurement
            const segDist = isHoriz
                ? Math.abs(pts[i + 1].x - pts[i].x)
                : Math.abs(pts[i + 1].y - pts[i].y);
            const dimText = segDist >= 1000 ? (segDist / 1000).toFixed(segDist >= 10000 ? 1 : 2) + ' m' : Math.round(segDist) + '';
            const midX = (d1.x + d2.x) / 2, midY = (d1.y + d2.y) / 2;
            const fontSize = Math.max(7, 2 * zoom);
            ctx.font = `${fontSize}px "Segoe UI", Arial, sans-serif`;
            ctx.textAlign = 'center';
            if (isHoriz) {
                ctx.textBaseline = 'bottom';
                ctx.fillText(dimText, midX, midY - 1.5);
            } else {
                ctx.save();
                ctx.translate(midX, midY);
                ctx.rotate(-Math.PI / 2);
                ctx.textBaseline = 'bottom';
                ctx.fillText(dimText, 0, -1.5);
                ctx.restore();
            }
        }

        // Overall dimension (above the chain)
        if (screenPts.length > 2) {
            const p1 = screenPts[0], p2 = screenPts[screenPts.length - 1];
            let o1, o2;
            if (isHoriz) {
                o1 = { x: p1.x, y: p1.y - overallOffsetPx };
                o2 = { x: p2.x, y: p2.y - overallOffsetPx };
            } else {
                o1 = { x: p1.x - overallOffsetPx, y: p1.y };
                o2 = { x: p2.x - overallOffsetPx, y: p2.y };
            }

            ctx.lineWidth = Math.max(0.5, 0.18 * zoom);
            ctx.beginPath(); ctx.moveTo(o1.x, o1.y); ctx.lineTo(o2.x, o2.y); ctx.stroke();

            const tk = 2.5 * zoom;
            ctx.lineWidth = Math.max(1, 0.25 * zoom);
            for (const tp of [o1, o2]) {
                ctx.beginPath(); ctx.moveTo(tp.x - tk, tp.y + tk); ctx.lineTo(tp.x + tk, tp.y - tk); ctx.stroke();
            }

            const totalDist = isHoriz
                ? Math.abs(pts[pts.length-1].x - pts[0].x)
                : Math.abs(pts[pts.length-1].y - pts[0].y);
            const totalText = totalDist >= 1000 ? (totalDist / 1000).toFixed(1) + ' m' : Math.round(totalDist) + '';
            const omx = (o1.x + o2.x) / 2, omy = (o1.y + o2.y) / 2;
            const fontSize2 = Math.max(8, 2.2 * zoom);
            ctx.font = `bold ${fontSize2}px "Segoe UI", Arial, sans-serif`;
            ctx.textAlign = 'center';
            if (isHoriz) {
                ctx.textBaseline = 'bottom';
                ctx.fillText(totalText, omx, omy - 1.5);
            } else {
                ctx.save(); ctx.translate(omx, omy); ctx.rotate(-Math.PI / 2);
                ctx.textBaseline = 'bottom'; ctx.fillText(totalText, 0, -1.5);
                ctx.restore();
            }
        }

        ctx.textAlign = 'left'; ctx.textBaseline = 'top';
    }

    // Chain dim preview
    if (activeTool === 'chaindim' && chainDimState.points.length > 0) {
        const allPts = [...chainDimState.points];
        if (chainDimState.currentEnd) allPts.push(chainDimState.currentEnd);
        if (allPts.length >= 2) {
            ctx.globalAlpha = 0.5;
            ctx.strokeStyle = '#2B7CD0';
            ctx.lineWidth = 1;
            ctx.setLineDash([4, 3]);
            for (let i = 0; i < allPts.length - 1; i++) {
                const p1 = coords.sheetToScreen(allPts[i].x, allPts[i].y);
                const p2 = coords.sheetToScreen(allPts[i+1].x, allPts[i+1].y);
                ctx.beginPath(); ctx.moveTo(p1.x, p1.y); ctx.lineTo(p2.x, p2.y); ctx.stroke();
            }
            ctx.setLineDash([]);
            ctx.globalAlpha = 1.0;
            // Dots
            ctx.fillStyle = '#2B7CD0';
            for (const pt of chainDimState.points) {
                const sp = coords.sheetToScreen(pt.x, pt.y);
                ctx.beginPath(); ctx.arc(sp.x, sp.y, 3, 0, Math.PI * 2); ctx.fill();
            }
        }
    }
}

engine.onRender(drawChainDimensions);

// Chain dim hit-testing
// ── Auto Columns at Grid Intersections ───────────────────
// NOTE: Auto-Cols button removed and replaced with Wall tool. The auto-column functionality
// remains available but the button is hidden. To restore the button, uncomment the following handler:

/*
document.getElementById('btn-auto-cols').addEventListener('click', () => {
    const vGrids = structuralGrids.filter(g => g.axis === 'V');
    const hGrids = structuralGrids.filter(g => g.axis === 'H');

    if (vGrids.length === 0 || hGrids.length === 0) {
        alert('Place at least one vertical and one horizontal grid line first.');
        return;
    }

    const memberSize = memberSizeSelect.value || '89x89x5SHS';
    const memberCat = memberCatSelect.value || 'SHS';
    const isBottomFloor = levelSystem.activeLevelIndex === 0;
    const newCols = [];

    for (const vg of vGrids) {
        for (const hg of hGrids) {
            // Check if a column already exists at this intersection on this level
            const tolerance = 50; // 50mm tolerance
            const exists = project.elements.some(el =>
                el.type === 'column' &&
                el.level === getActiveLevel().id &&
                Math.abs(el.x - vg.position) < tolerance &&
                Math.abs(el.y - hg.position) < tolerance
            );
            if (exists) continue;

            const tag = 'SC' + (_colNum++) + '-' + memberSize;
            newCols.push({
                id: generateId(),
                type: 'column',
                layer: 'S-COLS',
                level: getActiveLevel().id,
                x: vg.position,
                y: hg.position,
                size: parseInt(memberSize) || 89,
                tag: tag,
                extends: isBottomFloor ? 'above' : 'below',
                memberSize: memberSize,
                memberCategory: memberCat,
            });
        }
    }

    if (newCols.length === 0) {
        alert('Columns already exist at all grid intersections on this level.');
        return;
    }

    history.execute({
        description: 'Auto-place ' + newCols.length + ' columns at grid intersections',
        execute() { for (const c of newCols) project.elements.push(c); },
        undo() { for (const c of newCols) { const i = project.elements.indexOf(c); if (i !== -1) project.elements.splice(i, 1); } }
    });

    engine.requestRender();
    console.log('[Auto Cols] Placed ' + newCols.length + ' columns at grid intersections');
});
*/

// ── Beam-to-Column Snap ──────────────────────────────────
// Patch the snap engine to include column centres as snap points

const prevFindSnap2 = findSnap;
findSnap = function(screenX, screenY) {
    const best = prevFindSnap2(screenX, screenY);
    if (!snapState.enabled) return best;

    const sheetPos = engine.coords.screenToSheet(screenX, screenY);
    const radiusMM = snapState.snapRadius / engine.viewport.zoom;
    let bestDist = best ? Math.sqrt(Math.pow(sheetPos.x - best.x, 2) + Math.pow(sheetPos.y - best.y, 2)) : radiusMM;
    let result = best;

    // Snap to column centres on the current level
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'column') continue;
        const cp = engine.coords.realToSheet(el.x, el.y);
        const d = Math.sqrt(Math.pow(sheetPos.x - cp.x, 2) + Math.pow(sheetPos.y - cp.y, 2));
        if (d < bestDist) {
            bestDist = d;
            result = { x: cp.x, y: cp.y, type: SNAP_TYPES.ENDPOINT };
        }
    }

    // Also snap to inherited columns from adjacent levels
    const inherited = getInheritedColumns();
    for (const col of inherited) {
        const cp = engine.coords.realToSheet(col.x, col.y);
        const d = Math.sqrt(Math.pow(sheetPos.x - cp.x, 2) + Math.pow(sheetPos.y - cp.y, 2));
        if (d < bestDist) {
            bestDist = d;
            result = { x: cp.x, y: cp.y, type: SNAP_TYPES.ENDPOINT };
        }
    }

    // Snap to wall endpoints (for joining walls at corners)
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'wall') continue;
        for (const wp of [
            engine.coords.realToSheet(el.x1, el.y1),
            engine.coords.realToSheet(el.x2, el.y2)
        ]) {
            const d = Math.sqrt(Math.pow(sheetPos.x - wp.x, 2) + Math.pow(sheetPos.y - wp.y, 2));
            if (d < bestDist) {
                bestDist = d;
                result = { x: wp.x, y: wp.y, type: SNAP_TYPES.ENDPOINT };
            }
        }
    }

    return result;
};

// ══════════════════════════════════════════════════════════
// ── PHASE A: NOTES PANEL, TABLE, SLAB CALLOUT TOOLS ───────
// ══════════════════════════════════════════════════════════

// State for tool placement
let notesPlacementPending = null;
let tablePlacementPending = null;
let slabCalloutPlacementPending = null;

// ── NOTES PANEL TOOL ──────────────────────────────────────

const notesPanelModal = document.getElementById('notes-panel-modal');
const notesHeadingInput = document.getElementById('notes-heading');
const notesBodyInput = document.getElementById('notes-body');
const notesCancelBtn = document.getElementById('notes-cancel');
const notesApplyBtn = document.getElementById('notes-apply');

document.getElementById('btn-notes-panel').addEventListener('click', () => setActiveTool('notes-panel'));

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'notes-panel') return;

    const sheetPos = engine.getSheetPos(e);
    notesPlacementPending = { x: sheetPos.x, y: sheetPos.y };

    notesHeadingInput.value = '';
    notesBodyInput.value = '';
    notesPanelModal.classList.remove('hidden');
    notesHeadingInput.focus();
});

notesCancelBtn.addEventListener('click', () => {
    notesPanelModal.classList.add('hidden');
    notesPlacementPending = null;
    setActiveTool('select');
});

notesApplyBtn.addEventListener('click', () => {
    if (!notesPlacementPending) return;
    const heading = notesHeadingInput.value.trim() || 'NOTES';
    const body = notesBodyInput.value.trim() || 'Add notes here';

    const newNotesBox = {
        id: generateId(),
        type: 'notesbox',
        layer: 'S-ANNO',
        level: getActiveLevel().id,
        x: notesPlacementPending.x,
        y: notesPlacementPending.y,
        heading: heading,
        body: body,
        fontSize: 3.5,
        width: 60,
        height: 40
    };

    history.execute({
        description: 'Add notes panel',
        execute() { project.elements.push(newNotesBox); },
        undo() {
            const i = project.elements.indexOf(newNotesBox);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    notesPanelModal.classList.add('hidden');
    notesPlacementPending = null;
    engine.requestRender();
    setActiveTool('select');
});

// Render notes panel on canvas
engine.onRender((ctx, eng) => {
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'notesbox') continue;
        const pos = eng.coords.realToSheet(el.x, el.y);
        const sp = eng.coords.sheetToScreen(pos.x, pos.y);

        const zoom = eng.viewport.zoom;
        const w = el.width * zoom;
        const h = el.height * zoom;
        const fontSize = Math.max(7, (el.fontSize || 3.5) * zoom);
        const padding = 4 * zoom;

        // Border
        ctx.strokeStyle = isElementSelected(el) ? '#2B7CD0' : '#333333';
        ctx.lineWidth = 1.5;
        ctx.strokeRect(sp.x, sp.y, w, h);

        // Fill
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(sp.x, sp.y, w, h);

        // Heading
        ctx.fillStyle = '#000000';
        ctx.font = `bold ${fontSize}px "Segoe UI", Arial, sans-serif`;
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';
        ctx.fillText(el.heading || 'NOTES', sp.x + padding, sp.y + padding);

        // Body text (multi-line)
        const bodyLines = (el.body || '').split('\n');
        const lineHeight = fontSize * 1.3;
        ctx.font = `${fontSize * 0.85}px "Segoe UI", Arial, sans-serif`;
        for (let i = 0; i < bodyLines.length && (sp.y + padding + lineHeight * (i + 1.5)) < (sp.y + h - padding); i++) {
            ctx.fillText(bodyLines[i], sp.x + padding, sp.y + padding + fontSize * 1.4 + lineHeight * i);
        }
    }
});

// ── TABLE TOOL ────────────────────────────────────────────

const tableModal = document.getElementById('table-modal');
const tableTemplateSelect = document.getElementById('table-template');
const tableDataInput = document.getElementById('table-data');
const tableCancelBtn = document.getElementById('table-cancel');
const tableApplyBtn = document.getElementById('table-apply');

document.getElementById('btn-table').addEventListener('click', () => setActiveTool('table'));

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'table') return;

    const sheetPos = engine.getSheetPos(e);
    tablePlacementPending = { x: sheetPos.x, y: sheetPos.y };

    tableTemplateSelect.value = 'member';
    tableDataInput.value = 'Mark,Member\nB1,UB 305x127\nB2,UB 254x102';
    tableModal.classList.remove('hidden');
    tableDataInput.focus();
});

tableCancelBtn.addEventListener('click', () => {
    tableModal.classList.add('hidden');
    tablePlacementPending = null;
    setActiveTool('select');
});

tableApplyBtn.addEventListener('click', () => {
    if (!tablePlacementPending) return;
    const csvData = tableDataInput.value.trim();
    const rows = csvData.split('\n').map(r => r.split(',').map(c => c.trim()));

    const newTable = {
        id: generateId(),
        type: 'table',
        layer: 'S-ANNO',
        level: getActiveLevel().id,
        x: tablePlacementPending.x,
        y: tablePlacementPending.y,
        rows: rows,
        template: tableTemplateSelect.value,
        cellWidth: 25,
        cellHeight: 8
    };

    history.execute({
        description: 'Add table',
        execute() { project.elements.push(newTable); },
        undo() {
            const i = project.elements.indexOf(newTable);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    tableModal.classList.add('hidden');
    tablePlacementPending = null;
    engine.requestRender();
    setActiveTool('select');
});

// Render table on canvas
engine.onRender((ctx, eng) => {
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'table') continue;
        const pos = eng.coords.realToSheet(el.x, el.y);
        const sp = eng.coords.sheetToScreen(pos.x, pos.y);

        const zoom = eng.viewport.zoom;
        const cellW = el.cellWidth * zoom;
        const cellH = el.cellHeight * zoom;
        const fontSize = Math.max(6, 2.5 * zoom);
        const rows = el.rows || [];

        // Draw table borders
        ctx.strokeStyle = isElementSelected(el) ? '#2B7CD0' : '#333333';
        ctx.lineWidth = 1;
        ctx.fillStyle = '#FFFFFF';

        for (let row = 0; row < rows.length; row++) {
            const cols = rows[row];
            for (let col = 0; col < cols.length; col++) {
                const x = sp.x + col * cellW;
                const y = sp.y + row * cellH;

                ctx.strokeRect(x, y, cellW, cellH);
                ctx.fillRect(x, y, cellW, cellH);

                // Fill header row with light gray
                if (row === 0) {
                    ctx.fillStyle = '#E8E8E8';
                    ctx.fillRect(x, y, cellW, cellH);
                    ctx.fillStyle = '#FFFFFF';
                }

                // Draw text
                ctx.fillStyle = '#000000';
                ctx.font = `${row === 0 ? 'bold ' : ''}${fontSize}px "Segoe UI", Arial, sans-serif`;
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.fillText(cols[col] || '', x + cellW / 2, y + cellH / 2);
            }
        }
    }
});

// Hit test for tables
// ── SLAB CALLOUT TOOL ─────────────────────────────────────

const slabCalloutModal = document.getElementById('slab-callout-modal');
const slabThicknessInput = document.getElementById('slab-thickness');
const slabCancelBtn = document.getElementById('slab-cancel');
const slabApplyBtn = document.getElementById('slab-apply');

document.getElementById('btn-slab-callout').addEventListener('click', () => setActiveTool('slab-callout'));

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'slab-callout') return;

    const sheetPos = engine.getSheetPos(e);
    slabCalloutPlacementPending = { x: sheetPos.x, y: sheetPos.y };

    slabThicknessInput.value = '';
    slabCalloutModal.classList.remove('hidden');
    slabThicknessInput.focus();
});

slabCancelBtn.addEventListener('click', () => {
    slabCalloutModal.classList.add('hidden');
    slabCalloutPlacementPending = null;
    setActiveTool('select');
});

slabApplyBtn.addEventListener('click', () => {
    if (!slabCalloutPlacementPending) return;
    const thickness = slabThicknessInput.value.trim() || '250';

    const newSlabCallout = {
        id: generateId(),
        type: 'slabcallout',
        layer: 'S-SLAB',
        level: getActiveLevel().id,
        x: slabCalloutPlacementPending.x,
        y: slabCalloutPlacementPending.y,
        thickness: thickness,
        width: 16,
        height: 12
    };

    history.execute({
        description: 'Add slab callout',
        execute() { project.elements.push(newSlabCallout); },
        undo() {
            const i = project.elements.indexOf(newSlabCallout);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    slabCalloutModal.classList.add('hidden');
    slabCalloutPlacementPending = null;
    engine.requestRender();
    setActiveTool('select');
});

// Render slab callout on canvas
engine.onRender((ctx, eng) => {
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'slabcallout') continue;
        const pos = eng.coords.realToSheet(el.x, el.y);
        const sp = eng.coords.sheetToScreen(pos.x, pos.y);

        const zoom = eng.viewport.zoom;
        const w = el.width * zoom;
        const h = el.height * zoom;
        const fontSize = Math.max(8, 3 * zoom);

        // Border box
        ctx.strokeStyle = isElementSelected(el) ? '#2B7CD0' : '#333333';
        ctx.lineWidth = 1.5;
        ctx.strokeRect(sp.x, sp.y, w, h);

        // Fill
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(sp.x, sp.y, w, h);

        // Text
        ctx.fillStyle = '#000000';
        ctx.font = `bold ${fontSize}px "Courier New", monospace`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('[' + el.thickness + ']', sp.x + w / 2, sp.y + h / 2);
    }
});

// Hit test for slab callouts
// ══════════════════════════════════════════════════════════
// PHASE B: SLAB ZONE TOOLS (Edge, Fall, Step)
// ══════════════════════════════════════════════════════════

// ── Edge Thickening Tool ──────────────────────────────────
// Draws a thick line on S-FTNG layer representing slab edge thickening
// Auto-increments ET1, ET2, etc. designation

let _edgeNextNum = 1;
const edgeToolState = {
    placing: false,
    startPoint: null,
    currentEnd: null
};

document.getElementById('btn-edge').addEventListener('click', () => setActiveTool('edge'));

function getEdgeToolPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    if (edgeToolState.placing && edgeToolState.startPoint) {
        pos = applyOrtho(pos.x, pos.y, edgeToolState.startPoint.x, edgeToolState.startPoint.y);
    }
    return pos;
}

container.addEventListener('mousemove', (e) => {
    if (activeTool === 'edge') {
        edgeToolState.currentEnd = getEdgeToolPos(e);
        engine.requestRender();
    }
});

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'edge') return;
    const pos = getEdgeToolPos(e);
    if (!edgeToolState.placing) {
        edgeToolState.placing = true;
        edgeToolState.startPoint = pos;
    } else {
        const start = edgeToolState.startPoint;
        const end = pos;
        const dist = Math.sqrt(Math.pow(end.x - start.x, 2) + Math.pow(end.y - start.y, 2));
        if (dist < 0.5) return;
        const realStart = engine.coords.sheetToReal(start.x, start.y);
        const realEnd = engine.coords.sheetToReal(end.x, end.y);
        const edgeTag = 'ET' + (_edgeNextNum++);
        const newEdge = {
            id: generateId(),
            type: 'edge',
            layer: 'S-FTNG',
            x1: realStart.x, y1: realStart.y,
            x2: realEnd.x, y2: realEnd.y,
            tag: edgeTag,
            lineWeight: 0.7  // 0.7mm for edge beam
        };
        history.execute({
            description: 'Draw edge thickening',
            execute() { project.elements.push(newEdge); },
            undo() {
                const i = project.elements.indexOf(newEdge);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });
        edgeToolState.startPoint = { x: end.x, y: end.y };
        engine.requestRender();
    }
});

container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'edge' && edgeToolState.placing) {
        e.preventDefault();
        edgeToolState.placing = false;
        edgeToolState.startPoint = null;
        engine.requestRender();
    }
});

// ── Fall Arrow Tool ───────────────────────────────────────
// Click start → click end: draws arrow from high to low with FALL label

const fallState = {
    placing: false,
    startPoint: null,
    currentEnd: null
};

document.getElementById('btn-fall').addEventListener('click', () => setActiveTool('fall'));

function getFallPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    if (fallState.placing && fallState.startPoint) {
        pos = applyOrtho(pos.x, pos.y, fallState.startPoint.x, fallState.startPoint.y);
    }
    return pos;
}

container.addEventListener('mousemove', (e) => {
    if (activeTool === 'fall') {
        fallState.currentEnd = getFallPos(e);
        engine.requestRender();
    }
});

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'fall') return;
    const pos = getFallPos(e);
    if (!fallState.placing) {
        fallState.placing = true;
        fallState.startPoint = pos;
    } else {
        const start = fallState.startPoint;
        const end = pos;
        const dist = Math.sqrt(Math.pow(end.x - start.x, 2) + Math.pow(end.y - start.y, 2));
        if (dist < 0.5) return;
        const realStart = engine.coords.sheetToReal(start.x, start.y);
        const realEnd = engine.coords.sheetToReal(end.x, end.y);
        const newFall = {
            id: generateId(),
            type: 'fallarrow',
            layer: 'S-ANNO',
            x1: realStart.x, y1: realStart.y,
            x2: realEnd.x, y2: realEnd.y,
            label: 'FALL'
        };
        history.execute({
            description: 'Draw fall arrow',
            execute() { project.elements.push(newFall); },
            undo() {
                const i = project.elements.indexOf(newFall);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });
        fallState.placing = false;
        fallState.startPoint = null;
        engine.requestRender();
    }
});

container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'fall' && fallState.placing) {
        e.preventDefault();
        fallState.placing = false;
        fallState.startPoint = null;
        engine.requestRender();
    }
});

// ── Step-in-Slab Line Tool ────────────────────────────────
// Draws a thick dashed line representing slab level changes

const stepState = {
    placing: false,
    startPoint: null,
    currentEnd: null
};

document.getElementById('btn-step').addEventListener('click', () => setActiveTool('step'));

function getStepPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    if (stepState.placing && stepState.startPoint) {
        pos = applyOrtho(pos.x, pos.y, stepState.startPoint.x, stepState.startPoint.y);
    }
    return pos;
}

container.addEventListener('mousemove', (e) => {
    if (activeTool === 'step') {
        stepState.currentEnd = getStepPos(e);
        engine.requestRender();
    }
});

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'step') return;
    const pos = getStepPos(e);
    if (!stepState.placing) {
        stepState.placing = true;
        stepState.startPoint = pos;
    } else {
        const start = stepState.startPoint;
        const end = pos;
        const dist = Math.sqrt(Math.pow(end.x - start.x, 2) + Math.pow(end.y - start.y, 2));
        if (dist < 0.5) return;
        const realStart = engine.coords.sheetToReal(start.x, start.y);
        const realEnd = engine.coords.sheetToReal(end.x, end.y);
        const newStep = {
            id: generateId(),
            type: 'step',
            layer: 'S-ANNO',
            x1: realStart.x, y1: realStart.y,
            x2: realEnd.x, y2: realEnd.y
        };
        history.execute({
            description: 'Draw step line',
            execute() { project.elements.push(newStep); },
            undo() {
                const i = project.elements.indexOf(newStep);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });
        stepState.startPoint = { x: end.x, y: end.y };
        engine.requestRender();
    }
});

container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'step' && stepState.placing) {
        e.preventDefault();
        stepState.placing = false;
        stepState.startPoint = null;
        engine.requestRender();
    }
});

// Consolidated setActiveTool to handle all Phase B tools
// ── Rendering for Edge, Fall, and Step tools ──────────────

// Patch the main render to draw previews and elements
const origRender = engine.render;
engine.render = function(eng) {
    origRender.call(this, eng);
    const ctx = eng.canvas.getContext('2d');
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    // Edge preview
    if (activeTool === 'edge' && edgeToolState.placing && edgeToolState.startPoint && edgeToolState.currentEnd) {
        const start = coords.sheetToScreen(edgeToolState.startPoint.x, edgeToolState.startPoint.y);
        const end = coords.sheetToScreen(edgeToolState.currentEnd.x, edgeToolState.currentEnd.y);
        ctx.strokeStyle = '#FF6600';
        ctx.lineWidth = Math.max(1, 0.7 * zoom);
        ctx.setLineDash([]);
        ctx.globalAlpha = 0.6;
        ctx.beginPath();
        ctx.moveTo(start.x, start.y);
        ctx.lineTo(end.x, end.y);
        ctx.stroke();
        ctx.globalAlpha = 1;
    }

    // Fall arrow preview
    if (activeTool === 'fall' && fallState.placing && fallState.startPoint && fallState.currentEnd) {
        const start = coords.sheetToScreen(fallState.startPoint.x, fallState.startPoint.y);
        const end = coords.sheetToScreen(fallState.currentEnd.x, fallState.currentEnd.y);
        ctx.strokeStyle = '#2B7CD0';
        ctx.lineWidth = Math.max(1, 0.25 * zoom);
        ctx.setLineDash([]);
        ctx.globalAlpha = 0.6;
        ctx.beginPath();
        ctx.moveTo(start.x, start.y);
        ctx.lineTo(end.x, end.y);
        ctx.stroke();
        // Draw arrowhead at end
        const angle = Math.atan2(end.y - start.y, end.x - start.x);
        const arrowLen = Math.max(6, 2.5 * zoom);
        const arrowW = Math.PI / 7;
        ctx.fillStyle = '#2B7CD0';
        ctx.beginPath();
        ctx.moveTo(end.x, end.y);
        ctx.lineTo(end.x - arrowLen * Math.cos(angle - arrowW), end.y - arrowLen * Math.sin(angle - arrowW));
        ctx.lineTo(end.x - arrowLen * Math.cos(angle + arrowW), end.y - arrowLen * Math.sin(angle + arrowW));
        ctx.closePath();
        ctx.fill();
        ctx.globalAlpha = 1;
    }

    // Step preview
    if (activeTool === 'step' && stepState.placing && stepState.startPoint && stepState.currentEnd) {
        const start = coords.sheetToScreen(stepState.startPoint.x, stepState.startPoint.y);
        const end = coords.sheetToScreen(stepState.currentEnd.x, stepState.currentEnd.y);
        ctx.strokeStyle = '#8844AA';
        ctx.lineWidth = Math.max(1, 0.5 * zoom);
        ctx.setLineDash([4 * zoom, 3 * zoom]);
        ctx.globalAlpha = 0.6;
        ctx.beginPath();
        ctx.moveTo(start.x, start.y);
        ctx.lineTo(end.x, end.y);
        ctx.stroke();
        ctx.globalAlpha = 1;
    }

    // Render edge elements
    for (const el of project.getVisibleElements()) {
        if (el.type === 'edge') {
            const layer = project.layers[el.layer];
            if (!layer || !layer.visible) continue;
            const p1 = coords.realToScreen(el.x1, el.y1);
            const p2 = coords.realToScreen(el.x2, el.y2);
            const color = layer.color;
            ctx.strokeStyle = color;
            ctx.lineWidth = Math.max(1, (el.lineWeight || 0.7) * zoom);
            ctx.setLineDash([]);
            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.stroke();
            // Label near midpoint
            if (el.tag) {
                const mid = { x: (p1.x + p2.x) / 2, y: (p1.y + p2.y) / 2 };
                const fontSize = Math.max(8, 3 * zoom);
                ctx.font = `${fontSize}px Arial, sans-serif`;
                ctx.fillStyle = color;
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                ctx.fillText(el.tag, mid.x, mid.y - 4);
            }
        }
    }

    // Render fall arrow elements
    for (const el of project.getVisibleElements()) {
        if (el.type === 'fallarrow') {
            const layer = project.layers[el.layer];
            if (!layer || !layer.visible) continue;
            const p1 = coords.realToScreen(el.x1, el.y1);
            const p2 = coords.realToScreen(el.x2, el.y2);
            const color = layer.color;
            ctx.strokeStyle = color;
            ctx.lineWidth = Math.max(1, 0.25 * zoom);
            ctx.setLineDash([]);
            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.stroke();
            // Arrowhead at p2 (end)
            const angle = Math.atan2(p2.y - p1.y, p2.x - p1.x);
            const arrowLen = Math.max(6, 2.5 * zoom);
            const arrowW = Math.PI / 7;
            ctx.fillStyle = color;
            ctx.beginPath();
            ctx.moveTo(p2.x, p2.y);
            ctx.lineTo(p2.x - arrowLen * Math.cos(angle - arrowW), p2.y - arrowLen * Math.sin(angle - arrowW));
            ctx.lineTo(p2.x - arrowLen * Math.cos(angle + arrowW), p2.y - arrowLen * Math.sin(angle + arrowW));
            ctx.closePath();
            ctx.fill();
            // Label "FALL" at midpoint
            if (el.label) {
                const mid = { x: (p1.x + p2.x) / 2, y: (p1.y + p2.y) / 2 };
                const fontSize = Math.max(7, 2.5 * zoom);
                ctx.font = `bold ${fontSize}px Arial, sans-serif`;
                ctx.fillStyle = color;
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                ctx.fillText(el.label, mid.x, mid.y - 4);
            }
        }
    }

    // Render step elements
    for (const el of project.getVisibleElements()) {
        if (el.type === 'step') {
            const layer = project.layers[el.layer];
            if (!layer || !layer.visible) continue;
            const p1 = coords.realToScreen(el.x1, el.y1);
            const p2 = coords.realToScreen(el.x2, el.y2);
            const color = layer.color;
            ctx.strokeStyle = color;
            ctx.lineWidth = Math.max(1, 0.5 * zoom);
            ctx.setLineDash([4 * zoom, 3 * zoom]);
            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.stroke();
            ctx.setLineDash([]);
        }
    }

    // Selection rendering for Phase B tools
    for (const el of selectedElements) {
        if (el === selectedElement) continue;
        if (el.type === 'edge' || el.type === 'fallarrow' || el.type === 'step') {
            ctx.strokeStyle = '#2B7CD0';
            ctx.lineWidth = 2;
            ctx.setLineDash([4, 3]);
            const p1 = coords.realToScreen(el.x1, el.y1);
            const p2 = coords.realToScreen(el.x2, el.y2);
            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.stroke();
            ctx.setLineDash([]);
        }
    }
};

// ── Hit testing for Phase B tools ──────────────────────────

// ── PHASE C: RL LEVEL MARKER TOOL ──────────────────────────

document.getElementById('btn-rl-marker').addEventListener('click', () => setActiveTool('rlmarker'));
document.getElementById('btn-sec-ref').addEventListener('click', () => setActiveTool('secref'));
document.getElementById('btn-ref-tag').addEventListener('click', () => setActiveTool('reftag'));

// RLs panel — prominent button opens Levels & RL editor
if (document.getElementById('btn-levels-editor')) {
    document.getElementById('btn-levels-editor').addEventListener('click', () => openLevelsRLEditor());
}

// Click handler for RL Marker placement
container.addEventListener('click', (e) => {
    if (activeTool !== 'rlmarker') return;
    if (engine._spaceDown || engine._isPanning) return;

    const sheetPos = engine.getSheetPos(e);
    const modal = document.getElementById('rl-marker-modal');
    modal.classList.remove('hidden');

    const btnConfirm = modal.querySelector('.btn-confirm');
    const btnCancel = modal.querySelector('.btn-cancel');
    const levelIdInput = document.getElementById('rl-level-id');
    const rlValueInput = document.getElementById('rl-value');

    function closeModal() {
        modal.classList.add('hidden');
        levelIdInput.value = '';
        rlValueInput.value = '';
        btnConfirm.removeEventListener('click', handleConfirm);
        btnCancel.removeEventListener('click', handleCancel);
        document.removeEventListener('keydown', handleKeydown);
    }

    function handleConfirm() {
        const levelId = levelIdInput.value.trim();
        const rlValue = rlValueInput.value.trim();
        if (!levelId || !rlValue) {
            alert('Please fill in all fields');
            return;
        }

        const realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);
        const newMarker = {
            id: generateId(),
            type: 'rlmarker',
            layer: 'S-ANNO',
            x: realPos.x,
            y: realPos.y,
            levelId: levelId,
            rlValue: rlValue
        };

        history.execute({
            description: 'Add RL marker: ' + levelId,
            execute() { project.elements.push(newMarker); },
            undo() {
                const i = project.elements.indexOf(newMarker);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });

        closeModal();
        engine.requestRender();
    }

    function handleCancel() {
        closeModal();
    }

    function handleKeydown(e) {
        if (e.key === 'Escape') {
            e.preventDefault();
            handleCancel();
        }
    }

    btnConfirm.addEventListener('click', handleConfirm);
    btnCancel.addEventListener('click', handleCancel);
    document.addEventListener('keydown', handleKeydown);

    levelIdInput.focus();
});

// ── PHASE C: SECTION REFERENCE MARKER TOOL ─────────────────

container.addEventListener('click', (e) => {
    if (activeTool !== 'secref') return;
    if (engine._spaceDown || engine._isPanning) return;

    const sheetPos = engine.getSheetPos(e);
    const modal = document.getElementById('sec-ref-modal');
    modal.classList.remove('hidden');

    const btnConfirm = modal.querySelector('.btn-confirm');
    const btnCancel = modal.querySelector('.btn-cancel');
    const numberInput = document.getElementById('sec-number');
    const refInput = document.getElementById('sec-ref-drawing');
    const scaleInput = document.getElementById('sec-scale');

    function closeModal() {
        modal.classList.add('hidden');
        numberInput.value = '';
        refInput.value = '';
        scaleInput.value = '';
        btnConfirm.removeEventListener('click', handleConfirm);
        btnCancel.removeEventListener('click', handleCancel);
        document.removeEventListener('keydown', handleKeydown);
    }

    function handleConfirm() {
        const number = numberInput.value.trim();
        const reference = refInput.value.trim();
        const scale = scaleInput.value.trim();
        if (!number || !reference || !scale) {
            alert('Please fill in all fields');
            return;
        }

        const realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);
        const newMarker = {
            id: generateId(),
            type: 'secref',
            layer: 'S-ANNO',
            x: realPos.x,
            y: realPos.y,
            sectionNumber: number,
            refDrawing: reference,
            scale: scale
        };

        history.execute({
            description: 'Add section reference: ' + number,
            execute() { project.elements.push(newMarker); },
            undo() {
                const i = project.elements.indexOf(newMarker);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });

        closeModal();
        engine.requestRender();
    }

    function handleCancel() {
        closeModal();
    }

    function handleKeydown(e) {
        if (e.key === 'Escape') {
            e.preventDefault();
            handleCancel();
        }
    }

    btnConfirm.addEventListener('click', handleConfirm);
    btnCancel.addEventListener('click', handleCancel);
    document.addEventListener('keydown', handleKeydown);

    numberInput.focus();
});

// ── PHASE C: REFERENCE TAG TOOL ────────────────────────────

container.addEventListener('click', (e) => {
    if (activeTool !== 'reftag') return;
    if (engine._spaceDown || engine._isPanning) return;

    const sheetPos = engine.getSheetPos(e);
    const modal = document.getElementById('ref-tag-modal');
    modal.classList.remove('hidden');

    const btnConfirm = modal.querySelector('.btn-confirm');
    const btnCancel = modal.querySelector('.btn-cancel');
    const typeSelect = document.getElementById('ref-tag-type');
    const refInput = document.getElementById('ref-tag-ref');

    function closeModal() {
        modal.classList.add('hidden');
        typeSelect.value = 'REFER PLAN';
        refInput.value = '';
        btnConfirm.removeEventListener('click', handleConfirm);
        btnCancel.removeEventListener('click', handleCancel);
        document.removeEventListener('keydown', handleKeydown);
    }

    function handleConfirm() {
        const type = typeSelect.value;
        const ref = refInput.value.trim();

        let text = type;
        if (type === 'REFER DRG' && ref) {
            text = 'REFER DRG ' + ref;
        } else if (type === 'REFER DETAIL' && ref) {
            text = 'REFER DETAIL ' + ref;
        }

        const realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);
        const newMarker = {
            id: generateId(),
            type: 'text',
            layer: 'S-ANNO',
            x: realPos.x,
            y: realPos.y,
            text: text,
            fontSize: 3.5
        };

        history.execute({
            description: 'Add reference tag: ' + text,
            execute() { project.elements.push(newMarker); },
            undo() {
                const i = project.elements.indexOf(newMarker);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });

        closeModal();
        engine.requestRender();
    }

    function handleCancel() {
        closeModal();
    }

    function handleKeydown(e) {
        if (e.key === 'Escape') {
            e.preventDefault();
            handleCancel();
        }
    }

    btnConfirm.addEventListener('click', handleConfirm);
    btnCancel.addEventListener('click', handleCancel);
    document.addEventListener('keydown', handleKeydown);

    refInput.focus();
});

// ── PHASE C: RENDERING FOR MARKERS ────────────────────────

const cbIdx7_find = engine._renderCallbacks.indexOf(drawElements);
const prevPhase7Draw = cbIdx7_find !== -1 ? engine._renderCallbacks[cbIdx7_find] : drawElements;
const wrappedDrawElements7 = function(ctx, eng) {
    prevPhase7Draw(ctx, eng);

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    // Render RL Markers
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'rlmarker') continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const cp = coords.realToScreen(el.x, el.y);
        const circleRadius = Math.max(12, 6 * zoom);
        const isSelected = (selectedElement === el);

        // Circle background
        ctx.fillStyle = isSelected ? 'rgba(43,124,208,0.3)' : 'rgba(255,255,255,0.9)';
        ctx.beginPath();
        ctx.arc(cp.x, cp.y, circleRadius, 0, Math.PI * 2);
        ctx.fill();

        // Circle border
        ctx.strokeStyle = isSelected ? '#2B7CD0' : layer.color;
        ctx.lineWidth = Math.max(1, 1.5 * zoom);
        ctx.stroke();

        // Level ID text inside circle
        const fontSize = Math.max(8, 4 * zoom);
        ctx.font = `bold ${fontSize}px "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = isSelected ? '#2B7CD0' : '#000000';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(el.levelId, cp.x, cp.y);

        // RL value below circle
        ctx.font = `${Math.max(7, 3 * zoom)}px "Segoe UI", Arial, sans-serif`;
        ctx.textBaseline = 'top';
        ctx.fillText(el.rlValue, cp.x, cp.y + circleRadius + 2);
    }

    // Render Section References
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'secref') continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const bp = coords.realToScreen(el.x, el.y);
        const isSelected = (selectedElement === el);
        const boxWidth = Math.max(30, 8 * zoom);
        const boxHeight = Math.max(24, 6 * zoom);

        // Box background
        ctx.fillStyle = isSelected ? 'rgba(43,124,208,0.2)' : 'rgba(255,255,255,0.95)';
        ctx.fillRect(bp.x - boxWidth / 2, bp.y - boxHeight / 2, boxWidth, boxHeight);

        // Box border
        ctx.strokeStyle = isSelected ? '#2B7CD0' : layer.color;
        ctx.lineWidth = Math.max(1, 1 * zoom);
        ctx.strokeRect(bp.x - boxWidth / 2, bp.y - boxHeight / 2, boxWidth, boxHeight);

        // Horizontal divider
        ctx.beginPath();
        ctx.moveTo(bp.x - boxWidth / 2, bp.y);
        ctx.lineTo(bp.x + boxWidth / 2, bp.y);
        ctx.stroke();

        // Section number (top half, bold)
        const numFont = Math.max(8, 3.5 * zoom);
        ctx.font = `bold ${numFont}px "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = isSelected ? '#2B7CD0' : '#000000';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(el.sectionNumber, bp.x, bp.y - boxHeight / 4);

        // Reference drawing (bottom half)
        ctx.font = `${Math.max(7, 3 * zoom)}px "Segoe UI", Arial, sans-serif`;
        ctx.fillText(el.refDrawing, bp.x, bp.y + boxHeight / 4);

        // Scale below box
        ctx.font = `${Math.max(7, 2.5 * zoom)}px "Segoe UI", Arial, sans-serif`;
        ctx.textBaseline = 'top';
        ctx.fillText(el.scale, bp.x, bp.y + boxHeight / 2 + 2);
    }
};

// Replace the drawElements in render callbacks for Phase C
if (cbIdx7_find !== -1) engine._renderCallbacks[cbIdx7_find] = wrappedDrawElements7;

// ── Hit testing for Phase C tools ──────────────────────────

// ── PHASE D: FOOTING TOOL ──────────────────────────────────

let footingPlacementPending = null;
const footingModal = document.getElementById('footing-modal');
const footingMarkInput = document.getElementById('footing-mark');
const footingWidthInput = document.getElementById('footing-width');
const footingDepthInput = document.getElementById('footing-depth');
const footingReinfInput = document.getElementById('footing-reinforcement');

let _ftgNextNum = 1;

// Schedule type references for placement
let placementTypeRef = {
    footing: 'PF1',
    stripfooting: 'SF1',
    beam: 'SB1',
    column: 'SC1',
    wall: 'BW1',
    bracingWall: 'BR1'
};

document.getElementById('btn-footing').addEventListener('click', () => setActiveTool('footing'));

// Single-click to place pad footing instantly (no modal)
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'footing') return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    const sheetPos = snap ? { x: snap.x, y: snap.y } : engine.getSheetPos(e);

    const realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);
    const typeRef = placementTypeRef.footing || 'PF1';
    const typeData = project.scheduleTypes.padfooting[typeRef] || { width: 600, length: 600, depth: 300, reo: '', setdown: 200 };

    const newFooting = {
        id: generateId(),
        type: 'footing',
        layer: 'S-FTNG',
        level: getActiveLevel().id,
        x: realPos.x,
        y: realPos.y,
        mark: typeRef,
        typeRef: typeRef,
        footingWidth: typeData.width || 1000,  // mm — defaults from schedule
        footingDepth: typeData.depth || 300,   // mm — defaults from schedule
        reinforcement: typeData.reo || ''
    };

    history.execute({
        description: 'Add pad footing: ' + typeRef,
        execute() { project.elements.push(newFooting); },
        undo() {
            const i = project.elements.indexOf(newFooting);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    engine.requestRender();
    // Stay in footing tool for rapid placement
});

// Double-click pad footing to show TYPE REASSIGNMENT PICKER
container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'select') return;
    if (!selectedElement || selectedElement.type !== 'footing') return;

    const el = selectedElement;
    const rect = container.getBoundingClientRect();
    const screenX = e.clientX - rect.left;
    const screenY = e.clientY - rect.top;
    showTypeReassignmentPicker(el, 'padfooting', screenX, screenY);
});

// Footing modal — Cancel
footingModal.querySelector('.btn-cancel').addEventListener('click', () => {
    footingModal.classList.add('hidden');
    footingPlacementPending = null;
});

// Footing modal — Save (edit existing)
footingModal.querySelector('.btn-confirm').addEventListener('click', () => {
    if (!footingPlacementPending) return;

    const el = footingPlacementPending;
    el.mark = footingMarkInput.value.trim() || el.mark;
    el.footingWidth = parseFloat(footingWidthInput.value) || 1000;
    el.footingDepth = parseFloat(footingDepthInput.value) || 400;
    // Sync legacy fields
    el.width = el.footingWidth;
    el.depth = el.footingDepth;
    el.reinforcement = footingReinfInput.value.trim() || '';
    el.depthBelowFSL = parseFloat(document.getElementById('footing-setdown').value) || 200;

    footingModal.classList.add('hidden');
    footingPlacementPending = null;
    engine.requestRender();
});

// ── TYPE REASSIGNMENT PICKER ─────────────────────────────
// Shows a compact popup near the element with all schedule types
let typePickerEl = null;
function showTypeReassignmentPicker(el, category, screenX, screenY) {
    hideTypeReassignmentPicker();
    const types = project.scheduleTypes[category] || {};
    if (Object.keys(types).length === 0) return;

    const popup = document.createElement('div');
    popup.id = 'type-reassign-picker';
    popup.style.cssText = `
        position:absolute; left:${screenX + 8}px; top:${screenY - 10}px; z-index:50;
        background:#fff; border:1px solid #c0c4cc; border-radius:6px;
        box-shadow:0 4px 16px rgba(0,0,0,0.18); padding:4px; min-width:180px;
        font-family:var(--font-ui); font-size:11px; max-height:300px; overflow-y:auto;
    `;
    const currentRef = el.typeRef || el.mark || el.tag || '';
    for (const [typeRef, typeData] of Object.entries(types)) {
        const isActive = typeRef === currentRef;
        const dims = getTypeDimsSummary(category, typeData);
        const btn = document.createElement('button');
        btn.style.cssText = `
            display:flex; align-items:center; gap:6px; width:100%; padding:5px 8px;
            border:none; background:${isActive ? 'rgba(43,124,208,0.08)' : 'transparent'};
            cursor:pointer; border-radius:4px; font-family:var(--font-ui); font-size:11px;
            text-align:left; color:${isActive ? '#1a5fa0' : '#374151'}; font-weight:${isActive ? '700' : '500'};
        `;
        btn.innerHTML = `
            <div style="width:10px;height:10px;border-radius:2px;background:${typeData.color};flex-shrink:0;"></div>
            <span style="font-weight:700;min-width:28px;">${typeRef}</span>
            <span style="color:#6b7280;font-size:10px;">${dims}</span>
        `;
        btn.addEventListener('mouseenter', () => { if (!isActive) btn.style.background = '#f3f4f6'; });
        btn.addEventListener('mouseleave', () => { if (!isActive) btn.style.background = 'transparent'; });
        btn.addEventListener('click', () => {
            reassignElementType(el, category, typeRef);
            hideTypeReassignmentPicker();
        });
        popup.appendChild(btn);
    }
    container.appendChild(popup);
    typePickerEl = popup;

    // Close on click outside
    setTimeout(() => {
        const closer = (ev) => {
            if (!popup.contains(ev.target)) {
                hideTypeReassignmentPicker();
                document.removeEventListener('mousedown', closer);
            }
        };
        document.addEventListener('mousedown', closer);
    }, 50);
}

function hideTypeReassignmentPicker() {
    if (typePickerEl && typePickerEl.parentNode) {
        typePickerEl.parentNode.removeChild(typePickerEl);
    }
    typePickerEl = null;
}

function getTypeDimsSummary(category, typeData) {
    if (category === 'padfooting') {
        const w = typeData.width || 600;
        const l = typeData.rect ? (typeData.length || w) : w;
        const sizeStr = typeData.rect ? `${w}×${l}` : `${w}SQ`;
        return `${sizeStr} × ${typeData.depth || 300}D`;
    }
    if (category === 'stripfooting') return `${typeData.width || 300}W × ${typeData.depth || 500}D`;
    if (category === 'beam') return typeData.size || '(unassigned)';
    if (category === 'column') return typeData.size || '(unassigned)';
    if (category === 'wall') return typeData.wallType || '(unassigned)';
    return '';
}

function reassignElementType(el, category, newTypeRef) {
    const typeData = project.scheduleTypes[category][newTypeRef];
    if (!typeData) return;

    el.typeRef = newTypeRef;

    // Update the mark/tag to match the type ref
    if (category === 'padfooting') {
        el.mark = newTypeRef;
        el.footingWidth = typeData.width || el.footingWidth;
        el.footingDepth = typeData.depth || el.footingDepth;
        el.reinforcement = typeData.reo || el.reinforcement;
        el.depthBelowFSL = typeData.setdown || el.depthBelowFSL;
    } else if (category === 'stripfooting') {
        el.tag = newTypeRef;
        el.footingWidth = typeData.width || el.footingWidth;
        el.footingDepth = typeData.depth || el.footingDepth;
        el.reinforcement = typeData.reo || el.reinforcement;
    } else if (category === 'beam') {
        el.tag = newTypeRef;
        el.memberCategory = typeData.sectionType || el.memberCategory;
        el.memberSize = typeData.size || el.memberSize;
    } else if (category === 'column') {
        el.tag = newTypeRef;
        el.memberCategory = typeData.sectionType || el.memberCategory;
        el.memberSize = typeData.size || el.memberSize;
    } else if (category === 'wall') {
        el.tag = newTypeRef;
        el.wallType = typeData.wallType || el.wallType;
        el.thickness = typeData.thickness || el.thickness;
    } else if (category === 'bracingWall') {
        el.tag = newTypeRef;
    }

    engine.requestRender();
    console.log(`[Schedule] Reassigned ${el.type} to ${newTypeRef}`);
}

// ── NUMBER KEY SHORTCUT: Quick type selection while tool active or element selected ──
window.addEventListener('keydown', (e) => {
    const tag = (document.activeElement || {}).tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
    const num = parseInt(e.key);
    if (isNaN(num) || num < 1 || num > 9) return;

    // If a structural element is selected, reassign its type
    if (activeTool === 'select' && selectedElement) {
        const el = selectedElement;
        let category = null;
        let prefix = '';
        if (el.type === 'footing') { category = 'padfooting'; prefix = 'PF'; }
        else if (el.type === 'stripFooting') { category = 'stripfooting'; prefix = 'SF'; }
        else if (el.type === 'line' && el.layer === 'S-BEAM') { category = 'beam'; prefix = 'SB'; }
        else if (el.type === 'column') { category = 'column'; prefix = 'SC'; }
        else if (el.type === 'wall') { category = 'wall'; prefix = 'BW'; }
        else if (el.type === 'bracingWall') { category = 'bracingWall'; prefix = 'BR'; }

        if (category) {
            const newRef = prefix + num;
            if (project.scheduleTypes[category][newRef]) {
                reassignElementType(el, category, newRef);
            }
            return;
        }
    }

    // If a placement tool is active, switch the type being placed
    if (activeTool === 'footing') {
        const ref = 'PF' + num;
        if (project.scheduleTypes.padfooting[ref]) {
            placementTypeRef.footing = ref;
            console.log(`[Schedule] Placing: ${ref}`);
        }
    } else if (activeTool === 'stripFooting') {
        const ref = 'SF' + num;
        if (project.scheduleTypes.stripfooting[ref]) {
            placementTypeRef.stripfooting = ref;
            console.log(`[Schedule] Placing: ${ref}`);
        }
    } else if (activeTool === 'line' || activeTool === 'beam') {
        const ref = 'SB' + num;
        if (project.scheduleTypes.beam[ref]) {
            placementTypeRef.beam = ref;
            console.log(`[Schedule] Placing: ${ref}`);
        }
    } else if (activeTool === 'column') {
        const ref = 'SC' + num;
        if (project.scheduleTypes.column[ref]) {
            placementTypeRef.column = ref;
            console.log(`[Schedule] Placing: ${ref}`);
        }
    } else if (activeTool === 'wall') {
        const ref = 'BW' + num;
        if (project.scheduleTypes.wall[ref]) {
            placementTypeRef.wall = ref;
            console.log(`[Schedule] Placing: ${ref}`);
        }
    } else if (activeTool === 'bracingWall') {
        const ref = 'BR' + num;
        if (project.scheduleTypes.bracingWall && project.scheduleTypes.bracingWall[ref]) {
            placementTypeRef.bracingWall = ref;
            // Update dropdown
            const sel = document.getElementById('bracing-type-select');
            if (sel) sel.value = ref;
            console.log(`[Schedule] Placing: ${ref}`);
        }
    }
    // Update status bar and re-render to show new type preview
    updateStatusBar();
    engine.requestRender();
});

// ── STATUS BAR: Show active placement type ──
// Patch the updateStatusBar to show which type is being placed
const _origUpdateStatusForSchedule = updateStatusBar;
updateStatusBar = function(e) {
    _origUpdateStatusForSchedule(e);
    const toolNameSpan = document.getElementById('tool-name');
    if (!toolNameSpan) return;
    const toolText = toolNameSpan.textContent;
    let typeLabel = '';
    if (activeTool === 'footing') {
        const ref = placementTypeRef.footing || 'PF1';
        const td = project.scheduleTypes.padfooting[ref];
        if (td && td.width) typeLabel = ` → ${ref} (${td.width}×${td.length || td.width}×${td.depth}D)`;
        else typeLabel = ` → ${ref}`;
    } else if (activeTool === 'stripFooting') {
        const ref = placementTypeRef.stripfooting || 'SF1';
        const td = project.scheduleTypes.stripfooting[ref];
        if (td && td.width) typeLabel = ` → ${ref} (${td.width}W×${td.depth}D)`;
        else typeLabel = ` → ${ref}`;
    } else if (activeTool === 'line') {
        const ref = placementTypeRef.beam || 'SB1';
        const td = project.scheduleTypes.beam[ref];
        if (td && td.size) typeLabel = ` → ${ref} (${td.size})`;
        else typeLabel = ` → ${ref}`;
    } else if (activeTool === 'column') {
        const ref = placementTypeRef.column || 'SC1';
        const td = project.scheduleTypes.column[ref];
        if (td && td.size) typeLabel = ` → ${ref} (${td.size})`;
        else typeLabel = ` → ${ref}`;
    } else if (activeTool === 'wall') {
        const ref = placementTypeRef.wall || 'BW1';
        const td = project.scheduleTypes.wall[ref];
        if (td && td.wallType) typeLabel = ` → ${ref} (${td.wallType})`;
        else typeLabel = ` → ${ref}`;
    }
    if (typeLabel) toolNameSpan.textContent = toolText + typeLabel;
};

// ── DOUBLE-CLICK REASSIGNMENT for other element types ──
container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'select') return;
    if (!selectedElement) return;
    const el = selectedElement;
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;

    if (el.type === 'stripFooting') {
        showTypeReassignmentPicker(el, 'stripfooting', sx, sy);
    } else if (el.type === 'line' && el.layer === 'S-BEAM') {
        showTypeReassignmentPicker(el, 'beam', sx, sy);
    } else if (el.type === 'column') {
        showTypeReassignmentPicker(el, 'column', sx, sy);
    } else if (el.type === 'wall') {
        showTypeReassignmentPicker(el, 'wall', sx, sy);
    } else if (el.type === 'bracingWall') {
        showTypeReassignmentPicker(el, 'bracingWall', sx, sy);
    }
});

// Render footing elements
engine.onRender((ctx, eng) => {
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'footing') continue;

        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const pos = eng.coords.realToScreen(el.x, el.y);
        const zoom = eng.viewport.zoom;

        // Look up dimensions from schedule type (schedule is source of truth)
        const typeRef = el.typeRef || el.mark || 'PF1';
        const typeData = project.scheduleTypes.padfooting[typeRef] || {};
        const fW = typeData.width || el.footingWidth || 1000;
        const halfWidth = (fW / 2 / CONFIG.drawingScale) * zoom;
        const halfDepth = halfWidth; // square pad in plan

        // Draw thin dashed rectangle with colour coding
        const typeColor = typeData.color;
        const color = isElementSelected(el) ? '#2B7CD0' : layer.color;
        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(0.5, 0.25 * zoom);
        ctx.setLineDash([1.5 * zoom, 1 * zoom]);
        ctx.strokeRect(pos.x - halfWidth, pos.y - halfDepth, halfWidth * 2, halfDepth * 2);
        ctx.setLineDash([]);

        // Fill with type colour at 15% opacity
        if (typeColor && !isElementSelected(el)) {
            ctx.fillStyle = typeColor + '14'; // ~8% opacity — subtle fill
            ctx.fillRect(pos.x - halfWidth, pos.y - halfDepth, halfWidth * 2, halfDepth * 2);
        }

        // Draw footing mark as text near the footing
        if (el.mark) {
            const fontSize = Math.max(1, 2.5 * zoom);
            // Apply custom tag offset if user has dragged it
            const tagDx = (el._tagOffsetX || 0) / CONFIG.drawingScale * zoom;
            const tagDy = (el._tagOffsetY || 0) / CONFIG.drawingScale * zoom;
            ctx.font = `${fontSize}px Arial, sans-serif`;
            ctx.fillStyle = color;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'top';
            ctx.fillText(el.mark, pos.x + tagDx, pos.y + halfDepth + 2 + tagDy);
        }
    }
});

// Hit testing for footing elements
// Drag support for footings
const dragStateOrig = dragState;
const origMouseMoveHandler = window.onmousemove;

// ══════════════════════════════════════════════════════════
// ── STRIP FOOTING TOOL ──────────────────────────────────
// ══════════════════════════════════════════════════════════

const stripFtgState = {
    placing: false,
    startPoint: null,   // { x, y } in sheet-mm
    currentEnd: null,   // { x, y } in sheet-mm (preview)
};

// All strip footings default to SF1 — user changes to SF2, SF3 etc. only when specs differ

// Ground Floor constraint — only enable on GF
document.getElementById('btn-strip-footing').addEventListener('click', () => {
    const activeLevel = getActiveLevel();
    if (activeLevel.id !== 'GF') {
        alert('Strip footings can only be placed on the Ground Floor.');
        return;
    }
    setActiveTool('stripFooting');
});

// ── Strip Footing: Node-to-Node Drawing ──
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'stripFooting') return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    // Snap to nodes
    const snap = findSnap(sx, sy);
    const sheetPos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);

    if (!stripFtgState.placing) {
        // First click: start
        stripFtgState.placing = true;
        stripFtgState.startPoint = { ...sheetPos };
        stripFtgState.currentEnd = { ...sheetPos };
        engine.requestRender();
    } else {
        // Second click: commit the strip footing
        const realStart = engine.coords.sheetToReal(stripFtgState.startPoint.x, stripFtgState.startPoint.y);
        const realEnd = engine.coords.sheetToReal(stripFtgState.currentEnd.x, stripFtgState.currentEnd.y);

        const typeRef = placementTypeRef.stripfooting || 'SF1';
        const typeData = project.scheduleTypes.stripfooting[typeRef] || { width: 300, depth: 500, reo: '', setdown: 200 };

        const newStripFtg = {
            id: generateId(),
            type: 'stripFooting',
            layer: 'S-FTNG',
            level: 'GF',  // Always Ground Floor
            x1: realStart.x, y1: realStart.y,
            x2: realEnd.x, y2: realEnd.y,
            tag: typeRef,
            typeRef: typeRef,
            footingWidth: typeData.width || 300,  // mm — from schedule
            footingDepth: typeData.depth || 500,  // mm — from schedule
            reinforcement: typeData.reo || '',
            wallRef: null,      // null = drawn manually, string = linked wall ID
        };

        history.execute({
            description: 'Draw strip footing: ' + typeRef,
            execute() { project.elements.push(newStripFtg); },
            undo() {
                const i = project.elements.indexOf(newStripFtg);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });

        // Continue drawing (stay in placing mode for chaining)
        stripFtgState.startPoint = { ...stripFtgState.currentEnd };
        engine.requestRender();
    }
});

// Mousemove — preview
container.addEventListener('mousemove', (e) => {
    if (activeTool !== 'stripFooting' || !stripFtgState.placing) return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    let sheetPos = engine.coords.screenToSheet(sx, sy);
    const snap = findSnap(sx, sy);
    if (snap) sheetPos = { x: snap.x, y: snap.y };

    // Ortho lock or Shift: constrain to horizontal/vertical (or 45°)
    if (stripFtgState.startPoint) {
        if (snapState.orthoLock) {
            sheetPos = applyOrtho(sheetPos.x, sheetPos.y, stripFtgState.startPoint.x, stripFtgState.startPoint.y);
        } else if (e.shiftKey) {
            const dx = sheetPos.x - stripFtgState.startPoint.x;
            const dy = sheetPos.y - stripFtgState.startPoint.y;
            const angle = Math.atan2(dy, dx);
            const snapAngle = Math.round(angle / (Math.PI / 4)) * (Math.PI / 4);
            const dist = Math.sqrt(dx * dx + dy * dy);
            sheetPos = {
                x: stripFtgState.startPoint.x + dist * Math.cos(snapAngle),
                y: stripFtgState.startPoint.y + dist * Math.sin(snapAngle)
            };
        }
    }

    stripFtgState.currentEnd = { ...sheetPos };
    engine.requestRender();
});

// Right-click / Escape: cancel
container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'stripFooting' && stripFtgState.placing) {
        e.preventDefault();
        stripFtgState.placing = false;
        stripFtgState.startPoint = null;
        stripFtgState.currentEnd = null;
        engine.requestRender();
    }
});

// ── Strip Footing: 2D Rendering ──
// Draw strip footing preview (while placing)
engine.onRender((ctx, eng) => {
    if (activeTool !== 'stripFooting' || !stripFtgState.placing || !stripFtgState.startPoint || !stripFtgState.currentEnd) return;

    const zoom = eng.viewport.zoom;
    const p1 = eng.coords.sheetToScreen(stripFtgState.startPoint.x, stripFtgState.startPoint.y);
    const p2 = eng.coords.sheetToScreen(stripFtgState.currentEnd.x, stripFtgState.currentEnd.y);

    // Preview dashed centreline
    ctx.strokeStyle = '#886644';
    ctx.lineWidth = Math.max(0.5, 0.25 * zoom);
    ctx.setLineDash([1.5 * zoom, 1 * zoom]);
    ctx.beginPath();
    ctx.moveTo(p1.x, p1.y);
    ctx.lineTo(p2.x, p2.y);
    ctx.stroke();
    ctx.setLineDash([]);

    // Preview width outline — use current placement type
    const typeRef = placementTypeRef.stripfooting || 'SF1';
    const typeData = project.scheduleTypes.stripfooting[typeRef] || { width: 300 };
    const halfW = (typeData.width / 2 / CONFIG.drawingScale) * zoom;
    const dx = p2.x - p1.x, dy = p2.y - p1.y;
    const len = Math.sqrt(dx * dx + dy * dy);
    if (len > 0.5) {
        const nx = -dy / len * halfW, ny = dx / len * halfW;
        ctx.strokeStyle = 'rgba(136, 102, 68, 0.4)';
        ctx.lineWidth = Math.max(0.5, 0.25 * zoom);
        ctx.setLineDash([1.5 * zoom, 1 * zoom]);
        ctx.beginPath();
        ctx.moveTo(p1.x + nx, p1.y + ny);
        ctx.lineTo(p2.x + nx, p2.y + ny);
        ctx.lineTo(p2.x - nx, p2.y - ny);
        ctx.lineTo(p1.x - nx, p1.y - ny);
        ctx.closePath();
        ctx.stroke();
        ctx.setLineDash([]);
    }
});

// Draw committed strip footings (with mitre joins at corners)
engine.onRender((ctx, eng) => {
    const zoom = eng.viewport.zoom;
    const allSF = [];
    for (const el of project.getVisibleElements()) {
        if (el.type === 'stripFooting') allSF.push(el);
    }

    // Helper: line-line intersection
    function llIntersect(p1x,p1y,p2x,p2y, p3x,p3y,p4x,p4y) {
        const d = (p1x-p2x)*(p3y-p4y) - (p1y-p2y)*(p3x-p4x);
        if (Math.abs(d) < 0.0001) return null;
        const t = ((p1x-p3x)*(p3y-p4y) - (p1y-p3y)*(p3x-p4x)) / d;
        return { x: p1x + t*(p2x-p1x), y: p1y + t*(p2y-p1y) };
    }

    for (const el of allSF) {
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const p1 = eng.coords.realToScreen(el.x1, el.y1);
        const p2 = eng.coords.realToScreen(el.x2, el.y2);
        const dx = p2.x - p1.x, dy = p2.y - p1.y;
        const len = Math.sqrt(dx * dx + dy * dy);
        if (len < 0.5) continue;

        // Look up dimensions from schedule type
        const typeRef = el.typeRef || el.tag || 'SF1';
        const typeData = project.scheduleTypes.stripfooting[typeRef] || {};
        const fW = el.footingWidth || typeData.width || 300;
        const halfW = (fW / 2 / CONFIG.drawingScale) * zoom;
        const ux = dx / len, uy = dy / len;
        const nx = -uy * halfW, ny = ux * halfW;

        // Base corner points
        let s1a = { x: p1.x + nx, y: p1.y + ny }; // start side A
        let s1b = { x: p1.x - nx, y: p1.y - ny }; // start side B
        let s2a = { x: p2.x + nx, y: p2.y + ny }; // end side A
        let s2b = { x: p2.x - nx, y: p2.y - ny }; // end side B

        let startJoined = false, endJoined = false;

        // Find connected strip footings and compute mitre
        const joinTol = (el.footingWidth / CONFIG.drawingScale) * zoom * 0.6;

        for (const other of allSF) {
            if (other === el) continue;
            const op1 = eng.coords.realToScreen(other.x1, other.y1);
            const op2 = eng.coords.realToScreen(other.x2, other.y2);
            const odx = op2.x - op1.x, ody = op2.y - op1.y;
            const olen = Math.sqrt(odx * odx + ody * ody);
            if (olen < 0.5) continue;

            const oHalfW = (other.footingWidth / 2 / CONFIG.drawingScale) * zoom;
            const oux = odx / olen, ouy = ody / olen;
            const onx = -ouy * oHalfW, ony = oux * oHalfW;

            // Check if start of this footing connects to either end of other
            const d1s = Math.sqrt(Math.pow(p1.x - op1.x, 2) + Math.pow(p1.y - op1.y, 2));
            const d1e = Math.sqrt(Math.pow(p1.x - op2.x, 2) + Math.pow(p1.y - op2.y, 2));
            if (d1s < joinTol || d1e < joinTol) {
                startJoined = true;
                // Mitre at start end
                const oA1 = { x: op1.x + onx, y: op1.y + ony };
                const oA2 = { x: op2.x + onx, y: op2.y + ony };
                const oB1 = { x: op1.x - onx, y: op1.y - ony };
                const oB2 = { x: op2.x - onx, y: op2.y - ony };
                const intA = llIntersect(s1a.x,s1a.y, s2a.x,s2a.y, oA1.x,oA1.y, oA2.x,oA2.y)
                          || llIntersect(s1a.x,s1a.y, s2a.x,s2a.y, oB1.x,oB1.y, oB2.x,oB2.y);
                const intB = llIntersect(s1b.x,s1b.y, s2b.x,s2b.y, oB1.x,oB1.y, oB2.x,oB2.y)
                          || llIntersect(s1b.x,s1b.y, s2b.x,s2b.y, oA1.x,oA1.y, oA2.x,oA2.y);
                if (intA) s1a = intA;
                if (intB) s1b = intB;
            }

            // Check if end of this footing connects to either end of other
            const d2s = Math.sqrt(Math.pow(p2.x - op1.x, 2) + Math.pow(p2.y - op1.y, 2));
            const d2e = Math.sqrt(Math.pow(p2.x - op2.x, 2) + Math.pow(p2.y - op2.y, 2));
            if (d2s < joinTol || d2e < joinTol) {
                endJoined = true;
                const oA1 = { x: op1.x + onx, y: op1.y + ony };
                const oA2 = { x: op2.x + onx, y: op2.y + ony };
                const oB1 = { x: op1.x - onx, y: op1.y - ony };
                const oB2 = { x: op2.x - onx, y: op2.y - ony };
                const intA = llIntersect(s2a.x,s2a.y, s1a.x,s1a.y, oA1.x,oA1.y, oA2.x,oA2.y)
                          || llIntersect(s2a.x,s2a.y, s1a.x,s1a.y, oB1.x,oB1.y, oB2.x,oB2.y);
                const intB = llIntersect(s2b.x,s2b.y, s1b.x,s1b.y, oB1.x,oB1.y, oB2.x,oB2.y)
                          || llIntersect(s2b.x,s2b.y, s1b.x,s1b.y, oA1.x,oA1.y, oA2.x,oA2.y);
                if (intA) s2a = intA;
                if (intB) s2b = intB;
            }
        }

        const selected = isElementSelected(el);
        const color = selected ? '#2B7CD0' : layer.color;

        // Draw only outside lines — hide internal end caps at joined corners
        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(0.5, 0.25 * zoom);
        ctx.setLineDash([1.5 * zoom, 1 * zoom]);

        // Side A (always drawn): s1a → s2a
        ctx.beginPath();
        ctx.moveTo(s1a.x, s1a.y);
        ctx.lineTo(s2a.x, s2a.y);
        ctx.stroke();

        // End cap at end: s2a → s2b (only if end is NOT joined)
        if (!endJoined) {
            ctx.beginPath();
            ctx.moveTo(s2a.x, s2a.y);
            ctx.lineTo(s2b.x, s2b.y);
            ctx.stroke();
        }

        // Side B (always drawn): s2b → s1b
        ctx.beginPath();
        ctx.moveTo(s2b.x, s2b.y);
        ctx.lineTo(s1b.x, s1b.y);
        ctx.stroke();

        // End cap at start: s1b → s1a (only if start is NOT joined)
        if (!startJoined) {
            ctx.beginPath();
            ctx.moveTo(s1b.x, s1b.y);
            ctx.lineTo(s1a.x, s1a.y);
            ctx.stroke();
        }

        ctx.setLineDash([]);

        // Fill with type colour at 15% opacity
        const typeColor = typeData.color;
        if (typeColor && !selected) {
            ctx.fillStyle = typeColor + '14'; // ~8% opacity — subtle fill
            ctx.beginPath();
            ctx.moveTo(s1a.x, s1a.y);
            ctx.lineTo(s2a.x, s2a.y);
            ctx.lineTo(s2b.x, s2b.y);
            ctx.lineTo(s1b.x, s1b.y);
            ctx.closePath();
            ctx.fill();
        }

        // Tag label (SF1 etc.) at midpoint — rotated to match strip footing angle
        if (el.tag) {
            const mx = (p1.x + p2.x) / 2, my = (p1.y + p2.y) / 2;
            const fontSize = Math.max(1, 2.5 * zoom);
            ctx.save();
            ctx.font = `bold ${fontSize}px "Segoe UI", Arial, sans-serif`;
            ctx.fillStyle = color;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';

            // Offset perpendicular to the footing centreline
            const tagNx = -dy / len, tagNy = dx / len;
            const tagOffset = halfW + 2 * zoom;
            // Apply custom tag offset if user has dragged it
            const elTagDx = (el._tagOffsetX || 0) / CONFIG.drawingScale * zoom;
            const elTagDy = (el._tagOffsetY || 0) / CONFIG.drawingScale * zoom;
            const tx = mx + tagNx * tagOffset + elTagDx;
            const ty = my + tagNy * tagOffset + elTagDy;

            // Compute rotation angle aligned with the footing
            let angle = Math.atan2(dy, dx);
            // Keep text readable (never upside-down)
            if (angle > Math.PI / 2) angle -= Math.PI;
            if (angle < -Math.PI / 2) angle += Math.PI;

            ctx.translate(tx, ty);
            ctx.rotate(angle);
            ctx.fillText(el.tag, 0, 0);
            ctx.restore();
        }
    }
});

// ── Strip Footing: Hit Testing ──
// Add to the hitTestElement function — insert strip footing check
const _origHitTest = hitTestElement;
hitTestElement = function(sheetPos, tolerance) {
    // Check strip footings first
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'stripFooting') continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const p1 = engine.coords.realToSheet(el.x1, el.y1);
        const p2 = engine.coords.realToSheet(el.x2, el.y2);
        const dx = p2.x - p1.x, dy = p2.y - p1.y;
        const len = Math.sqrt(dx * dx + dy * dy);
        if (len < 0.1) continue;

        const halfT = (el.footingWidth / CONFIG.drawingScale / 2) + tolerance;
        const t = Math.max(0, Math.min(1, ((sheetPos.x - p1.x) * dx + (sheetPos.y - p1.y) * dy) / (len * len)));
        const closestX = p1.x + t * dx;
        const closestY = p1.y + t * dy;
        const dist = Math.sqrt(Math.pow(sheetPos.x - closestX, 2) + Math.pow(sheetPos.y - closestY, 2));

        if (dist < halfT) return el;
    }
    // Fall through to original hit test
    return _origHitTest(sheetPos, tolerance);
};

// ── Strip Footing: Drag Support ──
// Extend the existing drag handler to recognise stripFooting type
const _origDragStartCheck = 'stripFooting';
// The main drag handler already checks for el.type containing x1,y1,x2,y2
// Let's ensure stripFooting is included in the drag origCoords list

// ── Strip Footing: Wall-to-Footing Workflow ──
// Double-click on a wall while in Select mode → offer to add strip footing
const sfModal = document.getElementById('strip-footing-modal');
let _sfWallTarget = null;

container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'select') return;
    if (!selectedElement || selectedElement.type !== 'wall') return;

    // Only on Ground Floor
    if (getActiveLevel().id !== 'GF') {
        alert('Strip footings can only be added under walls on the Ground Floor.');
        return;
    }

    _sfWallTarget = selectedElement;
    document.getElementById('sf-width').value = '300';
    document.getElementById('sf-depth').value = '500';
    document.getElementById('sf-reinforcement').value = '';
    document.getElementById('sf-setdown').value = '200';
    sfModal.classList.remove('hidden');
});

sfModal.querySelector('.btn-cancel').addEventListener('click', () => {
    sfModal.classList.add('hidden');
    _sfWallTarget = null;
});

// ── Double-click strip footing to edit properties ──
let _sfEditTarget = null;
container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'select') return;
    if (!selectedElement || selectedElement.type !== 'stripFooting') return;

    _sfEditTarget = selectedElement;
    document.getElementById('sf-width').value = _sfEditTarget.footingWidth || 300;
    document.getElementById('sf-depth').value = _sfEditTarget.footingDepth || 500;
    document.getElementById('sf-reinforcement').value = _sfEditTarget.reinforcement || '';
    document.getElementById('sf-setdown').value = _sfEditTarget.depthBelowFSL || 200;
    _sfWallTarget = null; // not a wall-to-footing flow
    sfModal.classList.remove('hidden');
});

// Override strip footing confirm to handle both create and edit
const _origSfConfirm = sfModal.querySelector('.btn-confirm');
const _newSfConfirm = _origSfConfirm.cloneNode(true);
_origSfConfirm.parentNode.replaceChild(_newSfConfirm, _origSfConfirm);
_newSfConfirm.addEventListener('click', () => {
    if (_sfEditTarget) {
        // Editing existing strip footing
        _sfEditTarget.footingWidth = parseFloat(document.getElementById('sf-width').value) || 300;
        _sfEditTarget.footingDepth = parseFloat(document.getElementById('sf-depth').value) || 500;
        _sfEditTarget.reinforcement = document.getElementById('sf-reinforcement').value.trim() || '';
        _sfEditTarget.depthBelowFSL = parseFloat(document.getElementById('sf-setdown').value) || 200;
        _sfEditTarget = null;
        sfModal.classList.add('hidden');
        engine.requestRender();
        return;
    }
    if (!_sfWallTarget) { sfModal.classList.add('hidden'); return; }
    const wall = _sfWallTarget;

    const tag = 'SF1';
    const newStripFtg = {
        id: generateId(),
        type: 'stripFooting',
        layer: 'S-FTNG',
        level: 'GF',
        x1: wall.x1, y1: wall.y1,
        x2: wall.x2, y2: wall.y2,
        tag: tag,
        footingWidth: parseFloat(document.getElementById('sf-width').value) || 300,
        footingDepth: parseFloat(document.getElementById('sf-depth').value) || 500,
        reinforcement: document.getElementById('sf-reinforcement').value.trim() || '',
        depthBelowFSL: parseFloat(document.getElementById('sf-setdown').value) || 200,
        wallRef: wall.id,
    };

    history.execute({
        description: 'Add strip footing under wall: ' + tag,
        execute() { project.elements.push(newStripFtg); },
        undo() {
            const i = project.elements.indexOf(newStripFtg);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    sfModal.classList.add('hidden');
    _sfWallTarget = null;
    engine.requestRender();
});

// ── Multi-Select Footing Set-Down (right-click on selected footings) ──
container.addEventListener('contextmenu', (e) => {
    if (activeTool !== 'select') return;

    // Check if we have multiple footings selected
    const footingSel = selectedElements.filter(el =>
        el.type === 'footing' || el.type === 'stripFooting'
    );
    // Also include single selected footing
    if (footingSel.length === 0 && selectedElement &&
        (selectedElement.type === 'footing' || selectedElement.type === 'stripFooting')) {
        footingSel.push(selectedElement);
    }
    if (footingSel.length === 0) return;

    e.preventDefault();
    const currentVal = footingSel[0].depthBelowFSL || 200;
    const input = prompt(
        `Set depth below FSL for ${footingSel.length} footing(s).\n` +
        `Default is 200mm below Finished Surface Level.\n` +
        `Current: ${currentVal}mm\n\nEnter depth below FSL (mm):`,
        String(currentVal)
    );
    if (input === null) return;
    const val = parseFloat(input) || 400;
    for (const el of footingSel) {
        el.depthBelowFSL = val;
    }
    engine.requestRender();
});

// ── Auto Footing Schedule ───────────────────────────────────

// btn-ftg-schedule removed — now handled by unified Schedules button
const _ftgSchedBtn = document.getElementById('btn-ftg-schedule');
if (_ftgSchedBtn) _ftgSchedBtn.addEventListener('click', () => {
    const lvId = getActiveLevel().id;
    const padFootings = project.elements.filter(el => el.type === 'footing' && el.level === lvId);
    const stripFootings = project.elements.filter(el => el.type === 'stripFooting' && el.level === lvId);

    if (padFootings.length === 0 && stripFootings.length === 0) {
        alert('No footings placed on this level.');
        return;
    }

    // Group by mark and collect unique marks
    const marks = {};
    for (const f of padFootings) {
        if (!marks[f.mark]) {
            const fw = f.footingWidth || f.width || 1000;
            const fd = f.footingDepth || f.depth || 400;
            marks[f.mark] = {
                mark: f.mark,
                size: `${fw}x${fw}x${fd}D PAD`,
                reinforcement: f.reinforcement || ''
            };
        }
    }
    for (const sf of stripFootings) {
        if (sf.tag && !marks[sf.tag]) {
            marks[sf.tag] = {
                mark: sf.tag,
                size: `${sf.footingWidth}W x ${sf.footingDepth}D STRIP`,
                reinforcement: sf.reinforcement || ''
            };
        }
    }

    // Build table rows: [MARK, SIZE, REINFORCEMENT]
    const rows = [['MARK', 'SIZE', 'REINFORCEMENT']];
    for (const mark of Object.keys(marks).sort()) {
        const data = marks[mark];
        rows.push([data.mark, data.size, data.reinforcement]);
    }

    // Create table element
    const newTable = {
        id: generateId(),
        type: 'table',
        layer: 'S-ANNO',
        level: getActiveLevel().id,
        x: 15, // Bottom-left region in real coordinates
        y: 10,
        rows: rows,
        template: 'footing',
        cellWidth: 35,
        cellHeight: 8
    };

    history.execute({
        description: 'Auto-generate footing schedule',
        execute() { project.elements.push(newTable); },
        undo() {
            const i = project.elements.indexOf(newTable);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    engine.requestRender();
});

// ══════════════════════════════════════════════════════════
