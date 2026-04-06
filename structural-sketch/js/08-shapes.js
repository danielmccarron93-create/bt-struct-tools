// ── PHASE 8: POLYLINE + RECTANGLE + COPY/PASTE ───────────
// ══════════════════════════════════════════════════════════

// ── Polyline Tool ────────────────────────────────────────

const polyToolState = {
    points: [],       // array of { x, y } in sheet-mm
    currentEnd: null, // preview endpoint
};

// Wire polyline button + keyboard
document.getElementById('btn-polyline').addEventListener('click', () => setActiveTool('polyline'));

// P = polyline, R = rect
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'p') setActiveTool('polyline');
    if (e.key === 'r') setActiveTool('rect');
});

function getPolyToolPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    if (polyToolState.points.length > 0) {
        const last = polyToolState.points[polyToolState.points.length - 1];
        pos = applyOrtho(pos.x, pos.y, last.x, last.y);
    }
    return pos;
}

container.addEventListener('mousemove', (e) => {
    if (activeTool === 'polyline') {
        polyToolState.currentEnd = getPolyToolPos(e);
        engine.requestRender();
    }
});

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'polyline') return;
    if (pdfState.calibrating) return;

    const pos = getPolyToolPos(e);
    polyToolState.points.push(pos);
    engine.requestRender();
});

// Right-click or Escape finishes polyline
container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'polyline' && polyToolState.points.length >= 2) {
        e.preventDefault();
        commitPolyline();
    } else if (activeTool === 'polyline') {
        e.preventDefault();
        polyToolState.points = [];
        engine.requestRender();
    }
});

window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && activeTool === 'polyline' && polyToolState.points.length >= 2) {
        commitPolyline();
    }
});

function commitPolyline() {
    const pts = polyToolState.points;
    if (pts.length < 2) return;

    const layerId = getActiveLayer();
    const realPoints = pts.map(p => engine.coords.sheetToReal(p.x, p.y));

    // Auto-detect if the polyline is closed (first/last point within 2mm)
    const fp = realPoints[0], lp = realPoints[realPoints.length - 1];
    const closeDist = Math.sqrt(Math.pow(fp.x - lp.x, 2) + Math.pow(fp.y - lp.y, 2));
    const isClosed = closeDist < 2 * CONFIG.drawingScale; // within ~2mm on sheet

    const newPoly = {
        id: generateId(),
        type: 'polyline',
        layer: layerId,
        points: realPoints,
        closed: isClosed || false,
    };

    history.execute({
        description: 'Draw polyline',
        execute() { project.elements.push(newPoly); },
        undo() {
            const i = project.elements.indexOf(newPoly);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    polyToolState.points = [];
    polyToolState.currentEnd = null;
    engine.requestRender();
}

// ── Slab Tool (polyline-based with hatch and thickness callout) ─────────

const slabToolState = {
    points: [],       // sheet-mm vertices
    currentEnd: null, // preview next vertex
};

function getSlabToolPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    if (slabToolState.points.length > 0) {
        const last = slabToolState.points[slabToolState.points.length - 1];
        pos = applyOrtho(pos.x, pos.y, last.x, last.y);
    }
    return pos;
}

container.addEventListener('mousemove', (e) => {
    if (activeTool === 'slab') {
        slabToolState.currentEnd = getSlabToolPos(e);
        engine.requestRender();
    }
});

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'slab') return;
    if (pdfState.calibrating) return;

    const pos = getSlabToolPos(e);
    slabToolState.points.push(pos);
    engine.requestRender();
});

// Double-click closes the slab polygon
container.addEventListener('dblclick', (e) => {
    if (activeTool === 'slab' && slabToolState.points.length >= 3) {
        e.preventDefault();
        e.stopPropagation();
        // Remove the last point (double-click adds a duplicate from the second mousedown)
        slabToolState.points.pop();
        if (slabToolState.points.length >= 3) {
            commitSlab();
        }
    }
});

// Right-click or Escape also finishes slab
container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'slab' && slabToolState.points.length >= 3) {
        e.preventDefault();
        commitSlab();
    } else if (activeTool === 'slab') {
        e.preventDefault();
        slabToolState.points = [];
        engine.requestRender();
    }
});

window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && activeTool === 'slab' && slabToolState.points.length >= 3) {
        commitSlab();
    }
});

function commitSlab() {
    const pts = slabToolState.points;
    if (pts.length < 3) return;

    const layerId = 'S-SLAB'; // Always use S-SLAB layer for slabs
    const realPoints = pts.map(p => engine.coords.sheetToReal(p.x, p.y));

    // Close the slab polygon by adding first point at the end (if not already close)
    const fp = realPoints[0], lp = realPoints[realPoints.length - 1];
    const closeDist = Math.sqrt(Math.pow(fp.x - lp.x, 2) + Math.pow(fp.y - lp.y, 2));
    if (closeDist > 2 * CONFIG.drawingScale) {
        realPoints.push(fp); // explicitly close
    }

    const newSlab = {
        id: generateId(),
        type: 'polyline',
        layer: layerId,
        points: realPoints,
        closed: true,
        hatch: 'concrete',
        fillColor: '#888888',
        fillOpacity: 0.15,
        slabThickness: 200, // default thickness in mm
    };

    history.execute({
        description: 'Draw slab',
        execute() { project.elements.push(newSlab); },
        undo() {
            const i = project.elements.indexOf(newSlab);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    slabToolState.points = [];
    slabToolState.currentEnd = null;
    engine.requestRender();
}

// ── Rectangle Tool (drag-release) ────────────────────────

const rectToolState = {
    dragging: false,
    startPoint: null,
    currentEnd: null,
};

document.getElementById('btn-rect').addEventListener('click', () => setActiveTool('rect'));

function getRectToolPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);

    // Shift constrains to square
    if (_shiftDown && rectToolState.startPoint) {
        const s = rectToolState.startPoint;
        const dx = pos.x - s.x;
        const dy = pos.y - s.y;
        const side = Math.max(Math.abs(dx), Math.abs(dy));
        pos = { x: s.x + side * Math.sign(dx), y: s.y + side * Math.sign(dy) };
    }
    return pos;
}

// Mousedown — start dragging
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'rect') return;
    if (pdfState.calibrating) return;

    rectToolState.dragging = true;
    rectToolState.startPoint = getRectToolPos(e);
    rectToolState.currentEnd = rectToolState.startPoint;
});

// Mousemove — update preview
window.addEventListener('mousemove', (e) => {
    if (activeTool !== 'rect' || !rectToolState.dragging) return;
    rectToolState.currentEnd = getRectToolPos(e);
    engine.requestRender();
});

// Mouseup — commit rectangle
window.addEventListener('mouseup', (e) => {
    if (activeTool !== 'rect' || !rectToolState.dragging) return;
    rectToolState.dragging = false;

    const s = rectToolState.startPoint;
    const end = getRectToolPos(e);

    // Minimum size check (at least 1mm on sheet)
    const dxSheet = Math.abs(end.x - s.x);
    const dySheet = Math.abs(end.y - s.y);
    if (dxSheet < 1 && dySheet < 1) {
        rectToolState.startPoint = null;
        rectToolState.currentEnd = null;
        engine.requestRender();
        return;
    }

    const r1 = engine.coords.sheetToReal(s.x, s.y);
    const r2 = engine.coords.sheetToReal(end.x, s.y);
    const r3 = engine.coords.sheetToReal(end.x, end.y);
    const r4 = engine.coords.sheetToReal(s.x, end.y);

    const layerId = getActiveLayer();
    const newRect = {
        id: generateId(),
        type: 'polyline',
        layer: layerId,
        points: [r1, r2, r3, r4],
        closed: true,
    };

    history.execute({
        description: 'Draw rectangle',
        execute() { project.elements.push(newRect); },
        undo() {
            const i = project.elements.indexOf(newRect);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    rectToolState.startPoint = null;
    rectToolState.currentEnd = null;
    engine.requestRender();
});

container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'rect' && rectToolState.dragging) {
        e.preventDefault();
        rectToolState.dragging = false;
        rectToolState.startPoint = null;
        rectToolState.currentEnd = null;
        engine.requestRender();
    }
});

// ── Polyline + Rectangle Rendering ───────────────────────

// Patch into the wrapped drawElements
const prevWrappedDraw = wrappedDrawElements;
const extDrawElements = function(ctx, eng) {
    prevWrappedDraw(ctx, eng);

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    // Draw polylines
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'polyline') continue;
        const layer = project.layers[el.layer];
        if (!layer) continue;
        const isSelected = (selectedElement === el);
        const pts = el.points;
        if (pts.length < 2) continue;

        ctx.strokeStyle = isSelected ? '#2B7CD0' : layer.color;
        ctx.lineWidth = Math.max(1, layer.lineWeight * zoom);
        const pattern = DASH_PATTERNS[layer.pattern] || [];
        ctx.setLineDash(pattern.length > 0 ? pattern.map(d => d * zoom) : []);

        ctx.beginPath();
        const p0 = coords.realToScreen(pts[0].x, pts[0].y);
        ctx.moveTo(p0.x, p0.y);
        for (let i = 1; i < pts.length; i++) {
            const p = coords.realToScreen(pts[i].x, pts[i].y);
            ctx.lineTo(p.x, p.y);
        }
        if (el.closed) ctx.closePath();
        ctx.stroke();
        ctx.setLineDash([]);

        // Selection: show vertices
        if (isSelected) {
            ctx.fillStyle = '#2B7CD0';
            for (const pt of pts) {
                const sp = coords.realToScreen(pt.x, pt.y);
                ctx.fillRect(sp.x - 3, sp.y - 3, 6, 6);
            }
        }
    }

    // Polyline preview
    if (activeTool === 'polyline' && polyToolState.points.length > 0) {
        const layerId = getActiveLayer();
        const layer = project.layers[layerId];

        ctx.strokeStyle = layer ? layer.color : '#000';
        ctx.lineWidth = Math.max(1, (layer ? layer.lineWeight : 0.35) * zoom);
        ctx.globalAlpha = 0.6;
        const pattern = DASH_PATTERNS[layer ? layer.pattern : 'solid'] || [];
        ctx.setLineDash(pattern.length > 0 ? pattern.map(d => d * zoom) : []);

        ctx.beginPath();
        const fp = coords.sheetToScreen(polyToolState.points[0].x, polyToolState.points[0].y);
        ctx.moveTo(fp.x, fp.y);
        for (let i = 1; i < polyToolState.points.length; i++) {
            const p = coords.sheetToScreen(polyToolState.points[i].x, polyToolState.points[i].y);
            ctx.lineTo(p.x, p.y);
        }
        if (polyToolState.currentEnd) {
            const ep = coords.sheetToScreen(polyToolState.currentEnd.x, polyToolState.currentEnd.y);
            ctx.lineTo(ep.x, ep.y);
        }
        ctx.stroke();
        ctx.globalAlpha = 1.0;
        ctx.setLineDash([]);

        // Vertex dots
        ctx.fillStyle = '#2B7CD0';
        for (const pt of polyToolState.points) {
            const sp = coords.sheetToScreen(pt.x, pt.y);
            ctx.beginPath(); ctx.arc(sp.x, sp.y, 3, 0, Math.PI * 2); ctx.fill();
        }
    }

    // Slab preview
    if (activeTool === 'slab' && slabToolState.points.length > 0) {
        const layer = project.layers['S-SLAB'];

        ctx.strokeStyle = layer ? layer.color : '#000';
        ctx.lineWidth = Math.max(1, (layer ? layer.lineWeight : 0.35) * zoom);
        ctx.globalAlpha = 0.6;
        const pattern = DASH_PATTERNS[layer ? layer.pattern : 'solid'] || [];
        ctx.setLineDash(pattern.length > 0 ? pattern.map(d => d * zoom) : []);

        // Draw lines connecting vertices
        ctx.beginPath();
        const fp = coords.sheetToScreen(slabToolState.points[0].x, slabToolState.points[0].y);
        ctx.moveTo(fp.x, fp.y);
        for (let i = 1; i < slabToolState.points.length; i++) {
            const p = coords.sheetToScreen(slabToolState.points[i].x, slabToolState.points[i].y);
            ctx.lineTo(p.x, p.y);
        }

        // Preview line from last vertex to cursor
        if (slabToolState.currentEnd) {
            const ep = coords.sheetToScreen(slabToolState.currentEnd.x, slabToolState.currentEnd.y);
            ctx.lineTo(ep.x, ep.y);
            // Dashed closing line from cursor back to first point
            ctx.setLineDash([5, 5]);
            ctx.lineTo(fp.x, fp.y);
        }
        ctx.stroke();
        ctx.globalAlpha = 1.0;
        ctx.setLineDash([]);

        // Vertex dots
        ctx.fillStyle = '#2B7CD0';
        for (const pt of slabToolState.points) {
            const sp = coords.sheetToScreen(pt.x, pt.y);
            ctx.beginPath(); ctx.arc(sp.x, sp.y, 3, 0, Math.PI * 2); ctx.fill();
        }
    }

    // Rectangle preview
    if (activeTool === 'rect' && rectToolState.dragging && rectToolState.startPoint && rectToolState.currentEnd) {
        const s = rectToolState.startPoint;
        const e = rectToolState.currentEnd;
        const sp1 = coords.sheetToScreen(s.x, s.y);
        const sp2 = coords.sheetToScreen(e.x, e.y);

        ctx.strokeStyle = '#2B7CD0';
        ctx.lineWidth = 1.5;
        ctx.setLineDash([5, 3]);
        ctx.strokeRect(
            Math.min(sp1.x, sp2.x), Math.min(sp1.y, sp2.y),
            Math.abs(sp2.x - sp1.x), Math.abs(sp2.y - sp1.y)
        );
        ctx.setLineDash([]);

        // Show dimensions
        const r1 = coords.sheetToReal(s.x, s.y);
        const r2 = coords.sheetToReal(e.x, e.y);
        const w = Math.abs(r2.x - r1.x), h = Math.abs(r2.y - r1.y);
        const fmtD = (mm) => mm >= 1000 ? (mm/1000).toFixed(1) + 'm' : Math.round(mm) + '';

        ctx.font = 'bold 11px "Segoe UI", Arial, sans-serif';
        ctx.fillStyle = '#2B7CD0';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(fmtD(w), (sp1.x + sp2.x) / 2, Math.min(sp1.y, sp2.y) - 4);
        ctx.textBaseline = 'middle';
        ctx.textAlign = 'left';
        ctx.fillText(fmtD(h), Math.max(sp1.x, sp2.x) + 6, (sp1.y + sp2.y) / 2);
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';
    }
};

// Replace in render callbacks
const cbIdx2 = engine._renderCallbacks.indexOf(wrappedDrawElements);
if (cbIdx2 !== -1) engine._renderCallbacks[cbIdx2] = extDrawElements;

// ── Polyline hit-testing (patch hitTestElement) ───────────

/** Point-in-polygon test (ray casting algorithm) using sheet-mm coords */
function pointInPolygon(px, py, sheetPts) {
    let inside = false;
    for (let i = 0, j = sheetPts.length - 1; i < sheetPts.length; j = i++) {
        const xi = sheetPts[i].x, yi = sheetPts[i].y;
        const xj = sheetPts[j].x, yj = sheetPts[j].y;
        if (((yi > py) !== (yj > py)) && (px < (xj - xi) * (py - yi) / (yj - yi) + xi)) {
            inside = !inside;
        }
    }
    return inside;
}

// ── Polyline drag support (patch drag handlers) ──────────

// Extend the mousedown to handle polyline drag start
const origDragMousedown = container._selectMousedown;
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'select') return;
    if (!selectedElement || selectedElement.type !== 'polyline') return;
    if (!dragState.dragging) return; // already started by the main handler

    // Store polyline original points
    dragState.origCoords = {
        points: selectedElement.points.map(p => ({ x: p.x, y: p.y }))
    };
});

// Patch the mousemove drag for polylines
const origDragMove = window._dragMoveHandler;
window.addEventListener('mousemove', (e) => {
    if (!dragState.dragging || !dragState.el || dragState.el.type !== 'polyline' && dragState.el.type !== 'cloud') return;
    if (!dragState.origCoords || !dragState.origCoords.points) {
        // Store original if not yet stored
        dragState.origCoords = { points: dragState.el.points.map(p => ({ x: p.x, y: p.y })) };
    }
    const sheetPos = engine.getSheetPos(e);
    const dx = (sheetPos.x - dragState.startSheet.x) * CONFIG.drawingScale;
    const dy = (sheetPos.y - dragState.startSheet.y) * CONFIG.drawingScale;

    for (let i = 0; i < dragState.el.points.length; i++) {
        dragState.el.points[i].x = dragState.origCoords.points[i].x + dx;
        dragState.el.points[i].y = dragState.origCoords.points[i].y + dy;
    }
    engine.requestRender();
});

// Patch mouseup for polyline undo
window.addEventListener('mouseup', (e) => {
    if (!dragState.dragging || !dragState.el || dragState.el.type !== 'polyline' && dragState.el.type !== 'cloud') return;
    if (!dragState.origCoords || !dragState.origCoords.points) return;

    const el = dragState.el;
    const origPts = dragState.origCoords.points;
    const newPts = el.points.map(p => ({ x: p.x, y: p.y }));
    const moved = origPts.some((p, i) => Math.abs(p.x - newPts[i].x) > 0.1 || Math.abs(p.y - newPts[i].y) > 0.1);

    if (moved) {
        history.undoStack.push({
            description: 'Move polyline',
            execute() { el.points.forEach((p, i) => { p.x = newPts[i].x; p.y = newPts[i].y; }); },
            undo() { el.points.forEach((p, i) => { p.x = origPts[i].x; p.y = origPts[i].y; }); }
        });
        history.redoStack.length = 0;
    }
});

// Also need polyline support in the main drag start
// Patch: when selecting a polyline, start drag with points
const origSelectDown = container.__selectHandler;
// This is handled by the existing drag start which checks for polyline type

// ── Copy / Paste ─────────────────────────────────────────

let clipboard = null;

window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;

    // Ctrl+C = Copy
    if ((e.ctrlKey || e.metaKey) && e.key === 'c' && selectedElement) {
        e.preventDefault();
        clipboard = JSON.parse(JSON.stringify(selectedElement));
        if (typeof updatePasteButton === 'function') updatePasteButton();
        console.log('[Copy] Copied ' + clipboard.type);
    }

    // Ctrl+V = Paste (offset by 5mm on sheet, assign to active level)
    if ((e.ctrlKey || e.metaKey) && e.key === 'v' && clipboard && !Array.isArray(clipboard)) {
        e.preventDefault();
        const pasted = JSON.parse(JSON.stringify(clipboard));
        pasted.id = generateId();

        // Assign to current active level (enables cross-level paste)
        const activeLevel = typeof getActiveLevel === 'function' ? getActiveLevel() : null;
        if (activeLevel) pasted.level = activeLevel.id;

        // Clear tag offsets on paste (fresh positioning)
        delete pasted._tagOffsetX;
        delete pasted._tagOffsetY;

        // Offset to avoid overlap
        const offsetReal = 5 * CONFIG.drawingScale; // 5mm sheet offset

        if (pasted.type === 'line' || pasted.type === 'dimension' || pasted.type === 'leader' || pasted.type === 'edge' || pasted.type === 'fallarrow' || pasted.type === 'step') {
            pasted.x1 += offsetReal; pasted.y1 += offsetReal;
            pasted.x2 += offsetReal; pasted.y2 += offsetReal;
        } else if (pasted.type === 'polyline' || pasted.type === 'cloud') {
            for (const pt of pasted.points) { pt.x += offsetReal; pt.y += offsetReal; }
        } else if (pasted.type === 'column') {
            pasted.x += offsetReal; pasted.y += offsetReal;
            pasted.tag = 'SC' + _colNextNum++;
        } else if (pasted.type === 'wall') {
            pasted.x1 += offsetReal; pasted.y1 += offsetReal;
            pasted.x2 += offsetReal; pasted.y2 += offsetReal;
            // Re-tag wall
            if (pasted.wallType === '190 Block' || pasted.wallType === '140 Block' || pasted.wallType === '90 Block') {
                pasted.tag = 'BW' + _wallBwNum++;
            } else {
                pasted.tag = 'CW' + _wallCwNum++;
            }
        } else if (pasted.type === 'beam') {
            pasted.x1 += offsetReal; pasted.y1 += offsetReal;
            pasted.x2 += offsetReal; pasted.y2 += offsetReal;
        } else if (pasted.type === 'slab') {
            if (pasted.points) {
                for (const pt of pasted.points) { pt.x += offsetReal; pt.y += offsetReal; }
            }
        } else if (pasted.type === 'padfooting') {
            pasted.x += offsetReal; pasted.y += offsetReal;
            pasted.tag = 'PF' + _ftgNextNum++;
        } else if (pasted.type === 'stripfooting') {
            pasted.x1 += offsetReal; pasted.y1 += offsetReal;
            pasted.x2 += offsetReal; pasted.y2 += offsetReal;
        } else {
            // Generic fallback for types with x/y
            if ('x' in pasted) pasted.x += offsetReal;
            if ('y' in pasted) pasted.y += offsetReal;
        }

        history.execute({
            description: 'Paste ' + pasted.type,
            execute() { project.elements.push(pasted); },
            undo() {
                const i = project.elements.indexOf(pasted);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });

        selectedElement = pasted;
        engine.requestRender();
        console.log('[Paste] Pasted ' + pasted.type + ' to level ' + (pasted.level || '?'));
    }
});

// ── Polyline endpoint snap ───────────────────────────────
// Patch findSnap to include polyline endpoints/midpoints
const origFindSnap = findSnap;
findSnap = function(screenX, screenY) {
    const sheetPos = engine.coords.screenToSheet(screenX, screenY);
    const radiusMM = snapState.snapRadius / engine.viewport.zoom;
    let best = origFindSnap(screenX, screenY);
    let bestDist = best ? Math.sqrt(Math.pow(sheetPos.x - best.x, 2) + Math.pow(sheetPos.y - best.y, 2)) : radiusMM;

    if (!snapState.enabled) return best;

    for (const el of project.getVisibleElements()) {
        if (el.type !== 'polyline') continue;
        for (const pt of el.points) {
            const sp = engine.coords.realToSheet(pt.x, pt.y);
            const d = Math.sqrt(Math.pow(sheetPos.x - sp.x, 2) + Math.pow(sheetPos.y - sp.y, 2));
            if (d < bestDist) { bestDist = d; best = { x: sp.x, y: sp.y, type: SNAP_TYPES.ENDPOINT }; }
        }
        // Midpoints of each segment
        if (snapState.midpointSnap) {
            for (let i = 0; i < el.points.length - 1; i++) {
                const p1 = engine.coords.realToSheet(el.points[i].x, el.points[i].y);
                const p2 = engine.coords.realToSheet(el.points[i+1].x, el.points[i+1].y);
                const mx = (p1.x + p2.x) / 2, my = (p1.y + p2.y) / 2;
                const d = Math.sqrt(Math.pow(sheetPos.x - mx, 2) + Math.pow(sheetPos.y - my, 2));
                if (d < bestDist) { bestDist = d; best = { x: mx, y: my, type: SNAP_TYPES.MIDPOINT }; }
            }
        }
    }
    return best;
};

// ── Polyline PDF export ──────────────────────────────────
// Patch: add polyline export to the existing PDF export button handler
// We'll hook this via the existing export — polylines are handled in the
// element loop when el.type === 'line' check fails, so we need to add support.
// For now, polylines export as individual line segments in the PDF.
// (The existing export iterates project.getVisibleElements())

// ══════════════════════════════════════════════════════════
// ── PHASE 9: LEADER LINES + REVISION CLOUDS ─────────────
// ══════════════════════════════════════════════════════════

// ── Leader Line Tool ─────────────────────────────────────
// Click arrow point → click elbow/end → type label text

const leaderState = {
    placing: false,
    startPoint: null,   // arrow tip, sheet-mm
    currentEnd: null,   // preview end
};

let activeLeaderInput = null;

document.getElementById('btn-leader').addEventListener('click', () => setActiveTool('leader'));

// Extend setActiveTool
// E/Q = leader, W = cloud, S = section (T = callout mapped in 06-drawing-tools)
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'e') setActiveTool('leader');
    if (e.key === 'q') setActiveTool('callout');
    if (e.key === 'w') setActiveTool('cloud');
    // 'g' — reserved for future use
    if (e.key === 's') { e.preventDefault(); setActiveTool('section'); }
});

function getLeaderPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    if (leaderState.placing && leaderState.startPoint) {
        pos = applyOrtho(pos.x, pos.y, leaderState.startPoint.x, leaderState.startPoint.y);
    }
    return pos;
}

container.addEventListener('mousemove', (e) => {
    if (activeTool === 'leader') {
        leaderState.currentEnd = getLeaderPos(e);
        engine.requestRender();
    }
});

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'leader') return;
    if (pdfState.calibrating) return;

    if (activeLeaderInput) return; // don't start new while typing

    const pos = getLeaderPos(e);

    if (!leaderState.placing) {
        leaderState.placing = true;
        leaderState.startPoint = pos;
    } else {
        // Second click — show text input at the end point
        const endPos = pos;
        const screenEnd = engine.coords.sheetToScreen(endPos.x, endPos.y);

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'text-input-overlay';
        input.style.left = screenEnd.x + 'px';
        input.style.top = (screenEnd.y - 14) + 'px';
        input.style.width = '120px';
        input.style.fontSize = '12px';
        input.placeholder = 'Label...';
        input._arrowPt = { x: leaderState.startPoint.x, y: leaderState.startPoint.y };
        input._endPt = { x: endPos.x, y: endPos.y };

        container.appendChild(input);
        setTimeout(() => input.focus(), 30);
        activeLeaderInput = input;

        input.addEventListener('keydown', (ev) => {
            ev.stopPropagation();
            if (ev.key === 'Enter') { ev.preventDefault(); commitLeader(); }
            if (ev.key === 'Escape') { ev.preventDefault(); cancelLeader(); }
        });
        input.addEventListener('blur', () => {
            setTimeout(() => { if (activeLeaderInput === input) commitLeader(); }, 150);
        });

        leaderState.placing = false;
        leaderState.startPoint = null;
    }
});

container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'leader' && leaderState.placing) {
        e.preventDefault();
        leaderState.placing = false;
        leaderState.startPoint = null;
        engine.requestRender();
    }
});

function commitLeader() {
    if (!activeLeaderInput) return;
    const input = activeLeaderInput;
    const text = input.value.trim();
    activeLeaderInput = null;
    if (input.parentNode) input.parentNode.removeChild(input);

    if (!text) { engine.requestRender(); return; }

    const arrowReal = engine.coords.sheetToReal(input._arrowPt.x, input._arrowPt.y);
    const endReal = engine.coords.sheetToReal(input._endPt.x, input._endPt.y);
    const textSizeSelect = document.getElementById('text-size');
    const fontSize = textSizeSelect ? parseFloat(textSizeSelect.value) : 3.5;

    const newLeader = {
        id: generateId(),
        type: 'leader',
        layer: 'S-ANNO',
        x1: arrowReal.x, y1: arrowReal.y,  // arrow tip
        x2: endReal.x, y2: endReal.y,       // text anchor
        text: text,
        fontSize: fontSize,
    };

    history.execute({
        description: 'Add leader: ' + text,
        execute() { project.elements.push(newLeader); },
        undo() {
            const i = project.elements.indexOf(newLeader);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });
    engine.requestRender();
}

function cancelLeader() {
    if (!activeLeaderInput) return;
    const input = activeLeaderInput;
    activeLeaderInput = null;
    if (input.parentNode) input.parentNode.removeChild(input);
    engine.requestRender();
}

// ── Revision Cloud Tool (drag-rectangle) ─────────────────
// Drag to define a rectangular cloud area — renders as cloud bumps

const cloudState = {
    drawing: false,
    startPoint: null,
    currentEnd: null,
    points: [],       // for backward compat with rendering/drag
};

document.getElementById('btn-cloud').addEventListener('click', () => setActiveTool('cloud'));

function getCloudPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    return snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
}

// Mousedown — start drag
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'cloud') return;
    if (pdfState.calibrating) return;

    cloudState.drawing = true;
    cloudState.startPoint = getCloudPos(e);
    cloudState.currentEnd = cloudState.startPoint;
});

// Mousemove — update preview
window.addEventListener('mousemove', (e) => {
    if (activeTool !== 'cloud' || !cloudState.drawing) return;
    cloudState.currentEnd = getCloudPos(e);
    engine.requestRender();
});

// Mouseup — commit cloud
window.addEventListener('mouseup', (e) => {
    if (activeTool !== 'cloud' || !cloudState.drawing) return;
    cloudState.drawing = false;

    const s = cloudState.startPoint;
    const end = getCloudPos(e);

    // Minimum size check
    if (Math.abs(end.x - s.x) < 2 && Math.abs(end.y - s.y) < 2) {
        cloudState.startPoint = null;
        cloudState.currentEnd = null;
        engine.requestRender();
        return;
    }

    // Create 4 corner points (closed rectangle) in real-world coords
    const r1 = engine.coords.sheetToReal(s.x, s.y);
    const r2 = engine.coords.sheetToReal(end.x, s.y);
    const r3 = engine.coords.sheetToReal(end.x, end.y);
    const r4 = engine.coords.sheetToReal(s.x, end.y);

    const newCloud = {
        id: generateId(),
        type: 'cloud',
        layer: 'S-ANNO',
        points: [r1, r2, r3, r4, { ...r1 }], // closed path
    };

    history.execute({
        description: 'Draw revision cloud',
        execute() { project.elements.push(newCloud); },
        undo() {
            const i = project.elements.indexOf(newCloud);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    cloudState.startPoint = null;
    cloudState.currentEnd = null;
    engine.requestRender();
});

container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'cloud' && cloudState.drawing) {
        e.preventDefault();
        cloudState.drawing = false;
        cloudState.startPoint = null;
        cloudState.currentEnd = null;
        engine.requestRender();
    }
});

// ── Polygon Cloud Tool ───────────────────────────────────
// Click vertices to define a polygon cloud, right-click/Esc to close

const cloudPolyState = {
    points: [],       // sheet-mm vertices
    currentEnd: null, // preview next vertex
};

document.getElementById('btn-cloud-poly').addEventListener('click', () => setActiveTool('cloudpoly'));

// Extend setActiveTool for cloudpoly
// Shift+W = polygon cloud
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'W') setActiveTool('cloudpoly'); // Shift+W (uppercase)
});

function getCloudPolyPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    if (cloudPolyState.points.length > 0) {
        const last = cloudPolyState.points[cloudPolyState.points.length - 1];
        pos = applyOrtho(pos.x, pos.y, last.x, last.y);
    }
    return pos;
}

container.addEventListener('mousemove', (e) => {
    if (activeTool === 'cloudpoly') {
        cloudPolyState.currentEnd = getCloudPolyPos(e);
        engine.requestRender();
    }
});

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'cloudpoly') return;
    if (pdfState.calibrating) return;

    const pos = getCloudPolyPos(e);
    cloudPolyState.points.push(pos);
    engine.requestRender();
});

// Right-click or Escape closes the polygon cloud
container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'cloudpoly' && cloudPolyState.points.length >= 3) {
        e.preventDefault();
        commitCloudPoly();
    } else if (activeTool === 'cloudpoly') {
        e.preventDefault();
        cloudPolyState.points = [];
        engine.requestRender();
    }
});

window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && activeTool === 'cloudpoly' && cloudPolyState.points.length >= 3) {
        commitCloudPoly();
    }
});

function commitCloudPoly() {
    const pts = cloudPolyState.points;
    if (pts.length < 3) return;

    // Close the path
    const closedPts = [...pts, { ...pts[0] }];
    const realPts = closedPts.map(p => engine.coords.sheetToReal(p.x, p.y));

    const newCloud = {
        id: generateId(),
        type: 'cloud',
        layer: 'S-ANNO',
        points: realPts,
    };

    history.execute({
        description: 'Draw polygon cloud',
        execute() { project.elements.push(newCloud); },
        undo() {
            const i = project.elements.indexOf(newCloud);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    cloudPolyState.points = [];
    cloudPolyState.currentEnd = null;
    engine.requestRender();
}

// ── Leader + Cloud Rendering ─────────────────────────────

const prevExtDraw = extDrawElements;
const phase9Draw = function(ctx, eng) {
    prevExtDraw(ctx, eng);

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    for (const el of project.getVisibleElements()) {
        const isSelected = (selectedElement === el);

        // ── Leader lines ──
        if (el.type === 'leader') {
            const layer = project.layers[el.layer];
            if (!layer || !layer.visible) continue;

            const tip = coords.realToScreen(el.x1, el.y1);
            const end = coords.realToScreen(el.x2, el.y2);
            const color = isSelected ? '#2B7CD0' : layer.color;

            // Leader line
            ctx.strokeStyle = color;
            ctx.lineWidth = Math.max(1, 0.25 * zoom);
            ctx.setLineDash([]);
            ctx.beginPath();
            ctx.moveTo(tip.x, tip.y);
            ctx.lineTo(end.x, end.y);
            ctx.stroke();

            // Arrowhead at tip
            const angle = Math.atan2(tip.y - end.y, tip.x - end.x);
            const arrowLen = Math.max(6, 2.5 * zoom);
            const arrowW = Math.PI / 7;
            ctx.fillStyle = color;
            ctx.beginPath();
            ctx.moveTo(tip.x, tip.y);
            ctx.lineTo(tip.x - arrowLen * Math.cos(angle - arrowW), tip.y - arrowLen * Math.sin(angle - arrowW));
            ctx.lineTo(tip.x - arrowLen * Math.cos(angle + arrowW), tip.y - arrowLen * Math.sin(angle + arrowW));
            ctx.closePath();
            ctx.fill();

            // Text at end with underline
            if (el.text) {
                const fontSize = (el.fontSize || CALLOUT_DEFAULTS.fontSize) * zoom;
                ctx.font = `${fontSize}px "Architects Daughter", cursive`;
                ctx.fillStyle = color;

                // Position text: to the right or left of endpoint depending on direction
                const textLeft = end.x > tip.x;
                ctx.textAlign = textLeft ? 'left' : 'right';
                ctx.textBaseline = 'bottom';
                const textX = textLeft ? end.x + 3 : end.x - 3;
                ctx.fillText(el.text, textX, end.y - 2);

                // Underline
                const metrics = ctx.measureText(el.text);
                const ulX1 = textLeft ? end.x : end.x - metrics.width - 3;
                const ulX2 = textLeft ? end.x + metrics.width + 3 : end.x;
                ctx.strokeStyle = color;
                ctx.lineWidth = Math.max(0.5, 0.18 * zoom);
                ctx.beginPath();
                ctx.moveTo(ulX1, end.y);
                ctx.lineTo(ulX2, end.y);
                ctx.stroke();
            }

            // Selection handles
            if (isSelected) {
                ctx.fillStyle = '#2B7CD0';
                ctx.fillRect(tip.x - 3, tip.y - 3, 6, 6);
                ctx.fillRect(end.x - 3, end.y - 3, 6, 6);
            }
        }

        // ── Revision clouds ──
        if (el.type === 'cloud') {
            const layer = project.layers[el.layer];
            if (!layer || !layer.visible) continue;

            const pts = el.points;
            if (pts.length < 3) continue;
            const color = isSelected ? '#2B7CD0' : '#CC0000'; // clouds typically red

            ctx.strokeStyle = color;
            ctx.lineWidth = Math.max(1.5, 0.35 * zoom);
            ctx.setLineDash([]);

            // Draw cloud bumps: arcs between consecutive points
            const bumpSize = 3 * zoom; // radius of each arc bump
            ctx.beginPath();
            for (let i = 0; i < pts.length - 1; i++) {
                const p1 = coords.realToScreen(pts[i].x, pts[i].y);
                const p2 = coords.realToScreen(pts[i + 1].x, pts[i + 1].y);

                const segDx = p2.x - p1.x;
                const segDy = p2.y - p1.y;
                const segLen = Math.sqrt(segDx * segDx + segDy * segDy);
                if (segLen < 2) continue;

                // Number of bumps for this segment
                const numBumps = Math.max(1, Math.round(segLen / (bumpSize * 1.8)));
                const stepX = segDx / numBumps;
                const stepY = segDy / numBumps;

                for (let b = 0; b < numBumps; b++) {
                    const bx1 = p1.x + stepX * b;
                    const by1 = p1.y + stepY * b;
                    const bx2 = p1.x + stepX * (b + 1);
                    const by2 = p1.y + stepY * (b + 1);
                    const mx = (bx1 + bx2) / 2;
                    const my = (by1 + by2) / 2;

                    // Arc bulging outward (perpendicular to segment direction)
                    const perpX = -stepY / numBumps;
                    const perpY = stepX / numBumps;
                    const perpLen = Math.sqrt(perpX * perpX + perpY * perpY);
                    const bulge = bumpSize * 0.6;
                    const cpx = mx + (perpX / perpLen) * bulge;
                    const cpy = my + (perpY / perpLen) * bulge;

                    if (b === 0 && i === 0) ctx.moveTo(bx1, by1);
                    ctx.quadraticCurveTo(cpx, cpy, bx2, by2);
                }
            }
            ctx.stroke();

            if (isSelected) {
                ctx.strokeStyle = '#2B7CD0';
                ctx.lineWidth = 1;
                ctx.setLineDash([4, 3]);
                // Bounding box
                let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
                for (const pt of pts) {
                    const sp = coords.realToScreen(pt.x, pt.y);
                    minX = Math.min(minX, sp.x); minY = Math.min(minY, sp.y);
                    maxX = Math.max(maxX, sp.x); maxY = Math.max(maxY, sp.y);
                }
                ctx.strokeRect(minX - 4, minY - 4, maxX - minX + 8, maxY - minY + 8);
                ctx.setLineDash([]);
            }
        }
    }

    // Leader preview
    if (activeTool === 'leader' && leaderState.placing && leaderState.startPoint && leaderState.currentEnd) {
        const tip = coords.sheetToScreen(leaderState.startPoint.x, leaderState.startPoint.y);
        const end = coords.sheetToScreen(leaderState.currentEnd.x, leaderState.currentEnd.y);

        ctx.strokeStyle = '#2B7CD0';
        ctx.lineWidth = 1.5;
        ctx.globalAlpha = 0.6;
        ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(tip.x, tip.y);
        ctx.lineTo(end.x, end.y);
        ctx.stroke();

        // Preview arrowhead
        const angle = Math.atan2(tip.y - end.y, tip.x - end.x);
        const arrowLen = 8;
        ctx.fillStyle = '#2B7CD0';
        ctx.beginPath();
        ctx.moveTo(tip.x, tip.y);
        ctx.lineTo(tip.x - arrowLen * Math.cos(angle - 0.4), tip.y - arrowLen * Math.sin(angle - 0.4));
        ctx.lineTo(tip.x - arrowLen * Math.cos(angle + 0.4), tip.y - arrowLen * Math.sin(angle + 0.4));
        ctx.closePath();
        ctx.fill();
        ctx.globalAlpha = 1.0;

        ctx.fillStyle = '#2B7CD0';
        ctx.beginPath(); ctx.arc(tip.x, tip.y, 3, 0, Math.PI * 2); ctx.fill();
    }

    // Cloud preview while dragging rectangle
    if (activeTool === 'cloud' && cloudState.drawing && cloudState.startPoint && cloudState.currentEnd) {
        const s = cloudState.startPoint;
        const e = cloudState.currentEnd;
        const sp1 = coords.sheetToScreen(s.x, s.y);
        const sp2 = coords.sheetToScreen(e.x, e.y);

        // Draw cloud bumps along rectangle edges
        const x1 = Math.min(sp1.x, sp2.x), y1 = Math.min(sp1.y, sp2.y);
        const x2 = Math.max(sp1.x, sp2.x), y2 = Math.max(sp1.y, sp2.y);
        const corners = [{x:x1,y:y1},{x:x2,y:y1},{x:x2,y:y2},{x:x1,y:y2},{x:x1,y:y1}];

        ctx.strokeStyle = '#CC0000';
        ctx.lineWidth = 2;
        ctx.globalAlpha = 0.5;
        ctx.setLineDash([]);

        const bumpR = 3 * zoom;
        ctx.beginPath();
        for (let ci = 0; ci < corners.length - 1; ci++) {
            const cp1 = corners[ci], cp2 = corners[ci+1];
            const sdx = cp2.x - cp1.x, sdy = cp2.y - cp1.y;
            const sLen = Math.sqrt(sdx*sdx + sdy*sdy);
            if (sLen < 2) continue;
            const nBumps = Math.max(1, Math.round(sLen / (bumpR * 1.8)));
            const stX = sdx/nBumps, stY = sdy/nBumps;
            for (let b = 0; b < nBumps; b++) {
                const bx1 = cp1.x + stX*b, by1 = cp1.y + stY*b;
                const bx2 = cp1.x + stX*(b+1), by2 = cp1.y + stY*(b+1);
                const mx = (bx1+bx2)/2, my = (by1+by2)/2;
                const pxDir = -stY/nBumps, pyDir = stX/nBumps;
                const pLen = Math.sqrt(pxDir*pxDir + pyDir*pyDir) || 1;
                const cpx = mx + (pxDir/pLen) * bumpR * 0.6;
                const cpy = my + (pyDir/pLen) * bumpR * 0.6;
                if (b === 0 && ci === 0) ctx.moveTo(bx1, by1);
                ctx.quadraticCurveTo(cpx, cpy, bx2, by2);
            }
        }
        ctx.stroke();
        ctx.globalAlpha = 1.0;

        // Show dimensions
        const r1 = coords.sheetToReal(s.x, s.y);
        const r2 = coords.sheetToReal(e.x, e.y);
        const cw = Math.abs(r2.x - r1.x), ch = Math.abs(r2.y - r1.y);
        const fmtCD = (mm) => mm >= 1000 ? (mm/1000).toFixed(1) + 'm' : Math.round(mm) + '';
        ctx.font = 'bold 11px "Segoe UI", Arial, sans-serif';
        ctx.fillStyle = '#CC0000';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(fmtCD(cw), (sp1.x+sp2.x)/2, Math.min(sp1.y,sp2.y) - 4);
        ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
        ctx.fillText(fmtCD(ch), Math.max(sp1.x,sp2.x) + 6, (sp1.y+sp2.y)/2);
        ctx.textAlign = 'left'; ctx.textBaseline = 'top';
    }

    // Polygon cloud preview
    if (activeTool === 'cloudpoly' && cloudPolyState.points.length > 0) {
        const bumpR = 3 * zoom;
        const allPts = [...cloudPolyState.points];
        if (cloudPolyState.currentEnd) allPts.push(cloudPolyState.currentEnd);

        // Draw placed edges as cloud bumps
        ctx.strokeStyle = '#CC0000';
        ctx.lineWidth = 2;
        ctx.globalAlpha = 0.5;
        ctx.setLineDash([]);
        ctx.beginPath();
        let first = true;
        for (let ci = 0; ci < allPts.length - 1; ci++) {
            const cp1s = coords.sheetToScreen(allPts[ci].x, allPts[ci].y);
            const cp2s = coords.sheetToScreen(allPts[ci+1].x, allPts[ci+1].y);
            const sdx = cp2s.x - cp1s.x, sdy = cp2s.y - cp1s.y;
            const sLen = Math.sqrt(sdx*sdx + sdy*sdy);
            if (sLen < 2) continue;
            const nBumps = Math.max(1, Math.round(sLen / (bumpR * 1.8)));
            const stX = sdx/nBumps, stY = sdy/nBumps;
            for (let b = 0; b < nBumps; b++) {
                const bx1 = cp1s.x + stX*b, by1 = cp1s.y + stY*b;
                const bx2 = cp1s.x + stX*(b+1), by2 = cp1s.y + stY*(b+1);
                const mx = (bx1+bx2)/2, my = (by1+by2)/2;
                const pxD = -stY/nBumps, pyD = stX/nBumps;
                const pLen = Math.sqrt(pxD*pxD + pyD*pyD) || 1;
                const cpx = mx + (pxD/pLen) * bumpR * 0.6;
                const cpy = my + (pyD/pLen) * bumpR * 0.6;
                if (first) { ctx.moveTo(bx1, by1); first = false; }
                ctx.quadraticCurveTo(cpx, cpy, bx2, by2);
            }
        }
        // Dashed closing line back to start
        if (cloudPolyState.points.length >= 3 && cloudPolyState.currentEnd) {
            const lastPt = coords.sheetToScreen(cloudPolyState.currentEnd.x, cloudPolyState.currentEnd.y);
            const firstPt = coords.sheetToScreen(cloudPolyState.points[0].x, cloudPolyState.points[0].y);
            ctx.stroke(); // finish the bumps
            ctx.setLineDash([4, 3]);
            ctx.beginPath();
            ctx.moveTo(lastPt.x, lastPt.y);
            ctx.lineTo(firstPt.x, firstPt.y);
            ctx.stroke();
            ctx.setLineDash([]);
        } else {
            ctx.stroke();
        }
        ctx.globalAlpha = 1.0;

        // Vertex dots
        ctx.fillStyle = '#CC0000';
        for (const pt of cloudPolyState.points) {
            const sp = coords.sheetToScreen(pt.x, pt.y);
            ctx.beginPath(); ctx.arc(sp.x, sp.y, 3, 0, Math.PI * 2); ctx.fill();
        }
    }
};

// Replace in render callbacks
const cbIdx3 = engine._renderCallbacks.indexOf(extDrawElements);
if (cbIdx3 !== -1) engine._renderCallbacks[cbIdx3] = phase9Draw;

// ── Leader + Cloud drag support ──────────────────────────

// Leaders use x1,y1,x2,y2 like lines — already handled by existing drag.
// Clouds use points array like polylines — handled by polyline drag patch.
// Just need to ensure the drag start stores the right coords:
const origDragDown2 = container._origDragDown;
// Patch: in the mousedown for select, polyline origCoords is already handled
// for clouds since they have .points like polylines. Leaders have x1/y1/x2/y2
// like lines, so they're already handled. Nothing extra needed.

// ── Leader + Cloud PDF export ────────────────────────────
// Patch into the existing PDF export element loop

// We need to hook the export. The simplest approach is to note that
// the export already loops project.getVisibleElements() — we just need
// to add handlers for 'leader' and 'cloud' types there.

// Find and patch the export click handler's element loop:
const origExportBtn = document.getElementById('btn-export-pdf');
const origExportHandler = origExportBtn._origHandler;
// Since we can't easily patch inline, leaders and clouds will be exported
// in a future update. For now they render on screen and save/load in JSON.

// ══════════════════════════════════════════════════════════
// ── PHASE 10: HATCH/FILL PATTERNS ────────────────────────
// ══════════════════════════════════════════════════════════

// ── Hatch Pattern Definitions ────────────────────────────

/**
 * Draw a hatch pattern inside a polygon defined by screen-space points.
 * Each pattern function takes (ctx, points[], zoom, color, opacity).
 */
const HATCH_RENDERERS = {

    /** Concrete: Australian standard - scattered triangles */
    concrete(ctx, screenPts, zoom, color, opacity) {
        ctx.save();
        clipToPolygon(ctx, screenPts);
        ctx.globalAlpha = opacity;
        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(0.5, 0.12 * zoom);

        const bounds = polyBounds(screenPts);
        const spacing = Math.max(6, 3.5 * zoom);
        const triSize = Math.max(1.5, 0.8 * zoom);

        for (let x = bounds.x; x < bounds.x + bounds.w; x += spacing) {
            for (let y = bounds.y; y < bounds.y + bounds.h; y += spacing) {
                // Pseudo-random offset and rotation
                const ox = ((x * 7 + y * 13) % spacing) * 0.4 - spacing * 0.2;
                const oy = ((x * 11 + y * 3) % spacing) * 0.4 - spacing * 0.2;
                const angle = ((x * 17 + y * 23) % 628) / 100; // 0 to ~2PI
                const px = x + ox, py = y + oy;

                // Draw small triangle
                ctx.beginPath();
                ctx.moveTo(px + triSize * Math.cos(angle), py + triSize * Math.sin(angle));
                ctx.lineTo(px + triSize * Math.cos(angle + 2.1), py + triSize * Math.sin(angle + 2.1));
                ctx.lineTo(px + triSize * Math.cos(angle + 4.2), py + triSize * Math.sin(angle + 4.2));
                ctx.closePath();
                ctx.stroke();
            }
        }
        ctx.restore();
    },

    /** Blockwork: 45° cross-hatch (both directions) */
    blockwork(ctx, screenPts, zoom, color, opacity) {
        ctx.save();
        clipToPolygon(ctx, screenPts);
        ctx.globalAlpha = opacity;
        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(0.5, 0.15 * zoom);

        const bounds = polyBounds(screenPts);
        const spacing = Math.max(5, 3 * zoom);
        const diag = bounds.w + bounds.h;

        // 45° lines
        ctx.beginPath();
        for (let d = -diag; d < diag; d += spacing) {
            ctx.moveTo(bounds.x + d, bounds.y);
            ctx.lineTo(bounds.x + d + bounds.h, bounds.y + bounds.h);
        }
        // -45° lines
        for (let d = -diag; d < diag; d += spacing) {
            ctx.moveTo(bounds.x + d, bounds.y + bounds.h);
            ctx.lineTo(bounds.x + d + bounds.h, bounds.y);
        }
        ctx.stroke();
        ctx.restore();
    },

    /** Timber: wavy grain lines at ~0° (horizontal) */
    timber(ctx, screenPts, zoom, color, opacity) {
        ctx.save();
        clipToPolygon(ctx, screenPts);
        ctx.globalAlpha = opacity;
        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(0.5, 0.15 * zoom);

        const bounds = polyBounds(screenPts);
        const spacing = Math.max(4, 2 * zoom);
        const waveAmp = spacing * 0.25;
        const waveFreq = 0.03 / Math.max(0.3, zoom * 0.3);

        ctx.beginPath();
        for (let y = bounds.y; y < bounds.y + bounds.h; y += spacing) {
            ctx.moveTo(bounds.x, y);
            for (let x = bounds.x; x < bounds.x + bounds.w; x += 2) {
                const wy = y + Math.sin((x + y * 0.5) * waveFreq) * waveAmp;
                ctx.lineTo(x, wy);
            }
        }
        ctx.stroke();
        ctx.restore();
    },

    /** Solid fill */
    solid(ctx, screenPts, zoom, color, opacity) {
        ctx.save();
        ctx.globalAlpha = opacity;
        ctx.fillStyle = color;
        ctx.beginPath();
        ctx.moveTo(screenPts[0].x, screenPts[0].y);
        for (let i = 1; i < screenPts.length; i++) {
            ctx.lineTo(screenPts[i].x, screenPts[i].y);
        }
        ctx.closePath();
        ctx.fill();
        ctx.restore();
    }
};

/** Clip canvas to a polygon defined by screen-space points */
function clipToPolygon(ctx, pts) {
    ctx.beginPath();
    ctx.moveTo(pts[0].x, pts[0].y);
    for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
    ctx.closePath();
    ctx.clip();
}

/** Get bounding box of screen-space points */
function polyBounds(pts) {
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    for (const p of pts) {
        if (p.x < minX) minX = p.x; if (p.y < minY) minY = p.y;
        if (p.x > maxX) maxX = p.x; if (p.y > maxY) maxY = p.y;
    }
    return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
}

/** Calculate centroid of polygon points (in same coordinate system as input) */
function polyCentroid(pts) {
    if (!pts || pts.length === 0) return { x: 0, y: 0 };
    let cx = 0, cy = 0;
    for (const p of pts) {
        cx += p.x;
        cy += p.y;
    }
    return { x: cx / pts.length, y: cy / pts.length };
}

// ── Hatch/Fill Rendering ─────────────────────────────────

function drawHatchFills(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    for (const el of project.getVisibleElements()) {
        if (!el.hatch || el.hatch === 'none') continue;
        if (el.type !== 'polyline') continue;
        if (!el.points || el.points.length < 3) continue;

        const renderer = HATCH_RENDERERS[el.hatch];
        if (!renderer) continue;

        // Convert to screen points
        const screenPts = el.points.map(p => coords.realToScreen(p.x, p.y));
        const color = el.fillColor || '#CCCCCC';
        const opacity = el.fillOpacity !== undefined ? el.fillOpacity : 0.3;

        renderer(ctx, screenPts, zoom, color, opacity);
    }
}

/** Draw thickness callouts for slab elements */
function drawSlabCallouts(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    for (const el of project.getVisibleElements()) {
        // Only draw callouts for slab polylines with thickness
        if (el.type !== 'polyline' || el.layer !== 'S-SLAB' || !el.slabThickness) continue;
        if (!el.points || el.points.length < 3) continue;

        // Calculate centroid in real coords
        const centroid = polyCentroid(el.points);
        const screenCentroid = coords.realToScreen(centroid.x, centroid.y);

        // Callout — clean square box with bold italic number, no brackets
        const thickness = el.slabThickness;
        const text = String(thickness);
        const fontSize = Math.max(10, 5 * zoom);
        ctx.font = `bold italic ${fontSize}px "Segoe UI", Arial, sans-serif`;
        const metrics = ctx.measureText(text);
        const textWidth = metrics.width;

        // Square box sized to fit text with generous padding
        const boxSize = Math.max(textWidth, fontSize) + Math.max(6, 3.5 * zoom);

        ctx.save();

        // White background — clean square
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(screenCentroid.x - boxSize / 2, screenCentroid.y - boxSize / 2, boxSize, boxSize);

        // Thin black border
        ctx.strokeStyle = '#000000';
        ctx.lineWidth = Math.max(0.5, 0.18 * zoom);
        ctx.strokeRect(screenCentroid.x - boxSize / 2, screenCentroid.y - boxSize / 2, boxSize, boxSize);

        // Number text — bold italic, centred
        ctx.fillStyle = '#000000';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(text, screenCentroid.x, screenCentroid.y);

        ctx.restore();
    }
}

// Register hatch renderer — draw BEFORE elements (so outlines are on top)
// Insert at position 3 (after pdf, statusbar, grids but before elements)
engine._renderCallbacks.splice(3, 0, drawHatchFills);

// Register slab callout renderer — draw AFTER elements (on top)
engine._renderCallbacks.push(drawSlabCallouts);

// ── Hatch/Fill UI Controls ───────────────────────────────

const hatchSelect = document.getElementById('hatch-pattern');
const fillColorInput = document.getElementById('fill-color');
const fillOpacitySlider = document.getElementById('fill-opacity');
const fillOpacityVal = document.getElementById('fill-opacity-val');

// Apply hatch to selected element (polyline with 3+ points)
hatchSelect.addEventListener('change', () => {
    if (!selectedElement || selectedElement.type !== 'polyline' || !selectedElement.points || selectedElement.points.length < 3) return;
    const el = selectedElement;
    const oldHatch = el.hatch;
    const newHatch = hatchSelect.value;
    const color = fillColorInput.value;
    const opacity = parseInt(fillOpacitySlider.value) / 100;

    history.execute({
        description: 'Set hatch: ' + newHatch,
        execute() {
            el.hatch = newHatch;
            el.fillColor = color;
            el.fillOpacity = opacity;
        },
        undo() {
            el.hatch = oldHatch;
            if (!oldHatch || oldHatch === 'none') {
                delete el.fillColor;
                delete el.fillOpacity;
            }
        }
    });
    engine.requestRender();
});

// Update fill colour on selected element
fillColorInput.addEventListener('input', () => {
    if (!selectedElement || selectedElement.type !== 'polyline' || !selectedElement.points || selectedElement.points.length < 3) return;
    if (!selectedElement.hatch || selectedElement.hatch === 'none') return;
    selectedElement.fillColor = fillColorInput.value;
    engine.requestRender();
});

// Update fill opacity
fillOpacitySlider.addEventListener('input', () => {
    fillOpacityVal.textContent = fillOpacitySlider.value + '%';
    if (!selectedElement || selectedElement.type !== 'polyline' || !selectedElement.points || selectedElement.points.length < 3) return;
    if (!selectedElement.hatch || selectedElement.hatch === 'none') return;
    selectedElement.fillOpacity = parseInt(fillOpacitySlider.value) / 100;
    engine.requestRender();
});

// Sync UI when selecting a polyline with hatch
engine.onRender(() => {
    if (selectedElement && selectedElement.type === 'polyline' && selectedElement.hatch) {
        hatchSelect.value = selectedElement.hatch;
        if (selectedElement.fillColor) fillColorInput.value = selectedElement.fillColor;
        if (selectedElement.fillOpacity !== undefined) {
            fillOpacitySlider.value = Math.round(selectedElement.fillOpacity * 100);
            fillOpacityVal.textContent = Math.round(selectedElement.fillOpacity * 100) + '%';
        }
    }
});

// ══════════════════════════════════════════════════════════
// ── Callout Tool (T) ─────────────────────────────────────
// Click arrow tip → click text box position → type label → Enter
// Renders: bordered text box with word-wrapped text + leader line
//          with filled arrowhead. Double-click to edit/resize.
// Drag behaviour (Bluebeam-style):
//   - Drag body = move text box only, arrow tip stays fixed
//   - Drag arrow tip handle = re-aim arrow, text box stays fixed
// ══════════════════════════════════════════════════════════

const CALLOUT_DEFAULTS = {
    boxWidth: 30,       // sheet-mm — default text box width
    minBoxWidth: 10,    // sheet-mm — minimum resize width
    minBoxHeight: 6,    // sheet-mm — minimum resize height
    padding: 2,         // sheet-mm — inner padding
    fontSize: 5,        // sheet-mm (was 3.5, bumped for readability)
};

const calloutState = {
    phase: 0,           // 0=idle, 1=placed arrow tip, 2=dragging box
    arrowPt: null,      // arrow tip, sheet-mm (set on first click)
    boxStart: null,     // box top-left, sheet-mm (set on second mousedown)
    boxEnd: null,       // box bottom-right, sheet-mm (updated on mousemove)
    currentEnd: null,   // preview endpoint while in phase 1
};
let activeCalloutInput = null;

// ── Callout edit mode (double-click) ─────────────────────
let calloutEditState = {
    editing: false,
    el: null,           // element being edited
    textarea: null,     // DOM textarea overlay
    resizing: false,    // true while dragging a resize handle
    resizeCorner: null, // which corner is being dragged
    resizeStart: null,  // { x, y } screen coords at drag start
    origBoxWidth: 0,    // original boxWidth before resize
    origBoxHeight: 0,   // original boxHeight before resize
};

function getCalloutPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    return snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
}

// ── Helper: wrap text into lines that fit a given pixel width ──
function wrapText(ctx, text, maxWidth) {
    const lines = [];
    const paragraphs = text.split('\n');
    for (const para of paragraphs) {
        const words = para.split(/\s+/);
        if (words.length === 0 || (words.length === 1 && words[0] === '')) {
            lines.push('');
            continue;
        }
        let currentLine = words[0];
        for (let i = 1; i < words.length; i++) {
            const test = currentLine + ' ' + words[i];
            if (ctx.measureText(test).width <= maxWidth) {
                currentLine = test;
            } else {
                lines.push(currentLine);
                currentLine = words[i];
            }
        }
        lines.push(currentLine);
    }
    return lines;
}

// ── Helper: compute callout box geometry in screen coords ──
function getCalloutBoxGeom(el, coords, zoom, ctx) {
    const end = coords.realToScreen(el.x2, el.y2);
    const tip = coords.realToScreen(el.x1, el.y1);
    const fontSize = (el.fontSize || CALLOUT_DEFAULTS.fontSize) * zoom;
    const pad = (el.padding || CALLOUT_DEFAULTS.padding) * zoom;
    const boxWidthMM = el.boxWidth || CALLOUT_DEFAULTS.boxWidth;
    const boxW = boxWidthMM * zoom;

    ctx.font = `${el.fontBold ? 'bold ' : ''}${fontSize}px "Architects Daughter", cursive`;
    const innerW = boxW - pad * 2;
    const lines = wrapText(ctx, el.text || '', innerW > 10 ? innerW : 10);
    const lineHeight = fontSize * 1.3;

    // If boxHeight is fixed (user resized), use it; otherwise auto-height
    let boxH;
    if (el.boxHeight) {
        boxH = el.boxHeight * zoom;
    } else {
        boxH = Math.max((CALLOUT_DEFAULTS.minBoxHeight) * zoom, lines.length * lineHeight + pad * 2);
    }

    // Anchor: text box sits with its nearest edge touching the endpoint
    const textLeft = end.x > tip.x;
    const boxX = textLeft ? end.x : end.x - boxW;
    const boxY = end.y - boxH / 2;

    // Smart leader landing: connect to nearest box edge midpoint
    const boxCX = boxX + boxW / 2;
    const boxCY = boxY + boxH / 2;
    const candidates = [
        { x: boxX, y: boxCY },           // left edge mid
        { x: boxX + boxW, y: boxCY },    // right edge mid
        { x: boxCX, y: boxY },           // top edge mid
        { x: boxCX, y: boxY + boxH },    // bottom edge mid
    ];
    let leaderEnd = end;
    let bestDist = Infinity;
    for (const c of candidates) {
        const d = Math.sqrt(Math.pow(tip.x - c.x, 2) + Math.pow(tip.y - c.y, 2));
        if (d < bestDist) { bestDist = d; leaderEnd = c; }
    }

    return { boxX, boxY, boxW, boxH, pad, fontSize, lines, lineHeight, leaderEnd, textLeft };
}

// Mousemove — update preview (phase 1: leader preview, phase 2: box drag)
container.addEventListener('mousemove', (e) => {
    if (activeTool !== 'callout') return;
    if (calloutState.phase === 1) {
        calloutState.currentEnd = getCalloutPos(e);
        engine.requestRender();
    } else if (calloutState.phase === 2) {
        calloutState.boxEnd = getCalloutPos(e);
        engine.requestRender();
    }
});

// Mousedown — phase 0: set arrow tip, phase 1: start box drag
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'callout') return;
    if (pdfState.calibrating) return;
    if (activeCalloutInput) return;

    const pos = getCalloutPos(e);

    if (calloutState.phase === 0) {
        // First click — set arrow tip
        calloutState.phase = 1;
        calloutState.arrowPt = pos;
    } else if (calloutState.phase === 1) {
        // Second mousedown — start dragging box from this corner
        calloutState.phase = 2;
        calloutState.boxStart = pos;
        calloutState.boxEnd = pos;
    }
});

// Mouseup — if in phase 2, finish box drag and show textarea
window.addEventListener('mouseup', (e) => {
    if (activeTool !== 'callout' || calloutState.phase !== 2) return;
    if (activeCalloutInput) return;

    const boxStart = calloutState.boxStart;
    const boxEnd = calloutState.boxEnd || boxStart;

    // Calculate box bounds in sheet-mm
    const x1 = Math.min(boxStart.x, boxEnd.x);
    const y1 = Math.min(boxStart.y, boxEnd.y);
    const x2 = Math.max(boxStart.x, boxEnd.x);
    const y2 = Math.max(boxStart.y, boxEnd.y);
    const draggedW = x2 - x1;
    const draggedH = y2 - y1;

    // Use dragged size if large enough, otherwise use defaults
    const boxWidthMM = draggedW > 3 ? draggedW : CALLOUT_DEFAULTS.boxWidth;
    const boxHeightMM = draggedH > 3 ? draggedH : null;

    // Anchor point for the callout is the top-left of the box
    const anchorPt = { x: x1, y: y1 };
    const screenTL = engine.coords.sheetToScreen(x1, y1);
    const screenBR = engine.coords.sheetToScreen(x2, y2);

    const ta = document.createElement('textarea');
    ta.className = 'text-input-overlay';
    ta.style.left = screenTL.x + 'px';
    ta.style.top = screenTL.y + 'px';
    ta.style.width = Math.max(100, screenBR.x - screenTL.x) + 'px';
    ta.style.height = Math.max(40, screenBR.y - screenTL.y) + 'px';
    ta.style.fontSize = (CALLOUT_DEFAULTS.fontSize * engine.viewport.zoom) + 'px';
    ta.style.resize = 'none';
    ta.style.overflow = 'hidden';
    ta.style.lineHeight = '1.3';
    ta.placeholder = 'Callout text...';
    ta.style.border = '1px solid #2B7CD0';
    ta.style.background = 'rgba(255,255,255,0.95)';
    ta.style.fontFamily = '"Architects Daughter", cursive';
    ta.style.padding = (CALLOUT_DEFAULTS.padding * engine.viewport.zoom) + 'px';
    ta.style.boxSizing = 'border-box';
    ta._arrowPt = { x: calloutState.arrowPt.x, y: calloutState.arrowPt.y };
    ta._anchorPt = anchorPt;
    ta._boxWidthMM = boxWidthMM;
    ta._boxHeightMM = boxHeightMM;

    container.appendChild(ta);
    setTimeout(() => ta.focus(), 30);
    activeCalloutInput = ta;

    ta.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter' && !ev.shiftKey) { ev.preventDefault(); commitCallout(); }
        if (ev.key === 'Escape') { ev.preventDefault(); cancelCallout(); }
    });
    ta.addEventListener('blur', () => {
        setTimeout(() => { if (activeCalloutInput === ta) commitCallout(); }, 150);
    });

    calloutState.phase = 0;
    calloutState.arrowPt = null;
    calloutState.boxStart = null;
    calloutState.boxEnd = null;
});

// Right-click cancel
container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'callout' && calloutState.phase > 0) {
        e.preventDefault();
        calloutState.phase = 0;
        calloutState.arrowPt = null;
        calloutState.boxStart = null;
        calloutState.boxEnd = null;
        engine.requestRender();
    }
});

function commitCallout() {
    if (!activeCalloutInput) return;
    const input = activeCalloutInput;
    const text = input.value.trim();
    activeCalloutInput = null;
    if (input.parentNode) input.parentNode.removeChild(input);

    if (!text) { engine.requestRender(); return; }

    const arrowReal = engine.coords.sheetToReal(input._arrowPt.x, input._arrowPt.y);
    const anchorReal = engine.coords.sheetToReal(input._anchorPt.x, input._anchorPt.y);
    const textSizeSelect = document.getElementById('text-size');
    const fontSize = textSizeSelect ? parseFloat(textSizeSelect.value) : CALLOUT_DEFAULTS.fontSize;

    const newCallout = {
        id: generateId(),
        type: 'callout',
        layer: 'S-ANNO',
        x1: arrowReal.x, y1: arrowReal.y,  // arrow tip
        x2: anchorReal.x, y2: anchorReal.y, // text box anchor (top-left)
        text: text,
        fontSize: fontSize,
        fontBold: false,
        boxWidth: input._boxWidthMM || CALLOUT_DEFAULTS.boxWidth,
        boxHeight: input._boxHeightMM || null,
        padding: CALLOUT_DEFAULTS.padding,
    };

    history.execute({
        description: 'Add callout: ' + text,
        execute() { project.elements.push(newCallout); },
        undo() {
            const i = project.elements.indexOf(newCallout);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });
    engine.requestRender();
}

function cancelCallout() {
    if (!activeCalloutInput) return;
    const input = activeCalloutInput;
    activeCalloutInput = null;
    if (input.parentNode) input.parentNode.removeChild(input);
    engine.requestRender();
}

// ── Double-click to enter callout edit mode ──────────────
container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'select') return;
    const sheetPos = engine.getSheetPos(e);
    const tolerance = 4 / engine.viewport.zoom;
    const hit = hitTestElement(sheetPos, tolerance);
    if (!hit || hit.type !== 'callout') return;

    // Enter edit mode
    calloutEditState.editing = true;
    calloutEditState.el = hit;
    selectedElement = hit;

    const geom = getCalloutBoxGeom(hit, engine.coords, engine.viewport.zoom, engine.ctx);

    // Create textarea overlay positioned over the box
    const ta = document.createElement('textarea');
    ta.className = 'text-input-overlay';
    ta.style.position = 'absolute';
    ta.style.left = geom.boxX + 'px';
    ta.style.top = geom.boxY + 'px';
    ta.style.width = geom.boxW + 'px';
    ta.style.height = geom.boxH + 'px';
    ta.style.fontSize = geom.fontSize + 'px';
    ta.style.fontFamily = '"Architects Daughter", cursive';
    ta.style.fontWeight = hit.fontBold ? 'bold' : 'normal';
    ta.style.lineHeight = '1.3';
    ta.style.padding = geom.pad + 'px';
    ta.style.border = '2px solid #2B7CD0';
    ta.style.background = 'rgba(255,255,255,0.97)';
    ta.style.resize = 'none';
    ta.style.overflow = 'hidden';
    ta.style.boxSizing = 'border-box';
    ta.style.outline = 'none';
    ta.style.zIndex = '100';
    ta.value = hit.text;

    container.appendChild(ta);
    calloutEditState.textarea = ta;
    setTimeout(() => { ta.focus(); ta.select(); }, 30);

    // Prevent tool shortcuts while editing
    ta.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Escape') { ev.preventDefault(); exitCalloutEdit(true); }
    });
    ta.addEventListener('blur', () => {
        setTimeout(() => { if (calloutEditState.editing && calloutEditState.textarea === ta) exitCalloutEdit(true); }, 200);
    });

    engine.requestRender();
});

function exitCalloutEdit(commit) {
    if (!calloutEditState.editing) return;
    const el = calloutEditState.el;
    const ta = calloutEditState.textarea;

    if (commit && ta && el) {
        const newText = ta.value.trim();
        if (newText && newText !== el.text) {
            const oldText = el.text;
            history.execute({
                description: 'Edit callout text',
                execute() { el.text = newText; },
                undo() { el.text = oldText; }
            });
        }
        // Pick up font size from dropdown if changed during edit
        const textSizeSelect = document.getElementById('text-size');
        if (textSizeSelect) {
            const newSize = parseFloat(textSizeSelect.value);
            if (newSize && newSize !== el.fontSize) {
                const oldSize = el.fontSize;
                history.execute({
                    description: 'Change callout font size',
                    execute() { el.fontSize = newSize; },
                    undo() { el.fontSize = oldSize; }
                });
            }
        }
    }

    if (ta && ta.parentNode) ta.parentNode.removeChild(ta);
    calloutEditState.editing = false;
    calloutEditState.el = null;
    calloutEditState.textarea = null;
    calloutEditState.resizing = false;
    engine.requestRender();
}

// ── Resize handles: mousedown on corner starts resize ────
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (!calloutEditState.editing) return;
    const el = calloutEditState.el;
    if (!el) return;

    const geom = getCalloutBoxGeom(el, engine.coords, engine.viewport.zoom, engine.ctx);
    const mx = e.clientX - container.getBoundingClientRect().left;
    const my = e.clientY - container.getBoundingClientRect().top;
    const handleSize = 8;

    // Check corners: bottom-right and bottom-left
    const corners = [
        { name: 'br', x: geom.boxX + geom.boxW, y: geom.boxY + geom.boxH },
        { name: 'bl', x: geom.boxX, y: geom.boxY + geom.boxH },
        { name: 'tr', x: geom.boxX + geom.boxW, y: geom.boxY },
    ];

    for (const c of corners) {
        if (Math.abs(mx - c.x) < handleSize && Math.abs(my - c.y) < handleSize) {
            calloutEditState.resizing = true;
            calloutEditState.resizeCorner = c.name;
            calloutEditState.resizeStart = { x: mx, y: my };
            calloutEditState.origBoxWidth = el.boxWidth || CALLOUT_DEFAULTS.boxWidth;
            calloutEditState.origBoxHeight = el.boxHeight || (geom.boxH / engine.viewport.zoom);
            e.preventDefault();
            e.stopPropagation();
            return;
        }
    }
});

window.addEventListener('mousemove', (e) => {
    if (!calloutEditState.resizing) return;
    const el = calloutEditState.el;
    if (!el) return;

    const mx = e.clientX - container.getBoundingClientRect().left;
    const my = e.clientY - container.getBoundingClientRect().top;
    const zoom = engine.viewport.zoom;
    const dx = (mx - calloutEditState.resizeStart.x) / zoom;
    const dy = (my - calloutEditState.resizeStart.y) / zoom;

    const corner = calloutEditState.resizeCorner;
    if (corner === 'br' || corner === 'tr') {
        el.boxWidth = Math.max(CALLOUT_DEFAULTS.minBoxWidth, calloutEditState.origBoxWidth + dx);
    }
    if (corner === 'bl') {
        el.boxWidth = Math.max(CALLOUT_DEFAULTS.minBoxWidth, calloutEditState.origBoxWidth - dx);
    }
    if (corner === 'br' || corner === 'bl') {
        el.boxHeight = Math.max(CALLOUT_DEFAULTS.minBoxHeight, calloutEditState.origBoxHeight + dy);
    }
    if (corner === 'tr') {
        el.boxHeight = Math.max(CALLOUT_DEFAULTS.minBoxHeight, calloutEditState.origBoxHeight - dy);
    }

    // Update textarea overlay to match
    const geom = getCalloutBoxGeom(el, engine.coords, zoom, engine.ctx);
    const ta = calloutEditState.textarea;
    if (ta) {
        ta.style.left = geom.boxX + 'px';
        ta.style.top = geom.boxY + 'px';
        ta.style.width = geom.boxW + 'px';
        ta.style.height = geom.boxH + 'px';
    }

    engine.requestRender();
});

window.addEventListener('mouseup', (e) => {
    if (calloutEditState.resizing) {
        calloutEditState.resizing = false;
        calloutEditState.resizeCorner = null;
        engine.requestRender();
    }
});

// ── Callout Rendering + Preview ──────────────────────────

const prevPhase9Draw = phase9Draw;
const calloutDraw = function(ctx, eng) {
    prevPhase9Draw(ctx, eng);

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    // Render committed callout elements
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'callout') continue;

        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const isSelected = (selectedElement === el);
        const isEditing = (calloutEditState.editing && calloutEditState.el === el);
        const tip = coords.realToScreen(el.x1, el.y1);
        const color = isSelected ? '#2B7CD0' : layer.color;

        const geom = getCalloutBoxGeom(el, coords, zoom, ctx);

        // ── Leader line (tip to smart landing point on box edge) ──
        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(1, 0.25 * zoom);
        ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(tip.x, tip.y);
        ctx.lineTo(geom.leaderEnd.x, geom.leaderEnd.y);
        ctx.stroke();

        // ── Filled arrowhead at tip ──
        const angle = Math.atan2(tip.y - geom.leaderEnd.y, tip.x - geom.leaderEnd.x);
        const arrowLen = Math.max(6, 2.5 * zoom);
        const arrowW = Math.PI / 7;
        ctx.fillStyle = color;
        ctx.beginPath();
        ctx.moveTo(tip.x, tip.y);
        ctx.lineTo(tip.x - arrowLen * Math.cos(angle - arrowW), tip.y - arrowLen * Math.sin(angle - arrowW));
        ctx.lineTo(tip.x - arrowLen * Math.cos(angle + arrowW), tip.y - arrowLen * Math.sin(angle + arrowW));
        ctx.closePath();
        ctx.fill();

        // ── Text box (skip drawing text if in edit mode — textarea handles it) ──
        if (!isEditing) {
            // White fill
            ctx.fillStyle = 'rgba(255, 255, 255, 0.92)';
            ctx.fillRect(geom.boxX, geom.boxY, geom.boxW, geom.boxH);

            // Thin border
            ctx.strokeStyle = color;
            ctx.lineWidth = Math.max(0.5, 0.18 * zoom);
            ctx.setLineDash([]);
            ctx.strokeRect(geom.boxX, geom.boxY, geom.boxW, geom.boxH);

            // Word-wrapped text
            if (el.text) {
                ctx.fillStyle = color;
                ctx.font = `${el.fontBold ? 'bold ' : ''}${geom.fontSize}px "Architects Daughter", cursive`;
                ctx.textAlign = 'left';
                ctx.textBaseline = 'top';
                for (let i = 0; i < geom.lines.length; i++) {
                    ctx.fillText(geom.lines[i], geom.boxX + geom.pad, geom.boxY + geom.pad + i * geom.lineHeight);
                }
            }
        } else {
            // In edit mode: just draw the box border (textarea is on top)
            ctx.strokeStyle = '#2B7CD0';
            ctx.lineWidth = 2;
            ctx.setLineDash([]);
            ctx.strokeRect(geom.boxX, geom.boxY, geom.boxW, geom.boxH);
        }

        // ── Selection handles ──
        if (isSelected && !isEditing) {
            ctx.fillStyle = '#2B7CD0';
            // Arrow tip handle
            ctx.fillRect(tip.x - 3, tip.y - 3, 6, 6);
            // Text box anchor handle
            const end = coords.realToScreen(el.x2, el.y2);
            ctx.fillRect(end.x - 3, end.y - 3, 6, 6);
        }

        // ── Resize handles (in edit mode) ──
        if (isEditing) {
            const hs = 5; // handle half-size
            ctx.fillStyle = '#2B7CD0';
            // Bottom-right
            ctx.fillRect(geom.boxX + geom.boxW - hs, geom.boxY + geom.boxH - hs, hs * 2, hs * 2);
            // Bottom-left
            ctx.fillRect(geom.boxX - hs, geom.boxY + geom.boxH - hs, hs * 2, hs * 2);
            // Top-right
            ctx.fillRect(geom.boxX + geom.boxW - hs, geom.boxY - hs, hs * 2, hs * 2);
        }
    }

    // ── Callout preview while placing ──
    // Phase 1: arrow placed, showing leader preview to cursor
    if (activeTool === 'callout' && calloutState.phase === 1 && calloutState.arrowPt && calloutState.currentEnd) {
        const tip = coords.sheetToScreen(calloutState.arrowPt.x, calloutState.arrowPt.y);
        const end = coords.sheetToScreen(calloutState.currentEnd.x, calloutState.currentEnd.y);

        // Preview line
        ctx.strokeStyle = '#2B7CD0';
        ctx.lineWidth = 1.5;
        ctx.globalAlpha = 0.6;
        ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(tip.x, tip.y);
        ctx.lineTo(end.x, end.y);
        ctx.stroke();

        // Preview arrowhead
        const angle = Math.atan2(tip.y - end.y, tip.x - end.x);
        const arrowLen = 8;
        ctx.fillStyle = '#2B7CD0';
        ctx.beginPath();
        ctx.moveTo(tip.x, tip.y);
        ctx.lineTo(tip.x - arrowLen * Math.cos(angle - 0.4), tip.y - arrowLen * Math.sin(angle - 0.4));
        ctx.lineTo(tip.x - arrowLen * Math.cos(angle + 0.4), tip.y - arrowLen * Math.sin(angle + 0.4));
        ctx.closePath();
        ctx.fill();

        // Dot at arrow tip
        ctx.beginPath(); ctx.arc(tip.x, tip.y, 3, 0, Math.PI * 2); ctx.fill();
        ctx.globalAlpha = 1.0;
    }

    // Phase 2: dragging box — show leader + dashed box outline
    if (activeTool === 'callout' && calloutState.phase === 2 && calloutState.arrowPt && calloutState.boxStart && calloutState.boxEnd) {
        const tip = coords.sheetToScreen(calloutState.arrowPt.x, calloutState.arrowPt.y);
        const bs = calloutState.boxStart;
        const be = calloutState.boxEnd;
        const sx1 = coords.sheetToScreen(Math.min(bs.x, be.x), Math.min(bs.y, be.y));
        const sx2 = coords.sheetToScreen(Math.max(bs.x, be.x), Math.max(bs.y, be.y));
        const boxW = sx2.x - sx1.x;
        const boxH = sx2.y - sx1.y;
        const boxCX = sx1.x + boxW / 2;
        const boxCY = sx1.y + boxH / 2;

        // Preview leader line to nearest box edge
        ctx.strokeStyle = '#2B7CD0';
        ctx.lineWidth = 1.5;
        ctx.globalAlpha = 0.6;
        ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(tip.x, tip.y);
        ctx.lineTo(boxCX, boxCY);
        ctx.stroke();

        // Preview arrowhead
        const angle = Math.atan2(tip.y - boxCY, tip.x - boxCX);
        const arrowLen = 8;
        ctx.fillStyle = '#2B7CD0';
        ctx.beginPath();
        ctx.moveTo(tip.x, tip.y);
        ctx.lineTo(tip.x - arrowLen * Math.cos(angle - 0.4), tip.y - arrowLen * Math.sin(angle - 0.4));
        ctx.lineTo(tip.x - arrowLen * Math.cos(angle + 0.4), tip.y - arrowLen * Math.sin(angle + 0.4));
        ctx.closePath();
        ctx.fill();

        // Preview box outline (dashed)
        ctx.strokeStyle = '#2B7CD0';
        ctx.lineWidth = 1;
        ctx.setLineDash([4, 3]);
        ctx.strokeRect(sx1.x, sx1.y, boxW, boxH);
        ctx.setLineDash([]);

        // Dot at arrow tip
        ctx.beginPath(); ctx.arc(tip.x, tip.y, 3, 0, Math.PI * 2); ctx.fill();
        ctx.globalAlpha = 1.0;
    }
};

// Replace in render callbacks
const cbIdx_callout = engine._renderCallbacks.indexOf(phase9Draw);
if (cbIdx_callout !== -1) engine._renderCallbacks[cbIdx_callout] = calloutDraw;

// ── Callout hit-testing ──────────────────────────────────
const prevHitTest3 = hitTestElement;
hitTestElement = function(sheetPos) {
    const tolerance = 6 / engine.viewport.zoom;
    for (let i = project.elements.length - 1; i >= 0; i--) {
        const el = project.elements[i];
        if (el.type !== 'callout') continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const p1 = engine.coords.realToSheet(el.x1, el.y1);
        const p2 = engine.coords.realToSheet(el.x2, el.y2);

        // Hit on leader line
        if (pointToSegmentDist(sheetPos.x, sheetPos.y, p1.x, p1.y, p2.x, p2.y) < tolerance + 2)
            return el;

        // Hit on text box area (compute box bounds in sheet-mm)
        const zoom = engine.viewport.zoom;
        const geom = getCalloutBoxGeom(el, engine.coords, zoom, engine.ctx);
        // Convert screen box back to sheet coords
        const boxTL = engine.coords.screenToSheet(geom.boxX, geom.boxY);
        const boxBR = engine.coords.screenToSheet(geom.boxX + geom.boxW, geom.boxY + geom.boxH);

        if (sheetPos.x >= boxTL.x - tolerance && sheetPos.x <= boxBR.x + tolerance &&
            sheetPos.y >= boxTL.y - tolerance && sheetPos.y <= boxBR.y + tolerance)
            return el;
    }
    return prevHitTest3(sheetPos);
};

// ══════════════════════════════════════════════════════════
// ── Text Box Tool (T) ────────────────────────────────────
// Click to place → type text → Enter to commit
// Same bordered box as callout but NO leader line or arrow.
// Supports: drag, double-click edit, resize, word wrap.
// ══════════════════════════════════════════════════════════

let activeTextboxInput = null;
const textboxState = {
    phase: 0,        // 0=idle, 1=dragging box
    boxStart: null,  // top-left corner, sheet-mm
    boxEnd: null,    // bottom-right corner, sheet-mm
};

document.getElementById('btn-textbox').addEventListener('click', () => setActiveTool('textbox'));
document.getElementById('btn-callout').addEventListener('click', () => setActiveTool('callout'));

// ── Helper: compute textbox geometry (like callout but centered on anchor) ──
function getTextboxGeom(el, coords, zoom, ctx) {
    const anchor = coords.realToScreen(el.x, el.y);
    const fontSize = (el.fontSize || CALLOUT_DEFAULTS.fontSize) * zoom;
    const pad = (el.padding || CALLOUT_DEFAULTS.padding) * zoom;
    const boxWidthMM = el.boxWidth || CALLOUT_DEFAULTS.boxWidth;
    const boxW = boxWidthMM * zoom;

    ctx.font = `${el.fontBold ? 'bold ' : ''}${fontSize}px "Architects Daughter", cursive`;
    const innerW = boxW - pad * 2;
    const lines = wrapText(ctx, el.text || '', innerW > 10 ? innerW : 10);
    const lineHeight = fontSize * 1.3;

    let boxH;
    if (el.boxHeight) {
        boxH = el.boxHeight * zoom;
    } else {
        boxH = Math.max(CALLOUT_DEFAULTS.minBoxHeight * zoom, lines.length * lineHeight + pad * 2);
    }

    // Anchor is top-left of box
    const boxX = anchor.x;
    const boxY = anchor.y;

    return { boxX, boxY, boxW, boxH, pad, fontSize, lines, lineHeight };
}

// Helper to get sheet position for textbox tool
function getTextboxPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    return snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
}

// Mousedown — start dragging box from this corner
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'textbox') return;
    if (pdfState.calibrating) return;
    if (activeTextboxInput) return;

    const pos = getTextboxPos(e);
    textboxState.phase = 1;
    textboxState.boxStart = pos;
    textboxState.boxEnd = pos;
});

// Mousemove — update box preview while dragging
container.addEventListener('mousemove', (e) => {
    if (activeTool !== 'textbox' || textboxState.phase !== 1) return;
    textboxState.boxEnd = getTextboxPos(e);
    engine.requestRender();
});

// Mouseup — finish box drag and show textarea
window.addEventListener('mouseup', (e) => {
    if (activeTool !== 'textbox' || textboxState.phase !== 1) return;
    if (activeTextboxInput) return;

    const bs = textboxState.boxStart;
    const be = textboxState.boxEnd || bs;

    const x1 = Math.min(bs.x, be.x);
    const y1 = Math.min(bs.y, be.y);
    const x2 = Math.max(bs.x, be.x);
    const y2 = Math.max(bs.y, be.y);
    const draggedW = x2 - x1;
    const draggedH = y2 - y1;

    // Use dragged size if large enough, otherwise use defaults
    const boxWidthMM = draggedW > 3 ? draggedW : CALLOUT_DEFAULTS.boxWidth;
    const boxHeightMM = draggedH > 3 ? draggedH : null;

    const anchorPt = { x: x1, y: y1 };
    const screenTL = engine.coords.sheetToScreen(x1, y1);
    const screenBR = engine.coords.sheetToScreen(x2, y2);

    const ta = document.createElement('textarea');
    ta.className = 'text-input-overlay';
    ta.style.left = screenTL.x + 'px';
    ta.style.top = screenTL.y + 'px';
    ta.style.width = Math.max(100, screenBR.x - screenTL.x) + 'px';
    ta.style.height = Math.max(40, screenBR.y - screenTL.y) + 'px';
    ta.style.fontSize = (CALLOUT_DEFAULTS.fontSize * engine.viewport.zoom) + 'px';
    ta.style.resize = 'none';
    ta.style.overflow = 'hidden';
    ta.style.lineHeight = '1.3';
    ta.placeholder = 'Type text...';
    ta.style.border = '1px solid #2B7CD0';
    ta.style.background = 'rgba(255,255,255,0.95)';
    ta.style.fontFamily = '"Architects Daughter", cursive';
    ta.style.padding = (CALLOUT_DEFAULTS.padding * engine.viewport.zoom) + 'px';
    ta.style.boxSizing = 'border-box';
    ta._anchorPt = anchorPt;
    ta._boxWidthMM = boxWidthMM;
    ta._boxHeightMM = boxHeightMM;

    container.appendChild(ta);
    setTimeout(() => ta.focus(), 30);
    activeTextboxInput = ta;

    ta.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter' && !ev.shiftKey) { ev.preventDefault(); commitTextbox(); }
        if (ev.key === 'Escape') { ev.preventDefault(); cancelTextbox(); }
    });
    ta.addEventListener('blur', () => {
        setTimeout(() => { if (activeTextboxInput === ta) commitTextbox(); }, 150);
    });

    textboxState.phase = 0;
    textboxState.boxStart = null;
    textboxState.boxEnd = null;
});

// Right-click cancel for textbox
container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'textbox' && textboxState.phase > 0) {
        e.preventDefault();
        textboxState.phase = 0;
        textboxState.boxStart = null;
        textboxState.boxEnd = null;
        engine.requestRender();
    }
});

function commitTextbox() {
    if (!activeTextboxInput) return;
    const input = activeTextboxInput;
    const text = input.value.trim();
    activeTextboxInput = null;
    if (input.parentNode) input.parentNode.removeChild(input);

    if (!text) { engine.requestRender(); return; }

    const realPos = engine.coords.sheetToReal(input._anchorPt.x, input._anchorPt.y);
    const textSizeSelect = document.getElementById('text-size');
    const fontSize = textSizeSelect ? parseFloat(textSizeSelect.value) : CALLOUT_DEFAULTS.fontSize;

    const newTextbox = {
        id: generateId(),
        type: 'textbox',
        layer: 'S-ANNO',
        x: realPos.x, y: realPos.y,    // top-left anchor
        text: text,
        fontSize: fontSize,
        fontBold: false,
        boxWidth: input._boxWidthMM || CALLOUT_DEFAULTS.boxWidth,
        boxHeight: input._boxHeightMM || null,
        padding: CALLOUT_DEFAULTS.padding,
    };

    history.execute({
        description: 'Add text box: ' + text,
        execute() { project.elements.push(newTextbox); },
        undo() {
            const i = project.elements.indexOf(newTextbox);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });
    engine.requestRender();
}

function cancelTextbox() {
    if (!activeTextboxInput) return;
    const input = activeTextboxInput;
    activeTextboxInput = null;
    if (input.parentNode) input.parentNode.removeChild(input);
    engine.requestRender();
}

// ── Double-click to edit textbox ─────────────────────────
// (Extends the existing dblclick handler — check for textbox too)
container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'select') return;
    const sheetPos = engine.getSheetPos(e);
    const tolerance = 4 / engine.viewport.zoom;
    const hit = hitTestElement(sheetPos, tolerance);
    if (!hit || hit.type !== 'textbox') return;

    // Enter edit mode (reuse calloutEditState)
    calloutEditState.editing = true;
    calloutEditState.el = hit;
    selectedElement = hit;

    const geom = getTextboxGeom(hit, engine.coords, engine.viewport.zoom, engine.ctx);

    const ta = document.createElement('textarea');
    ta.className = 'text-input-overlay';
    ta.style.position = 'absolute';
    ta.style.left = geom.boxX + 'px';
    ta.style.top = geom.boxY + 'px';
    ta.style.width = geom.boxW + 'px';
    ta.style.height = geom.boxH + 'px';
    ta.style.fontSize = geom.fontSize + 'px';
    ta.style.fontFamily = '"Architects Daughter", cursive';
    ta.style.fontWeight = hit.fontBold ? 'bold' : 'normal';
    ta.style.lineHeight = '1.3';
    ta.style.padding = geom.pad + 'px';
    ta.style.border = '2px solid #2B7CD0';
    ta.style.background = 'rgba(255,255,255,0.97)';
    ta.style.resize = 'none';
    ta.style.overflow = 'hidden';
    ta.style.boxSizing = 'border-box';
    ta.style.outline = 'none';
    ta.style.zIndex = '100';
    ta.value = hit.text;

    container.appendChild(ta);
    calloutEditState.textarea = ta;
    setTimeout(() => { ta.focus(); ta.select(); }, 30);

    ta.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Escape') { ev.preventDefault(); exitCalloutEdit(true); }
    });
    ta.addEventListener('blur', () => {
        setTimeout(() => { if (calloutEditState.editing && calloutEditState.textarea === ta) exitCalloutEdit(true); }, 200);
    });

    engine.requestRender();
});

// ── Textbox Rendering (chain after calloutDraw) ──────────

const prevCalloutDraw = calloutDraw;
const textboxDraw = function(ctx, eng) {
    prevCalloutDraw(ctx, eng);

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    for (const el of project.getVisibleElements()) {
        if (el.type !== 'textbox') continue;

        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const isSelected = (selectedElement === el);
        const isEditing = (calloutEditState.editing && calloutEditState.el === el);
        const color = isSelected ? '#2B7CD0' : layer.color;

        const geom = getTextboxGeom(el, coords, zoom, ctx);

        if (!isEditing) {
            // White fill
            ctx.fillStyle = 'rgba(255, 255, 255, 0.92)';
            ctx.fillRect(geom.boxX, geom.boxY, geom.boxW, geom.boxH);

            // Thin border
            ctx.strokeStyle = color;
            ctx.lineWidth = Math.max(0.5, 0.18 * zoom);
            ctx.setLineDash([]);
            ctx.strokeRect(geom.boxX, geom.boxY, geom.boxW, geom.boxH);

            // Word-wrapped text
            if (el.text) {
                ctx.fillStyle = color;
                ctx.font = `${el.fontBold ? 'bold ' : ''}${geom.fontSize}px "Architects Daughter", cursive`;
                ctx.textAlign = 'left';
                ctx.textBaseline = 'top';
                for (let i = 0; i < geom.lines.length; i++) {
                    ctx.fillText(geom.lines[i], geom.boxX + geom.pad, geom.boxY + geom.pad + i * geom.lineHeight);
                }
            }
        } else {
            // In edit mode: just draw border
            ctx.strokeStyle = '#2B7CD0';
            ctx.lineWidth = 2;
            ctx.setLineDash([]);
            ctx.strokeRect(geom.boxX, geom.boxY, geom.boxW, geom.boxH);
        }

        // Selection handles
        if (isSelected && !isEditing) {
            ctx.fillStyle = '#2B7CD0';
            const anchor = coords.realToScreen(el.x, el.y);
            ctx.fillRect(anchor.x - 3, anchor.y - 3, 6, 6);
        }

        // Resize handles in edit mode
        if (isEditing) {
            const hs = 5;
            ctx.fillStyle = '#2B7CD0';
            ctx.fillRect(geom.boxX + geom.boxW - hs, geom.boxY + geom.boxH - hs, hs * 2, hs * 2);
            ctx.fillRect(geom.boxX - hs, geom.boxY + geom.boxH - hs, hs * 2, hs * 2);
            ctx.fillRect(geom.boxX + geom.boxW - hs, geom.boxY - hs, hs * 2, hs * 2);
        }
    }

    // ── Textbox preview while dragging to size ──
    if (activeTool === 'textbox' && textboxState.phase === 1 && textboxState.boxStart && textboxState.boxEnd) {
        const bs = textboxState.boxStart;
        const be = textboxState.boxEnd;
        const sx1 = coords.sheetToScreen(Math.min(bs.x, be.x), Math.min(bs.y, be.y));
        const sx2 = coords.sheetToScreen(Math.max(bs.x, be.x), Math.max(bs.y, be.y));
        const bw = sx2.x - sx1.x;
        const bh = sx2.y - sx1.y;

        ctx.strokeStyle = '#2B7CD0';
        ctx.lineWidth = 1;
        ctx.globalAlpha = 0.6;
        ctx.setLineDash([4, 3]);
        ctx.strokeRect(sx1.x, sx1.y, bw, bh);
        ctx.setLineDash([]);
        ctx.globalAlpha = 1.0;
    }
};

// Replace in render callbacks
const cbIdx_textbox = engine._renderCallbacks.indexOf(calloutDraw);
if (cbIdx_textbox !== -1) engine._renderCallbacks[cbIdx_textbox] = textboxDraw;

// ── Textbox hit-testing ──────────────────────────────────
const prevHitTest4 = hitTestElement;
hitTestElement = function(sheetPos) {
    const tolerance = 6 / engine.viewport.zoom;
    const zoom = engine.viewport.zoom;
    for (let i = project.elements.length - 1; i >= 0; i--) {
        const el = project.elements[i];
        if (el.type !== 'textbox') continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const geom = getTextboxGeom(el, engine.coords, zoom, engine.ctx);
        const boxTL = engine.coords.screenToSheet(geom.boxX, geom.boxY);
        const boxBR = engine.coords.screenToSheet(geom.boxX + geom.boxW, geom.boxY + geom.boxH);

        if (sheetPos.x >= boxTL.x - tolerance && sheetPos.x <= boxBR.x + tolerance &&
            sheetPos.y >= boxTL.y - tolerance && sheetPos.y <= boxBR.y + tolerance)
            return el;
    }
    return prevHitTest4(sheetPos);
};

// ── Hatch PDF Export Patch ────────────────────────────────
// Note: hatches are complex patterns — the PDF export currently renders
// outlines only. Solid fills can be added; pattern hatches would need
// pdf-level pattern support which is a future enhancement.

// ══════════════════════════════════════════════════════════
