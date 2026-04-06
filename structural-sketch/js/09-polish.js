// ── PHASE 11: POLISH ─────────────────────────────────────
// ══════════════════════════════════════════════════════════

// ── Multi-Select ─────────────────────────────────────────

let selectedElements = []; // array for multi-select

// Keep backward compat: selectedElement is the first in the set
// Patch selection rendering to handle multiple selections
function isElementSelected(el) {
    return el === selectedElement || selectedElements.includes(el);
}

// Override the select mousedown to support Shift+click for multi-select
// We need to intercept before the existing handler
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'select') return;
    if (pdfState.calibrating) return;
    if (dragState.dragging) return;

    const sheetPos = engine.getSheetPos(e);
    const tolerance = 4 / engine.viewport.zoom;
    const hit = hitTestElement(sheetPos, tolerance);

    if (e.shiftKey && hit) {
        // Shift+click: add/remove from multi-select
        e.stopImmediatePropagation(); // prevent the normal select handler
        if (selectedElements.includes(hit)) {
            selectedElements = selectedElements.filter(el => el !== hit);
            if (selectedElement === hit) selectedElement = selectedElements[0] || null;
        } else {
            selectedElements.push(hit);
            if (!selectedElement) selectedElement = hit;
        }
        engine.requestRender();
        if (typeof updateBulkPropsPanel === 'function') updateBulkPropsPanel();
        return;
    }

    if (!e.shiftKey && hit) {
        // Normal click on element: clear multi-select, set single
        selectedElements = [hit];
        if (typeof updateBulkPropsPanel === 'function') updateBulkPropsPanel();
    } else if (!hit) {
        // Click on nothing: might start drag-box
        selectedElements = [];
        if (typeof updateBulkPropsPanel === 'function') updateBulkPropsPanel();
    }
}, true); // capture phase so it runs before the existing handler

// ── Drag-Box Selection ───────────────────────────────────

const selBoxState = {
    active: false,
    startX: 0, startY: 0,
    endX: 0, endY: 0,
};
let selBoxDiv = null;

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'select') return;
    if (e.shiftKey) return; // shift+click is multi-select, not box

    // Only start box select if clicking on empty space
    const sheetPos = engine.getSheetPos(e);
    const tolerance = 4 / engine.viewport.zoom;
    const hit = hitTestElement(sheetPos, tolerance);
    if (hit) return; // clicked on element, normal select handles it

    selBoxState.active = true;
    const rect = container.getBoundingClientRect();
    selBoxState.startX = e.clientX - rect.left;
    selBoxState.startY = e.clientY - rect.top;
    selBoxState.endX = selBoxState.startX;
    selBoxState.endY = selBoxState.startY;

    // Create visual selection box
    selBoxDiv = document.createElement('div');
    selBoxDiv.className = 'selection-box';
    container.appendChild(selBoxDiv);
});

window.addEventListener('mousemove', (e) => {
    if (!selBoxState.active) return;
    const rect = container.getBoundingClientRect();
    selBoxState.endX = e.clientX - rect.left;
    selBoxState.endY = e.clientY - rect.top;

    const x = Math.min(selBoxState.startX, selBoxState.endX);
    const y = Math.min(selBoxState.startY, selBoxState.endY);
    const w = Math.abs(selBoxState.endX - selBoxState.startX);
    const h = Math.abs(selBoxState.endY - selBoxState.startY);

    selBoxDiv.style.left = x + 'px';
    selBoxDiv.style.top = y + 'px';
    selBoxDiv.style.width = w + 'px';
    selBoxDiv.style.height = h + 'px';

    // Window (L→R) = solid blue | Crossing (R→L) = dashed green
    const isCrossing = selBoxState.endX < selBoxState.startX;
    selBoxDiv.className = 'selection-box ' + (isCrossing ? 'crossing-select' : 'window-select');
});

window.addEventListener('mouseup', (e) => {
    if (!selBoxState.active) return;
    selBoxState.active = false;

    if (selBoxDiv && selBoxDiv.parentNode) selBoxDiv.parentNode.removeChild(selBoxDiv);
    selBoxDiv = null;

    const x1 = Math.min(selBoxState.startX, selBoxState.endX);
    const y1 = Math.min(selBoxState.startY, selBoxState.endY);
    const x2 = Math.max(selBoxState.startX, selBoxState.endX);
    const y2 = Math.max(selBoxState.startY, selBoxState.endY);

    // Minimum box size to count as drag-select (not just a click)
    if (Math.abs(x2 - x1) < 5 && Math.abs(y2 - y1) < 5) return;

    // Window (L→R): fully enclosed only | Crossing (R→L): any overlap
    const isCrossing = selBoxState.endX < selBoxState.startX;

    // Convert box corners to sheet-mm
    const sheetTL = engine.coords.screenToSheet(x1, y1);
    const sheetBR = engine.coords.screenToSheet(x2, y2);

    // Helper: is a point inside the selection box?
    const ptInBox = (px, py) =>
        px >= sheetTL.x && px <= sheetBR.x && py >= sheetTL.y && py <= sheetBR.y;

    // Helper: does a line segment intersect or overlap with the box?
    // Uses Cohen-Sutherland-style check for segment-AABB intersection
    const segIntersectsBox = (ax, ay, bx, by) => {
        // If either endpoint is inside, it intersects
        if (ptInBox(ax, ay) || ptInBox(bx, by)) return true;
        // Check if segment crosses any box edge
        const boxEdges = [
            [sheetTL.x, sheetTL.y, sheetBR.x, sheetTL.y], // top
            [sheetBR.x, sheetTL.y, sheetBR.x, sheetBR.y], // right
            [sheetTL.x, sheetBR.y, sheetBR.x, sheetBR.y], // bottom
            [sheetTL.x, sheetTL.y, sheetTL.x, sheetBR.y], // left
        ];
        for (const [cx, cy, dx, dy] of boxEdges) {
            if (segmentsIntersect(ax, ay, bx, by, cx, cy, dx, dy)) return true;
        }
        return false;
    };

    // Line-line intersection test
    function segmentsIntersect(x1, y1, x2, y2, x3, y3, x4, y4) {
        const denom = (x2 - x1) * (y4 - y3) - (y2 - y1) * (x4 - x3);
        if (Math.abs(denom) < 1e-10) return false;
        const t = ((x3 - x1) * (y4 - y3) - (y3 - y1) * (x4 - x3)) / denom;
        const u = ((x3 - x1) * (y2 - y1) - (y3 - y1) * (x2 - x1)) / denom;
        return t >= 0 && t <= 1 && u >= 0 && u <= 1;
    }

    selectedElements = [];
    for (const el of project.getVisibleElements()) {
        let selected = false;

        // ── Line-type elements (beams, walls, strip footings, dims, leaders, edges, etc.) ──
        if (el.type === 'line' || el.type === 'wall' || el.type === 'stripFooting' ||
            el.type === 'dimension' || el.type === 'leader' || el.type === 'edge' ||
            el.type === 'fallarrow' || el.type === 'step') {
            const p1 = engine.coords.realToSheet(el.x1, el.y1);
            const p2 = engine.coords.realToSheet(el.x2, el.y2);
            if (isCrossing) {
                // Crossing: any part of line touches the box
                selected = segIntersectsBox(p1.x, p1.y, p2.x, p2.y);
            } else {
                // Window: BOTH endpoints must be inside
                selected = ptInBox(p1.x, p1.y) && ptInBox(p2.x, p2.y);
            }
        }

        // ── Point-type elements (columns, text, footings, RL markers, etc.) ──
        else if (el.type === 'column' || el.type === 'text' || el.type === 'footing' ||
                 el.type === 'rlmarker' || el.type === 'secref' || el.type === 'reftag' ||
                 el.type === 'borehole' || el.type === 'slabcallout' || el.type === 'notesbox') {
            const p = engine.coords.realToSheet(el.x, el.y);
            // For both window and crossing, point elements are selected if centre is inside
            selected = ptInBox(p.x, p.y);
        }

        // ── Polygon-type elements (polylines, clouds, slabs) ──
        else if (el.type === 'polyline' || el.type === 'cloud') {
            const pts = (el.points || []).map(pt => engine.coords.realToSheet(pt.x, pt.y));
            if (isCrossing) {
                // Crossing: any vertex inside OR any edge intersects the box
                selected = pts.some(p => ptInBox(p.x, p.y));
                if (!selected) {
                    for (let j = 0; j < pts.length - 1; j++) {
                        if (segIntersectsBox(pts[j].x, pts[j].y, pts[j+1].x, pts[j+1].y)) {
                            selected = true; break;
                        }
                    }
                    // Check closing edge for closed shapes
                    if (!selected && el.closed && pts.length > 2) {
                        selected = segIntersectsBox(
                            pts[pts.length-1].x, pts[pts.length-1].y,
                            pts[0].x, pts[0].y
                        );
                    }
                }
            } else {
                // Window: ALL vertices must be inside
                selected = pts.length > 0 && pts.every(p => ptInBox(p.x, p.y));
            }
        }

        // ── Chain dimensions ──
        else if (el.type === 'chaindim') {
            const pts = (el.points || []).map(pt => engine.coords.realToSheet(pt.x, pt.y));
            if (isCrossing) {
                selected = pts.some(p => ptInBox(p.x, p.y));
            } else {
                selected = pts.length > 0 && pts.every(p => ptInBox(p.x, p.y));
            }
        }

        // ── Table ──
        else if (el.type === 'table') {
            const p = engine.coords.realToSheet(el.x, el.y);
            selected = ptInBox(p.x, p.y);
        }

        if (selected) selectedElements.push(el);
    }

    selectedElement = selectedElements[0] || null;
    engine.requestRender();
    // Trigger bulk properties panel update
    if (typeof updateBulkPropsPanel === 'function') updateBulkPropsPanel();
});

// ── Ctrl+A — Select All ──────────────────────────────────

window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'a' && document.activeElement === document.body) {
        e.preventDefault();
        selectedElements = [...project.getVisibleElements()];
        selectedElement = selectedElements[0] || null;
        engine.requestRender();
        if (typeof updateBulkPropsPanel === 'function') updateBulkPropsPanel();
    }
});

// ── Delete handles multi-select ──────────────────────────

window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if ((e.key === 'Delete' || e.key === 'Backspace') && selectedElements.length > 0 && activeTool === 'select') {
        e.preventDefault();
        const toRemove = [...selectedElements];
        history.execute({
            description: 'Delete ' + toRemove.length + ' element(s)',
            execute() {
                for (const el of toRemove) {
                    const i = project.elements.indexOf(el);
                    if (i !== -1) project.elements.splice(i, 1);
                }
            },
            undo() {
                for (const el of toRemove) project.elements.push(el);
            }
        });
        selectedElement = null;
        selectedElements = [];
        engine.requestRender();
    }
});

// ── Copy/Paste handles multi-select ──────────────────────

window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if ((e.ctrlKey || e.metaKey) && e.key === 'c' && selectedElements.length > 1) {
        e.preventDefault();
        clipboard = selectedElements.map(el => JSON.parse(JSON.stringify(el)));
        if (typeof updatePasteButton === 'function') updatePasteButton();
        console.log('[Copy] Copied ' + clipboard.length + ' elements');
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'v' && Array.isArray(clipboard) && clipboard.length > 0) {
        e.preventDefault();
        const offsetReal = 5 * CONFIG.drawingScale;
        const activeLevel = typeof getActiveLevel === 'function' ? getActiveLevel() : null;
        const pasted = clipboard.map(src => {
            const el = JSON.parse(JSON.stringify(src));
            el.id = generateId();
            // Assign to active level (cross-level paste)
            if (activeLevel) el.level = activeLevel.id;
            // Clear tag offsets
            delete el._tagOffsetX;
            delete el._tagOffsetY;

            if (el.type === 'line' || el.type === 'dimension' || el.type === 'leader' || el.type === 'edge' || el.type === 'fallarrow' || el.type === 'step' || el.type === 'beam' || el.type === 'stripfooting') {
                el.x1 += offsetReal; el.y1 += offsetReal;
                el.x2 += offsetReal; el.y2 += offsetReal;
            } else if (el.type === 'wall') {
                el.x1 += offsetReal; el.y1 += offsetReal;
                el.x2 += offsetReal; el.y2 += offsetReal;
                if (el.wallType === '190 Block' || el.wallType === '140 Block' || el.wallType === '90 Block') {
                    el.tag = 'BW' + _wallBwNum++;
                } else {
                    el.tag = 'CW' + _wallCwNum++;
                }
            } else if (el.type === 'polyline' || el.type === 'cloud' || el.type === 'slab') {
                if (el.points) for (const pt of el.points) { pt.x += offsetReal; pt.y += offsetReal; }
            } else if (el.type === 'column') {
                el.x += offsetReal; el.y += offsetReal;
                el.tag = 'SC' + _colNextNum++;
            } else if (el.type === 'padfooting') {
                el.x += offsetReal; el.y += offsetReal;
                el.tag = 'PF' + _ftgNextNum++;
            } else {
                if ('x' in el) el.x += offsetReal;
                if ('y' in el) el.y += offsetReal;
            }
            return el;
        });

        history.execute({
            description: 'Paste ' + pasted.length + ' elements',
            execute() { for (const el of pasted) project.elements.push(el); },
            undo() { for (const el of pasted) { const i = project.elements.indexOf(el); if (i !== -1) project.elements.splice(i, 1); } }
        });

        selectedElements = pasted;
        selectedElement = pasted[0];
        engine.requestRender();
    }
});

// ── Ctrl+D — Duplicate selected elements ─────────────────

window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (!(e.ctrlKey || e.metaKey) || e.key !== 'd') return;
    e.preventDefault();

    // Gather elements to duplicate (multi-select or single)
    const toDuplicate = selectedElements.length > 0 ? selectedElements :
                        selectedElement ? [selectedElement] : [];
    if (toDuplicate.length === 0) return;

    const offsetReal = 3 * CONFIG.drawingScale; // small offset so duplicates are visible
    const activeLevel = typeof getActiveLevel === 'function' ? getActiveLevel() : null;

    const duplicated = toDuplicate.map(src => {
        const el = JSON.parse(JSON.stringify(src));
        el.id = generateId();
        if (activeLevel) el.level = activeLevel.id;
        delete el._tagOffsetX;
        delete el._tagOffsetY;

        // Offset by type
        if (el.type === 'line' || el.type === 'dimension' || el.type === 'leader' ||
            el.type === 'edge' || el.type === 'fallarrow' || el.type === 'step' ||
            el.type === 'wall' || el.type === 'stripFooting') {
            el.x1 += offsetReal; el.y1 += offsetReal;
            el.x2 += offsetReal; el.y2 += offsetReal;
        } else if (el.type === 'polyline' || el.type === 'cloud') {
            if (el.points) for (const pt of el.points) { pt.x += offsetReal; pt.y += offsetReal; }
        } else if (el.type === 'column') {
            el.x += offsetReal; el.y += offsetReal;
            if (typeof _colNextNum !== 'undefined') el.tag = 'SC' + _colNextNum++;
        } else if (el.type === 'footing') {
            el.x += offsetReal; el.y += offsetReal;
            if (typeof _ftgNextNum !== 'undefined') el.tag = 'PF' + _ftgNextNum++;
        } else {
            if ('x' in el) el.x += offsetReal;
            if ('y' in el) el.y += offsetReal;
        }
        return el;
    });

    history.execute({
        description: 'Duplicate ' + duplicated.length + ' element(s)',
        execute() { for (const el of duplicated) project.elements.push(el); },
        undo() { for (const el of duplicated) { const i = project.elements.indexOf(el); if (i !== -1) project.elements.splice(i, 1); } }
    });

    selectedElements = duplicated;
    selectedElement = duplicated[0] || null;
    engine.requestRender();
    if (typeof updateBulkPropsPanel === 'function') updateBulkPropsPanel();
});

// Patch the isSelected check in element rendering
// Override selectedElement getter to work with multi-select
const origPhase9Draw = phase9Draw;
const phase11Draw = function(ctx, eng) {
    // Temporarily set selectedElement for each render check
    const origSel = selectedElement;
    origPhase9Draw(ctx, eng);
    selectedElement = origSel;

    // Draw multi-select highlights for additional elements
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;
    for (const el of selectedElements) {
        if (el === selectedElement) continue; // already drawn by main renderer
        // Draw a blue highlight
        ctx.strokeStyle = '#2B7CD0';
        ctx.lineWidth = 2;
        ctx.setLineDash([4, 3]);

        if (el.type === 'line' || el.type === 'leader' || el.type === 'wall' || el.type === 'stripFooting' ||
            el.type === 'dimension' || el.type === 'edge' || el.type === 'fallarrow' || el.type === 'step') {
            const p1 = coords.realToScreen(el.x1, el.y1);
            const p2 = coords.realToScreen(el.x2, el.y2);
            ctx.beginPath(); ctx.moveTo(p1.x, p1.y); ctx.lineTo(p2.x, p2.y); ctx.stroke();
            // Endpoint squares for line-type elements
            ctx.fillStyle = '#2B7CD0';
            const nr = 3;
            ctx.fillRect(p1.x - nr, p1.y - nr, nr * 2, nr * 2);
            ctx.fillRect(p2.x - nr, p2.y - nr, nr * 2, nr * 2);
        } else if (el.type === 'column') {
            const cp = coords.realToScreen(el.x, el.y);
            const r = Math.max(3, (el.size / CONFIG.drawingScale) * zoom / 2);
            ctx.strokeRect(cp.x - r - 2, cp.y - r - 2, (r+2)*2, (r+2)*2);
        } else if (el.type === 'footing') {
            const cp = coords.realToScreen(el.x, el.y);
            const fw = (el.footingWidth || 1000) / CONFIG.drawingScale * zoom;
            const r = fw / 2 + 2;
            ctx.strokeRect(cp.x - r, cp.y - r, r * 2, r * 2);
        } else if (el.type === 'text') {
            const tp = coords.realToScreen(el.x, el.y);
            const fs = Math.max(7, (el.fontSize || 3.5) * zoom);
            const tw = (el._renderWidth || el.text.length * (el.fontSize || 3.5) * 0.7) * zoom;
            ctx.strokeRect(tp.x - 2, tp.y - fs/2 - 2, tw + 4, fs + 4);
        } else if (el.type === 'polyline' || el.type === 'cloud') {
            for (let i = 0; i < el.points.length - 1; i++) {
                const p1 = coords.realToScreen(el.points[i].x, el.points[i].y);
                const p2 = coords.realToScreen(el.points[i+1].x, el.points[i+1].y);
                ctx.beginPath(); ctx.moveTo(p1.x, p1.y); ctx.lineTo(p2.x, p2.y); ctx.stroke();
            }
        }
        ctx.setLineDash([]);
    }
};

const cbIdx4 = engine._renderCallbacks.indexOf(phase9Draw);
if (cbIdx4 !== -1) engine._renderCallbacks[cbIdx4] = phase11Draw;

// ── Double-Click Text: Inline Edit ───────────────────────

container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'select') return;
    if (!selectedElement || selectedElement.type !== 'text') return;

    const el = selectedElement;
    const tp = engine.coords.realToScreen(el.x, el.y);
    const fontSize = Math.max(12, (el.fontSize || 3.5) * engine.viewport.zoom);

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'text-input-overlay';
    input.value = el.text;
    input.style.left = tp.x + 'px';
    input.style.top = (tp.y - fontSize / 2) + 'px';
    input.style.fontSize = fontSize + 'px';
    input.style.minWidth = '100px';

    container.appendChild(input);
    setTimeout(() => { input.focus(); input.select(); }, 30);

    const commitEdit = () => {
        const newText = input.value.trim();
        if (input.parentNode) input.parentNode.removeChild(input);
        if (!newText || newText === el.text) return;
        const oldText = el.text;
        history.execute({
            description: 'Edit text',
            execute() { el.text = newText; },
            undo() { el.text = oldText; }
        });
        engine.requestRender();
    };

    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); commitEdit(); }
        if (ev.key === 'Escape') { ev.preventDefault(); if (input.parentNode) input.parentNode.removeChild(input); }
    });
    input.addEventListener('blur', () => setTimeout(commitEdit, 100));
});

// ── Slab Thickness Editor (double-click on slab callout) ─────────────
container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'select') return;
    if (!selectedElement || selectedElement.type !== 'polyline' || selectedElement.layer !== 'S-SLAB' || !selectedElement.slabThickness) return;

    const el = selectedElement;
    const centroid = polyCentroid(el.points);
    const tp = engine.coords.realToScreen(centroid.x, centroid.y);
    const fontSize = Math.max(10, 7 * engine.viewport.zoom);

    const input = document.createElement('input');
    input.type = 'number';
    input.className = 'text-input-overlay';
    input.value = el.slabThickness;
    input.style.left = (tp.x - 30) + 'px';
    input.style.top = (tp.y - fontSize / 2) + 'px';
    input.style.fontSize = fontSize + 'px';
    input.style.width = '60px';

    container.appendChild(input);
    setTimeout(() => { input.focus(); input.select(); }, 30);

    const commitEdit = () => {
        const newThickness = parseInt(input.value);
        if (input.parentNode) input.parentNode.removeChild(input);
        if (!newThickness || newThickness === el.slabThickness) return;
        const oldThickness = el.slabThickness;
        history.execute({
            description: 'Edit slab thickness',
            execute() { el.slabThickness = newThickness; },
            undo() { el.slabThickness = oldThickness; }
        });
        engine.requestRender();
    };

    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') { ev.preventDefault(); commitEdit(); }
        if (ev.key === 'Escape') { ev.preventDefault(); if (input.parentNode) input.parentNode.removeChild(input); }
    });
    input.addEventListener('blur', () => setTimeout(commitEdit, 100));
});

// ── New Project ──────────────────────────────────────────

document.getElementById('btn-new').addEventListener('click', () => {
    if (!confirm('Start a new project? Any unsaved changes will be lost.')) return;
    project.elements.length = 0;
    structuralGrids.length = 0;
    selectedElement = null;
    selectedElements = [];
    _nextId = 1;
    _colNextNum = 1;
    _ftgNextNum = 1;
    gridLabelState.V.nextNum = 1; gridLabelState.V.nextAlpha = 0;
    gridLabelState.H.nextNum = 1; gridLabelState.H.nextAlpha = 0;
    project.projectInfo = {
        projectName: 'PROJECT NAME', address: 'SITE ADDRESS',
        drawingTitle: 'STRUCTURAL PLAN', drawingNumber: 'S-001',
        projectNumber: '', revision: 'A', scale: '1:100',
        status: 'FOR CONSTRUCTION', drawnBy: '', designedBy: '',
        checkedBy: '', approvedBy: '',
        date: new Date().toLocaleDateString('en-AU', { day:'2-digit', month:'2-digit', year:'numeric' }),
        revDesc: '', company: '', companySubtitle: 'STRUCTURAL ENGINEERS'
    };
    engine._titleBlockData = project.projectInfo;
    CONFIG.drawingScale = 100;
    document.getElementById('scale-select').value = '100';
    pdfState.loaded = false; pdfState.pdfDoc = null; pdfState.pageCanvases = {};
    // Hide PDF controls
    document.getElementById('btn-pdf-toggle').style.display = 'none';
    document.getElementById('pdf-page-select').style.display = 'none';
    document.getElementById('pdf-controls').style.display = 'none';
    document.getElementById('pdf-calib-group').style.display = 'none';
    history.clear();
    buildLayerPanel();
    engine.fitToView();
    updateStatusBar();
    console.log('[New] Project reset');
});

// Ctrl+N shortcut
window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'n') {
        e.preventDefault();
        document.getElementById('btn-new').click();
    }
});

// ── Right-Click Context Menu ─────────────────────────────

const ctxMenu = document.getElementById('context-menu');

function showContextMenu(x, y, items) {
    let html = '';
    for (const item of items) {
        if (item.divider) {
            html += '<div class="ctx-divider"></div>';
            continue;
        }
        const disabled = item.disabled ? ' disabled' : '';
        const shortcut = item.shortcut ? `<span class="ctx-shortcut">${item.shortcut}</span>` : '';
        html += `<div class="ctx-item${disabled}" data-action="${item.action}">${item.label}${shortcut}</div>`;
    }
    ctxMenu.innerHTML = html;
    ctxMenu.classList.remove('hidden');

    // Position: ensure menu stays within viewport
    const menuW = 180, menuH = items.length * 28;
    const posX = (x + menuW > window.innerWidth) ? x - menuW : x;
    const posY = (y + menuH > window.innerHeight) ? y - menuH : y;
    ctxMenu.style.left = Math.max(0, posX) + 'px';
    ctxMenu.style.top = Math.max(0, posY) + 'px';

    // Wire click handlers
    for (const el of ctxMenu.querySelectorAll('.ctx-item:not(.disabled)')) {
        el.addEventListener('click', () => {
            const action = el.dataset.action;
            ctxMenu.classList.add('hidden');
            executeContextAction(action);
        });
    }
}

function hideContextMenu() {
    ctxMenu.classList.add('hidden');
}

// Hide context menu on any click elsewhere
window.addEventListener('mousedown', (e) => {
    if (!ctxMenu.contains(e.target)) hideContextMenu();
});
window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') hideContextMenu();
});

// Intercept right-click on canvas
container.addEventListener('contextmenu', (e) => {
    // Don't override if a drawing tool is using right-click to cancel
    if (activeTool !== 'select') return;

    e.preventDefault();

    const sheetPos = engine.getSheetPos(e);
    const tolerance = 4 / engine.viewport.zoom;
    const hit = hitTestElement(sheetPos, tolerance);

    // If right-clicked on an element not in current selection, select it
    if (hit && !selectedElements.includes(hit)) {
        selectedElement = hit;
        selectedElements = [hit];
        engine.requestRender();
    }

    const hasSelection = selectedElements.length > 0 || selectedElement;
    const selCount = selectedElements.length || (selectedElement ? 1 : 0);
    const selLabel = selCount > 1 ? selCount + ' Elements' : (selectedElement ? (selectedElement.type === 'line' && selectedElement.layer === 'S-BEAM' ? 'Beam' : selectedElement.type.charAt(0).toUpperCase() + selectedElement.type.slice(1)) : '');

    const items = [];

    if (hasSelection) {
        items.push({ label: 'Duplicate', action: 'duplicate', shortcut: 'Ctrl+D' });
        items.push({ label: 'Copy', action: 'copy', shortcut: 'Ctrl+C' });
        items.push({ divider: true });
        items.push({ label: 'Delete' + (selCount > 1 ? ' (' + selCount + ')' : ''), action: 'delete', shortcut: 'Del' });
        items.push({ divider: true });
    }

    // Paste always available if clipboard has content
    const canPaste = (typeof clipboard !== 'undefined' && clipboard &&
                     (Array.isArray(clipboard) ? clipboard.length > 0 : true));
    items.push({ label: 'Paste', action: 'paste', shortcut: 'Ctrl+V', disabled: !canPaste });

    if (hasSelection) {
        items.push({ divider: true });
        // Type reassignment (for elements with typeRef)
        if (selCount === 1 && selectedElement && selectedElement.typeRef) {
            items.push({ label: 'Change Type...', action: 'changetype' });
        }
        items.push({ label: 'Select All', action: 'selectall', shortcut: 'Ctrl+A' });
    } else {
        items.push({ label: 'Select All', action: 'selectall', shortcut: 'Ctrl+A' });
    }

    showContextMenu(e.clientX, e.clientY, items);
});

function executeContextAction(action) {
    switch (action) {
        case 'duplicate': {
            // Trigger Ctrl+D
            const ev = new KeyboardEvent('keydown', { key: 'd', ctrlKey: true, bubbles: true });
            window.dispatchEvent(ev);
            break;
        }
        case 'copy': {
            // Copy to clipboard
            const toCopy = selectedElements.length > 0 ? selectedElements :
                           selectedElement ? [selectedElement] : [];
            if (toCopy.length > 0) {
                clipboard = toCopy.map(el => JSON.parse(JSON.stringify(el)));
                if (typeof updatePasteButton === 'function') updatePasteButton();
            }
            break;
        }
        case 'paste': {
            const ev = new KeyboardEvent('keydown', { key: 'v', ctrlKey: true, bubbles: true });
            window.dispatchEvent(ev);
            break;
        }
        case 'delete': {
            const ev = new KeyboardEvent('keydown', { key: 'Delete', bubbles: true });
            window.dispatchEvent(ev);
            break;
        }
        case 'selectall': {
            selectedElements = [...project.getVisibleElements()];
            selectedElement = selectedElements[0] || null;
            engine.requestRender();
            if (typeof updateBulkPropsPanel === 'function') updateBulkPropsPanel();
            break;
        }
        case 'changetype': {
            if (selectedElement && typeof showTypeReassignmentPicker === 'function') {
                // Determine category from element type
                let cat = 'beam';
                if (selectedElement.type === 'wall') cat = 'wall';
                else if (selectedElement.type === 'column') cat = 'column';
                else if (selectedElement.type === 'footing') cat = 'padfooting';
                else if (selectedElement.type === 'stripFooting') cat = 'stripfooting';
                showTypeReassignmentPicker(selectedElement, cat, 150, 80);
            }
            break;
        }
    }
}

// ── Shortcuts Panel ──────────────────────────────────────

const shortcutsPanel = document.getElementById('shortcuts-panel');

document.getElementById('btn-shortcuts').addEventListener('click', () => {
    shortcutsPanel.classList.toggle('hidden');
});

document.getElementById('shortcuts-close').addEventListener('click', () => {
    shortcutsPanel.classList.add('hidden');
});

shortcutsPanel.addEventListener('click', (e) => {
    if (e.target === shortcutsPanel) shortcutsPanel.classList.add('hidden');
});

// ? key opens shortcuts
window.addEventListener('keydown', (e) => {
    if (e.key === '?' && document.activeElement === document.body) {
        shortcutsPanel.classList.toggle('hidden');
    }
});

// ── Mirror Tool ──────────────────────────────────────────
// Mirrors selected elements about a user-drawn axis line.
// Click btn-mirror → tool enters mirror mode → user clicks 2 points for axis → mirrored copies placed.

let _mirrorState = { active: false, startPoint: null, previewEnd: null };

document.getElementById('btn-mirror').addEventListener('click', () => {
    const toMirror = selectedElements.length > 0 ? selectedElements : (selectedElement ? [selectedElement] : []);
    if (toMirror.length === 0) {
        alert('Select elements first, then click Mirror.');
        return;
    }
    _mirrorState = { active: true, startPoint: null, previewEnd: null, elements: toMirror.slice() };
    container.style.cursor = 'crosshair';
});

container.addEventListener('mousedown', (e) => {
    if (!_mirrorState.active || e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;

    const snap = findSnap(e.clientX - container.getBoundingClientRect().left, e.clientY - container.getBoundingClientRect().top);
    const sheetPos = snap ? { x: snap.x, y: snap.y } : engine.getSheetPos(e);

    if (!_mirrorState.startPoint) {
        _mirrorState.startPoint = { ...sheetPos };
        _mirrorState.previewEnd = { ...sheetPos };
    } else {
        // Second click — execute mirror
        const p1 = engine.coords.sheetToReal(_mirrorState.startPoint.x, _mirrorState.startPoint.y);
        const p2 = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);

        // Axis vector
        const axDx = p2.x - p1.x;
        const axDy = p2.y - p1.y;
        const axLen2 = axDx * axDx + axDy * axDy;
        if (axLen2 < 0.01) { _mirrorState.active = false; return; }

        /** Reflect a point about the axis line p1→p2 */
        function reflectPt(px, py) {
            const t = ((px - p1.x) * axDx + (py - p1.y) * axDy) / axLen2;
            const closestX = p1.x + t * axDx;
            const closestY = p1.y + t * axDy;
            return { x: 2 * closestX - px, y: 2 * closestY - py };
        }

        const mirrored = _mirrorState.elements.map(src => {
            const el = JSON.parse(JSON.stringify(src));
            el.id = generateId();
            delete el._tagOffsetX;
            delete el._tagOffsetY;
            delete el.wallRef;
            delete el.colRef;

            if (el.x1 !== undefined && el.y1 !== undefined) {
                const r1 = reflectPt(el.x1, el.y1);
                const r2 = reflectPt(el.x2, el.y2);
                el.x1 = r1.x; el.y1 = r1.y;
                el.x2 = r2.x; el.y2 = r2.y;
            } else if (el.points) {
                for (const pt of el.points) {
                    const r = reflectPt(pt.x, pt.y);
                    pt.x = r.x; pt.y = r.y;
                }
            } else if (el.x !== undefined && el.y !== undefined) {
                const r = reflectPt(el.x, el.y);
                el.x = r.x; el.y = r.y;
            }
            return el;
        });

        history.execute({
            description: 'Mirror ' + mirrored.length + ' element(s)',
            execute() { for (const el of mirrored) project.elements.push(el); },
            undo() { for (const el of mirrored) { const i = project.elements.indexOf(el); if (i !== -1) project.elements.splice(i, 1); } }
        });

        selectedElements = mirrored;
        selectedElement = mirrored[0] || null;
        _mirrorState = { active: false, startPoint: null, previewEnd: null };
        container.style.cursor = '';
        engine.requestRender();
        if (typeof updateBulkPropsPanel === 'function') updateBulkPropsPanel();
    }
}, true);

container.addEventListener('mousemove', (e) => {
    if (!_mirrorState.active || !_mirrorState.startPoint) return;
    const snap = findSnap(e.clientX - container.getBoundingClientRect().left, e.clientY - container.getBoundingClientRect().top);
    _mirrorState.previewEnd = snap ? { x: snap.x, y: snap.y } : engine.getSheetPos(e);
    engine.requestRender();
});

// Escape to cancel mirror
window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && _mirrorState.active) {
        _mirrorState = { active: false, startPoint: null, previewEnd: null };
        container.style.cursor = '';
        engine.requestRender();
    }
});

// Render mirror axis preview
engine.onRender((ctx, eng) => {
    if (!_mirrorState.active || !_mirrorState.startPoint || !_mirrorState.previewEnd) return;
    const p1 = eng.coords.sheetToScreen(_mirrorState.startPoint.x, _mirrorState.startPoint.y);
    const p2 = eng.coords.sheetToScreen(_mirrorState.previewEnd.x, _mirrorState.previewEnd.y);

    ctx.save();
    ctx.strokeStyle = '#e91e63';
    ctx.lineWidth = 1.5;
    ctx.setLineDash([8, 4, 2, 4]); // dash-dot pattern for mirror axis
    ctx.beginPath();

    // Extend the line beyond the two points for visibility
    const dx = p2.x - p1.x;
    const dy = p2.y - p1.y;
    const len = Math.sqrt(dx * dx + dy * dy);
    if (len > 1) {
        const ext = 2000; // extend 2000px in each direction
        const nx = dx / len, ny = dy / len;
        ctx.moveTo(p1.x - nx * ext, p1.y - ny * ext);
        ctx.lineTo(p2.x + nx * ext, p2.y + ny * ext);
    }
    ctx.stroke();
    ctx.setLineDash([]);

    // Draw axis endpoints
    for (const pt of [p1, p2]) {
        ctx.fillStyle = '#e91e63';
        ctx.beginPath();
        ctx.arc(pt.x, pt.y, 4, 0, Math.PI * 2);
        ctx.fill();
    }

    // Label
    ctx.fillStyle = '#e91e63';
    ctx.font = 'bold 11px sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('MIRROR AXIS', (p1.x + p2.x) / 2, Math.min(p1.y, p2.y) - 12);
    ctx.restore();
});

// M key shortcut for mirror
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'm' || e.key === 'M') {
        document.getElementById('btn-mirror').click();
    }
});


// ── Array Tool ───────────────────────────────────────────
// Linear array: select elements, click Array, enter count + spacing,
// creates copies along a user-defined direction.

document.getElementById('btn-array').addEventListener('click', () => {
    const toArray = selectedElements.length > 0 ? selectedElements : (selectedElement ? [selectedElement] : []);
    if (toArray.length === 0) {
        alert('Select elements first, then click Array.');
        return;
    }

    // Prompt for count and spacing
    const countStr = prompt('Number of copies (not including original):', '3');
    if (!countStr) return;
    const count = parseInt(countStr);
    if (isNaN(count) || count < 1 || count > 50) {
        alert('Please enter a number between 1 and 50.');
        return;
    }

    const spacingStr = prompt('Spacing between copies (mm):', '3000');
    if (!spacingStr) return;
    const spacing = parseFloat(spacingStr);
    if (isNaN(spacing) || spacing <= 0) {
        alert('Please enter a positive spacing in mm.');
        return;
    }

    // Direction: ask user
    const dirStr = prompt('Direction: H (horizontal), V (vertical), or angle in degrees:', 'H');
    if (!dirStr) return;
    let angleDeg = 0;
    if (dirStr.toUpperCase() === 'H') angleDeg = 0;
    else if (dirStr.toUpperCase() === 'V') angleDeg = 90;
    else angleDeg = parseFloat(dirStr);
    if (isNaN(angleDeg)) { alert('Invalid direction.'); return; }

    const angleRad = angleDeg * Math.PI / 180;
    const stepX = spacing * Math.cos(angleRad);
    const stepY = spacing * Math.sin(angleRad);

    const allCopies = [];

    for (let n = 1; n <= count; n++) {
        const offsetX = stepX * n;
        const offsetY = stepY * n;

        for (const src of toArray) {
            const el = JSON.parse(JSON.stringify(src));
            el.id = generateId();
            delete el._tagOffsetX;
            delete el._tagOffsetY;
            delete el.wallRef;
            delete el.colRef;

            if (el.x1 !== undefined && el.y1 !== undefined) {
                el.x1 += offsetX; el.y1 += offsetY;
                el.x2 += offsetX; el.y2 += offsetY;
            } else if (el.points) {
                for (const pt of el.points) { pt.x += offsetX; pt.y += offsetY; }
            } else if (el.x !== undefined && el.y !== undefined) {
                el.x += offsetX; el.y += offsetY;
            }
            allCopies.push(el);
        }
    }

    history.execute({
        description: 'Array ' + toArray.length + ' element(s) × ' + count,
        execute() { for (const el of allCopies) project.elements.push(el); },
        undo() { for (const el of allCopies) { const i = project.elements.indexOf(el); if (i !== -1) project.elements.splice(i, 1); } }
    });

    selectedElements = allCopies;
    selectedElement = allCopies[0] || null;
    engine.requestRender();
    if (typeof updateBulkPropsPanel === 'function') updateBulkPropsPanel();
});

// ══════════════════════════════════════════════════════════
