// ── PHASE 5: DIMENSIONS + PDF EXPORT ─────────────────────
// ══════════════════════════════════════════════════════════

// ── Dimension Tool ───────────────────────────────────────

const dimToolState = {
    placing: false,
    startPoint: null,  // sheet-mm
    currentEnd: null,  // sheet-mm
    offset: 5,         // mm offset from measured line (sheet-mm)
};

/** Get snapped + ortho position for dimension tool */
function getDimToolPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    if (dimToolState.placing && dimToolState.startPoint) {
        pos = applyOrtho(pos.x, pos.y, dimToolState.startPoint.x, dimToolState.startPoint.y);
    }
    return pos;
}

// Wire dim tool button + keyboard
document.getElementById('btn-dim').addEventListener('click', () => setActiveTool('dim'));

// Add D key
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'd') { e.preventDefault(); setActiveTool('dim'); }
});

// Mouse move — preview
container.addEventListener('mousemove', (e) => {
    if (activeTool !== 'dim') return;
    dimToolState.currentEnd = getDimToolPos(e);
    engine.requestRender();
});

// Mouse click — place dimension start/end
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'dim') return;
    if (pdfState.calibrating) return;

    const pos = getDimToolPos(e);

    if (!dimToolState.placing) {
        dimToolState.placing = true;
        dimToolState.startPoint = pos;
    } else {
        const start = dimToolState.startPoint;
        const end = pos;
        const dist = Math.sqrt(Math.pow(end.x - start.x, 2) + Math.pow(end.y - start.y, 2));
        if (dist < 0.5) return;

        const realStart = engine.coords.sheetToReal(start.x, start.y);
        const realEnd = engine.coords.sheetToReal(end.x, end.y);

        const newDim = {
            id: generateId(),
            type: 'dimension',
            layer: 'S-DIMS',
            x1: realStart.x, y1: realStart.y,
            x2: realEnd.x, y2: realEnd.y,
            offset: dimToolState.offset,
        };

        history.execute({
            description: 'Add dimension',
            execute() { project.elements.push(newDim); },
            undo() {
                const i = project.elements.indexOf(newDim);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });

        dimToolState.placing = false;
        dimToolState.startPoint = null;
        engine.requestRender();
    }
});

// Right-click cancels
container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'dim' && dimToolState.placing) {
        e.preventDefault();
        dimToolState.placing = false;
        dimToolState.startPoint = null;
        engine.requestRender();
    }
});

// ── Dimension Rendering ──────────────────────────────────

/**
 * Draw a single dimension (Revit style: thin line, tick marks at 45°, text above).
 * Supports horizontal and vertical dimensions with automatic offset.
 */
function drawDimension(ctx, eng, el, isSelected, isPreview) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    let p1s, p2s; // screen positions of the measured points
    if (isPreview) {
        p1s = coords.sheetToScreen(el.x1, el.y1);
        p2s = coords.sheetToScreen(el.x2, el.y2);
    } else {
        p1s = coords.realToScreen(el.x1, el.y1);
        p2s = coords.realToScreen(el.x2, el.y2);
    }

    // Determine if horizontal or vertical
    const dx = p2s.x - p1s.x;
    const dy = p2s.y - p1s.y;
    const isHoriz = Math.abs(dx) >= Math.abs(dy);

    // Offset the dimension line perpendicular to measurement direction
    const offsetPx = (el.offset || 5) * zoom;
    let ox = 0, oy = 0;
    if (isHoriz) {
        oy = -offsetPx; // offset above
    } else {
        ox = -offsetPx; // offset to left
    }

    // Dimension line endpoints (offset from measured points)
    const d1 = { x: p1s.x + ox, y: p1s.y + oy };
    const d2 = { x: p2s.x + ox, y: p2s.y + oy };

    const color = isSelected ? '#2B7CD0' : '#000000';
    const alpha = isPreview ? 0.5 : 1.0;

    ctx.save();
    ctx.globalAlpha = alpha;
    ctx.strokeStyle = color;
    ctx.fillStyle = color;
    ctx.lineWidth = Math.max(0.5, 0.18 * zoom);
    ctx.setLineDash([]);

    // Extension lines (from measured point to dimension line)
    const extGap = 1.5 * zoom; // small gap at measured point
    const extOver = 1.5 * zoom; // overshoot past dimension line

    if (isHoriz) {
        // Left extension
        ctx.beginPath();
        ctx.moveTo(p1s.x, p1s.y + (oy > 0 ? extGap : -extGap));
        ctx.lineTo(d1.x, d1.y - (oy > 0 ? -extOver : extOver));
        ctx.stroke();
        // Right extension
        ctx.beginPath();
        ctx.moveTo(p2s.x, p2s.y + (oy > 0 ? extGap : -extGap));
        ctx.lineTo(d2.x, d2.y - (oy > 0 ? -extOver : extOver));
        ctx.stroke();
    } else {
        // Top extension
        ctx.beginPath();
        ctx.moveTo(p1s.x + (ox > 0 ? extGap : -extGap), p1s.y);
        ctx.lineTo(d1.x - (ox > 0 ? -extOver : extOver), d1.y);
        ctx.stroke();
        // Bottom extension
        ctx.beginPath();
        ctx.moveTo(p2s.x + (ox > 0 ? extGap : -extGap), p2s.y);
        ctx.lineTo(d2.x - (ox > 0 ? -extOver : extOver), d2.y);
        ctx.stroke();
    }

    // Dimension line
    ctx.beginPath();
    ctx.moveTo(d1.x, d1.y);
    ctx.lineTo(d2.x, d2.y);
    ctx.stroke();

    // Tick marks (45° slash — Revit style, not arrows)
    const tickLen = 3 * zoom;
    for (const tp of [d1, d2]) {
        ctx.lineWidth = Math.max(1, 0.25 * zoom);
        ctx.beginPath();
        ctx.moveTo(tp.x - tickLen, tp.y + tickLen);
        ctx.lineTo(tp.x + tickLen, tp.y - tickLen);
        ctx.stroke();
    }

    // Measurement text
    let realDist;
    if (isPreview) {
        const r1 = coords.sheetToReal(el.x1, el.y1);
        const r2 = coords.sheetToReal(el.x2, el.y2);
        realDist = isHoriz
            ? Math.abs(r2.x - r1.x)
            : Math.abs(r2.y - r1.y);
    } else {
        realDist = isHoriz
            ? Math.abs(el.x2 - el.x1)
            : Math.abs(el.y2 - el.y1);
    }

    let dimText;
    if (realDist >= 1000) {
        dimText = (realDist / 1000).toFixed(realDist >= 10000 ? 1 : 2) + ' m';
    } else {
        dimText = Math.round(realDist) + '';
    }

    const midX = (d1.x + d2.x) / 2;
    const midY = (d1.y + d2.y) / 2;
    const fontSize = Math.max(8, 2.2 * zoom);
    ctx.font = `${fontSize}px "Segoe UI", Arial, sans-serif`;
    ctx.textAlign = 'center';

    if (isHoriz) {
        ctx.textBaseline = 'bottom';
        ctx.fillText(dimText, midX, midY - 2);
    } else {
        ctx.save();
        ctx.translate(midX, midY);
        ctx.rotate(-Math.PI / 2);
        ctx.textBaseline = 'bottom';
        ctx.fillText(dimText, 0, -2);
        ctx.restore();
    }

    ctx.restore();
}

// Hook into the drawElements renderer — add dimension rendering
const origDrawElements = drawElements;
const wrappedDrawElements = function(ctx, eng) {
    origDrawElements(ctx, eng);

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    // Draw placed dimensions
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'dimension') continue;
        drawDimension(ctx, eng, el, selectedElement === el, false);
    }

    // Draw dimension preview
    if (activeTool === 'dim' && dimToolState.placing && dimToolState.startPoint && dimToolState.currentEnd) {
        drawDimension(ctx, eng, {
            x1: dimToolState.startPoint.x, y1: dimToolState.startPoint.y,
            x2: dimToolState.currentEnd.x, y2: dimToolState.currentEnd.y,
            offset: dimToolState.offset,
        }, false, true);
    }
};

// Replace the drawElements in render callbacks
const cbIdx = engine._renderCallbacks.indexOf(drawElements);
if (cbIdx !== -1) engine._renderCallbacks[cbIdx] = wrappedDrawElements;

// ── PDF Export ───────────────────────────────────────────

document.getElementById('btn-export-pdf').addEventListener('click', async () => {
    if (typeof jspdf === 'undefined' && typeof window.jspdf === 'undefined') {
        alert('jsPDF library not loaded. Check your internet connection.');
        return;
    }

    const { jsPDF } = window.jspdf;
    const c = CONFIG;

    // Create A1 landscape PDF (dimensions in mm)
    const pdf = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: [c.SHEET_WIDTH_MM, c.SHEET_HEIGHT_MM],
        compress: true,
    });

    // ── Draw the drawing frame ──
    pdf.setDrawColor(0);
    pdf.setLineWidth(0.5);
    pdf.rect(c.MARGIN_LEFT, c.MARGIN_TOP,
        c.SHEET_WIDTH_MM - c.MARGIN_LEFT - c.MARGIN_RIGHT,
        c.SHEET_HEIGHT_MM - c.MARGIN_TOP - c.MARGIN_BOTTOM);

    // Title block divider
    const tbTop = c.SHEET_HEIGHT_MM - c.MARGIN_BOTTOM - c.TITLE_BLOCK_HEIGHT_MM;
    pdf.line(c.MARGIN_LEFT, tbTop,
        c.SHEET_WIDTH_MM - c.MARGIN_RIGHT, tbTop);

    // ── Draw structural grids ──
    pdf.setDrawColor(128);
    pdf.setLineWidth(0.18);
    const da = engine.coords.drawArea;

    for (const grid of structuralGrids) {
        if (isOrthoGrid(grid)) {
            // ── Classic orthogonal grid ──
            if (grid.axis === 'V') {
                const sx = da.left + grid.position / c.drawingScale;
                pdf.line(sx, da.top, sx, da.bottom);
                const by = da.top - 5;
                pdf.setDrawColor(128);
                pdf.setLineWidth(0.18);
                pdf.circle(sx, by, 3);
                pdf.setFontSize(6);
                pdf.setTextColor(0);
                pdf.text(grid.label, sx, by, { align: 'center', baseline: 'middle' });
            } else {
                const sy = da.top + grid.position / c.drawingScale;
                pdf.line(da.left, sy, da.right, sy);
                const bx = da.left - 5;
                pdf.setDrawColor(128);
                pdf.circle(bx, sy, 3);
                pdf.setFontSize(6);
                pdf.setTextColor(0);
                pdf.text(grid.label, bx, sy, { align: 'center', baseline: 'middle' });
            }
        } else if (isAngledGrid(grid)) {
            // ── Angled / finite-extent grid ──
            const ep = gridEndpoints(grid);
            pdf.line(ep.x1, ep.y1, ep.x2, ep.y2);
            // Bubble at the top-most endpoint
            const bubbleEnd = (ep.y1 <= ep.y2) ? { x: ep.x1, y: ep.y1 } : { x: ep.x2, y: ep.y2 };
            const dx = ep.x2 - ep.x1, dy = ep.y2 - ep.y1;
            const len = Math.sqrt(dx * dx + dy * dy);
            const ux = len > 0 ? dx / len : 0, uy = len > 0 ? dy / len : -1;
            const bx = bubbleEnd.x - ux * 5, by = bubbleEnd.y - uy * 5;
            pdf.setDrawColor(128);
            pdf.setLineWidth(0.18);
            pdf.circle(bx, by, 3);
            pdf.setFontSize(6);
            pdf.setTextColor(0);
            pdf.text(grid.label, bx, by, { align: 'center', baseline: 'middle' });
        }
    }

    // ── Draw elements ──
    for (const el of project.getVisibleElements()) {
        const layer = project.layers[el.layer];
        if (!layer) continue;

        // Parse colour
        const rgb = hexToRgb(layer.color);
        pdf.setDrawColor(rgb.r, rgb.g, rgb.b);
        pdf.setTextColor(rgb.r, rgb.g, rgb.b);

        if (el.type === 'line') {
            pdf.setLineWidth(layer.lineWeight);
            const p1 = engine.coords.realToSheet(el.x1, el.y1);
            const p2 = engine.coords.realToSheet(el.x2, el.y2);

            // Dash pattern
            const pattern = DASH_PATTERNS[layer.pattern];
            if (pattern && pattern.length > 0) {
                pdf.setLineDashPattern(pattern, 0);
            } else {
                pdf.setLineDashPattern([], 0);
            }
            pdf.line(p1.x, p1.y, p2.x, p2.y);
            pdf.setLineDashPattern([], 0);
        }

        if (el.type === 'polyline') {
            pdf.setLineWidth(layer.lineWeight);
            const pattern = DASH_PATTERNS[layer.pattern];
            if (pattern && pattern.length > 0) {
                pdf.setLineDashPattern(pattern, 0);
            } else {
                pdf.setLineDashPattern([], 0);
            }
            const pts = el.points.map(p => engine.coords.realToSheet(p.x, p.y));
            if (pts.length >= 2) {
                for (let pi = 0; pi < pts.length - 1; pi++) {
                    pdf.line(pts[pi].x, pts[pi].y, pts[pi+1].x, pts[pi+1].y);
                }
                if (el.closed) {
                    pdf.line(pts[pts.length-1].x, pts[pts.length-1].y, pts[0].x, pts[0].y);
                }
            }
            pdf.setLineDashPattern([], 0);
        }

        if (el.type === 'column') {
            const cp = engine.coords.realToSheet(el.x, el.y);
            const halfSize = el.size / c.drawingScale / 2;
            pdf.setLineWidth(0.35);
            pdf.setFillColor(200);
            pdf.rect(cp.x - halfSize, cp.y - halfSize, halfSize * 2, halfSize * 2, 'FD');
            // X cross
            pdf.line(cp.x - halfSize, cp.y - halfSize, cp.x + halfSize, cp.y + halfSize);
            pdf.line(cp.x + halfSize, cp.y - halfSize, cp.x - halfSize, cp.y + halfSize);
            // Tag
            if (el.tag) {
                pdf.setFontSize(5);
                pdf.setTextColor(0);
                pdf.text(el.tag, cp.x, cp.y + halfSize + 2, { align: 'center' });
            }
        }

        if (el.type === 'text') {
            const tp = engine.coords.realToSheet(el.x, el.y);
            pdf.setFontSize((el.fontSize || 2.5) * 2.8);
            pdf.setFont('helvetica', 'bold');
            pdf.text(el.text, tp.x, tp.y);
            pdf.setFont('helvetica', 'normal');
        }

        if (el.type === 'dimension') {
            const p1 = engine.coords.realToSheet(el.x1, el.y1);
            const p2 = engine.coords.realToSheet(el.x2, el.y2);
            const isHoriz = Math.abs(p2.x - p1.x) >= Math.abs(p2.y - p1.y);
            const off = el.offset || 5;

            let d1, d2;
            if (isHoriz) {
                d1 = { x: p1.x, y: p1.y - off };
                d2 = { x: p2.x, y: p2.y - off };
            } else {
                d1 = { x: p1.x - off, y: p1.y };
                d2 = { x: p2.x - off, y: p2.y };
            }

            pdf.setDrawColor(0);
            pdf.setLineWidth(0.13);

            // Extension lines
            if (isHoriz) {
                pdf.line(p1.x, p1.y - 1, d1.x, d1.y - 1);
                pdf.line(p2.x, p2.y - 1, d2.x, d2.y - 1);
            } else {
                pdf.line(p1.x - 1, p1.y, d1.x - 1, d1.y);
                pdf.line(p2.x - 1, p2.y, d2.x - 1, d2.y);
            }

            // Dim line
            pdf.line(d1.x, d1.y, d2.x, d2.y);

            // Tick marks
            const tk = 1.5;
            pdf.setLineWidth(0.2);
            for (const tp of [d1, d2]) {
                pdf.line(tp.x - tk, tp.y + tk, tp.x + tk, tp.y - tk);
            }

            // Measurement text
            const realDist = isHoriz
                ? Math.abs(el.x2 - el.x1)
                : Math.abs(el.y2 - el.y1);
            const dimText = realDist >= 1000
                ? (realDist / 1000).toFixed(realDist >= 10000 ? 1 : 2) + ' m'
                : Math.round(realDist) + '';

            pdf.setFontSize(5);
            pdf.setTextColor(0);
            const midX = (d1.x + d2.x) / 2;
            const midY = (d1.y + d2.y) / 2;

            if (isHoriz) {
                pdf.text(dimText, midX, midY - 1, { align: 'center' });
            } else {
                pdf.text(dimText, midX - 1, midY, { align: 'center', angle: 90 });
            }
        }

        // ── Bracing Walls (rectangle with X + label) ──
        if (el.type === 'bracingWall') {
            const p1 = engine.coords.realToSheet(el.x1, el.y1);
            const p2 = engine.coords.realToSheet(el.x2, el.y2);
            const dx = p2.x - p1.x, dy = p2.y - p1.y;
            const len = Math.sqrt(dx * dx + dy * dy);
            if (len < 0.1) continue;

            const ux = dx / len, uy = dy / len;
            const nx = -uy, ny = ux;
            const panelW = typeof BRACE_PANEL_WIDTH_MM !== 'undefined' ? BRACE_PANEL_WIDTH_MM : 120;
            const halfW = panelW / 2 / c.drawingScale;

            // Four corners
            const c1x = p1.x + nx * halfW, c1y = p1.y + ny * halfW;
            const c2x = p1.x - nx * halfW, c2y = p1.y - ny * halfW;
            const c3x = p2.x - nx * halfW, c3y = p2.y - ny * halfW;
            const c4x = p2.x + nx * halfW, c4y = p2.y + ny * halfW;

            pdf.setDrawColor(0);
            pdf.setLineWidth(0.25);
            pdf.setLineDashPattern([], 0);

            // Rectangle outline
            pdf.lines([[c4x - c1x, c4y - c1y], [c3x - c4x, c3y - c4y], [c2x - c3x, c2y - c3y], [c1x - c2x, c1y - c2y]], c1x, c1y);

            // X diagonals
            pdf.line(c1x, c1y, c3x, c3y);
            pdf.line(c2x, c2y, c4x, c4y);

            // Label: "1600 (6.4)"
            const typeRef = el.typeRef || 'BR1';
            const typeData = (typeof project !== 'undefined' && project.scheduleTypes.bracingWall) ? project.scheduleTypes.bracingWall[typeRef] : {};
            const baseCap = typeData?.capacity || 0;
            const realLenMm = Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2));
            const labelText = `${Math.round(realLenMm)} (${baseCap.toFixed(1)})`;

            const midX = (p1.x + p2.x) / 2;
            const midY = (p1.y + p2.y) / 2;
            pdf.setFontSize(4);
            pdf.setTextColor(0);
            pdf.text(labelText, midX, midY - halfW - 1, { align: 'center' });
        }
    }

    // ── Title block text ──
    const tb = project.projectInfo;
    const tbLeft = c.MARGIN_LEFT;
    const tbRight = c.SHEET_WIDTH_MM - c.MARGIN_RIGHT;
    const tbW = tbRight - tbLeft;
    const tbH = c.TITLE_BLOCK_HEIGHT_MM;
    const revW = 120, compW = 110, projW = 175, titleW = 160;
    const signW = tbW - revW - compW - projW - titleW;

    // Zone dividers
    pdf.setDrawColor(0);
    pdf.setLineWidth(0.3);
    pdf.line(tbLeft + revW, tbTop, tbLeft + revW, tbTop + tbH);
    pdf.line(tbLeft + revW + compW, tbTop, tbLeft + revW + compW, tbTop + tbH);
    pdf.line(tbLeft + revW + compW + projW, tbTop, tbLeft + revW + compW + projW, tbTop + tbH);
    pdf.line(tbLeft + revW + compW + projW + titleW, tbTop, tbLeft + revW + compW + projW + titleW, tbTop + tbH);

    // Company
    pdf.setFontSize(14);
    pdf.setFont('helvetica', 'bold');
    pdf.setTextColor(0);
    pdf.text(tb.company || '', tbLeft + revW + compW / 2, tbTop + tbH * 0.35, { align: 'center' });
    pdf.setFontSize(7);
    pdf.setFont('helvetica', 'normal');
    pdf.setTextColor(100);
    pdf.text(tb.companySubtitle || '', tbLeft + revW + compW / 2, tbTop + tbH * 0.58, { align: 'center' });

    // Project
    pdf.setFontSize(6);
    pdf.setTextColor(128);
    pdf.text('Project Details', tbLeft + revW + compW + 3, tbTop + 3);
    pdf.setFontSize(12);
    pdf.setFont('helvetica', 'bold');
    pdf.setTextColor(0);
    pdf.text(tb.projectName || '', tbLeft + revW + compW + 3, tbTop + 10);
    pdf.setFontSize(8);
    pdf.setFont('helvetica', 'normal');
    pdf.setTextColor(100);
    pdf.text(tb.address || '', tbLeft + revW + compW + 3, tbTop + 17);

    // Drawing title
    pdf.setFontSize(6);
    pdf.setTextColor(128);
    const ttx = tbLeft + revW + compW + projW;
    pdf.text('Drawing Title:', ttx + 3, tbTop + 3);
    pdf.setFontSize(10);
    pdf.setFont('helvetica', 'bold');
    pdf.setTextColor(0);
    pdf.text(tb.drawingTitle || '', ttx + 3, tbTop + 10);

    // Status
    if (tb.status) {
        pdf.setFontSize(10);
        pdf.setFont('helvetica', 'bold');
        pdf.setTextColor(200, 0, 0);
        pdf.text(tb.status, ttx + titleW / 2, tbTop + tbH * 0.72, { align: 'center' });
    }

    // Scale, Dwg No, Rev in bottom row of sign-off zone
    const ssx = ttx + titleW;
    const metaCellW = signW / 5;
    const metaY = tbTop + tbH * 0.8;
    pdf.setFontSize(5);
    pdf.setTextColor(128);
    pdf.setFont('helvetica', 'normal');
    const metaLabels = ['Scale:', 'Project No:', 'Drawing No:', 'Sheet Size:', 'Rev:'];
    const metaVals = [tb.scale || '1:100', tb.projectNumber || '', tb.drawingNumber || 'S-001', 'A1', tb.revision || 'A'];
    for (let mi = 0; mi < 5; mi++) {
        const mx = ssx + mi * metaCellW;
        pdf.text(metaLabels[mi], mx + 1, metaY);
        pdf.setFontSize(8);
        pdf.setFont('helvetica', 'bold');
        pdf.setTextColor(0);
        pdf.text(metaVals[mi], mx + 1, metaY + 4);
        pdf.setFontSize(5);
        pdf.setFont('helvetica', 'normal');
        pdf.setTextColor(128);
        pdf.line(mx, metaY - 2, mx, tbTop + tbH);
    }

    // Save
    const filename = (tb.drawingNumber || 'drawing') + '.pdf';
    pdf.save(filename);
    console.log('[Export] Saved ' + filename);
});

function hexToRgb(hex) {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    return { r, g, b };
}

// ══════════════════════════════════════════════════════════
// ── PHASE 6: SAVE/LOAD + MEASURE ─────────────────────────
// ══════════════════════════════════════════════════════════

// ── Save Project ─────────────────────────────────────────

function saveProject() {
    const json = JSON.stringify(JSON.parse(getProjectJSON()), null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const name = (project.projectInfo.drawingNumber || 'drawing').replace(/[^a-zA-Z0-9_-]/g, '_');
    a.href = url;
    a.download = name + '.sdraw';
    a.click();
    URL.revokeObjectURL(url);
    console.log('[Save] Project saved as ' + name + '.sdraw');
}

document.getElementById('btn-save').addEventListener('click', saveProject);

// Ctrl+S shortcut
window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        saveProject();
    }
});

// ── Auto-Save & Recovery ─────────────────────────────────
// Saves project to localStorage every 60 seconds (if changes detected).
// On page load, checks for unsaved recovery data and offers to restore.

const AUTOSAVE_KEY = 'structuralsketch_autosave';
const AUTOSAVE_TIMESTAMP_KEY = 'structuralsketch_autosave_ts';
const AUTOSAVE_INTERVAL_MS = 60000; // 60 seconds
let _lastAutoSaveHash = '';

/** Serialize project state to JSON string (same format as saveProject) */
function getProjectJSON() {
    const data = {
        version: 2,
        appName: 'StructuralSketch',
        savedAt: new Date().toISOString(),
        config: {
            drawingScale: CONFIG.drawingScale,
            gridVisible: CONFIG.gridVisible,
        },
        projectInfo: project.projectInfo,
        elements: project.elements,
        layers: project.layers,
        activeLayerId: project.activeLayerId,
        structuralGrids: structuralGrids.map(g => serialiseGrid(g)),
        gridZones: gridZones.map(z => ({ id: z.id, name: z.name, angle: z.angle, visible: z.visible, color: z.color })),
        gridLabelState: gridLabelState,
        nextId: _nextId,
        colNextNum: _colNextNum,
        ftgNextNum: _ftgNextNum,
    };
    // Compact JSON for localStorage (no pretty-printing — saves space)
    return JSON.stringify(data);
}

/** Auto-save to localStorage if project has changed */
function autoSave() {
    try {
        const json = getProjectJSON();
        // Simple change detection: compare hash (length + first/last chars)
        const hash = json.length + ':' + json.slice(0, 100) + json.slice(-100);
        if (hash === _lastAutoSaveHash) return; // no changes
        _lastAutoSaveHash = hash;

        localStorage.setItem(AUTOSAVE_KEY, json);
        localStorage.setItem(AUTOSAVE_TIMESTAMP_KEY, new Date().toISOString());
        console.log('[AutoSave] Saved to localStorage (' + (json.length / 1024).toFixed(1) + ' KB)');
        // Update status bar auto-save indicator
        const statusAS = document.getElementById('status-autosave');
        if (statusAS) {
            const now = new Date();
            statusAS.textContent = 'Saved ' + now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        }
    } catch (e) {
        // localStorage might be full or unavailable
        console.warn('[AutoSave] Failed:', e.message);
    }
}

/** Check for recovery data on page load */
function checkAutoSaveRecovery() {
    try {
        const saved = localStorage.getItem(AUTOSAVE_KEY);
        if (!saved) return;
        const ts = localStorage.getItem(AUTOSAVE_TIMESTAMP_KEY);
        const tsStr = ts ? new Date(ts).toLocaleString() : 'unknown time';

        // Only offer recovery if there are elements in the saved data
        const peek = JSON.parse(saved);
        if (!peek.elements || peek.elements.length === 0) return;

        const msg = 'StructuralSketch found unsaved work from ' + tsStr +
                    ' (' + peek.elements.length + ' elements).\n\n' +
                    'Would you like to recover it?';
        if (confirm(msg)) {
            loadProject(saved);
            console.log('[AutoSave] Recovered project from localStorage');
        } else {
            // Clear old recovery data
            localStorage.removeItem(AUTOSAVE_KEY);
            localStorage.removeItem(AUTOSAVE_TIMESTAMP_KEY);
        }
    } catch (e) {
        console.warn('[AutoSave] Recovery check failed:', e.message);
    }
}

// Start auto-save interval
setInterval(autoSave, AUTOSAVE_INTERVAL_MS);

// Clear auto-save data after a manual save (file download)
const _origSaveProject = saveProject;
saveProject = function() {
    _origSaveProject();
    // Update hash so auto-save doesn't re-save immediately
    try { _lastAutoSaveHash = getProjectJSON().length + ':' + getProjectJSON().slice(0, 100) + getProjectJSON().slice(-100); } catch(e) {}
};

// Check for recovery after a short delay (let the app fully initialise)
setTimeout(checkAutoSaveRecovery, 500);

// ── Load Project ─────────────────────────────────────────

function loadProject(json) {
    try {
        const data = JSON.parse(json);
        if (data.appName !== 'StructuralSketch') {
            alert('This file is not a StructuralSketch project.');
            return;
        }

        // Restore config
        if (data.config) {
            CONFIG.drawingScale = data.config.drawingScale || 100;
            CONFIG.gridVisible = data.config.gridVisible || false;
            updateGridBtn();
        }

        // Restore project info
        if (data.projectInfo) {
            project.projectInfo = { ...project.projectInfo, ...data.projectInfo };
            engine._titleBlockData = project.projectInfo;
        }

        // Restore elements
        project.elements = data.elements || [];

        // Restore layers (merge to keep defaults for any new layers)
        if (data.layers) {
            for (const [id, layerData] of Object.entries(data.layers)) {
                if (project.layers[id]) {
                    Object.assign(project.layers[id], layerData);
                }
            }
        }

        // Restore structural grids (normalise for backward compat)
        structuralGrids.length = 0;
        if (data.structuralGrids) {
            for (const g of data.structuralGrids) {
                structuralGrids.push(normaliseGrid(g));
            }
        }

        // Restore grid zones
        gridZones.length = 0;
        if (data.gridZones) {
            for (const z of data.gridZones) {
                gridZones.push(z);
            }
        }
        // Ensure main zone exists (backward compat — old saves won't have zones)
        ensureMainZone();

        // Restore label state
        if (data.gridLabelState) {
            Object.assign(gridLabelState.V, data.gridLabelState.V || {});
            Object.assign(gridLabelState.H, data.gridLabelState.H || {});
        }

        // Restore counters
        if (data.nextId) _nextId = data.nextId;
        if (data.colNextNum) _colNextNum = data.colNextNum;
        if (data.ftgNextNum) _ftgNextNum = data.ftgNextNum;

        // Restore scale dropdown
        const scaleSelect = document.getElementById('scale-select');
        const scaleVal = String(CONFIG.drawingScale);
        let found = false;
        for (const opt of scaleSelect.options) {
            if (opt.value === scaleVal) { found = true; break; }
        }
        if (!found) {
            const opt = document.createElement('option');
            opt.value = scaleVal;
            opt.textContent = '1:' + scaleVal;
            scaleSelect.appendChild(opt);
        }
        scaleSelect.value = scaleVal;

        // Clear history
        history.clear();
        selectedElement = null;

        // Rebuild UI
        buildLayerPanel();
        engine.requestRender();
        updateStatusBar();

        console.log('[Load] Project loaded — ' + project.elements.length + ' elements, ' +
            structuralGrids.length + ' grids');
    } catch (err) {
        alert('Error loading project: ' + err.message);
        console.error('[Load]', err);
    }
}

document.getElementById('btn-load').addEventListener('click', () => {
    document.getElementById('load-file-input').click();
});

document.getElementById('load-file-input').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => loadProject(reader.result);
    reader.readAsText(file);
    e.target.value = '';
});

// Ctrl+O shortcut
window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'o') {
        e.preventDefault();
        document.getElementById('load-file-input').click();
    }
});

// ── Measure Tool ─────────────────────────────────────────

const measureState = {
    placing: false,
    startPoint: null,
    currentEnd: null,
};

const measureReadout = document.getElementById('measure-readout');

function getMeasurePos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    if (measureState.placing && measureState.startPoint) {
        pos = applyOrtho(pos.x, pos.y, measureState.startPoint.x, measureState.startPoint.y);
    }
    return pos;
}

// Wire measure tool
document.getElementById('btn-measure').addEventListener('click', () => setActiveTool('measure'));

// Extend setActiveTool for measure
// M key
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'm') setActiveTool('measure');
});

// Mouse move
container.addEventListener('mousemove', (e) => {
    if (activeTool !== 'measure') return;
    measureState.currentEnd = getMeasurePos(e);

    if (measureState.placing && measureState.startPoint) {
        // Show floating readout near cursor
        const realStart = engine.coords.sheetToReal(measureState.startPoint.x, measureState.startPoint.y);
        const realEnd = engine.coords.sheetToReal(measureState.currentEnd.x, measureState.currentEnd.y);
        const totalDist = Math.sqrt(Math.pow(realEnd.x - realStart.x, 2) + Math.pow(realEnd.y - realStart.y, 2));
        const dx = Math.abs(realEnd.x - realStart.x);
        const dy = Math.abs(realEnd.y - realStart.y);

        const fmtDist = (mm) => mm >= 1000 ? (mm / 1000).toFixed(mm >= 10000 ? 1 : 2) + ' m' : Math.round(mm) + ' mm';

        let text = fmtDist(totalDist);
        // Show dx/dy breakdown if not purely ortho
        if (dx > 10 && dy > 10) {
            text += '  (Δx: ' + fmtDist(dx) + '  Δy: ' + fmtDist(dy) + ')';
        }

        measureReadout.textContent = text;
        measureReadout.style.display = '';

        const rect = container.getBoundingClientRect();
        measureReadout.style.left = (e.clientX - rect.left + 16) + 'px';
        measureReadout.style.top = (e.clientY - rect.top - 30) + 'px';
    }

    engine.requestRender();
});

// Mouse click
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'measure') return;

    const pos = getMeasurePos(e);

    if (!measureState.placing) {
        measureState.placing = true;
        measureState.startPoint = pos;
    } else {
        // Second click — keep the measurement visible, reset for next
        measureState.placing = false;
        measureState.startPoint = null;
        // Readout stays visible until next action
    }
});

// Right-click cancels
container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'measure' && measureState.placing) {
        e.preventDefault();
        measureState.placing = false;
        measureState.startPoint = null;
        measureReadout.style.display = 'none';
        engine.requestRender();
    }
});

// Draw measure visualization (temporary, not saved)
function drawMeasure(ctx, eng) {
    if (activeTool !== 'measure' || !measureState.placing || !measureState.startPoint || !measureState.currentEnd) return;

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    const sp = coords.sheetToScreen(measureState.startPoint.x, measureState.startPoint.y);
    const ep = coords.sheetToScreen(measureState.currentEnd.x, measureState.currentEnd.y);

    // Dashed measurement line
    ctx.strokeStyle = '#2B7CD0';
    ctx.lineWidth = 1.5;
    ctx.setLineDash([6, 4]);
    ctx.beginPath();
    ctx.moveTo(sp.x, sp.y);
    ctx.lineTo(ep.x, ep.y);
    ctx.stroke();
    ctx.setLineDash([]);

    // Start point dot
    ctx.fillStyle = '#2B7CD0';
    ctx.beginPath();
    ctx.arc(sp.x, sp.y, 4, 0, Math.PI * 2);
    ctx.fill();

    // End point dot
    ctx.beginPath();
    ctx.arc(ep.x, ep.y, 4, 0, Math.PI * 2);
    ctx.fill();

    // If ortho, show the constraint guides
    if (snapState.orthoLock) {
        ctx.strokeStyle = 'rgba(43,124,208,0.25)';
        ctx.lineWidth = 1;
        ctx.setLineDash([3, 3]);
        // Horizontal guide
        ctx.beginPath();
        ctx.moveTo(sp.x - 50, sp.y);
        ctx.lineTo(sp.x + 50, sp.y);
        ctx.stroke();
        // Vertical guide
        ctx.beginPath();
        ctx.moveTo(sp.x, sp.y - 50);
        ctx.lineTo(sp.x, sp.y + 50);
        ctx.stroke();
        ctx.setLineDash([]);
    }
}

// Register measure renderer (on top of everything)
engine.onRender(drawMeasure);

// ══════════════════════════════════════════════════════════
// ── PHASE 7: EDITABLE TITLE BLOCK ────────────────────────
// ══════════════════════════════════════════════════════════

const tbDialog = document.getElementById('tb-edit-dialog');
const tbFields = [
    'projectName', 'address', 'drawingTitle', 'company', 'companySubtitle',
    'drawingNumber', 'projectNumber', 'revision', 'drawnBy', 'designedBy',
    'checkedBy', 'approvedBy', 'date', 'status', 'revDesc'
];

function openTitleBlockEditor() {
    // Populate fields from project info
    for (const key of tbFields) {
        const input = document.getElementById('tb-' + key);
        if (input) input.value = project.projectInfo[key] || '';
    }
    tbDialog.classList.remove('hidden');
    // Focus the first field
    document.getElementById('tb-projectName').focus();
}

function closeTitleBlockEditor() {
    tbDialog.classList.add('hidden');
}

function applyTitleBlockEdits() {
    const oldInfo = { ...project.projectInfo };
    const newInfo = {};
    for (const key of tbFields) {
        const input = document.getElementById('tb-' + key);
        if (input) newInfo[key] = input.value;
    }

    history.execute({
        description: 'Edit title block',
        execute() {
            Object.assign(project.projectInfo, newInfo);
            engine._titleBlockData = project.projectInfo;
        },
        undo() {
            Object.assign(project.projectInfo, oldInfo);
            engine._titleBlockData = project.projectInfo;
        }
    });

    closeTitleBlockEditor();
    engine.requestRender();
}

document.getElementById('tb-apply').addEventListener('click', applyTitleBlockEdits);
document.getElementById('tb-cancel').addEventListener('click', closeTitleBlockEditor);

// Enter in any field applies, Escape cancels
tbDialog.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { e.preventDefault(); closeTitleBlockEditor(); }
    if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); applyTitleBlockEdits(); }
});

// Detect double-click on title block area to open editor
container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'select') return;
    const sheetPos = engine.getSheetPos(e);
    const c = CONFIG;
    const tbTop = c.SHEET_HEIGHT_MM - c.MARGIN_BOTTOM - c.TITLE_BLOCK_HEIGHT_MM;
    const tbBottom = c.SHEET_HEIGHT_MM - c.MARGIN_BOTTOM;
    const tbLeft = c.MARGIN_LEFT;
    const tbRight = c.SHEET_WIDTH_MM - c.MARGIN_RIGHT;

    if (sheetPos.x >= tbLeft && sheetPos.x <= tbRight &&
        sheetPos.y >= tbTop && sheetPos.y <= tbBottom) {
        openTitleBlockEditor();
    }

    // Double-click on text to edit it
    if (selectedElement && selectedElement.type === 'text') {
        const el = selectedElement;
        const newText = prompt('Edit text:', el.text);
        if (newText !== null && newText !== el.text) {
            const oldText = el.text;
            history.execute({
                description: 'Edit text',
                execute() { el.text = newText; },
                undo() { el.text = oldText; }
            });
            engine.requestRender();
        }
    }
});

// Also add a title block edit button somewhere accessible
// We'll make the title block hint text clickable by showing a tooltip on hover
// For now, the double-click approach is clean and unobtrusive

// ══════════════════════════════════════════════════════════
