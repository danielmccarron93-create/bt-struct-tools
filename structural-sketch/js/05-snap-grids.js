// ── PHASE 3: SNAP ENGINE + STRUCTURAL GRIDS ──────────────
// ══════════════════════════════════════════════════════════

// ── Snap & Ortho State ───────────────────────────────────

const snapState = {
    enabled: true,         // master snap toggle
    gridSnap: true,        // snap to background grid
    endpointSnap: true,    // snap to element endpoints
    midpointSnap: true,    // snap to midpoints
    intersectionSnap: true,// snap to intersections
    gridLineSnap: true,    // snap to structural grid lines
    extensionSnap: true,   // snap to H/V extension lines from element endpoints

    orthoLock: false,      // orthogonal constraint
    snapRadius: 12,        // pixels — snap search radius on screen

    // Current snap result (updated each mouse move)
    activeSnap: null,      // { x, y, type, description } in sheet-mm, or null

    // Extension guide lines for visual feedback (populated by findSnap)
    extensionGuides: [],   // [{ from: {x,y}, to: {x,y}, axis: 'H'|'V' }] in sheet-mm
};

const SNAP_TYPES = {
    GRID: { color: '#2B7CD0', symbol: 'square', label: 'Grid' },
    ENDPOINT: { color: '#FF3300', symbol: 'square', label: 'Endpoint' },
    MIDPOINT: { color: '#2E8B57', symbol: 'triangle', label: 'Midpoint' },
    INTERSECTION: { color: '#E68A00', symbol: 'cross', label: 'Intersection' },
    GRIDLINE: { color: '#808080', symbol: 'diamond', label: 'Grid Line' },
    PERPENDICULAR: { color: '#9B59B6', symbol: 'circle', label: 'Perpendicular' },
    NEAREST: { color: '#D35400', symbol: 'cross', label: 'Nearest' },
    EXTENSION: { color: '#00ACC1', symbol: 'diamond', label: 'Extension' },
};

/**
 * Find the best snap point near a screen position.
 * Returns { x, y, type } in sheet-mm, or null.
 *
 * Priority system (higher priority wins if within radius):
 *   P1: Grid intersections (2-axis lock) — boosted radius × 1.5
 *   P2: Element endpoints (columns, line ends, footings) — includes ghost level
 *   P3: Element midpoints
 *   P4: Single-axis grid line snap
 *   P5: Background grid (only if nothing else found)
 *
 * When the active tool is 'line' (beam/wall), columns and structural element
 * endpoints get a boosted snap radius for easier connection.
 */
function findSnap(screenX, screenY) {
    if (!snapState.enabled) return null;

    const sheetPos = engine.coords.screenToSheet(screenX, screenY);
    const radiusMM = snapState.snapRadius / engine.viewport.zoom;

    // Boosted radius for high-priority snaps (grid intersections, columns when drawing beams)
    const boostRadius = radiusMM * 1.6;

    // Whether the active tool would benefit from boosted member snap
    const isStructTool = (activeTool === 'line' || activeTool === 'wall' || activeTool === 'column');

    // Collect candidates with priority levels
    let candidates = []; // { x, y, type, dist, priority }

    const dist2d = (x1, y1, x2, y2) => Math.sqrt((x1 - x2) ** 2 + (y1 - y2) ** 2);

    // ── Structural grid intersections (P1 — highest, boosted radius) ──
    if (snapState.gridLineSnap && snapState.intersectionSnap) {
        for (const g1 of structuralGrids) {
            if (g1.axis !== 'V') continue;
            for (const g2 of structuralGrids) {
                if (g2.axis !== 'H') continue;
                const ix = engine.coords.drawArea.left + g1.position / CONFIG.drawingScale;
                const iy = engine.coords.drawArea.top + g2.position / CONFIG.drawingScale;
                const d = dist2d(sheetPos.x, sheetPos.y, ix, iy);
                if (d < boostRadius) {
                    candidates.push({ x: ix, y: iy, type: SNAP_TYPES.INTERSECTION, dist: d, priority: 1 });
                }
            }
        }
    }

    // ── Element endpoints: columns, line ends, footings (P2) ──
    // Collect from visible elements AND ghost elements in one pass
    const allSnapEls = [];
    if (snapState.endpointSnap || snapState.midpointSnap || snapState.extensionSnap) {
        allSnapEls.push(...project.getVisibleElements());
    }
    if ((snapState.endpointSnap || snapState.extensionSnap) && typeof project.getGhostElements === 'function') {
        allSnapEls.push(...project.getGhostElements());
    }

    /** Nearest point on line segment p1-p2 from point q. Returns { x, y, t } where t is 0..1 parameter */
    function nearestOnSegment(p1x, p1y, p2x, p2y, qx, qy) {
        const dx = p2x - p1x, dy = p2y - p1y;
        const lenSq = dx * dx + dy * dy;
        if (lenSq < 0.0001) return { x: p1x, y: p1y, t: 0 };
        let t = ((qx - p1x) * dx + (qy - p1y) * dy) / lenSq;
        t = Math.max(0, Math.min(1, t)); // clamp to segment
        return { x: p1x + t * dx, y: p1y + t * dy, t };
    }

    // Track line-type elements for intersection snaps later
    const lineEls = [];

    for (const el of allSnapEls) {
        // Line-type elements (beams, walls, strip footings)
        if (el.type === 'line' || el.type === 'wall' || el.type === 'stripFooting') {
            const p1 = engine.coords.realToSheet(el.x1, el.y1);
            const p2 = engine.coords.realToSheet(el.x2, el.y2);
            lineEls.push({ p1, p2, el });

            if (snapState.endpointSnap) {
                // Boost endpoint radius when drawing structural members
                const epRadius = isStructTool ? boostRadius : radiusMM;
                const d1 = dist2d(sheetPos.x, sheetPos.y, p1.x, p1.y);
                if (d1 < epRadius) candidates.push({ x: p1.x, y: p1.y, type: SNAP_TYPES.ENDPOINT, dist: d1, priority: 2 });
                const d2 = dist2d(sheetPos.x, sheetPos.y, p2.x, p2.y);
                if (d2 < epRadius) candidates.push({ x: p2.x, y: p2.y, type: SNAP_TYPES.ENDPOINT, dist: d2, priority: 2 });
            }
            if (snapState.midpointSnap) {
                const mx = (p1.x + p2.x) / 2, my = (p1.y + p2.y) / 2;
                const dm = dist2d(sheetPos.x, sheetPos.y, mx, my);
                if (dm < radiusMM) candidates.push({ x: mx, y: my, type: SNAP_TYPES.MIDPOINT, dist: dm, priority: 3 });
            }

            // Nearest point on line — useful when approaching a beam at an angle
            // Only activate when using structural tools (beam/wall/column drawing)
            if (isStructTool && snapState.endpointSnap) {
                const np = nearestOnSegment(p1.x, p1.y, p2.x, p2.y, sheetPos.x, sheetPos.y);
                // Only snap to interior of line (not too close to endpoints — those are caught above)
                if (np.t > 0.02 && np.t < 0.98) {
                    const dn = dist2d(sheetPos.x, sheetPos.y, np.x, np.y);
                    if (dn < radiusMM) {
                        candidates.push({ x: np.x, y: np.y, type: SNAP_TYPES.NEAREST, dist: dn, priority: 3 });
                    }
                }
            }
        }

        // Columns — snap to centre point (boosted when drawing beams)
        if (el.type === 'column' && snapState.endpointSnap) {
            const cp = engine.coords.realToSheet(el.x, el.y);
            const colRadius = isStructTool ? boostRadius : radiusMM;
            const dc = dist2d(sheetPos.x, sheetPos.y, cp.x, cp.y);
            if (dc < colRadius) candidates.push({ x: cp.x, y: cp.y, type: SNAP_TYPES.ENDPOINT, dist: dc, priority: 2 });
        }

        // Pad footings
        if (el.type === 'footing' && snapState.endpointSnap) {
            const fp = engine.coords.realToSheet(el.x, el.y);
            const df = dist2d(sheetPos.x, sheetPos.y, fp.x, fp.y);
            if (df < radiusMM) candidates.push({ x: fp.x, y: fp.y, type: SNAP_TYPES.ENDPOINT, dist: df, priority: 2 });
        }
    }

    // ── Line-line intersection snaps (where beams cross each other) ──
    // When actively drawing a line, check where the in-progress line would cross existing elements
    if (isStructTool && typeof lineToolState !== 'undefined' && lineToolState.placing && lineToolState.startPoint) {
        const refStart = lineToolState.startPoint;
        for (const le of lineEls) {
            // Compute intersection of the line being drawn with each existing line
            const a1x = refStart.x, a1y = refStart.y, a2x = sheetPos.x, a2y = sheetPos.y;
            const b1x = le.p1.x, b1y = le.p1.y, b2x = le.p2.x, b2y = le.p2.y;

            const dax = a2x - a1x, day = a2y - a1y;
            const dbx = b2x - b1x, dby = b2y - b1y;
            const denom = dax * dby - day * dbx;

            if (Math.abs(denom) < 0.0001) continue; // parallel

            const t = ((b1x - a1x) * dby - (b1y - a1y) * dbx) / denom;
            const s = ((b1x - a1x) * day - (b1y - a1y) * dax) / denom;

            // s must be on the existing segment [0,1]; t can extend a bit beyond start
            if (s >= 0 && s <= 1 && t > 0.1) {
                const ix = a1x + t * dax;
                const iy = a1y + t * day;
                const di = dist2d(sheetPos.x, sheetPos.y, ix, iy);
                if (di < boostRadius) {
                    candidates.push({ x: ix, y: iy, type: SNAP_TYPES.INTERSECTION, dist: di, priority: 2 });
                }
            }
        }
    }

    // ── Extension line snaps (P3.5) ──────────────────────────
    // When actively drawing a line-type element, project H/V extension lines
    // from ALL visible element endpoints & column centres across the page.
    // If the cursor is near a projected H or V alignment, snap to it.
    // If two extension lines cross (H from one source, V from another), snap
    // to that intersection — this is the "across the page" alignment snap.
    snapState.extensionGuides = []; // clear each frame

    const isDrawing = (
        (activeTool === 'line' && typeof lineToolState !== 'undefined' && lineToolState.placing) ||
        (activeTool === 'wall' && typeof wallToolState !== 'undefined' && wallToolState.placing) ||
        (activeTool === 'stripFooting' && typeof stripFtgState !== 'undefined' && stripFtgState.placing)
    );

    if (snapState.extensionSnap && isDrawing) {
        // Collect all snap-worthy points (endpoints, column centres, footing centres)
        const extPoints = []; // { x, y } in sheet-mm
        for (const el of allSnapEls) {
            if (el.type === 'line' || el.type === 'wall' || el.type === 'stripFooting') {
                const p1 = engine.coords.realToSheet(el.x1, el.y1);
                const p2 = engine.coords.realToSheet(el.x2, el.y2);
                extPoints.push(p1, p2);
            }
            if (el.type === 'column') {
                extPoints.push(engine.coords.realToSheet(el.x, el.y));
            }
            if (el.type === 'footing') {
                extPoints.push(engine.coords.realToSheet(el.x, el.y));
            }
        }

        // Also include structural grid intersections as extension sources
        // (already handled by grid snap, so skip here to avoid duplicates)

        // Extension snap radius — slightly wider than normal to feel helpful
        const extRadius = radiusMM * 1.2;

        // Collect H and V extension hits separately for cross-detection
        const hHits = []; // { y, srcPt } — cursor is near horizontal extension from srcPt
        const vHits = []; // { x, srcPt } — cursor is near vertical extension from srcPt

        for (const pt of extPoints) {
            // Skip points that are very close to cursor (those are already endpoint snaps)
            const dToPt = dist2d(sheetPos.x, sheetPos.y, pt.x, pt.y);
            if (dToPt < radiusMM * 0.5) continue;

            // Horizontal extension: same Y as source point, cursor X is anywhere
            const dy = Math.abs(sheetPos.y - pt.y);
            if (dy < extRadius) {
                hHits.push({ y: pt.y, srcPt: pt, dist: dy });
            }

            // Vertical extension: same X as source point, cursor Y is anywhere
            const dx = Math.abs(sheetPos.x - pt.x);
            if (dx < extRadius) {
                vHits.push({ x: pt.x, srcPt: pt, dist: dx });
            }
        }

        // ── Cross-extension intersection (H from one source × V from another) ──
        // This is the highest-value snap: aligning to two different elements simultaneously
        for (const h of hHits) {
            for (const v of vHits) {
                // Don't intersect extensions from the same source point
                if (Math.abs(h.srcPt.x - v.srcPt.x) < 0.01 && Math.abs(h.srcPt.y - v.srcPt.y) < 0.01) continue;

                const ix = v.x;
                const iy = h.y;
                const di = dist2d(sheetPos.x, sheetPos.y, ix, iy);
                if (di < extRadius * 1.5) {
                    candidates.push({
                        x: ix, y: iy,
                        type: SNAP_TYPES.EXTENSION,
                        dist: di,
                        priority: 2,  // high priority — two-axis lock
                        _extGuides: [
                            { from: h.srcPt, to: { x: ix, y: iy }, axis: 'H' },
                            { from: v.srcPt, to: { x: ix, y: iy }, axis: 'V' },
                        ]
                    });
                }
            }
        }

        // ── Single-axis extension snaps (P3 — just H or just V alignment) ──
        // Only add if cursor is reasonably close to the projected line
        for (const h of hHits) {
            const snapPt = { x: sheetPos.x, y: h.y };
            candidates.push({
                x: snapPt.x, y: snapPt.y,
                type: SNAP_TYPES.EXTENSION,
                dist: h.dist,
                priority: 3,
                _extGuides: [{ from: h.srcPt, to: snapPt, axis: 'H' }]
            });
        }
        for (const v of vHits) {
            const snapPt = { x: v.x, y: sheetPos.y };
            candidates.push({
                x: snapPt.x, y: snapPt.y,
                type: SNAP_TYPES.EXTENSION,
                dist: v.dist,
                priority: 3,
                _extGuides: [{ from: v.srcPt, to: snapPt, axis: 'V' }]
            });
        }
    }

    // ── Single-axis structural grid line snaps (P4) ──
    if (snapState.gridLineSnap) {
        for (const grid of structuralGrids) {
            if (grid.axis === 'V') {
                const gx = engine.coords.drawArea.left + grid.position / CONFIG.drawingScale;
                const d = Math.abs(sheetPos.x - gx);
                if (d < radiusMM) {
                    candidates.push({ x: gx, y: sheetPos.y, type: SNAP_TYPES.GRIDLINE, dist: d, priority: 4 });
                }
            } else {
                const gy = engine.coords.drawArea.top + grid.position / CONFIG.drawingScale;
                const d = Math.abs(sheetPos.y - gy);
                if (d < radiusMM) {
                    candidates.push({ x: sheetPos.x, y: gy, type: SNAP_TYPES.GRIDLINE, dist: d, priority: 4 });
                }
            }
        }
    }

    // Sort: priority first (lower = better), then distance
    candidates.sort((a, b) => {
        if (a.priority !== b.priority) return a.priority - b.priority;
        return a.dist - b.dist;
    });

    if (candidates.length > 0) {
        const winner = candidates[0];
        // Capture extension guide lines from winning candidate for visual feedback
        if (winner._extGuides) {
            snapState.extensionGuides = winner._extGuides;
            delete winner._extGuides; // clean up internal property
        } else {
            snapState.extensionGuides = [];
        }
        return winner;
    }

    // ── Background grid snap (P5 — only if nothing else) ──
    if (snapState.gridSnap && CONFIG.gridVisible) {
        const minor = CONFIG.GRID_MINOR_MM;
        const da = engine.coords.drawArea;
        const snapX = Math.round((sheetPos.x - da.left) / minor) * minor + da.left;
        const snapY = Math.round((sheetPos.y - da.top) / minor) * minor + da.top;
        const d = dist2d(sheetPos.x, sheetPos.y, snapX, snapY);
        if (d < radiusMM) {
            return { x: snapX, y: snapY, type: SNAP_TYPES.GRID };
        }
    }

    return null;
}

/**
 * Apply orthogonal constraint relative to a reference point.
 * Forces the point to be purely horizontal or vertical from ref.
 */
// Track Shift key state globally
let _shiftDown = false;
window.addEventListener('keydown', (e) => { if (e.key === 'Shift') _shiftDown = true; });
window.addEventListener('keyup', (e) => { if (e.key === 'Shift') _shiftDown = false; });

/**
 * Apply angle constraint relative to a reference point.
 * - Ortho lock (O toggle): constrains to 0° or 90° (pure H/V)
 * - Shift held: constrains to nearest 45° increment (0°, 45°, 90°, 135°, etc.)
 * - Neither: no constraint
 */
function applyOrtho(sheetX, sheetY, refX, refY) {
    if (refX === undefined || refY === undefined) return { x: sheetX, y: sheetY };

    const useOrtho = snapState.orthoLock;
    const useShift = _shiftDown;

    if (!useOrtho && !useShift) return { x: sheetX, y: sheetY };

    const dx = sheetX - refX;
    const dy = sheetY - refY;
    const dist = Math.sqrt(dx * dx + dy * dy);
    if (dist < 0.001) return { x: sheetX, y: sheetY };

    if (useOrtho && !useShift) {
        // Pure H/V constraint (0° / 90°)
        if (Math.abs(dx) >= Math.abs(dy)) {
            return { x: sheetX, y: refY };
        } else {
            return { x: refX, y: sheetY };
        }
    }

    // Shift (or Shift+Ortho): snap to nearest 45° increment
    const angle = Math.atan2(dy, dx);
    const snap45 = Math.round(angle / (Math.PI / 4)) * (Math.PI / 4);
    return {
        x: refX + dist * Math.cos(snap45),
        y: refY + dist * Math.sin(snap45)
    };
}

/** Draw the snap indicator at the current snap point */
function drawSnapIndicator(ctx, eng) {
    const snap = snapState.activeSnap;
    if (!snap) return;

    const sp = eng.coords.sheetToScreen(snap.x, snap.y);
    const r = 6;

    ctx.save();
    ctx.strokeStyle = snap.type.color;
    ctx.fillStyle = snap.type.color;
    ctx.lineWidth = 2;

    switch (snap.type.symbol) {
        case 'square':
            ctx.strokeRect(sp.x - r, sp.y - r, r * 2, r * 2);
            break;
        case 'triangle':
            ctx.beginPath();
            ctx.moveTo(sp.x, sp.y - r);
            ctx.lineTo(sp.x + r, sp.y + r);
            ctx.lineTo(sp.x - r, sp.y + r);
            ctx.closePath();
            ctx.stroke();
            break;
        case 'cross':
            ctx.beginPath();
            ctx.moveTo(sp.x - r, sp.y - r); ctx.lineTo(sp.x + r, sp.y + r);
            ctx.moveTo(sp.x + r, sp.y - r); ctx.lineTo(sp.x - r, sp.y + r);
            ctx.stroke();
            break;
        case 'diamond':
            ctx.beginPath();
            ctx.moveTo(sp.x, sp.y - r);
            ctx.lineTo(sp.x + r, sp.y);
            ctx.lineTo(sp.x, sp.y + r);
            ctx.lineTo(sp.x - r, sp.y);
            ctx.closePath();
            ctx.stroke();
            break;
        case 'circle':
            ctx.beginPath();
            ctx.arc(sp.x, sp.y, r, 0, Math.PI * 2);
            ctx.stroke();
            break;
    }

    // Label
    if (eng.viewport.zoom > 0.5) {
        ctx.font = '9px "Segoe UI", Arial, sans-serif';
        ctx.fillStyle = snap.type.color;
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';
        ctx.fillText(snap.type.label, sp.x + r + 3, sp.y + r + 2);
    }

    ctx.restore();

    // ── Extension guide lines ──────────────────────────────
    // Draw subtle dashed lines from source element to snap point
    if (snapState.extensionGuides.length > 0) {
        ctx.save();
        ctx.setLineDash([4, 4]);
        ctx.lineWidth = 1;
        ctx.globalAlpha = 0.55;

        for (const guide of snapState.extensionGuides) {
            const from = eng.coords.sheetToScreen(guide.from.x, guide.from.y);
            const to = eng.coords.sheetToScreen(guide.to.x, guide.to.y);

            ctx.strokeStyle = SNAP_TYPES.EXTENSION.color;
            ctx.beginPath();
            ctx.moveTo(from.x, from.y);
            ctx.lineTo(to.x, to.y);
            ctx.stroke();

            // Small dot at source point to show what we're extending from
            ctx.fillStyle = SNAP_TYPES.EXTENSION.color;
            ctx.beginPath();
            ctx.arc(from.x, from.y, 3, 0, Math.PI * 2);
            ctx.fill();
        }

        ctx.restore();
    }
}

// Track mouse position for cursor previews
const cursorPos = { screenX: 0, screenY: 0 };

// Update snap on mouse move
container.addEventListener('mousemove', (e) => {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;
    cursorPos.screenX = sx;
    cursorPos.screenY = sy;
    snapState.activeSnap = findSnap(sx, sy);
    // If the winning snap isn't an extension type, clear guides
    if (!snapState.activeSnap || snapState.activeSnap.type !== SNAP_TYPES.EXTENSION) {
        if (!snapState.activeSnap || !snapState.activeSnap._extGuides) {
            snapState.extensionGuides = [];
        }
    }
    // Request render for cursor preview update
    if (activeTool === 'column' || activeTool === 'footing') {
        engine.requestRender();
    }
});

// ── Snap & Ortho Toolbar ─────────────────────────────────

const snapBtn = document.getElementById('btn-snap');
const orthoBtn = document.getElementById('btn-ortho');
const statusOrtho = document.getElementById('status-ortho');
const statusSnap = document.getElementById('status-snap');

function updateSnapOrthoUI() {
    if (snapState.enabled) {
        snapBtn.classList.add('active');
        statusSnap.className = 'ortho-badge on';
    } else {
        snapBtn.classList.remove('active');
        statusSnap.className = 'ortho-badge off';
    }
    if (snapState.orthoLock) {
        orthoBtn.classList.add('active');
        statusOrtho.className = 'ortho-badge on';
    } else {
        orthoBtn.classList.remove('active');
        statusOrtho.className = 'ortho-badge off';
    }
}

// Default: snap on, ortho off
snapState.enabled = true;
snapState.orthoLock = false;
updateSnapOrthoUI();

snapBtn.addEventListener('click', () => {
    snapState.enabled = !snapState.enabled;
    updateSnapOrthoUI();
    engine.requestRender();
});

orthoBtn.addEventListener('click', () => {
    snapState.orthoLock = !snapState.orthoLock;
    updateSnapOrthoUI();
});

// O = toggle ortho (snap toggle via ribbon button only)
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'o') {
        snapState.orthoLock = !snapState.orthoLock;
        updateSnapOrthoUI();
    }
});

// ── Auto-Link Toggle (status bar badge) ──────────────────
const statusAutoLink = document.getElementById('status-autolink');
if (statusAutoLink) {
    statusAutoLink.addEventListener('click', () => {
        autoLinkSettings.enabled = !autoLinkSettings.enabled;
        statusAutoLink.className = 'ortho-badge ' + (autoLinkSettings.enabled ? 'on' : 'off');
        statusAutoLink.title = autoLinkSettings.enabled
            ? 'Auto-link ON: walls auto-create strip footings, columns auto-create pad footings (click to disable)'
            : 'Auto-link OFF: no automatic footing creation (click to enable)';
    });
}

// ── Structural Grid Data ─────────────────────────────────

/**
 * Structural grids. Each grid is:
 * { id, axis: 'V'|'H', position: real-world mm, label: string }
 */
const structuralGrids = [];

// Per-axis label counters and scheme ('num' or 'alpha')
const gridLabelState = {
    V: { scheme: 'num', nextNum: 1, nextAlpha: 0 },
    H: { scheme: 'alpha', nextNum: 1, nextAlpha: 0 },
};

function nextGridLabel(axis) {
    const st = gridLabelState[axis];
    if (st.scheme === 'num') {
        return String(st.nextNum++);
    } else {
        const label = String.fromCharCode(65 + st.nextAlpha);
        st.nextAlpha++;
        return label;
    }
}

// ── Structural Grid Rendering ────────────────────────────

function drawStructuralGrids(ctx, eng) {
    if (structuralGrids.length === 0 && !gridToolState.active) return;

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;
    const da = coords.drawArea;
    const bubbleR = 4; // mm
    const bubbleOffset = 7; // mm outside drawing frame

    ctx.save();

    for (const grid of structuralGrids) {
        const isSelected = (gridToolState.selectedGrid === grid);

        // Grid line style — Revit convention: grey, thin, solid
        ctx.strokeStyle = isSelected ? '#2B7CD0' : '#808080';
        ctx.lineWidth = Math.max(0.5, (isSelected ? 0.3 : 0.18) * zoom);
        ctx.setLineDash([]);

        if (grid.axis === 'V') {
            const sx = da.left + grid.position / CONFIG.drawingScale;
            if (sx < da.left || sx > da.right) continue;

            const p1 = coords.sheetToScreen(sx, da.top);
            const p2 = coords.sheetToScreen(sx, da.bottom);

            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.stroke();

            // Bubble at top
            const bubbleCenter = coords.sheetToScreen(sx, da.top - bubbleOffset);
            drawGridBubble(ctx, bubbleCenter.x, bubbleCenter.y, bubbleR * zoom, grid.label, isSelected, zoom);

        } else {
            const sy = da.top + grid.position / CONFIG.drawingScale;
            if (sy < da.top || sy > da.bottom) continue;

            const p1 = coords.sheetToScreen(da.left, sy);
            const p2 = coords.sheetToScreen(da.right, sy);

            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.stroke();

            // Bubble at left
            const bubbleCenter = coords.sheetToScreen(da.left - bubbleOffset, sy);
            drawGridBubble(ctx, bubbleCenter.x, bubbleCenter.y, bubbleR * zoom, grid.label, isSelected, zoom);
        }
    }

    // Draw preview line if in grid placement mode
    if (gridToolState.active && gridToolState.previewPos !== null) {
        ctx.strokeStyle = '#2B7CD0';
        ctx.lineWidth = Math.max(1, 0.25 * zoom);
        ctx.setLineDash([4, 3]);

        if (gridToolState.axis === 'V') {
            const p1 = coords.sheetToScreen(gridToolState.previewPos, da.top);
            const p2 = coords.sheetToScreen(gridToolState.previewPos, da.bottom);
            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.stroke();
        } else {
            const p1 = coords.sheetToScreen(da.left, gridToolState.previewPos);
            const p2 = coords.sheetToScreen(da.right, gridToolState.previewPos);
            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.stroke();
        }
        ctx.setLineDash([]);
    }

    ctx.restore();
}

function drawGridBubble(ctx, cx, cy, r, label, isSelected, zoom) {
    // Circle
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.fillStyle = isSelected ? '#2B7CD0' : '#FFFFFF';
    ctx.fill();
    ctx.strokeStyle = isSelected ? '#2B7CD0' : '#808080';
    ctx.lineWidth = Math.max(0.8, 0.25 * zoom);
    ctx.stroke();

    // Label
    const fontSize = Math.max(7, r * 1.2);
    ctx.font = `bold ${fontSize}px "Segoe UI", Arial, sans-serif`;
    ctx.fillStyle = isSelected ? '#FFFFFF' : '#000000';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(label, cx, cy + 0.5);

    // Reset
    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';
}

// Register grid renderer
engine.onRender(drawStructuralGrids);
// Register snap indicator renderer (draws last, on top)
engine.onRender(drawSnapIndicator);

// ── Grid Placement Tool ──────────────────────────────────

const gridToolState = {
    active: false,
    axis: 'V',             // 'V' or 'H'
    previewPos: null,      // sheet-mm coordinate of preview line
    selectedGrid: null,    // for future editing
};

const gridBanner = document.getElementById('grid-banner');
const gridAxisLabel = document.getElementById('grid-axis-label');
const gridPlaceBtn = document.getElementById('btn-grid-place');
const selectBtn = document.getElementById('btn-select');
const gridLabelSelect = document.getElementById('grid-label-scheme');

/** Sync the label scheme dropdown to the current axis */
function syncGridLabelUI() {
    gridAxisLabel.textContent = gridToolState.axis === 'V' ? 'vertical' : 'horizontal';
    gridLabelSelect.value = gridLabelState[gridToolState.axis].scheme;
}

// Label scheme dropdown change
gridLabelSelect.addEventListener('change', () => {
    gridLabelState[gridToolState.axis].scheme = gridLabelSelect.value;
});

function activateGridTool() {
    gridToolState.active = true;
    gridBanner.classList.remove('hidden');
    syncGridLabelUI();
    container.style.cursor = 'crosshair';
    gridPlaceBtn.classList.add('active');
    selectBtn.classList.remove('active');
    document.getElementById('status-tool').textContent = 'Grids';
    engine.requestRender();
}

function deactivateGridTool() {
    gridToolState.active = false;
    gridToolState.previewPos = null;
    gridBanner.classList.add('hidden');
    container.style.cursor = '';
    gridPlaceBtn.classList.remove('active');
    selectBtn.classList.add('active');
    document.getElementById('status-tool').textContent = 'Select';
    engine.requestRender();
}

gridPlaceBtn.addEventListener('click', () => {
    if (gridToolState.active) {
        deactivateGridTool();
    } else {
        activateGridTool();
    }
});

selectBtn.addEventListener('click', () => {
    deactivateGridTool();
});

document.getElementById('grid-toggle-axis').addEventListener('click', () => {
    gridToolState.axis = gridToolState.axis === 'V' ? 'H' : 'V';
    syncGridLabelUI();
    engine.requestRender();
});

document.getElementById('grid-done').addEventListener('click', () => {
    deactivateGridTool();
});

// Keyboard shortcuts for grid tool
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;

    // Escape exits grid tool
    if (e.key === 'Escape' && gridToolState.active) {
        deactivateGridTool();
        return;
    }

    // Tab switches axis while in grid tool
    if (e.key === 'Tab' && gridToolState.active) {
        e.preventDefault();
        gridToolState.axis = gridToolState.axis === 'V' ? 'H' : 'V';
        syncGridLabelUI();
        engine.requestRender();
        return;
    }

    // Delete selected grid
    if ((e.key === 'Delete' || e.key === 'Backspace') && gridToolState.selectedGrid) {
        const idx = structuralGrids.indexOf(gridToolState.selectedGrid);
        if (idx !== -1) {
            const removed = gridToolState.selectedGrid;
            history.execute({
                description: 'Delete grid ' + removed.label,
                execute() {
                    const i = structuralGrids.indexOf(removed);
                    if (i !== -1) structuralGrids.splice(i, 1);
                },
                undo() { structuralGrids.push(removed); }
            });
            gridToolState.selectedGrid = null;
            engine.requestRender();
        }
    }
});

/**
 * PDF Line Snap — scans the PDF raster to find dark lines near the cursor.
 * For vertical grids: scans columns of pixels to find the darkest vertical line.
 * For horizontal grids: scans rows to find the darkest horizontal line.
 * Returns the sheet-mm position of the detected line, or null.
 */
function findPdfLineSnap(screenX, screenY, axis) {
    if (!pdfState.loaded || !pdfState.visible) return null;

    const cached = pdfState.pageCanvases[pdfState.currentPage];
    if (!cached) return null;

    const coords = engine.coords;
    const sheetPos = coords.screenToSheet(screenX, screenY);

    // Convert sheet position to PDF pixel position
    const pdfPixelX = ((sheetPos.x - pdfState.sheetX) / pdfState.sheetWidth) * cached.nativeWidth;
    const pdfPixelY = ((sheetPos.y - pdfState.sheetY) / pdfState.sheetHeight) * cached.nativeHeight;

    // Check bounds
    if (pdfPixelX < 0 || pdfPixelX >= cached.nativeWidth ||
        pdfPixelY < 0 || pdfPixelY >= cached.nativeHeight) return null;

    const pdfCtx = cached.canvas.getContext('2d');
    const searchRadiusPx = 20; // pixels in PDF image to search
    const darknessThreshold = 100; // pixel brightness below this = "dark"

    if (axis === 'V') {
        // Scan columns around cursor X, find darkest column
        const startCol = Math.max(0, Math.floor(pdfPixelX - searchRadiusPx));
        const endCol = Math.min(cached.nativeWidth - 1, Math.ceil(pdfPixelX + searchRadiusPx));
        const scanHeight = Math.min(60, cached.nativeHeight); // sample strip height
        const scanY = Math.max(0, Math.floor(pdfPixelY - scanHeight / 2));

        let imgData;
        try {
            imgData = pdfCtx.getImageData(startCol, scanY, endCol - startCol + 1, scanHeight);
        } catch (e) { return null; }

        const w = endCol - startCol + 1;
        let bestCol = -1;
        let bestDarkCount = 0;

        for (let col = 0; col < w; col++) {
            let darkCount = 0;
            for (let row = 0; row < scanHeight; row++) {
                const idx = (row * w + col) * 4;
                const brightness = (imgData.data[idx] + imgData.data[idx + 1] + imgData.data[idx + 2]) / 3;
                if (brightness < darknessThreshold) darkCount++;
            }
            if (darkCount > bestDarkCount && darkCount > scanHeight * 0.15) {
                bestDarkCount = darkCount;
                bestCol = col;
            }
        }

        if (bestCol >= 0) {
            // Convert back to sheet-mm
            const pdfX = startCol + bestCol;
            const sheetX = pdfState.sheetX + (pdfX / cached.nativeWidth) * pdfState.sheetWidth;
            return sheetX;
        }

    } else {
        // Scan rows around cursor Y, find darkest row
        const startRow = Math.max(0, Math.floor(pdfPixelY - searchRadiusPx));
        const endRow = Math.min(cached.nativeHeight - 1, Math.ceil(pdfPixelY + searchRadiusPx));
        const scanWidth = Math.min(60, cached.nativeWidth);
        const scanX = Math.max(0, Math.floor(pdfPixelX - scanWidth / 2));

        let imgData;
        try {
            imgData = pdfCtx.getImageData(scanX, startRow, scanWidth, endRow - startRow + 1);
        } catch (e) { return null; }

        const h = endRow - startRow + 1;
        let bestRow = -1;
        let bestDarkCount = 0;

        for (let row = 0; row < h; row++) {
            let darkCount = 0;
            for (let col = 0; col < scanWidth; col++) {
                const idx = (row * scanWidth + col) * 4;
                const brightness = (imgData.data[idx] + imgData.data[idx + 1] + imgData.data[idx + 2]) / 3;
                if (brightness < darknessThreshold) darkCount++;
            }
            if (darkCount > bestDarkCount && darkCount > scanWidth * 0.15) {
                bestDarkCount = darkCount;
                bestRow = row;
            }
        }

        if (bestRow >= 0) {
            const pdfY = startRow + bestRow;
            const sheetY = pdfState.sheetY + (pdfY / cached.nativeHeight) * pdfState.sheetHeight;
            return sheetY;
        }
    }

    return null;
}

// Mouse move — update preview line position
container.addEventListener('mousemove', (e) => {
    if (!gridToolState.active) return;

    const sheetPos = engine.getSheetPos(e);
    const da = engine.coords.drawArea;
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    // Priority: 1. Existing structural grid snap, 2. PDF line snap, 3. Raw cursor
    const snap = findSnap(sx, sy);

    if (gridToolState.axis === 'V') {
        if (snap && snap.type === SNAP_TYPES.GRIDLINE) {
            gridToolState.previewPos = snap.x;
        } else {
            // Try PDF line snap
            const pdfSnap = findPdfLineSnap(sx, sy, 'V');
            gridToolState.previewPos = pdfSnap !== null ? pdfSnap : sheetPos.x;
        }
        gridToolState.previewPos = Math.max(da.left, Math.min(da.right, gridToolState.previewPos));
    } else {
        if (snap && snap.type === SNAP_TYPES.GRIDLINE) {
            gridToolState.previewPos = snap.y;
        } else {
            const pdfSnap = findPdfLineSnap(sx, sy, 'H');
            gridToolState.previewPos = pdfSnap !== null ? pdfSnap : sheetPos.y;
        }
        gridToolState.previewPos = Math.max(da.top, Math.min(da.bottom, gridToolState.previewPos));
    }

    engine.requestRender();
});

// Mouse click — place grid or select existing
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (pdfState.calibrating) return;

    if (gridToolState.active) {
        // Place a new grid line
        const da = engine.coords.drawArea;
        let realPos;

        if (gridToolState.axis === 'V') {
            const sheetX = gridToolState.previewPos || engine.getSheetPos(e).x;
            realPos = (sheetX - da.left) * CONFIG.drawingScale;
        } else {
            const sheetY = gridToolState.previewPos || engine.getSheetPos(e).y;
            realPos = (sheetY - da.top) * CONFIG.drawingScale;
        }

        if (realPos < 0) return; // outside drawing area

        const newGrid = {
            id: generateId(),
            axis: gridToolState.axis,
            position: realPos,
            label: nextGridLabel(gridToolState.axis)
        };

        history.execute({
            description: 'Place grid ' + newGrid.label,
            execute() { structuralGrids.push(newGrid); },
            undo() {
                const i = structuralGrids.indexOf(newGrid);
                if (i !== -1) structuralGrids.splice(i, 1);
            }
        });

        engine.requestRender();
        return;
    }

    // Select mode — check if clicking near a grid line
    if (!gridToolState.active) {
        const sheetPos = engine.getSheetPos(e);
        const da = engine.coords.drawArea;
        const tolerance = 3 / engine.viewport.zoom; // 3px tolerance

        gridToolState.selectedGrid = null;

        for (const grid of structuralGrids) {
            if (grid.axis === 'V') {
                const sx = da.left + grid.position / CONFIG.drawingScale;
                if (Math.abs(sheetPos.x - sx) < tolerance &&
                    sheetPos.y >= da.top && sheetPos.y <= da.bottom) {
                    gridToolState.selectedGrid = grid;
                    break;
                }
            } else {
                const sy = da.top + grid.position / CONFIG.drawingScale;
                if (Math.abs(sheetPos.y - sy) < tolerance &&
                    sheetPos.x >= da.left && sheetPos.x <= da.right) {
                    gridToolState.selectedGrid = grid;
                    break;
                }
            }
        }

        engine.requestRender();
    }
});

// Expose grids to global app state
window._app.structuralGrids = structuralGrids;
window._app.snapState = snapState;

// ══════════════════════════════════════════════════════════
