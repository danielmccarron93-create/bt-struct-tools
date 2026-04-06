// ── 18-DETAILS-SYSTEM.JS ─────────────────────────────────
// Section Cuts, Detail Callouts & Details Sheet System
// ══════════════════════════════════════════════════════════
console.log('[18-details-system] LOADING...');

// ── PHASE 1: SHEET MODEL & DETAILS TAB ──────────────────

/**
 * The details system manages:
 * 1. Section markers on plans (AS 1100.501 style)
 * 2. Detail callouts on plans (circle + leader + label)
 * 3. A dedicated DETAILS sheet with viewports at 1:10
 */

const detailsSystem = {
    // All section cuts and detail callouts
    viewports: [],      // { id, type:'section'|'detail', number/letter, sourceLevel, ... }

    // Auto-incrementing counters
    nextSectionNumber: 1,
    nextDetailLetter: 'A',

    // Details sheet config
    detailScale: 10,    // 1:10 detail scale
    sheetNum: 'S201',   // Default details sheet number

    // Viewport layout config (A1 sheet)
    viewportCols: 3,
    viewportRows: 2,
    viewportPadding: 15,    // mm padding between viewports on sheet
    viewportTitleHeight: 12, // mm for title below each viewport

    // State
    activeSheet: 'plan',  // 'plan' or 'details'
};

// ── VIEWPORT EDITING STATE ──────────────────────────────────
// Interaction state machine for crop box editing, positioning, and scale/title changes on details sheet

const _detVpEditState = {
    mode: 'none',           // 'none' | 'dragging' | 'cropping' | 'titleEdit' | 'scaleEdit'
    activeViewport: null,    // viewport being edited
    activeVpIndex: -1,       // index in viewports array
    dragOffset: null,        // { dx, dy } offset from mouse to viewport origin in sheet-mm
    cropHandle: null,        // which handle is being dragged: 'n','s','e','w','ne','nw','se','sw'
    cropStartBox: null,      // crop box at start of drag
    cropStartMouse: null,    // mouse position at start of drag { x, y } in screen px
    titleEditEl: null,       // temporary input element for title editing
    scaleEditEl: null,       // temporary dropdown for scale editing
};

/** Get next detail letter and advance */
function getNextDetailLetter() {
    const letter = detailsSystem.nextDetailLetter;
    // Advance: A->B->...->Z->AA->AB...
    if (letter.length === 1 && letter < 'Z') {
        detailsSystem.nextDetailLetter = String.fromCharCode(letter.charCodeAt(0) + 1);
    } else if (letter === 'Z') {
        detailsSystem.nextDetailLetter = 'AA';
    } else {
        // Multi-char: increment last char
        const last = letter.slice(-1);
        if (last < 'Z') {
            detailsSystem.nextDetailLetter = letter.slice(0, -1) + String.fromCharCode(last.charCodeAt(0) + 1);
        }
    }
    return letter;
}

/** Get next section number and advance */
function getNextSectionNumber() {
    return detailsSystem.nextSectionNumber++;
}

/** Compute viewport grid positions on A1 details sheet */
function computeViewportLayout() {
    // CRITICAL: Use sheet dimensions directly, NOT engine.coords.drawArea
    // The drawArea in screen coords would be wrong for the details sheet
    const sheetW = CONFIG.SHEET_WIDTH_MM;
    const sheetH = CONFIG.SHEET_HEIGHT_MM;
    const cols = detailsSystem.viewportCols;
    const rows = detailsSystem.viewportRows;
    const pad = detailsSystem.viewportPadding;
    const titleH = detailsSystem.viewportTitleHeight;

    const totalW = sheetW - pad * 2;
    const totalH = sheetH - pad * 2;
    const cellW = (totalW - (cols - 1) * pad) / cols;
    const cellH = (totalH - (rows - 1) * pad) / rows;
    const drawH = cellH - titleH; // drawable area inside viewport (excluding title)

    const positions = [];
    for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
            positions.push({
                index: r * cols + c,
                x: pad + c * (cellW + pad),
                y: pad + r * (cellH + pad),
                width: cellW,
                height: drawH,
                titleY: pad + r * (cellH + pad) + drawH,
                titleHeight: titleH
            });
        }
    }
    return positions;
}

// ── PHASE 1: DETAILS TAB IN LEVEL BAR ───────────────────

// Add a "DETAILS" tab after the level tabs
function addDetailsTab() {
    const levelBar = document.getElementById('level-bar');
    if (!levelBar) return;

    // Check if details tab already exists
    if (document.getElementById('tab-details-sheet')) return;

    // Find the 3D view tab and insert before it
    const tab3d = document.getElementById('tab-3d-view');

    // Create separator
    const sep = document.createElement('div');
    sep.style.cssText = 'width:1px;height:22px;background:var(--border-default);margin:0 4px;';
    sep.id = 'details-tab-sep';

    // Create Details tab
    const detailsTab = document.createElement('button');
    detailsTab.id = 'tab-details-sheet';
    detailsTab.className = 'level-tab';
    detailsTab.title = 'Details Sheet (Sections & Details at 1:10)';
    detailsTab.innerHTML = 'DETAILS <span class="level-ht" id="details-count-badge" style="display:none;">0</span>';
    detailsTab.style.cssText = 'border-radius:var(--r-sm);border:1px solid var(--border-default);font-weight:600;';

    detailsTab.addEventListener('click', () => {
        if (detailsSystem.activeSheet === 'details') {
            // Switch back to plan
            switchToSheet('plan');
        } else {
            switchToSheet('details');
        }
    });

    if (tab3d) {
        levelBar.insertBefore(sep, tab3d);
        levelBar.insertBefore(detailsTab, tab3d);
    } else {
        levelBar.appendChild(sep);
        levelBar.appendChild(detailsTab);
    }

    updateDetailsCountBadge();
}

/** Update the badge showing viewport count */
function updateDetailsCountBadge() {
    const badge = document.getElementById('details-count-badge');
    if (!badge) return;
    const count = detailsSystem.viewports.length;
    if (count > 0) {
        badge.textContent = count;
        badge.style.display = '';
    } else {
        badge.style.display = 'none';
    }
}

/** Switch between plan view and details sheet */
function switchToSheet(mode) {
    detailsSystem.activeSheet = mode;

    const detailsTab = document.getElementById('tab-details-sheet');
    const levelTabs = document.querySelectorAll('.level-tab');

    if (mode === 'details') {
        // Deactivate all level tabs
        levelTabs.forEach(t => t.classList.remove('active'));
        if (detailsTab) detailsTab.classList.add('active');

        // Disable drawing tools
        setActiveTool('select');

        // Re-render with details sheet content
        engine.requestRender();
        updateStatusBar();
    } else {
        // Restore plan view
        if (detailsTab) detailsTab.classList.remove('active');
        buildLevelTabs(); // re-activates the correct level tab
        engine.requestRender();
        updateStatusBar();
    }
}

// Initialise the details tab on load
setTimeout(() => {
    try {
        addDetailsTab();
        console.log('[18-details-system] DETAILS tab created successfully');
    } catch(err) {
        console.error('[18-details-system] Error creating DETAILS tab:', err);
    }
}, 100);


// ══════════════════════════════════════════════════════════
// ── PHASE 2: SECTION MARKER TOOL (AS 1100.501) ──────────
// ══════════════════════════════════════════════════════════

/**
 * Upgraded section tool:
 * - Two-click placement (start + end of cut line)
 * - Creates a persistent section-marker element on the plan
 * - Auto-creates a viewport on the DETAILS sheet
 * - Renders AS 1100.501 style: long-dash-dot line, direction arrows,
 *   numbered circles at each end
 */

const sectionMarkerState = {
    placing: false,
    startPoint: null,    // sheet-mm
    currentEnd: null,    // sheet-mm (preview)
};

// Hook into existing section button — we're replacing the old section tool
// The old tool in 11-3d-engine.js opens a popup overlay.
// We'll override it to create persistent markers instead.

// Override the section tool activation
const _origSetActiveTool_det = setActiveTool;
// We'll intercept 'section' tool to use our new marker system
// but keep a 'section-legacy' option for the old popup

// Section mousedown handler (replaces old one in 11-3d-engine.js)
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'section') return;
    if (detailsSystem.activeSheet === 'details') return; // not on details sheet

    // Prevent the old section handler from firing
    e.stopImmediatePropagation();

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);

    if (sectionMarkerState.placing && sectionMarkerState.startPoint) {
        pos = applyOrtho(pos.x, pos.y, sectionMarkerState.startPoint.x, sectionMarkerState.startPoint.y);
    }

    if (!sectionMarkerState.placing) {
        // First click — start point
        sectionMarkerState.placing = true;
        sectionMarkerState.startPoint = pos;
    } else {
        // Second click — commit section marker
        const start = sectionMarkerState.startPoint;
        const end = pos;
        sectionMarkerState.placing = false;
        sectionMarkerState.startPoint = null;
        sectionMarkerState.currentEnd = null;

        commitSectionMarker(start, end);
        setActiveTool('select');
    }
}, true); // Use capture phase to fire BEFORE old handler

// Track mouse for section marker preview
container.addEventListener('mousemove', (e) => {
    if (activeTool === 'section' && sectionMarkerState.placing) {
        const rect = container.getBoundingClientRect();
        const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
        const snap = findSnap(sx, sy);
        let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
        if (sectionMarkerState.startPoint) {
            pos = applyOrtho(pos.x, pos.y, sectionMarkerState.startPoint.x, sectionMarkerState.startPoint.y);
        }
        sectionMarkerState.currentEnd = pos;
        engine.requestRender();
    }
}, true);

// Escape cancels section marker placement
window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && activeTool === 'section' && sectionMarkerState.placing) {
        sectionMarkerState.placing = false;
        sectionMarkerState.startPoint = null;
        sectionMarkerState.currentEnd = null;
        engine.requestRender();
    }
});

/** Commit a section marker to the plan and create a viewport */
function commitSectionMarker(startSheet, endSheet) {
    const realStart = engine.coords.sheetToReal(startSheet.x, startSheet.y);
    const realEnd = engine.coords.sheetToReal(endSheet.x, endSheet.y);

    // Check minimum length
    const dx = realEnd.x - realStart.x, dy = realEnd.y - realStart.y;
    const len = Math.sqrt(dx * dx + dy * dy);
    if (len < 200) return; // too short — at least 200mm real

    const sectionNum = getNextSectionNumber();
    const activeLevel = getActiveLevel();

    // Create the section marker element on the plan
    const marker = {
        id: generateId(),
        type: 'section-marker',
        layer: 'S-ANNO',
        level: activeLevel.id,
        x1: realStart.x, y1: realStart.y,
        x2: realEnd.x,   y2: realEnd.y,
        sectionNumber: sectionNum,
        sheetRef: detailsSystem.sheetNum,
        // View direction: perpendicular to line, pointing "right" when looking from start to end
        viewDirection: 'right',  // can be toggled
    };

    // Create the corresponding viewport
    const viewport = {
        id: generateId(),
        type: 'section',
        number: sectionNum,
        label: 'SECTION ' + sectionNum,
        markerId: marker.id,
        sourceLevel: activeLevel.id,
        cutStart: { x: realStart.x, y: realStart.y },
        cutEnd: { x: realEnd.x, y: realEnd.y },
        scale: detailsSystem.detailScale,

        // New properties for free positioning and crop box editing
        cropBox: null,      // { minAlong, maxAlong, minElev, maxElev } in real-world mm, null = auto-crop
        sheetPos: null,     // { x, y } in sheet-mm (top-left of viewport content area), null = use auto grid
        customTitle: null,  // custom title string, null = use default label
    };

    history.execute({
        description: 'Add section cut ' + sectionNum,
        execute() {
            project.elements.push(marker);
            detailsSystem.viewports.push(viewport);
        },
        undo() {
            const mi = project.elements.indexOf(marker);
            if (mi !== -1) project.elements.splice(mi, 1);
            const vi = detailsSystem.viewports.indexOf(viewport);
            if (vi !== -1) detailsSystem.viewports.splice(vi, 1);
        }
    });

    updateDetailsCountBadge();
    engine.requestRender();
    console.log('[Section] Created Section ' + sectionNum + ' on ' + activeLevel.name);
}


// ══════════════════════════════════════════════════════════
// ── PHASE 3: DETAIL CALLOUT TOOL ────────────────────────
// ══════════════════════════════════════════════════════════

/**
 * Detail callout: click area of interest, then click to place label.
 * Creates:
 * - A dashed circle on the plan around the area
 * - A leader line to a detail bubble
 * - A viewport entry on the DETAILS sheet
 */

const detailCalloutState = {
    placing: false,
    step: 0,            // 0=waiting, 1=placed centre (picking label pos), 2=done
    centrePoint: null,   // sheet-mm — centre of detail circle
    labelPoint: null,    // sheet-mm — where the label bubble goes
    currentEnd: null,    // preview
    radius: null,        // real-world mm radius of detail circle
};

// ── Section split-button: dropdown toggle + click handlers ──
(function initSectionSplitButton() {
    const splitContainer = document.getElementById('btn-section-split');
    const mainBtn = document.getElementById('btn-section');
    const dropdownArrow = document.getElementById('btn-section-dropdown');
    const splitMenu = document.getElementById('section-split-menu');
    const menuSectionBtn = document.getElementById('btn-section-menu-section');
    const menuDetailBtn = document.getElementById('btn-detail-callout');

    if (!splitContainer || !mainBtn || !dropdownArrow || !splitMenu) {
        console.warn('[18-details-system] Split button elements not found');
        return;
    }

    // Main button click → activate section tool directly
    mainBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        splitMenu.classList.remove('open');
        setActiveTool('section');
    });

    // Dropdown arrow → toggle menu
    dropdownArrow.addEventListener('click', (e) => {
        e.stopPropagation();
        splitMenu.classList.toggle('open');
    });

    // Menu item: Section Cut
    if (menuSectionBtn) {
        menuSectionBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            splitMenu.classList.remove('open');
            setActiveTool('section');
        });
    }

    // Menu item: Detail Circle
    if (menuDetailBtn) {
        menuDetailBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            splitMenu.classList.remove('open');
            setActiveTool('detail');
        });
    }

    // Close menu when clicking elsewhere
    document.addEventListener('click', () => {
        splitMenu.classList.remove('open');
    });

    console.log('[18-details-system] Section split-button initialised');
})();

// Shift+D = detail tool (reuse D for chain dim, Shift+D for detail)
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'D' && e.shiftKey) {
        e.preventDefault();
        setActiveTool('detail');
    }
});

// Detail callout mouse handler
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'detail') return;
    if (detailsSystem.activeSheet === 'details') return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    const pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);

    if (detailCalloutState.step === 0) {
        // First click: centre of detail area
        detailCalloutState.placing = true;
        detailCalloutState.step = 1;
        detailCalloutState.centrePoint = pos;
        detailCalloutState.radius = 1500; // default 1500mm real-world radius
    } else if (detailCalloutState.step === 1) {
        // Second click: label position
        detailCalloutState.labelPoint = pos;
        detailCalloutState.step = 0;
        detailCalloutState.placing = false;

        commitDetailCallout(detailCalloutState.centrePoint, pos, detailCalloutState.radius);
        detailCalloutState.centrePoint = null;
        detailCalloutState.labelPoint = null;
        detailCalloutState.currentEnd = null;
        setActiveTool('select');
    }
});

// Mouse move for detail preview
container.addEventListener('mousemove', (e) => {
    if (activeTool === 'detail' && detailCalloutState.placing) {
        const rect = container.getBoundingClientRect();
        const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
        const snap = findSnap(sx, sy);
        detailCalloutState.currentEnd = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
        engine.requestRender();
    }
});

// Escape cancels
window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && activeTool === 'detail' && detailCalloutState.placing) {
        detailCalloutState.placing = false;
        detailCalloutState.step = 0;
        detailCalloutState.centrePoint = null;
        detailCalloutState.currentEnd = null;
        engine.requestRender();
    }
});

// Scroll wheel adjusts detail circle radius during placement
container.addEventListener('wheel', (e) => {
    if (activeTool === 'detail' && detailCalloutState.step === 1) {
        e.preventDefault();
        const delta = e.deltaY > 0 ? -200 : 200; // 200mm increments
        detailCalloutState.radius = Math.max(500, Math.min(5000, detailCalloutState.radius + delta));
        engine.requestRender();
    }
}, { passive: false });

/** Commit a detail callout */
function commitDetailCallout(centreSheet, labelSheet, radiusReal) {
    const realCentre = engine.coords.sheetToReal(centreSheet.x, centreSheet.y);
    const realLabel = engine.coords.sheetToReal(labelSheet.x, labelSheet.y);
    const activeLevel = getActiveLevel();

    const detailLetter = getNextDetailLetter();

    // Create the detail callout element
    const callout = {
        id: generateId(),
        type: 'detail-callout',
        layer: 'S-ANNO',
        level: activeLevel.id,
        cx: realCentre.x, cy: realCentre.y,     // circle centre
        lx: realLabel.x,  ly: realLabel.y,       // label position
        radius: radiusReal,
        detailLetter: detailLetter,
        sheetRef: detailsSystem.sheetNum,
    };

    // Create viewport
    const viewport = {
        id: generateId(),
        type: 'detail',
        letter: detailLetter,
        label: 'DETAIL ' + detailLetter,
        markerId: callout.id,
        sourceLevel: activeLevel.id,
        centre: { x: realCentre.x, y: realCentre.y },
        radius: radiusReal,
        scale: detailsSystem.detailScale,

        // New properties for free positioning and crop box editing
        cropBox: null,      // { minAlong, maxAlong, minElev, maxElev } for detail viewports not typically used
        sheetPos: null,     // { x, y } in sheet-mm (top-left of viewport content area), null = use auto grid
        customTitle: null,  // custom title string, null = use default label
    };

    history.execute({
        description: 'Add detail callout ' + detailLetter,
        execute() {
            project.elements.push(callout);
            detailsSystem.viewports.push(viewport);
        },
        undo() {
            const ci = project.elements.indexOf(callout);
            if (ci !== -1) project.elements.splice(ci, 1);
            const vi = detailsSystem.viewports.indexOf(viewport);
            if (vi !== -1) detailsSystem.viewports.splice(vi, 1);
        }
    });

    updateDetailsCountBadge();
    engine.requestRender();
    console.log('[Detail] Created Detail ' + detailLetter + ' on ' + activeLevel.name);
}


// ══════════════════════════════════════════════════════════
// ── RENDERING: SECTION MARKERS ON PLAN ──────────────────
// ══════════════════════════════════════════════════════════

/**
 * Draw section markers in AS 1100.501 style:
 * - Long dash-dot-dot line (heavy lineweight)
 * - Short perpendicular tails with arrowheads showing view direction
 * - Numbered circles at each end: section number (top), sheet ref (bottom)
 *
 * CRITICAL: Only draws when activeSheet === 'plan'
 */
function drawSectionMarkers(ctx, eng) {
    // AIRTIGHT GUARD: if not on plan, do nothing
    if (detailsSystem.activeSheet !== 'plan') return;

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;
    const activeId = getActiveLevel().id;

    // Draw persistent section markers
    for (const el of project.elements) {
        if (el.type !== 'section-marker') continue;
        if (el.level !== activeId) continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const p1 = coords.realToScreen(el.x1, el.y1);
        const p2 = coords.realToScreen(el.x2, el.y2);

        const dx = p2.x - p1.x, dy = p2.y - p1.y;
        const len = Math.sqrt(dx * dx + dy * dy);
        if (len < 5) continue;

        // Unit vectors
        const ux = dx / len, uy = dy / len;         // along line
        const nx = -uy, ny = ux;                      // perpendicular (view direction)

        const lineColor = '#CC0000';
        const circleRadius = Math.max(5, 3.5 * zoom);
        const tailLen = Math.max(10, 6 * zoom);
        const arrowSize = Math.max(4, 2.5 * zoom);
        const isSelected = (selectedElement === el || (selectedElements && selectedElements.includes(el)));

        // ── Section cut line: long-dash-dot pattern ──
        ctx.strokeStyle = isSelected ? '#2B7CD0' : lineColor;
        ctx.lineWidth = Math.max(1.5, 0.7 * zoom);

        // Extend line slightly past the circles
        const extP1 = { x: p1.x - ux * (circleRadius + 2), y: p1.y - uy * (circleRadius + 2) };
        const extP2 = { x: p2.x + ux * (circleRadius + 2), y: p2.y + uy * (circleRadius + 2) };

        // Dash-dot-dot pattern
        const dashLen = Math.max(8, 4 * zoom);
        const dotLen = Math.max(1.5, 0.8 * zoom);
        const gap = Math.max(3, 1.5 * zoom);
        ctx.setLineDash([dashLen, gap, dotLen, gap, dotLen, gap]);
        ctx.beginPath();
        ctx.moveTo(extP1.x, extP1.y);
        ctx.lineTo(extP2.x, extP2.y);
        ctx.stroke();
        ctx.setLineDash([]);

        // ── Perpendicular tails with arrowheads at each end ──
        const viewDir = el.viewDirection === 'left' ? -1 : 1;

        // Tail at start
        const t1Start = { x: p1.x, y: p1.y };
        const t1End = { x: p1.x + nx * tailLen * viewDir, y: p1.y + ny * tailLen * viewDir };
        ctx.lineWidth = Math.max(1.5, 0.7 * zoom);
        ctx.beginPath();
        ctx.moveTo(t1Start.x, t1Start.y);
        ctx.lineTo(t1End.x, t1End.y);
        ctx.stroke();

        // Arrowhead at tail end (start)
        drawArrowhead(ctx, t1Start.x, t1Start.y, t1End.x, t1End.y, arrowSize, isSelected ? '#2B7CD0' : lineColor);

        // Tail at end
        const t2Start = { x: p2.x, y: p2.y };
        const t2End = { x: p2.x + nx * tailLen * viewDir, y: p2.y + ny * tailLen * viewDir };
        ctx.beginPath();
        ctx.moveTo(t2Start.x, t2Start.y);
        ctx.lineTo(t2End.x, t2End.y);
        ctx.stroke();

        drawArrowhead(ctx, t2Start.x, t2Start.y, t2End.x, t2End.y, arrowSize, isSelected ? '#2B7CD0' : lineColor);

        // ── Numbered circles at each end ──
        const circleOffset = circleRadius + tailLen + 4;
        const c1 = { x: p1.x + nx * circleOffset * viewDir, y: p1.y + ny * circleOffset * viewDir };
        const c2 = { x: p2.x + nx * circleOffset * viewDir, y: p2.y + ny * circleOffset * viewDir };

        drawSectionCircle(ctx, c1.x, c1.y, circleRadius, el.sectionNumber, el.sheetRef, isSelected ? '#2B7CD0' : lineColor, zoom);
        drawSectionCircle(ctx, c2.x, c2.y, circleRadius, el.sectionNumber, el.sheetRef, isSelected ? '#2B7CD0' : lineColor, zoom);
    }

    // ── Preview for section marker being placed ──
    if (activeTool === 'section' && sectionMarkerState.placing && sectionMarkerState.startPoint && sectionMarkerState.currentEnd) {
        const sp = coords.sheetToScreen(sectionMarkerState.startPoint.x, sectionMarkerState.startPoint.y);
        const ep = coords.sheetToScreen(sectionMarkerState.currentEnd.x, sectionMarkerState.currentEnd.y);

        // Preview line
        ctx.strokeStyle = '#CC000080';
        ctx.lineWidth = 2;
        const dashLen = Math.max(8, 4 * zoom);
        const dotLen = Math.max(1.5, 0.8 * zoom);
        const gap = Math.max(3, 1.5 * zoom);
        ctx.setLineDash([dashLen, gap, dotLen, gap, dotLen, gap]);
        ctx.beginPath();
        ctx.moveTo(sp.x, sp.y);
        ctx.lineTo(ep.x, ep.y);
        ctx.stroke();
        ctx.setLineDash([]);

        // Preview circles
        const previewRadius = Math.max(5, 3.5 * zoom);
        const nextNum = detailsSystem.nextSectionNumber;
        drawSectionCircle(ctx, sp.x, sp.y - previewRadius - 8, previewRadius, nextNum, detailsSystem.sheetNum, '#CC000080', zoom);
        drawSectionCircle(ctx, ep.x, ep.y - previewRadius - 8, previewRadius, nextNum, detailsSystem.sheetNum, '#CC000080', zoom);

        // Start dot
        ctx.fillStyle = '#CC0000';
        ctx.beginPath();
        ctx.arc(sp.x, sp.y, 3, 0, Math.PI * 2);
        ctx.fill();
    }
}

/** Draw an arrowhead at the tip of a line */
function drawArrowhead(ctx, x1, y1, x2, y2, size, color) {
    const dx = x2 - x1, dy = y2 - y1;
    const len = Math.sqrt(dx * dx + dy * dy);
    if (len < 1) return;
    const ux = dx / len, uy = dy / len;

    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    ctx.lineTo(x2 - ux * size - uy * size * 0.4, y2 - uy * size + ux * size * 0.4);
    ctx.lineTo(x2 - ux * size + uy * size * 0.4, y2 - uy * size - ux * size * 0.4);
    ctx.closePath();
    ctx.fill();
}

/** Draw a section/detail reference circle with number and sheet ref */
function drawSectionCircle(ctx, cx, cy, radius, topText, bottomText, color, zoom) {
    // Background circle
    ctx.fillStyle = '#FFFFFF';
    ctx.beginPath();
    ctx.arc(cx, cy, radius, 0, Math.PI * 2);
    ctx.fill();

    // Circle outline
    ctx.strokeStyle = color;
    ctx.lineWidth = Math.max(1, 0.5 * zoom);
    ctx.beginPath();
    ctx.arc(cx, cy, radius, 0, Math.PI * 2);
    ctx.stroke();

    // Horizontal divider
    ctx.beginPath();
    ctx.moveTo(cx - radius, cy);
    ctx.lineTo(cx + radius, cy);
    ctx.stroke();

    // Top text (section number / detail letter) — bold
    const topFont = Math.max(7, 2.8 * zoom);
    ctx.font = `bold ${topFont}px "Segoe UI", Arial, sans-serif`;
    ctx.fillStyle = color;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(String(topText), cx, cy - radius * 0.32);

    // Bottom text (sheet reference)
    const botFont = Math.max(6, 2.2 * zoom);
    ctx.font = `${botFont}px "Segoe UI", Arial, sans-serif`;
    ctx.fillText(String(bottomText), cx, cy + radius * 0.35);
}


// ══════════════════════════════════════════════════════════
// ── RENDERING: DETAIL CALLOUTS ON PLAN ──────────────────
// ══════════════════════════════════════════════════════════

/**
 * Draw detail callouts (dashed circles + leader + bubbles).
 *
 * CRITICAL: Only draws when activeSheet === 'plan'
 */
function drawDetailCallouts(ctx, eng) {
    // AIRTIGHT GUARD: if not on plan, do nothing
    if (detailsSystem.activeSheet !== 'plan') return;

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;
    const activeId = getActiveLevel().id;

    for (const el of project.elements) {
        if (el.type !== 'detail-callout') continue;
        if (el.level !== activeId) continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        const cp = coords.realToScreen(el.cx, el.cy);    // circle centre
        const lp = coords.realToScreen(el.lx, el.ly);    // label position

        const isSelected = (selectedElement === el || (selectedElements && selectedElements.includes(el)));
        const color = isSelected ? '#2B7CD0' : '#CC0000';

        // ── Dashed circle around area of interest ──
        const radiusSheet = el.radius / CONFIG.drawingScale;  // convert real mm to sheet mm
        const radiusPx = radiusSheet * zoom;

        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(1, 0.5 * zoom);
        ctx.setLineDash([4 * zoom / 2, 2 * zoom / 2]);
        ctx.beginPath();
        ctx.arc(cp.x, cp.y, radiusPx, 0, Math.PI * 2);
        ctx.stroke();
        ctx.setLineDash([]);

        // ── Leader line from circle edge to label ──
        const ldx = lp.x - cp.x, ldy = lp.y - cp.y;
        const lLen = Math.sqrt(ldx * ldx + ldy * ldy);
        if (lLen > radiusPx + 5) {
            const startX = cp.x + (ldx / lLen) * radiusPx;
            const startY = cp.y + (ldy / lLen) * radiusPx;

            ctx.strokeStyle = color;
            ctx.lineWidth = Math.max(0.8, 0.4 * zoom);
            ctx.beginPath();
            ctx.moveTo(startX, startY);
            ctx.lineTo(lp.x, lp.y);
            ctx.stroke();
        }

        // ── Detail bubble (circle with letter and sheet ref) ──
        const bubbleRadius = Math.max(5, 3.5 * zoom);
        drawSectionCircle(ctx, lp.x, lp.y, bubbleRadius, el.detailLetter, el.sheetRef, color, zoom);
    }

    // ── Preview for detail callout being placed ──
    if (activeTool === 'detail' && detailCalloutState.placing && detailCalloutState.centrePoint) {
        const cp = coords.sheetToScreen(detailCalloutState.centrePoint.x, detailCalloutState.centrePoint.y);

        // Preview dashed circle
        const radiusSheet = detailCalloutState.radius / CONFIG.drawingScale;
        const radiusPx = radiusSheet * zoom;

        ctx.strokeStyle = '#CC000060';
        ctx.lineWidth = 1.5;
        ctx.setLineDash([6, 3]);
        ctx.beginPath();
        ctx.arc(cp.x, cp.y, radiusPx, 0, Math.PI * 2);
        ctx.stroke();
        ctx.setLineDash([]);

        // Radius label
        ctx.font = `${Math.max(9, 3 * zoom)}px "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = '#CC0000';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText('r=' + detailCalloutState.radius + 'mm', cp.x, cp.y - radiusPx - 4);
        ctx.fillText('(scroll to resize)', cp.x, cp.y - radiusPx - 4 - Math.max(12, 3.5 * zoom));

        if (detailCalloutState.currentEnd && detailCalloutState.step === 1) {
            const ep = coords.sheetToScreen(detailCalloutState.currentEnd.x, detailCalloutState.currentEnd.y);

            // Preview leader line
            const ldx = ep.x - cp.x, ldy = ep.y - cp.y;
            const lLen = Math.sqrt(ldx * ldx + ldy * ldy);
            if (lLen > radiusPx) {
                const startX = cp.x + (ldx / lLen) * radiusPx;
                const startY = cp.y + (ldy / lLen) * radiusPx;
                ctx.strokeStyle = '#CC000060';
                ctx.lineWidth = 1;
                ctx.beginPath();
                ctx.moveTo(startX, startY);
                ctx.lineTo(ep.x, ep.y);
                ctx.stroke();
            }

            // Preview bubble
            const bubbleRadius = Math.max(5, 3.5 * zoom);
            drawSectionCircle(ctx, ep.x, ep.y, bubbleRadius, detailsSystem.nextDetailLetter, detailsSystem.sheetNum, '#CC000060', zoom);
        }
    }
}


// ══════════════════════════════════════════════════════════
// ── AUTO-CROP CALCULATION ───────────────────────────────────
// ══════════════════════════════════════════════════════════

/**
 * Calculate the auto-crop box for a section viewport.
 * Collects all elements cut by the section, finds their bounding box,
 * and adds margin on all sides.
 *
 * Returns { minAlong, maxAlong, minElev, maxElev } in real-world mm,
 * or null if no elements found.
 */
function calculateAutoCropBox(vp) {
    if (vp.type !== 'section') return null;

    const cutStart = vp.cutStart;
    const cutEnd = vp.cutEnd;

    // Section line direction
    const sdx = cutEnd.x - cutStart.x;
    const sdy = cutEnd.y - cutStart.y;
    const sectionLen = Math.sqrt(sdx * sdx + sdy * sdy);
    if (sectionLen < 100) return null;

    // Unit vector along section line
    const ux = sdx / sectionLen;
    const uy = sdy / sectionLen;
    const nx = -uy, ny = ux; // normal to section line
    const cutTolerance = 500 * CONFIG.drawingScale * 0.01;

    /** Test whether a line segment is crossed by the section line */
    function testLineCut(x1, y1, x2, y2) {
        const d1x = x1 - cutStart.x, d1y = y1 - cutStart.y;
        const d2x = x2 - cutStart.x, d2y = y2 - cutStart.y;
        const perp1 = d1x * nx + d1y * ny;
        const perp2 = d2x * nx + d2y * ny;
        const absPerp1 = Math.abs(perp1), absPerp2 = Math.abs(perp2);
        const crossesLine = (perp1 * perp2 <= 0);
        const nearLine = (absPerp1 < cutTolerance || absPerp2 < cutTolerance);
        if (!crossesLine && !nearLine) return { hit: false };
        let along;
        if (crossesLine && Math.abs(perp1 - perp2) > 0.01) {
            const t = perp1 / (perp1 - perp2);
            const ix = x1 + t * (x2 - x1), iy = y1 + t * (y2 - y1);
            along = (ix - cutStart.x) * ux + (iy - cutStart.y) * uy;
        } else {
            const along1 = d1x * ux + d1y * uy, along2 = d2x * ux + d2y * uy;
            along = absPerp1 < absPerp2 ? along1 : along2;
        }
        if (along < -cutTolerance || along > sectionLen + cutTolerance) return { hit: false };
        return { hit: true, along };
    }

    // Collect elements cut by this section (same logic as renderSectionViewport)
    const sourceLevelId = vp.sourceLevel;
    const sourceLv = levelSystem.levels.find(l => l.id === sourceLevelId);
    if (!sourceLv) return null;
    const sourceLvIdx = levelSystem.levels.findIndex(l => l.id === sourceLevelId);

    let minAlong = Infinity, maxAlong = -Infinity;
    let minElev = Infinity, maxElev = -Infinity;
    let hasElements = false;

    for (const el of project.elements) {
        if (el.level && el.level !== sourceLevelId) continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        if (el.type === 'column') {
            const dx = el.x - cutStart.x, dy = el.y - cutStart.y;
            const along = dx * ux + dy * uy;
            const perp = Math.abs(dx * (-uy) + dy * ux);

            if (perp < cutTolerance && along > -cutTolerance && along < sectionLen + cutTolerance) {
                const dir = el.extends || 'below';
                let botElev, topElev;
                if (dir === 'below' && sourceLvIdx > 0) {
                    botElev = levelSystem.levels[sourceLvIdx - 1].elevation;
                    topElev = sourceLv.elevation;
                } else if (dir === 'above' && sourceLvIdx < levelSystem.levels.length - 1) {
                    botElev = sourceLv.elevation;
                    topElev = levelSystem.levels[sourceLvIdx + 1].elevation;
                } else if (dir === 'both') {
                    botElev = sourceLvIdx > 0 ? levelSystem.levels[sourceLvIdx - 1].elevation : sourceLv.elevation;
                    topElev = sourceLvIdx < levelSystem.levels.length - 1 ? levelSystem.levels[sourceLvIdx + 1].elevation : sourceLv.elevation + sourceLv.height;
                } else {
                    botElev = sourceLv.elevation;
                    topElev = sourceLv.elevation + sourceLv.height;
                }
                const colHalfWidth = (el.size || 89) / 2;
                minAlong = Math.min(minAlong, along - colHalfWidth);
                maxAlong = Math.max(maxAlong, along + colHalfWidth);
                minElev = Math.min(minElev, botElev);
                maxElev = Math.max(maxElev, topElev);
                hasElements = true;
            }
        }

        // Beams: type 'line' on S-BEAM layer
        if (el.type === 'line' && el.layer === 'S-BEAM') {
            const hit = testLineCut(el.x1, el.y1, el.x2, el.y2);
            if (hit.hit) {
                minAlong = Math.min(minAlong, hit.along - 150);
                maxAlong = Math.max(maxAlong, hit.along + 150);
                minElev = Math.min(minElev, sourceLv.elevation);
                maxElev = Math.max(maxElev, sourceLv.elevation + 500);
                hasElements = true;
            }
        }

        // Walls: type 'wall' on S-WALL layer (also handle legacy 'line' on S-WALL)
        if ((el.type === 'wall' || (el.type === 'line' && el.layer === 'S-WALL'))) {
            const hit = testLineCut(el.x1, el.y1, el.x2, el.y2);
            if (hit.hit) {
                const thickness = project.scheduleTypes?.wall?.[el.typeRef]?.thickness || el.thickness || 200;
                const height = sourceLv.height * 0.8;
                minAlong = Math.min(minAlong, hit.along - thickness / 2);
                maxAlong = Math.max(maxAlong, hit.along + thickness / 2);
                minElev = Math.min(minElev, sourceLv.elevation);
                maxElev = Math.max(maxElev, sourceLv.elevation + height);
                hasElements = true;
            }
        }

        if (el.type === 'polyline' && el.layer === 'S-SLAB' && (el.closed || (el.points && el.points.length > 3))) {
            for (const pt of el.points) {
                const dx = pt.x - cutStart.x, dy = pt.y - cutStart.y;
                const along = dx * ux + dy * uy;
                const perp = Math.abs(dx * (-uy) + dy * ux);
                if (perp < cutTolerance && along > 0 && along < sectionLen) {
                    minAlong = Math.min(minAlong, 0);
                    maxAlong = Math.max(maxAlong, sectionLen);
                    minElev = Math.min(minElev, sourceLv.elevation);
                    maxElev = Math.max(maxElev, sourceLv.elevation + (el.slabThickness || 200));
                    hasElements = true;
                    break;
                }
            }
        }

        // Footings
        if (el.type === 'footing') {
            const dx = el.x - cutStart.x, dy = el.y - cutStart.y;
            const along = dx * ux + dy * uy;
            const perp = Math.abs(dx * (-uy) + dy * ux);
            if (perp < cutTolerance && along > -cutTolerance && along < sectionLen + cutTolerance) {
                const ftgType = project.scheduleTypes?.padfooting?.[el.typeRef] || {};
                const ftgW = ftgType.width || el.width || 600;
                const ftgD = ftgType.depth || el.depth || 300;
                const setdown = el.depthBelowFSL || ftgType.setdown || 200;
                minAlong = Math.min(minAlong, along - ftgW / 2);
                maxAlong = Math.max(maxAlong, along + ftgW / 2);
                minElev = Math.min(minElev, sourceLv.elevation - setdown);
                maxElev = Math.max(maxElev, sourceLv.elevation);
                hasElements = true;
            }
        }

        // Strip footings (type: 'stripFooting' on S-FTNG layer)
        if (el.type === 'stripFooting') {
            const hit = testLineCut(el.x1, el.y1, el.x2, el.y2);
            if (hit.hit) {
                const sfType = project.scheduleTypes?.stripfooting?.[el.typeRef] || {};
                const sfW = sfType.width || el.footingWidth || 400;
                minAlong = Math.min(minAlong, hit.along - sfW / 2);
                maxAlong = Math.max(maxAlong, hit.along + sfW / 2);
                minElev = Math.min(minElev, sourceLv.elevation - (sfType.setdown || 200));
                maxElev = Math.max(maxElev, sourceLv.elevation);
                hasElements = true;
            }
        }
    }

    if (!hasElements) return null;

    // Add margin: 300mm on all sides
    const margin = 300;
    return {
        minAlong: minAlong - margin,
        maxAlong: maxAlong + margin,
        minElev: minElev - margin,
        maxElev: maxElev + margin
    };
}


// ══════════════════════════════════════════════════════════
// ── PHASE 4: DETAILS SHEET RENDERER ─────────────────────
// ══════════════════════════════════════════════════════════

/**
 * When the DETAILS tab is active, this replaces the normal plan rendering.
 * It draws:
 * - Viewport grid layout on the A1 sheet
 * - Each viewport's content (section cross-sections or detail plan views)
 * - Viewport titles and scale labels
 *
 * CRITICAL: Only draws when activeSheet === 'details'
 */
function drawDetailsSheet(ctx, eng) {
    // AIRTIGHT GUARD: if not on details sheet, do nothing
    if (detailsSystem.activeSheet !== 'details') return;

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;
    const gridPositions = computeViewportLayout();

    if (detailsSystem.viewports.length === 0) {
        // Empty state — show instructions
        const centre = coords.sheetToScreen(CONFIG.SHEET_WIDTH_MM / 2, CONFIG.SHEET_HEIGHT_MM / 2 - 20);
        ctx.font = `${Math.max(14, 5 * zoom)}px "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = '#999999';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('No sections or details yet.', centre.x, centre.y);
        ctx.font = `${Math.max(11, 3.5 * zoom)}px "Segoe UI", Arial, sans-serif`;
        ctx.fillText('Use the Section or Detail tool on a plan level to create viewports.', centre.x, centre.y + Math.max(20, 7 * zoom));
        return;
    }

    // Render each viewport
    for (let i = 0; i < detailsSystem.viewports.length; i++) {
        const vp = detailsSystem.viewports[i];

        // Determine viewport position and size
        let vpSheetX, vpSheetY, vpW_mm, vpH_mm;

        if (vp.sheetPos) {
            // Viewport has a saved position — use it
            vpSheetX = vp.sheetPos.x;
            vpSheetY = vp.sheetPos.y;

            // Calculate size based on crop box or content
            if (vp.type === 'section') {
                if (vp.cropBox) {
                    // Size from crop box
                    const cropW_mm = (vp.cropBox.maxAlong - vp.cropBox.minAlong) / vp.scale;
                    const cropH_mm = (vp.cropBox.maxElev - vp.cropBox.minElev) / vp.scale;
                    vpW_mm = cropW_mm;
                    vpH_mm = cropH_mm;
                } else {
                    // Auto-calculate crop box first
                    if (!vp.cropBox) {
                        vp.cropBox = calculateAutoCropBox(vp) || {
                            minAlong: 0, maxAlong: 5000,
                            minElev: 0, maxElev: 3000
                        };
                    }
                    const cropW_mm = (vp.cropBox.maxAlong - vp.cropBox.minAlong) / vp.scale;
                    const cropH_mm = (vp.cropBox.maxElev - vp.cropBox.minElev) / vp.scale;
                    vpW_mm = cropW_mm;
                    vpH_mm = cropH_mm;
                }
            } else {
                // Detail viewport: fixed size based on radius
                vpW_mm = vp.radius * 2 / vp.scale;
                vpH_mm = vp.radius * 2 / vp.scale;
            }
        } else {
            // Use auto grid layout and save the position
            if (i < gridPositions.length) {
                const pos = gridPositions[i];
                vpSheetX = pos.x;
                vpSheetY = pos.y;
                vpW_mm = pos.width;
                vpH_mm = pos.height;
                vp.sheetPos = { x: vpSheetX, y: vpSheetY };
            } else {
                continue; // No position available
            }
        }

        // Convert viewport bounds to screen coords
        const tl = coords.sheetToScreen(vpSheetX, vpSheetY);
        const br = coords.sheetToScreen(vpSheetX + vpW_mm, vpSheetY + vpH_mm);
        const vpW_px = br.x - tl.x;
        const vpH_px = br.y - tl.y;

        if (vpW_px < 10 || vpH_px < 10) continue; // Too small to render

        // ── Viewport content (clipped to viewport bounds) ──
        ctx.save();
        ctx.beginPath();
        ctx.rect(tl.x, tl.y, vpW_px, vpH_px);
        ctx.clip();

        if (vp.type === 'section') {
            renderSectionViewport(ctx, eng, vp, tl.x, tl.y, vpW_px, vpH_px);
        } else if (vp.type === 'detail') {
            renderDetailViewport(ctx, eng, vp, tl.x, tl.y, vpW_px, vpH_px);
        }

        ctx.restore();

        // ── Viewport border ──
        const isActiveVp = (_detVpEditState.activeViewport === vp);
        if (!isActiveVp) {
            // Subtle border for non-selected viewports
            ctx.strokeStyle = '#CCCCCC';
            ctx.lineWidth = 0.5;
            ctx.strokeRect(tl.x, tl.y, vpW_px, vpH_px);
        }

        // ── Draw selection outline and crop handles on active viewport ──
        if (isActiveVp) {
            // Dashed selection border
            ctx.setLineDash([4, 3]);
            ctx.strokeStyle = '#2266FF';
            ctx.lineWidth = 1.5;
            ctx.strokeRect(tl.x, tl.y, vpW_px, vpH_px);
            ctx.setLineDash([]);

            // Draw 8 handles: 4 corners + 4 edge midpoints
            const handleSize = 7;
            const handles = [
                { id: 'nw', x: tl.x, y: tl.y },
                { id: 'n',  x: tl.x + vpW_px / 2, y: tl.y },
                { id: 'ne', x: tl.x + vpW_px, y: tl.y },
                { id: 'e',  x: tl.x + vpW_px, y: tl.y + vpH_px / 2 },
                { id: 'se', x: tl.x + vpW_px, y: tl.y + vpH_px },
                { id: 's',  x: tl.x + vpW_px / 2, y: tl.y + vpH_px },
                { id: 'sw', x: tl.x, y: tl.y + vpH_px },
                { id: 'w',  x: tl.x, y: tl.y + vpH_px / 2 },
            ];
            for (const h of handles) {
                // White filled square with blue border
                ctx.fillStyle = '#FFFFFF';
                ctx.fillRect(h.x - handleSize / 2, h.y - handleSize / 2, handleSize, handleSize);
                ctx.strokeStyle = '#2266FF';
                ctx.lineWidth = 1.2;
                ctx.strokeRect(h.x - handleSize / 2, h.y - handleSize / 2, handleSize, handleSize);
            }
        }

        // ── Viewport title and scale labels ──
        const titleFont = Math.max(8, 3 * zoom);
        const titleY_px = tl.y + vpH_px + Math.max(4, 1.5 * zoom);

        // Get title text
        const titleText = vp.customTitle || vp.label;

        ctx.font = `bold ${titleFont}px "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = '#000000';
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';
        ctx.fillText(titleText, tl.x, titleY_px);

        // Scale label (clickable in scaleEdit mode)
        ctx.font = `${Math.max(7, 2.5 * zoom)}px "Segoe UI", Arial, sans-serif`;
        const scaleText = '1:' + vp.scale;
        ctx.fillText(scaleText, tl.x, titleY_px + titleFont + 2);
    }
}


// ══════════════════════════════════════════════════════════
// ── PHASE 5: SECTION VIEWPORT RENDERING ─────────────────
// ══════════════════════════════════════════════════════════

/**
 * Render a section cross-section into a viewport rectangle.
 * Shows structural members in their true profile.
 * If vp.cropBox is set, uses those bounds; otherwise auto-calculates.
 */
function renderSectionViewport(ctx, eng, vp, vpX, vpY, vpW, vpH) {
    const cutStart = vp.cutStart;
    const cutEnd = vp.cutEnd;

    // Section line direction
    const sdx = cutEnd.x - cutStart.x;
    const sdy = cutEnd.y - cutStart.y;
    const sectionLen = Math.sqrt(sdx * sdx + sdy * sdy);
    if (sectionLen < 100) return;

    // Unit vector along section line
    const ux = sdx / sectionLen;
    const uy = sdy / sectionLen;
    // Normal to section line (perpendicular unit vector)
    const nx = -uy, ny = ux;
    const cutTolerance = 500 * CONFIG.drawingScale * 0.01;

    /**
     * Test whether a line element (beam/wall/strip) is intersected by the section line.
     * Returns { hit: true, along } if the section line crosses or comes within tolerance
     * of the element's line segment. 'along' is the position along the section line.
     */
    function testLineCut(x1, y1, x2, y2) {
        // Perpendicular distance from each endpoint to the section line
        const d1x = x1 - cutStart.x, d1y = y1 - cutStart.y;
        const d2x = x2 - cutStart.x, d2y = y2 - cutStart.y;
        const perp1 = d1x * nx + d1y * ny; // signed perpendicular distance
        const perp2 = d2x * nx + d2y * ny;

        // Check if section line crosses through the element
        // (endpoints on opposite sides, or one/both within tolerance)
        const absPerp1 = Math.abs(perp1), absPerp2 = Math.abs(perp2);
        const crossesLine = (perp1 * perp2 <= 0); // opposite signs = crosses
        const nearLine = (absPerp1 < cutTolerance || absPerp2 < cutTolerance);

        if (!crossesLine && !nearLine) return { hit: false };

        // Find the intersection/closest point along the section line
        let along;
        if (crossesLine && Math.abs(perp1 - perp2) > 0.01) {
            // Interpolate to find the crossing point
            const t = perp1 / (perp1 - perp2); // 0..1 parameter along element
            const ix = x1 + t * (x2 - x1);
            const iy = y1 + t * (y2 - y1);
            along = (ix - cutStart.x) * ux + (iy - cutStart.y) * uy;
        } else {
            // Near-parallel or both close — use the closer endpoint
            const along1 = d1x * ux + d1y * uy;
            const along2 = d2x * ux + d2y * uy;
            along = absPerp1 < absPerp2 ? along1 : along2;
        }

        // Check 'along' is within the section line extent
        if (along < -cutTolerance || along > sectionLen + cutTolerance) return { hit: false };

        return { hit: true, along };
    }

    // ── Collect elements cut by this section (SOURCE LEVEL ONLY) ──
    const sectionElements = [];
    const sourceLevelId = vp.sourceLevel;
    const sourceLv = levelSystem.levels.find(l => l.id === sourceLevelId);
    if (!sourceLv) return;
    const sourceLvIdx = levelSystem.levels.findIndex(l => l.id === sourceLevelId);

    for (const el of project.elements) {
        // Only include elements on the source level (treat missing level as matching)
        if (el.level && el.level !== sourceLevelId) continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        if (el.type === 'column') {
            const dx = el.x - cutStart.x, dy = el.y - cutStart.y;
            const along = dx * ux + dy * uy;
            const perp = Math.abs(dx * (-uy) + dy * ux);

            if (perp < cutTolerance && along > -cutTolerance && along < sectionLen + cutTolerance) {
                const dir = el.extends || 'below';
                let botElev, topElev;
                if (dir === 'below' && sourceLvIdx > 0) {
                    botElev = levelSystem.levels[sourceLvIdx - 1].elevation;
                    topElev = sourceLv.elevation;
                } else if (dir === 'above' && sourceLvIdx < levelSystem.levels.length - 1) {
                    botElev = sourceLv.elevation;
                    topElev = levelSystem.levels[sourceLvIdx + 1].elevation;
                } else if (dir === 'both') {
                    botElev = sourceLvIdx > 0 ? levelSystem.levels[sourceLvIdx - 1].elevation : sourceLv.elevation;
                    topElev = sourceLvIdx < levelSystem.levels.length - 1 ? levelSystem.levels[sourceLvIdx + 1].elevation : sourceLv.elevation + sourceLv.height;
                } else {
                    botElev = sourceLv.elevation;
                    topElev = sourceLv.elevation + sourceLv.height;
                }
                sectionElements.push({
                    type: 'column', along, tag: el.tag, size: el.size || 89,
                    botElev, topElev
                });
            }
        }

        // Beams: type 'line' on S-BEAM layer
        if (el.type === 'line' && el.layer === 'S-BEAM') {
            const hit = testLineCut(el.x1, el.y1, el.x2, el.y2);
            if (hit.hit) {
                sectionElements.push({
                    type: 'beam',
                    along: hit.along, elev: sourceLv.elevation,
                    height: 500,
                    width: 300,
                    typeRef: el.typeRef,
                });
            }
        }

        // Walls: type 'wall' on S-WALL layer (also handle legacy 'line' on S-WALL)
        if ((el.type === 'wall' || (el.type === 'line' && el.layer === 'S-WALL'))) {
            const hit = testLineCut(el.x1, el.y1, el.x2, el.y2);
            if (hit.hit) {
                const thickness = project.scheduleTypes?.wall?.[el.typeRef]?.thickness || el.thickness || 200;
                sectionElements.push({
                    type: 'wall',
                    along: hit.along, elev: sourceLv.elevation,
                    height: sourceLv.height * 0.8,
                    width: thickness,
                    typeRef: el.typeRef,
                });
            }
        }

        if (el.type === 'polyline' && el.layer === 'S-SLAB' && (el.closed || (el.points && el.points.length > 3))) {
            for (const pt of el.points) {
                const dx = pt.x - cutStart.x, dy = pt.y - cutStart.y;
                const along = dx * ux + dy * uy;
                const perp = Math.abs(dx * (-uy) + dy * ux);
                if (perp < cutTolerance && along > 0 && along < sectionLen) {
                    sectionElements.push({
                        type: 'slab', along: 0, width: sectionLen,
                        elev: sourceLv.elevation, thickness: el.slabThickness || 200,
                    });
                    break;
                }
            }
        }

        // Pad footings
        if (el.type === 'footing') {
            const dx = el.x - cutStart.x, dy = el.y - cutStart.y;
            const along = dx * ux + dy * uy;
            const perp = Math.abs(dx * (-uy) + dy * ux);
            if (perp < cutTolerance && along > -cutTolerance && along < sectionLen + cutTolerance) {
                const ftgType = project.scheduleTypes?.padfooting?.[el.typeRef] || {};
                const ftgW = ftgType.width || el.width || 600;
                const ftgD = ftgType.depth || el.depth || 300;
                const setdown = el.depthBelowFSL || ftgType.setdown || 200;
                sectionElements.push({
                    type: 'footing', along,
                    width: ftgW, depth: ftgD,
                    elev: sourceLv.elevation - setdown,
                    tag: el.tag || el.typeRef || '',
                });
            }
        }

        // Strip footings (type: 'stripFooting' on S-FTNG layer)
        if (el.type === 'stripFooting') {
            const hit = testLineCut(el.x1, el.y1, el.x2, el.y2);
            if (hit.hit) {
                const sfType = project.scheduleTypes?.stripfooting?.[el.typeRef] || {};
                sectionElements.push({
                    type: 'strip', along: hit.along,
                    width: sfType.width || el.footingWidth || 400,
                    depth: sfType.depth || el.footingDepth || 300,
                    elev: sourceLv.elevation - (sfType.setdown || 200),
                });
            }
        }
    }

    // ── Determine elevation and along ranges using crop box if available ──
    const detailScale = vp.scale || 10;  // 1:10
    const zoom = eng.viewport.zoom;      // pixels per sheet-mm

    let minAlong, maxAlong, minElev, maxElev;

    if (vp.cropBox) {
        // Use crop box bounds
        minAlong = vp.cropBox.minAlong;
        maxAlong = vp.cropBox.maxAlong;
        minElev = vp.cropBox.minElev;
        maxElev = vp.cropBox.maxElev;
    } else {
        // Auto-calculate from visible elements and level bounds
        const levelBelow = sourceLvIdx > 0 ? levelSystem.levels[sourceLvIdx - 1].elevation : sourceLv.elevation - 500;
        const levelAbove = sourceLvIdx < levelSystem.levels.length - 1
            ? levelSystem.levels[sourceLvIdx + 1].elevation
            : sourceLv.elevation + (sourceLv.height || 2700);
        const margin = 300; // 300mm real margin above/below
        minElev = levelBelow - margin;
        maxElev = levelAbove + margin;
        minAlong = 0;
        maxAlong = sectionLen;
    }

    // ── Render at detail scale ──
    // pixels per real-world mm = zoom / detailScale
    // e.g. at 1:10 and zoom=1, 10mm real = 1mm sheet = 1px
    const pxPerRealMm = zoom / detailScale;

    // True 1:10 scale
    let scale = pxPerRealMm;
    const contentW = (maxAlong - minAlong) * scale;
    const contentH = (maxElev - minElev) * scale;

    // If content overflows viewport, scale down to fit
    if (contentW > vpW * 0.95 || contentH > vpH * 0.95) {
        scale = Math.min((vpW * 0.9) / (maxAlong - minAlong), (vpH * 0.9) / (maxElev - minElev));
    }

    // Compute origin — centre content in viewport
    const finalW = (maxAlong - minAlong) * scale;
    const finalH = (maxElev - minElev) * scale;
    const finalOx = vpX + (vpW - finalW) / 2;
    const finalOy = vpY + (vpH + finalH) / 2;

    function toVP(along, elev) {
        return {
            x: finalOx + (along - minAlong) * scale,
            y: finalOy - (elev - minElev) * scale
        };
    }

    // ── Level lines (source level and adjacent only) ──
    const levelsToShow = [sourceLv];
    if (sourceLvIdx > 0) levelsToShow.push(levelSystem.levels[sourceLvIdx - 1]);
    if (sourceLvIdx < levelSystem.levels.length - 1) levelsToShow.push(levelSystem.levels[sourceLvIdx + 1]);

    for (const lv of levelsToShow) {
        const y = toVP(0, lv.elevation).y;
        ctx.strokeStyle = '#AAAAAA';
        ctx.lineWidth = 0.5;
        ctx.setLineDash([4, 3]);
        ctx.beginPath();
        ctx.moveTo(finalOx - 5, y);
        ctx.lineTo(finalOx + finalW + 5, y);
        ctx.stroke();
        ctx.setLineDash([]);

        // Level label
        const fontSize = Math.max(7, 2.5 * zoom);
        ctx.font = `${fontSize}px "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = '#888888';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(lv.name, finalOx - 8, y);
    }

    // ── Ground line (only if ground level 0 is within view range) ──
    if (minElev <= 0 && maxElev >= 0) {
        const groundY = toVP(0, 0).y;
        ctx.fillStyle = '#F0EDE8';
        ctx.fillRect(finalOx - 5, groundY, finalW + 10, finalOy - groundY);
        ctx.strokeStyle = '#333333';
        ctx.lineWidth = 1.5;
        ctx.beginPath();
        ctx.moveTo(finalOx - 5, groundY);
        ctx.lineTo(finalOx + finalW + 5, groundY);
        ctx.stroke();
    }

    // ── Section elements ──
    for (const el of sectionElements) {
        if (el.type === 'column') {
            const p1 = toVP(el.along, el.botElev);
            const p2 = toVP(el.along, el.topElev);
            const colW = Math.max(el.size * scale, 4);

            ctx.fillStyle = 'rgba(43,102,170,0.15)';
            ctx.fillRect(p1.x - colW / 2, p2.y, colW, p1.y - p2.y);
            ctx.strokeStyle = '#2266AA';
            ctx.lineWidth = 1;
            ctx.strokeRect(p1.x - colW / 2, p2.y, colW, p1.y - p2.y);

            // X cross
            ctx.beginPath();
            ctx.moveTo(p1.x - colW / 2, p2.y);
            ctx.lineTo(p1.x + colW / 2, p1.y);
            ctx.moveTo(p1.x + colW / 2, p2.y);
            ctx.lineTo(p1.x - colW / 2, p1.y);
            ctx.stroke();

            if (el.tag) {
                ctx.font = `bold ${Math.max(7, 2.5)}px "Segoe UI", Arial, sans-serif`;
                ctx.fillStyle = '#2266AA';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                ctx.fillText(el.tag, p1.x, p2.y - 2);
            }
        }

        if (el.type === 'beam') {
            const p = toVP(el.along, el.elev);

            // Look up exact section properties from Liberty catalogue
            const beamSize = project.scheduleTypes?.beam?.[el.typeRef]?.size || '';
            const catalogueProfile = (typeof lookupSteelSection === 'function') ? lookupSteelSection(beamSize) : null;

            if (catalogueProfile && typeof drawSteelSection === 'function') {
                // Use accurate catalogue-based drawing with root fillets
                // drawSteelSection expects cy at section CENTRE, but p.y is at beam soffit (bottom)
                // So shift up by half the section depth
                const sectionDepthPx = (catalogueProfile.d || catalogueProfile.od || el.height) * scale;
                drawSteelSection(ctx, p.x, p.y - sectionDepthPx / 2, catalogueProfile, scale);
            } else {
                // Fallback: old ratio-based drawing
                const bH = Math.max(el.height * scale, 3);
                const beamProfile = parseBeamSize(beamSize);
                const bW = Math.max(beamProfile.width * scale, 4);
                drawIBeamSection(ctx, p.x, p.y, bW, bH, beamProfile, scale);
            }
        }

        if (el.type === 'wall') {
            const p = toVP(el.along, el.elev);
            const wW = Math.max(el.width * scale, 3);
            const wH = Math.max(el.height * scale, 4);
            ctx.fillStyle = 'rgba(100,100,100,0.15)';
            ctx.fillRect(p.x - wW / 2, p.y - wH, wW, wH);
            ctx.strokeStyle = '#666666';
            ctx.lineWidth = 1;
            ctx.strokeRect(p.x - wW / 2, p.y - wH, wW, wH);

            // Concrete hatch (diagonal lines)
            ctx.save();
            ctx.beginPath();
            ctx.rect(p.x - wW / 2, p.y - wH, wW, wH);
            ctx.clip();
            ctx.strokeStyle = '#AAAAAA';
            ctx.lineWidth = 0.5;
            for (let d = -wH; d < wW + wH; d += 4) {
                ctx.beginPath();
                ctx.moveTo(p.x - wW / 2 + d, p.y);
                ctx.lineTo(p.x - wW / 2 + d + wH, p.y - wH);
                ctx.stroke();
            }
            ctx.restore();
        }

        if (el.type === 'slab') {
            const p = toVP(el.along, el.elev);
            const sW = sectionLen * scale;
            const sH = Math.max(el.thickness * scale, 3);
            ctx.fillStyle = 'rgba(180,180,180,0.2)';
            ctx.fillRect(finalOx, p.y - sH, sW, sH);
            ctx.strokeStyle = '#333333';
            ctx.lineWidth = 1;
            ctx.strokeRect(finalOx, p.y - sH, sW, sH);

            // Concrete hatch
            ctx.save();
            ctx.beginPath();
            ctx.rect(finalOx, p.y - sH, sW, sH);
            ctx.clip();
            ctx.strokeStyle = '#CCCCCC';
            ctx.lineWidth = 0.3;
            for (let d = 0; d < sW + sH; d += 4) {
                ctx.beginPath();
                ctx.moveTo(finalOx + d, p.y);
                ctx.lineTo(finalOx + d - sH * 1.5, p.y - sH);
                ctx.stroke();
            }
            ctx.restore();
        }

        if (el.type === 'footing') {
            const p = toVP(el.along, el.elev);
            const fW = Math.max(el.width * scale, 6);
            const fH = Math.max(el.depth * scale, 4);
            ctx.fillStyle = 'rgba(150,150,150,0.2)';
            ctx.fillRect(p.x - fW / 2, p.y, fW, fH);
            ctx.strokeStyle = '#555555';
            ctx.lineWidth = 1;
            ctx.strokeRect(p.x - fW / 2, p.y, fW, fH);

            if (el.tag) {
                ctx.font = `${Math.max(6, 2)}px "Segoe UI", Arial, sans-serif`;
                ctx.fillStyle = '#555555';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'top';
                ctx.fillText(el.tag, p.x, p.y + fH + 2);
            }
        }

        if (el.type === 'strip') {
            const p = toVP(el.along, el.elev);
            const sW = Math.max(el.width * scale, 4);
            const sH = Math.max(el.depth * scale, 3);
            ctx.fillStyle = 'rgba(150,150,150,0.2)';
            ctx.fillRect(p.x - sW / 2, p.y, sW, sH);
            ctx.strokeStyle = '#555555';
            ctx.lineWidth = 1;
            ctx.strokeRect(p.x - sW / 2, p.y, sW, sH);
        }
    }
}

/** Parse beam size string like "360UB56.7" and return profile data */
function parseBeamSize(sizeStr) {
    if (!sizeStr) return { type: 'unknown', depth: 200, width: 100, flangeThickness: 10, webThickness: 6 };

    const match = sizeStr.match(/^(\d+)(UB|UC|PFC|SHS|RHS)/i);
    if (!match) return { type: 'unknown', depth: 200, width: 100, flangeThickness: 10, webThickness: 6 };

    const depth = parseInt(match[1]);
    const typeCode = match[2].toUpperCase();

    if (typeCode === 'UB') {
        return {
            type: 'UB',
            depth: depth,
            width: depth * 0.55,
            flangeThickness: depth / 20,
            webThickness: depth / 40
        };
    } else if (typeCode === 'UC') {
        return {
            type: 'UC',
            depth: depth,
            width: depth,  // roughly square
            flangeThickness: depth / 15,
            webThickness: depth / 30
        };
    } else if (typeCode === 'PFC') {
        return {
            type: 'PFC',
            depth: depth,
            width: depth * 0.4,
            flangeThickness: depth / 20,
            webThickness: depth / 40
        };
    } else if (typeCode === 'SHS' || typeCode === 'RHS') {
        return {
            type: typeCode,
            depth: depth,
            width: depth * 0.8,
            flangeThickness: depth / 12,
            webThickness: depth / 12
        };
    }

    return { type: 'unknown', width: 100, flangeThickness: 10 };
}

/** Draw an I-beam section profile at the given position */
function drawIBeamSection(ctx, cx, cy, width, height, profile, scale) {
    const flangeT = profile.flangeThickness * scale;
    const webT = profile.webThickness * scale;
    const webW = Math.max(webT, 1);
    const flangeW = width;
    const h = height;

    ctx.fillStyle = 'rgba(170,102,43,0.15)';
    ctx.strokeStyle = '#AA6600';
    ctx.lineWidth = 0.8;

    // Top flange
    ctx.fillRect(cx - flangeW / 2, cy - h, flangeW, flangeT);
    ctx.strokeRect(cx - flangeW / 2, cy - h, flangeW, flangeT);

    // Web
    ctx.fillRect(cx - webW / 2, cy - h + flangeT, webW, h - 2 * flangeT);
    ctx.strokeRect(cx - webW / 2, cy - h + flangeT, webW, h - 2 * flangeT);

    // Bottom flange
    ctx.fillRect(cx - flangeW / 2, cy - flangeT, flangeW, flangeT);
    ctx.strokeRect(cx - flangeW / 2, cy - flangeT, flangeW, flangeT);
}


// ══════════════════════════════════════════════════════════
// ── PHASE 6: DETAIL VIEWPORT RENDERING ──────────────────
// ══════════════════════════════════════════════════════════

/**
 * Render a detail viewport — a zoomed-in plan view of the circular area.
 * Shows all elements within the detail radius at the viewport's scale.
 */
function renderDetailViewport(ctx, eng, vp, vpX, vpY, vpW, vpH) {
    const centre = vp.centre;
    const radius = vp.radius;

    // The detail shows a square region centred on the detail point
    // Side = 2 * radius in real-world mm
    const realSide = radius * 2;

    // Scale: how many px per real-world mm inside this viewport
    const pxPerMm = Math.min(vpW, vpH) / realSide;

    // Origin: top-left of the real-world region
    const realLeft = centre.x - radius;
    const realTop = centre.y - radius;

    function toVP(rx, ry) {
        return {
            x: vpX + (rx - realLeft) * pxPerMm,
            y: vpY + (ry - realTop) * pxPerMm
        };
    }

    // ── Draw clipping circle ──
    // Show a subtle dashed circle at the detail boundary
    const cpVP = toVP(centre.x, centre.y);
    const radiusPx = radius * pxPerMm;

    // Light background
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(vpX, vpY, vpW, vpH);

    // ── Draw elements from source level within the radius ──
    const levelId = vp.sourceLevel;

    for (const el of project.elements) {
        if (el.level !== levelId) continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        // Check if element is within the detail area
        if (el.type === 'column') {
            const distX = el.x - centre.x, distY = el.y - centre.y;
            const dist = Math.sqrt(distX * distX + distY * distY);
            if (dist > radius * 1.2) continue; // allow slight overflow

            const p = toVP(el.x, el.y);
            const colSize = (el.size || 89) * pxPerMm;

            // Column rectangle with X
            ctx.fillStyle = 'rgba(43,102,170,0.15)';
            ctx.fillRect(p.x - colSize / 2, p.y - colSize / 2, colSize, colSize);
            ctx.strokeStyle = '#2266AA';
            ctx.lineWidth = 1;
            ctx.strokeRect(p.x - colSize / 2, p.y - colSize / 2, colSize, colSize);

            ctx.beginPath();
            ctx.moveTo(p.x - colSize / 2, p.y - colSize / 2);
            ctx.lineTo(p.x + colSize / 2, p.y + colSize / 2);
            ctx.moveTo(p.x + colSize / 2, p.y - colSize / 2);
            ctx.lineTo(p.x - colSize / 2, p.y + colSize / 2);
            ctx.stroke();

            if (el.tag) {
                ctx.font = `bold ${Math.max(8, colSize * 0.4)}px "Segoe UI", Arial, sans-serif`;
                ctx.fillStyle = '#2266AA';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                ctx.fillText(el.tag, p.x, p.y - colSize / 2 - 2);
            }
        }

        if (el.type === 'line') {
            const mx = (el.x1 + el.x2) / 2, my = (el.y1 + el.y2) / 2;
            const dist = Math.sqrt((mx - centre.x) ** 2 + (my - centre.y) ** 2);
            if (dist > radius * 1.5) continue;

            const p1 = toVP(el.x1, el.y1);
            const p2 = toVP(el.x2, el.y2);

            ctx.strokeStyle = layer.color || '#333333';
            ctx.lineWidth = Math.max(1, (layer.lineWeight || 0.5) * pxPerMm);
            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.stroke();
        }

        if (el.type === 'footing') {
            const dist = Math.sqrt((el.x - centre.x) ** 2 + (el.y - centre.y) ** 2);
            if (dist > radius * 1.2) continue;

            const ftgType = project.scheduleTypes?.padfooting?.[el.typeRef] || {};
            const w = (ftgType.width || el.width || 600) * pxPerMm;
            const l = (ftgType.length || el.length || 600) * pxPerMm;
            const p = toVP(el.x, el.y);

            ctx.fillStyle = 'rgba(150,150,150,0.1)';
            ctx.fillRect(p.x - w / 2, p.y - l / 2, w, l);
            ctx.strokeStyle = '#555555';
            ctx.lineWidth = 1;
            ctx.strokeRect(p.x - w / 2, p.y - l / 2, w, l);

            // Dashed diagonal
            ctx.setLineDash([3, 2]);
            ctx.beginPath();
            ctx.moveTo(p.x - w / 2, p.y - l / 2);
            ctx.lineTo(p.x + w / 2, p.y + l / 2);
            ctx.moveTo(p.x + w / 2, p.y - l / 2);
            ctx.lineTo(p.x - w / 2, p.y + l / 2);
            ctx.stroke();
            ctx.setLineDash([]);
        }

        if (el.type === 'polyline' && el.points && el.points.length >= 2) {
            // Check if any point is within range
            const inRange = el.points.some(pt => {
                const d = Math.sqrt((pt.x - centre.x) ** 2 + (pt.y - centre.y) ** 2);
                return d < radius * 1.5;
            });
            if (!inRange) continue;

            const pts = el.points.map(pt => toVP(pt.x, pt.y));

            ctx.strokeStyle = layer.color || '#333333';
            ctx.lineWidth = Math.max(1, (layer.lineWeight || 0.5) * pxPerMm);

            if (el.layer === 'S-SLAB') {
                ctx.fillStyle = 'rgba(200,200,200,0.1)';
            }

            ctx.beginPath();
            ctx.moveTo(pts[0].x, pts[0].y);
            for (let j = 1; j < pts.length; j++) {
                ctx.lineTo(pts[j].x, pts[j].y);
            }
            if (el.closed) ctx.closePath();
            if (el.layer === 'S-SLAB') ctx.fill();
            ctx.stroke();
        }
    }

    // ── Detail boundary circle (dashed) ──
    ctx.strokeStyle = '#AAAAAA';
    ctx.lineWidth = 0.5;
    ctx.setLineDash([4, 3]);
    ctx.beginPath();
    ctx.arc(cpVP.x, cpVP.y, radiusPx, 0, Math.PI * 2);
    ctx.stroke();
    ctx.setLineDash([]);
}


// ══════════════════════════════════════════════════════════
// ── REGISTER RENDER CALLBACKS ───────────────────────────
// ══════════════════════════════════════════════════════════

// Draw section markers and detail callouts on plan
engine.onRender(drawSectionMarkers);
engine.onRender(drawDetailCallouts);

// Draw details sheet content (only active when on details tab)
engine.onRender(drawDetailsSheet);

// ── PATCH ENGINE RENDER: suppress plan-mode callbacks on details sheet ──
// When the details sheet is active, only our three callbacks above should fire.
// All other render callbacks (iso view, grids, schedules, beam tags, etc.)
// would draw on top of / interfere with the details sheet.
(function patchEngineRender() {
    const _detailsCallbacks = new Set([drawSectionMarkers, drawDetailCallouts, drawDetailsSheet]);
    const _origRender = engine.render.bind(engine);

    engine.render = function() {
        if (detailsSystem.activeSheet !== 'details') {
            // Normal plan mode — run all callbacks as usual
            _origRender();
            return;
        }

        // Details mode — replicate the engine's render() but only run our callbacks
        const ctx = engine.ctx;
        const dpr = engine.dpr;
        const w = engine.width;
        const h = engine.height;

        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        ctx.clearRect(0, 0, w, h);

        // Background
        ctx.fillStyle = engine.config.CANVAS_BG;
        ctx.fillRect(0, 0, w, h);

        // Draw sheet background and frame
        engine._drawSheet(ctx);
        engine._drawDrawingFrame(ctx);

        // Only run details-system callbacks (suppress iso view, grids, schedules, etc.)
        for (const fn of engine._renderCallbacks) {
            if (_detailsCallbacks.has(fn)) {
                fn(ctx, engine);
            }
        }
    };

    console.log('[18-details-system] Engine render patched for details sheet isolation');
})();


// ══════════════════════════════════════════════════════════
// ── PHASE 7: SAVE / LOAD INTEGRATION ───────────────────
// ══════════════════════════════════════════════════════════

// Patch project.toJSON to include details system data
const _detSysOrigToJSON = project.toJSON.bind(project);
project.toJSON = function () {
    const data = _detSysOrigToJSON();
    data.detailsSystem = {
        viewports: detailsSystem.viewports,
        nextSectionNumber: detailsSystem.nextSectionNumber,
        nextDetailLetter: detailsSystem.nextDetailLetter,
        sheetNum: detailsSystem.sheetNum,
    };
    return data;
};

// Patch ProjectData.fromJSON to restore details system data
const _detSysOrigFromJSON = ProjectData.fromJSON;
ProjectData.fromJSON = function (data) {
    const proj = _detSysOrigFromJSON(data);
    if (data.detailsSystem) {
        detailsSystem.viewports = data.detailsSystem.viewports || [];
        detailsSystem.nextSectionNumber = data.detailsSystem.nextSectionNumber || 1;
        detailsSystem.nextDetailLetter = data.detailsSystem.nextDetailLetter || 'A';
        detailsSystem.sheetNum = data.detailsSystem.sheetNum || 'S201';
    }
    updateDetailsCountBadge();
    return proj;
};


// ══════════════════════════════════════════════════════════
// ── HIT TESTING FOR SECTION MARKERS & DETAIL CALLOUTS ───
// ══════════════════════════════════════════════════════════

// Extend the existing hit-test logic to detect clicks on our new element types
// This allows selection, moving, and deletion of section markers and detail callouts

const _detSysOrigHitTest = typeof hitTestElement === 'function' ? hitTestElement : null;

function hitTestDetailsElements(sx, sy) {
    const coords = engine.coords;
    const activeId = getActiveLevel().id;
    const hitRadius = 8 / engine.viewport.zoom; // 8px tolerance in sheet-mm

    for (let i = project.elements.length - 1; i >= 0; i--) {
        const el = project.elements[i];
        if (el.level !== activeId) continue;
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        if (el.type === 'section-marker') {
            // Check proximity to the section line
            const p1 = coords.realToSheet(el.x1, el.y1);
            const p2 = coords.realToSheet(el.x2, el.y2);

            const dx = p2.x - p1.x, dy = p2.y - p1.y;
            const len2 = dx * dx + dy * dy;
            if (len2 < 1) continue;

            const t = Math.max(0, Math.min(1, ((sx - p1.x) * dx + (sy - p1.y) * dy) / len2));
            const closestX = p1.x + t * dx;
            const closestY = p1.y + t * dy;
            const dist = Math.sqrt((sx - closestX) ** 2 + (sy - closestY) ** 2);

            if (dist < hitRadius * 2) return el;
        }

        if (el.type === 'detail-callout') {
            // Check proximity to label bubble
            const lp = coords.realToSheet(el.lx, el.ly);
            const dist = Math.sqrt((sx - lp.x) ** 2 + (sy - lp.y) ** 2);
            if (dist < hitRadius * 2) return el;

            // Also check circle
            const cp = coords.realToSheet(el.cx, el.cy);
            const cdist = Math.sqrt((sx - cp.x) ** 2 + (sy - cp.y) ** 2);
            const radiusSheet = el.radius / CONFIG.drawingScale;
            if (Math.abs(cdist - radiusSheet) < hitRadius) return el;
        }
    }

    return null;
}

// Hook into the main mousedown for selection
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || activeTool !== 'select') return;
    if (detailsSystem.activeSheet === 'details') return;
    if (engine._spaceDown || engine._isPanning) return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const sheetPos = engine.coords.screenToSheet(sx, sy);

    const hit = hitTestDetailsElements(sheetPos.x, sheetPos.y);
    if (hit) {
        // Only take over selection if we actually hit something
        // The existing select handler will also fire — we set selectedElement here
        // so it gets picked up
        selectedElement = hit;
        selectedElements = [hit];
        engine.requestRender();
    }
});


// ══════════════════════════════════════════════════════════
// ── VIEWPORT INTERACTION HANDLERS (DETAILS SHEET) ────────
// ══════════════════════════════════════════════════════════

/**
 * Hit test against viewports on the details sheet.
 * Returns { viewport, vpIndex, hitZone, handleId } or null
 * hitZone: 'content', 'title', 'scale', 'handle', or null
 *
 * Edge/corner handles are always hit-testable on the selected (active) viewport.
 * Full edge zones (not just small handle squares) — you can grab anywhere along an edge.
 */
function _detVpEditHitTest(screenX, screenY) {
    if (detailsSystem.activeSheet !== 'details') return null;
    if (detailsSystem.viewports.length === 0) return null;

    const coords = engine.coords;
    const zoom = engine.viewport.zoom;

    // Check each viewport
    for (let i = 0; i < detailsSystem.viewports.length; i++) {
        const vp = detailsSystem.viewports[i];
        if (!vp.sheetPos) continue;

        const vpSheetX = vp.sheetPos.x;
        const vpSheetY = vp.sheetPos.y;

        // Determine viewport size
        let vpW_mm, vpH_mm;
        if (vp.type === 'section') {
            if (vp.cropBox) {
                vpW_mm = (vp.cropBox.maxAlong - vp.cropBox.minAlong) / vp.scale;
                vpH_mm = (vp.cropBox.maxElev - vp.cropBox.minElev) / vp.scale;
            } else {
                vpW_mm = 300; vpH_mm = 200;
            }
        } else {
            vpW_mm = vp.radius * 2 / vp.scale;
            vpH_mm = vp.radius * 2 / vp.scale;
        }

        // Convert to screen coords
        const tl = coords.sheetToScreen(vpSheetX, vpSheetY);
        const br = coords.sheetToScreen(vpSheetX + vpW_mm, vpSheetY + vpH_mm);
        const vpW_px = br.x - tl.x;
        const vpH_px = br.y - tl.y;

        if (vpW_px < 10 || vpH_px < 10) continue;

        // Edge grab zone in pixels (generous — 8px either side of the edge)
        const edgeGrab = 8;

        // Expanded bounds for edge detection
        const inExpandedBounds = (
            screenX >= tl.x - edgeGrab && screenX <= tl.x + vpW_px + edgeGrab &&
            screenY >= tl.y - edgeGrab && screenY <= tl.y + vpH_px + edgeGrab
        );

        if (!inExpandedBounds) {
            // Also check title area below viewport
            const titleY_px = tl.y + vpH_px + Math.max(4, 1.5 * zoom);
            if (screenY >= titleY_px && screenY < titleY_px + 30 &&
                screenX >= tl.x && screenX < tl.x + Math.max(150, vpW_px)) {
                if (screenX < tl.x + Math.min(150, vpW_px)) {
                    return { viewport: vp, vpIndex: i, hitZone: 'title' };
                }
            }
            continue;
        }

        // Check edge/corner handles on the ACTIVE viewport (selected or cropping)
        const isActive = (_detVpEditState.activeViewport === vp);
        if (isActive && vp.cropBox) {
            // Distance from each edge
            const distLeft   = Math.abs(screenX - tl.x);
            const distRight  = Math.abs(screenX - (tl.x + vpW_px));
            const distTop    = Math.abs(screenY - tl.y);
            const distBottom = Math.abs(screenY - (tl.y + vpH_px));

            const nearLeft   = distLeft   < edgeGrab;
            const nearRight  = distRight  < edgeGrab;
            const nearTop    = distTop    < edgeGrab;
            const nearBottom = distBottom < edgeGrab;

            // Corner priority (two edges meet)
            const cornerGrab = edgeGrab + 4; // slightly larger zone for corners
            const cornerLeft   = distLeft   < cornerGrab;
            const cornerRight  = distRight  < cornerGrab;
            const cornerTop    = distTop    < cornerGrab;
            const cornerBottom = distBottom < cornerGrab;

            if (cornerTop && cornerLeft)     return { viewport: vp, vpIndex: i, hitZone: 'handle', handleId: 'nw' };
            if (cornerTop && cornerRight)    return { viewport: vp, vpIndex: i, hitZone: 'handle', handleId: 'ne' };
            if (cornerBottom && cornerLeft)  return { viewport: vp, vpIndex: i, hitZone: 'handle', handleId: 'sw' };
            if (cornerBottom && cornerRight) return { viewport: vp, vpIndex: i, hitZone: 'handle', handleId: 'se' };

            // Edge hits (must be within the edge extent, not just near a corner)
            const withinX = screenX > tl.x + edgeGrab && screenX < tl.x + vpW_px - edgeGrab;
            const withinY = screenY > tl.y + edgeGrab && screenY < tl.y + vpH_px - edgeGrab;

            if (nearTop    && withinX) return { viewport: vp, vpIndex: i, hitZone: 'handle', handleId: 'n' };
            if (nearBottom && withinX) return { viewport: vp, vpIndex: i, hitZone: 'handle', handleId: 's' };
            if (nearLeft   && withinY) return { viewport: vp, vpIndex: i, hitZone: 'handle', handleId: 'w' };
            if (nearRight  && withinY) return { viewport: vp, vpIndex: i, hitZone: 'handle', handleId: 'e' };
        }

        // Check if actually inside viewport content area
        const insideBounds = (
            screenX >= tl.x && screenX <= tl.x + vpW_px &&
            screenY >= tl.y && screenY <= tl.y + vpH_px
        );

        if (!insideBounds) continue;

        // Check title area (below viewport)
        const titleY_px = tl.y + vpH_px + Math.max(4, 1.5 * zoom);
        if (screenY >= titleY_px && screenY < titleY_px + 30) {
            if (screenX >= tl.x && screenX < tl.x + Math.min(150, vpW_px)) {
                return { viewport: vp, vpIndex: i, hitZone: 'title' };
            }
            if (screenX >= tl.x && screenX < tl.x + 60) {
                return { viewport: vp, vpIndex: i, hitZone: 'scale' };
            }
        }

        // Otherwise it's content area
        return { viewport: vp, vpIndex: i, hitZone: 'content' };
    }

    return null;
}

// Mouse down on details sheet — select viewport, start dragging or crop-resizing
container.addEventListener('mousedown', (e) => {
    if (detailsSystem.activeSheet !== 'details') return;
    if (e.button !== 0) return;
    if (_detVpEditState.mode === 'titleEdit' || _detVpEditState.mode === 'scaleEdit') return;

    const rect = container.getBoundingClientRect();
    const screenX = e.clientX - rect.left;
    const screenY = e.clientY - rect.top;

    const hit = _detVpEditHitTest(screenX, screenY);

    // Clicking outside any viewport — deselect
    if (!hit) {
        if (_detVpEditState.activeViewport) {
            _detVpEditState.mode = 'none';
            _detVpEditState.activeViewport = null;
            _detVpEditState.activeVpIndex = -1;
            engine.requestRender();
        }
        return;
    }

    const { viewport: vp, vpIndex: vpIdx, hitZone } = hit;

    if (hitZone === 'handle') {
        // Start crop-resizing immediately (edge or corner grab)
        // Ensure crop box exists
        if (!vp.cropBox && vp.type === 'section') {
            vp.cropBox = calculateAutoCropBox(vp) || {
                minAlong: 0, maxAlong: 5000,
                minElev: 0, maxElev: 3000
            };
        }
        _detVpEditState.mode = 'cropping';
        _detVpEditState.activeViewport = vp;
        _detVpEditState.activeVpIndex = vpIdx;
        _detVpEditState.cropHandle = hit.handleId;
        _detVpEditState.cropStartBox = {
            minAlong: vp.cropBox.minAlong,
            maxAlong: vp.cropBox.maxAlong,
            minElev: vp.cropBox.minElev,
            maxElev: vp.cropBox.maxElev
        };
        _detVpEditState.cropStartMouse = { x: screenX, y: screenY };
        e.preventDefault();
    } else if (hitZone === 'content') {
        // Select viewport (show handles) and start dragging to reposition
        _detVpEditState.mode = 'dragging';
        _detVpEditState.activeViewport = vp;
        _detVpEditState.activeVpIndex = vpIdx;

        // Ensure crop box exists for this viewport so handles can show
        if (!vp.cropBox && vp.type === 'section') {
            vp.cropBox = calculateAutoCropBox(vp) || {
                minAlong: 0, maxAlong: 5000,
                minElev: 0, maxElev: 3000
            };
        }

        const coords = engine.coords;
        const sheetPos = coords.screenToSheet(screenX, screenY);
        _detVpEditState.dragOffset = {
            dx: sheetPos.x - vp.sheetPos.x,
            dy: sheetPos.y - vp.sheetPos.y
        };

        e.preventDefault();
    }
}, true);

// Mouse move — update dragging or crop handle
container.addEventListener('mousemove', (e) => {
    if (detailsSystem.activeSheet !== 'details') return;

    const rect = container.getBoundingClientRect();
    const screenX = e.clientX - rect.left;
    const screenY = e.clientY - rect.top;

    if (_detVpEditState.mode === 'dragging' && _detVpEditState.activeViewport) {
        const vp = _detVpEditState.activeViewport;
        const coords = engine.coords;
        const sheetPos = coords.screenToSheet(screenX, screenY);

        let newX = sheetPos.x - _detVpEditState.dragOffset.dx;
        let newY = sheetPos.y - _detVpEditState.dragOffset.dy;

        // Snap to 5mm grid
        const snapGrid = 5;
        newX = Math.round(newX / snapGrid) * snapGrid;
        newY = Math.round(newY / snapGrid) * snapGrid;

        // Clamp to sheet bounds
        newX = Math.max(0, Math.min(newX, CONFIG.SHEET_WIDTH_MM - 50));
        newY = Math.max(0, Math.min(newY, CONFIG.SHEET_HEIGHT_MM - 50));

        vp.sheetPos.x = newX;
        vp.sheetPos.y = newY;

        engine.requestRender();
    } else if ((_detVpEditState.mode === 'cropping' || _detVpEditState.mode === 'selected') && _detVpEditState.cropHandle) {
        const vp = _detVpEditState.activeViewport;
        const cropBox = vp.cropBox;
        const startBox = _detVpEditState.cropStartBox;
        const startMouse = _detVpEditState.cropStartMouse;

        const coords = engine.coords;
        const mouseDelta_px = { x: screenX - startMouse.x, y: screenY - startMouse.y };

        // Convert mouse delta from px to real-world mm using viewport scale
        const detailScale = vp.scale || 10;
        const zoom = engine.viewport.zoom;
        const pxPerMm = zoom / detailScale;

        const mouseDelta_mm = {
            along: mouseDelta_px.x / pxPerMm,
            elev: -mouseDelta_px.y / pxPerMm  // Y is inverted
        };

        // Update crop box based on handle
        const handle = _detVpEditState.cropHandle;
        let newBox = { ...startBox };

        if (handle === 'n' || handle === 'nw' || handle === 'ne') {
            newBox.maxElev = startBox.maxElev + mouseDelta_mm.elev;
        }
        if (handle === 's' || handle === 'sw' || handle === 'se') {
            newBox.minElev = startBox.minElev + mouseDelta_mm.elev;
        }
        if (handle === 'e' || handle === 'ne' || handle === 'se') {
            newBox.maxAlong = startBox.maxAlong + mouseDelta_mm.along;
        }
        if (handle === 'w' || handle === 'nw' || handle === 'sw') {
            newBox.minAlong = startBox.minAlong + mouseDelta_mm.along;
        }

        // Enforce minimum crop size (200mm × 200mm)
        const minSize = 200;
        if (newBox.maxAlong - newBox.minAlong >= minSize &&
            newBox.maxElev - newBox.minElev >= minSize) {
            vp.cropBox = newBox;
            engine.requestRender();
        }
    } else {
        // Update cursor based on hover
        const hit = _detVpEditHitTest(screenX, screenY);
        if (hit) {
            if (hit.hitZone === 'handle') {
                const id = hit.handleId;
                let cursor = 'default';
                if (id === 'n' || id === 's') cursor = 'ns-resize';
                else if (id === 'e' || id === 'w') cursor = 'ew-resize';
                else if (id === 'nw' || id === 'se') cursor = 'nwse-resize';
                else if (id === 'ne' || id === 'sw') cursor = 'nesw-resize';
                container.style.cursor = cursor;
            } else if (hit.hitZone === 'content') {
                container.style.cursor = 'grab';
            } else if (hit.hitZone === 'title') {
                container.style.cursor = 'text';
            } else if (hit.hitZone === 'scale') {
                container.style.cursor = 'pointer';
            } else {
                container.style.cursor = 'default';
            }
        } else {
            container.style.cursor = 'default';
        }
    }
}, true);

// Mouse up — finish dragging or crop handle editing (viewport stays selected)
container.addEventListener('mouseup', (e) => {
    if (detailsSystem.activeSheet !== 'details') return;

    if (_detVpEditState.mode === 'dragging') {
        // Finish drag — keep viewport selected (don't clear activeViewport)
        _detVpEditState.mode = 'selected';
        _detVpEditState.dragOffset = null;
        engine.requestRender();
    } else if (_detVpEditState.mode === 'cropping') {
        // Finish crop resize — keep viewport selected
        _detVpEditState.mode = 'selected';
        _detVpEditState.cropHandle = null;
        _detVpEditState.cropStartBox = null;
        _detVpEditState.cropStartMouse = null;
        engine.requestRender();
    }
}, true);

// Double-click on title — enter title edit mode
container.addEventListener('dblclick', (e) => {
    if (detailsSystem.activeSheet !== 'details') return;

    const rect = container.getBoundingClientRect();
    const screenX = e.clientX - rect.left;
    const screenY = e.clientY - rect.top;

    const hit = _detVpEditHitTest(screenX, screenY);
    if (!hit) return;

    // Double-click on content: select viewport (same as single click, just ensure selected)
    if (hit.hitZone === 'content') {
        const vp = hit.viewport;
        if (!vp.cropBox && vp.type === 'section') {
            vp.cropBox = calculateAutoCropBox(vp) || {
                minAlong: 0, maxAlong: 5000,
                minElev: 0, maxElev: 3000
            };
        }
        _detVpEditState.mode = 'selected';
        _detVpEditState.activeViewport = vp;
        _detVpEditState.activeVpIndex = hit.vpIndex;
        engine.requestRender();
        e.preventDefault();
    }
}, true);

// Escape key — exit crop mode or cancel edits
window.addEventListener('keydown', (e) => {
    if (detailsSystem.activeSheet !== 'details') return;
    if (_detVpEditState.mode === 'titleEdit' || _detVpEditState.mode === 'scaleEdit') return;

    if (e.key === 'Escape' && (_detVpEditState.mode === 'cropping' || _detVpEditState.mode === 'selected')) {
        _detVpEditState.mode = 'none';
        _detVpEditState.activeViewport = null;
        _detVpEditState.activeVpIndex = -1;
        _detVpEditState.cropHandle = null;
        _detVpEditState.cropStartBox = null;
        _detVpEditState.cropStartMouse = null;
        engine.requestRender();
    }
});


// ══════════════════════════════════════════════════════════
// ── TOOL STATE CLEANUP ──────────────────────────────────
// ══════════════════════════════════════════════════════════

// Patch setActiveTool to clean up detail states
// We need to add cleanup for detail tool
// This is handled by adding to the tool-clearing section
// The cleanup in 06-drawing-tools.js already has:
//   if (tool !== 'section') { sectionState.placing = false; ... }
// We add equivalent for our new states:

const _prevSetActiveTool = window.setActiveTool || setActiveTool;
// Note: setActiveTool is already defined in 06-drawing-tools.js
// We wrap it to add our cleanup

const __wrappedSetActiveTool = setActiveTool;
setActiveTool = function(tool) {
    // Clean up detail callout state
    if (tool !== 'detail') {
        detailCalloutState.placing = false;
        detailCalloutState.step = 0;
        detailCalloutState.centrePoint = null;
        detailCalloutState.currentEnd = null;
    }

    // Clean up section marker state
    if (tool !== 'section') {
        sectionMarkerState.placing = false;
        sectionMarkerState.startPoint = null;
        sectionMarkerState.currentEnd = null;
    }

    // Clean up viewport editing state
    _detVpEditState.mode = 'none';
    _detVpEditState.activeViewport = null;
    _detVpEditState.activeVpIndex = -1;
    _detVpEditState.dragOffset = null;
    _detVpEditState.cropHandle = null;
    _detVpEditState.cropStartBox = null;
    _detVpEditState.cropStartMouse = null;
    if (_detVpEditState.titleEditEl) {
        _detVpEditState.titleEditEl.remove();
        _detVpEditState.titleEditEl = null;
    }
    if (_detVpEditState.scaleEditEl) {
        _detVpEditState.scaleEditEl.remove();
        _detVpEditState.scaleEditEl = null;
    }

    // Call original
    __wrappedSetActiveTool(tool);

    // Update tool names in status bar
    if (tool === 'detail') {
        document.getElementById('status-tool').textContent = 'Detail Callout';
    }
};


// ══════════════════════════════════════════════════════════
// ── STATUS BAR UPDATE ───────────────────────────────────
// ══════════════════════════════════════════════════════════

// Patch updateStatusBar to show details sheet info
const _origUpdateStatusBar = typeof updateStatusBar === 'function' ? updateStatusBar : null;
if (_origUpdateStatusBar) {
    const _wrappedUpdateStatus = updateStatusBar;
    updateStatusBar = function(e) {
        _wrappedUpdateStatus(e);
        if (detailsSystem.activeSheet === 'details') {
            const scaleEl = document.getElementById('status-scale');
            if (scaleEl) scaleEl.textContent = '1:' + detailsSystem.detailScale;
            const toolEl = document.getElementById('status-tool');
            if (toolEl) toolEl.textContent = 'Details Sheet (' + detailsSystem.viewports.length + ' viewports)';
        }
    };
}


console.log('[Details System] Loaded — Section markers, detail callouts, and details sheet renderer ready');
