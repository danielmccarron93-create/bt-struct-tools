// ── PHASE 4: DRAWING TOOLS + LAYERS ──────────────────────
// ══════════════════════════════════════════════════════════

// ── Active Tool State ────────────────────────────────────

let activeTool = 'select'; // 'select', 'line', 'column', 'text', 'grids'

// ── Auto-Link Settings ─────────────────────────────────────
// When enabled, placing walls auto-creates strip footings and
// placing columns auto-creates pad footings on the ground level.
const autoLinkSettings = {
    enabled: true,   // Master toggle for auto-footing creation
};

/** Find linked footings for a given wall or column element */
function findLinkedFootings(el) {
    if (el.type === 'wall') {
        return project.elements.filter(f => f.type === 'stripFooting' && f.wallRef === el.id);
    }
    if (el.type === 'column') {
        return project.elements.filter(f => f.type === 'footing' && f.colRef === el.id);
    }
    return [];
}

function setActiveTool(tool) {
    // === STATE CLEARING (ALL TOOLS) ===

    // Grid tool deactivation
    if (activeTool === 'grids' && tool !== 'grids') {
        deactivateGridTool();
    }

    // Line tool
    if (tool !== 'line') {
        lineToolState.placing = false;
        lineToolState.startPoint = null;
    }

    // Polyline, Slab, and Rect tools
    if (tool !== 'polyline') {
        polyToolState.points = [];
        polyToolState.currentEnd = null;
    }
    if (tool !== 'slab') {
        slabToolState.points = [];
        slabToolState.currentEnd = null;
    }
    if (tool !== 'rect') {
        rectToolState.startPoint = null;
        rectToolState.currentEnd = null;
    }

    // Dimension tool
    if (tool !== 'dim') {
        dimToolState.placing = false;
        dimToolState.startPoint = null;
    }

    // Measure tool
    if (tool !== 'measure') {
        measureState.placing = false;
        measureState.startPoint = null;
        if (measureReadout) measureReadout.style.display = 'none';
    }

    // Leader tool
    if (tool !== 'leader') {
        leaderState.placing = false;
        leaderState.startPoint = null;
    }

    // Callout tool
    if (tool !== 'callout' && typeof calloutState !== 'undefined') {
        calloutState.placing = false;
        calloutState.startPoint = null;
        if (typeof activeCalloutInput !== 'undefined' && activeCalloutInput) {
            if (activeCalloutInput.parentNode) activeCalloutInput.parentNode.removeChild(activeCalloutInput);
            activeCalloutInput = null;
        }
    }

    // Cloud tool
    if (tool !== 'cloud') {
        cloudState.drawing = false;
        cloudState.points = [];
    }

    // Cloud Poly tool
    if (tool !== 'cloudpoly') {
        cloudPolyState.points = [];
        cloudPolyState.currentEnd = null;
    }

    // Section tool
    if (tool !== 'section') {
        sectionState.placing = false;
        sectionState.startPoint = null;
    }

    // V18: Detail callout tool
    if (tool !== 'detail' && typeof detailCalloutState !== 'undefined') {
        detailCalloutState.placing = false;
        detailCalloutState.step = 0;
        detailCalloutState.centrePoint = null;
        detailCalloutState.currentEnd = null;
    }

    // Chain Dimension tool
    if (tool !== 'chaindim') {
        chainDimState.points = [];
        chainDimState.currentEnd = null;
    }

    // Notes Panel tool
    if (tool !== 'notes-panel') {
        notesPlacementPending = null;
    }

    // Table tool
    if (tool !== 'table') {
        tablePlacementPending = null;
    }

    // Slab Callout tool
    if (tool !== 'slab-callout') {
        slabCalloutPlacementPending = null;
    }

    // Edge tool
    if (tool !== 'edge') {
        edgeToolState.placing = false;
        edgeToolState.startPoint = null;
    }

    // Fall/FallArrow tool
    if (tool !== 'fall') {
        fallState.placing = false;
        fallState.startPoint = null;
    }

    // Step/StepLine tool
    if (tool !== 'step') {
        stepState.placing = false;
        stepState.startPoint = null;
    }

    // Footing tool
    if (tool !== 'footing') {
        footingPlacementPending = null;
    }

    // Strip footing tool
    if (tool !== 'stripFooting') {
        if (typeof stripFtgState !== 'undefined') {
            stripFtgState.placing = false;
            stripFtgState.startPoint = null;
            stripFtgState.currentEnd = null;
        }
    }

    // Wall tool
    if (tool !== 'wall') {
        wallToolState.placing = false;
        wallToolState.startPoint = null;
        wallToolState.currentEnd = null;
    }

    // Bracing wall tool
    const bracingRelatedTools = ['bracingWall', 'buildingEnvelope', 'ridgeLine', 'roofSkeleton'];
    if (tool !== 'bracingWall') {
        if (typeof bracingWallToolState !== 'undefined') {
            bracingWallToolState.placing = false;
            bracingWallToolState.startPoint = null;
            bracingWallToolState.currentEnd = null;
        }
    }
    // Close bracing panel only when switching to a non-bracing-related tool
    if (!bracingRelatedTools.includes(tool)) {
        const bracingPanel = document.getElementById('bracing-summary-panel');
        if (bracingPanel && !bracingPanel.classList.contains('hidden')) {
            bracingPanel.classList.add('hidden');
        }
    }

    // Building envelope tool
    if (tool !== 'buildingEnvelope') {
        if (typeof envelopeToolState !== 'undefined') {
            envelopeToolState.points = [];
            envelopeToolState.currentEnd = null;
        }
    }

    // Ridge line tool
    if (tool !== 'ridgeLine') {
        if (typeof ridgeLineToolState !== 'undefined') {
            ridgeLineToolState.points = [];
            ridgeLineToolState.currentEnd = null;
        }
    }

    // Roof skeleton tool
    if (tool !== 'roofSkeleton') {
        if (typeof roofSkeletonState !== 'undefined') {
            roofSkeletonState.chainStartId = null;
            roofSkeletonState.hoverPoint = null;
            roofSkeletonState.snapTarget = null;
        }
    }

    // === UPDATE ACTIVE TOOL ===
    activeTool = tool;

    // === UPDATE TOOLBAR BUTTONS ===
    selectBtn.classList.toggle('active', tool === 'select');
    document.getElementById('btn-line').classList.toggle('active', tool === 'line');
    document.getElementById('btn-column').classList.toggle('active', tool === 'column');
    document.getElementById('btn-text').classList.toggle('active', tool === 'text');
    gridPlaceBtn.classList.toggle('active', tool === 'grids');
    if (document.getElementById('btn-grids-ribbon')) {
        document.getElementById('btn-grids-ribbon').classList.toggle('active', tool === 'grids');
    }

    // Polyline, Slab, and Rect buttons
    if (document.getElementById('btn-polyline')) {
        document.getElementById('btn-polyline').classList.toggle('active', tool === 'polyline');
    }
    if (document.getElementById('btn-slab')) {
        document.getElementById('btn-slab').classList.toggle('active', tool === 'slab');
    }
    if (document.getElementById('btn-rect')) {
        document.getElementById('btn-rect').classList.toggle('active', tool === 'rect');
    }

    // Dimension and Measure buttons
    if (document.getElementById('btn-dim')) {
        document.getElementById('btn-dim').classList.toggle('active', tool === 'dim');
    }
    if (document.getElementById('btn-measure')) {
        document.getElementById('btn-measure').classList.toggle('active', tool === 'measure');
    }

    // Leader, Cloud, Cloud Poly buttons
    if (document.getElementById('btn-leader')) {
        document.getElementById('btn-leader').classList.toggle('active', tool === 'leader');
    }
    if (document.getElementById('btn-callout')) {
        document.getElementById('btn-callout').classList.toggle('active', tool === 'callout');
    }
    if (document.getElementById('btn-cloud')) {
        document.getElementById('btn-cloud').classList.toggle('active', tool === 'cloud');
    }
    if (document.getElementById('btn-cloud-poly')) {
        document.getElementById('btn-cloud-poly').classList.toggle('active', tool === 'cloudpoly');
    }

    // Section split-button highlight (container highlights for section OR detail)
    if (document.getElementById('btn-section-split')) {
        document.getElementById('btn-section-split').classList.toggle('active', tool === 'section' || tool === 'detail');
    }
    if (document.getElementById('btn-section')) {
        document.getElementById('btn-section').classList.toggle('active', tool === 'section');
    }
    if (document.getElementById('btn-chain-dim')) {
        document.getElementById('btn-chain-dim').classList.toggle('active', tool === 'chaindim');
    }

    // Phase A buttons (Notes, Table, Slab Callout)
    if (document.getElementById('btn-notes-panel')) {
        document.getElementById('btn-notes-panel').classList.toggle('active', tool === 'notes-panel');
    }
    if (document.getElementById('btn-table')) {
        document.getElementById('btn-table').classList.toggle('active', tool === 'table');
    }
    if (document.getElementById('btn-slab-callout')) {
        document.getElementById('btn-slab-callout').classList.toggle('active', tool === 'slab-callout');
    }

    // Phase B buttons (Edge, Fall, Step)
    if (document.getElementById('btn-edge')) {
        document.getElementById('btn-edge').classList.toggle('active', tool === 'edge');
    }
    if (document.getElementById('btn-fall')) {
        document.getElementById('btn-fall').classList.toggle('active', tool === 'fall');
    }
    if (document.getElementById('btn-step')) {
        document.getElementById('btn-step').classList.toggle('active', tool === 'step');
    }

    // Phase C buttons (RL Marker, Sec Ref, Ref Tag)
    if (document.getElementById('btn-rl-marker')) {
        document.getElementById('btn-rl-marker').classList.toggle('active', tool === 'rlmarker');
        document.getElementById('btn-sec-ref').classList.toggle('active', tool === 'secref');
        document.getElementById('btn-ref-tag').classList.toggle('active', tool === 'reftag');
    }

    // Phase D buttons (Footing + Strip Footing)
    if (document.getElementById('btn-footing')) {
        document.getElementById('btn-footing').classList.toggle('active', tool === 'footing');
    }
    if (document.getElementById('btn-strip-footing')) {
        document.getElementById('btn-strip-footing').classList.toggle('active', tool === 'stripFooting');
    }

    // Wall button
    if (document.getElementById('btn-wall')) {
        document.getElementById('btn-wall').classList.toggle('active', tool === 'wall');
    }

    // Bracing wall button
    if (document.getElementById('btn-bracing-wall')) {
        document.getElementById('btn-bracing-wall').classList.toggle('active', tool === 'bracingWall');
    }

    // Building envelope button
    if (document.getElementById('btn-envelope')) {
        document.getElementById('btn-envelope').classList.toggle('active', tool === 'buildingEnvelope');
    }

    // === DIM SUB-TOOLBAR: show when dim-related tool is active ===
    const dimSubBar = document.getElementById('dim-sub-toolbar');
    if (dimSubBar) {
        const dimTools = ['dim', 'chaindim', 'measure', 'grids', 'rlmarker'];
        if (dimTools.includes(tool)) {
            dimSubBar.classList.remove('hidden');
        } else {
            dimSubBar.classList.add('hidden');
        }
    }

    // === UPDATE STRUCTURAL TOOL HIGHLIGHTS ===
    // Clear structural tool highlights unless we set them explicitly
    if (typeof beamBtn !== 'undefined' && tool !== 'line' && tool !== 'column') {
        beamBtn.classList.remove('active');
    }
    if (typeof slabBtn !== 'undefined' && tool !== 'rect') {
        slabBtn.classList.remove('active');
    }

    // === SET CURSOR ===
    // V18: Detail callout button highlight
    if (document.getElementById('btn-detail-callout')) {
        document.getElementById('btn-detail-callout').classList.toggle('active', tool === 'detail');
    }

    const cursorTools = ['line', 'text', 'polyline', 'rect', 'dim', 'measure',
                         'leader', 'cloud', 'cloudpoly', 'section', 'chaindim',
                         'notes-panel', 'table', 'slab-callout', 'edge', 'fall', 'step',
                         'rlmarker', 'secref', 'reftag', 'stripFooting', 'wall', 'detail', 'bracingWall', 'buildingEnvelope', 'ridgeLine', 'roofSkeleton', 'callout'];
    // Column and footing get 'none' cursor — canvas preview replaces it
    const previewTools = ['column', 'footing'];
    if (previewTools.includes(tool)) {
        container.style.cursor = 'none';
    } else {
        container.style.cursor = cursorTools.includes(tool) ? 'crosshair' : '';
    }

    // === UPDATE STATUS BAR ===
    const names = {
        select: 'Select',
        line: 'Line',
        column: 'Column',
        text: 'Text',
        grids: 'Grids',
        polyline: 'Polyline',
        rect: 'Rectangle',
        dim: 'Dimension',
        measure: 'Measure',
        leader: 'Leader',
        cloud: 'Cloud',
        cloudpoly: 'Cloud Poly',
        section: 'Section',
        chaindim: 'Chain Dim',
        'notes-panel': 'Notes Panel',
        table: 'Table',
        'slab-callout': 'Slab Callout',
        edge: 'Edge',
        fall: 'Fall Arrow',
        step: 'Step Line',
        rlmarker: 'RL Marker',
        secref: 'Sec Ref',
        reftag: 'Ref Tag',
        footing: 'Footing',
        stripFooting: 'Strip Footing',
        wall: 'Wall',
        bracingWall: 'Bracing Wall',
        buildingEnvelope: 'Building Envelope',
        ridgeLine: 'Ridge Line',
        roofSkeleton: 'Roof Skeleton',
        detail: 'Detail Callout',
        callout: 'Callout'
    };
    document.getElementById('status-tool').textContent = names[tool] || tool;

    // === REQUEST RENDER ===
    engine.requestRender();
}

// Wire tool buttons
selectBtn.addEventListener('click', () => setActiveTool('select'));
document.getElementById('btn-line').addEventListener('click', () => setActiveTool('line'));
document.getElementById('btn-column').addEventListener('click', () => setActiveTool('column'));
document.getElementById('btn-text').addEventListener('click', () => setActiveTool('text'));
document.getElementById('btn-wall').addEventListener('click', () => setActiveTool('wall'));
// Bracing wall button
if (document.getElementById('btn-bracing-wall')) {
    document.getElementById('btn-bracing-wall').addEventListener('click', () => {
        setActiveTool('bracingWall');
        // Show bracing panel automatically
        if (typeof toggleBracingPanel === 'function') {
            const panel = document.getElementById('bracing-summary-panel');
            if (panel && panel.classList.contains('hidden')) toggleBracingPanel();
        }
    });
}
// Building envelope button
if (document.getElementById('btn-envelope')) {
    document.getElementById('btn-envelope').addEventListener('click', () => {
        setActiveTool('buildingEnvelope');
    });
}
// Bracing type selector
if (document.getElementById('bracing-type-select')) {
    document.getElementById('bracing-type-select').addEventListener('change', (e) => {
        placementTypeRef.bracingWall = e.target.value;
    });
}
// Grid button
gridPlaceBtn.removeEventListener('click', gridPlaceBtn._handler);
gridPlaceBtn._handler = () => {
    if (activeTool === 'grids') {
        setActiveTool('select');
    } else {
        setActiveTool('grids');
        activateGridTool();
    }
};
gridPlaceBtn.addEventListener('click', gridPlaceBtn._handler);

// Ribbon Grids button (prominent)
const gridsRibbonBtn = document.getElementById('btn-grids-ribbon');
if (gridsRibbonBtn) {
    gridsRibbonBtn.addEventListener('click', () => {
        if (activeTool === 'grids') {
            setActiveTool('select');
        } else {
            setActiveTool('grids');
            activateGridTool();
        }
    });
}

// ══════════════════════════════════════════════════════════
// ── CONTEXT-SENSITIVE TOOLBAR + OVERFLOW DROPDOWN ────────
// ══════════════════════════════════════════════════════════

function updateToolbarContext(tool) {
    const ctxMember = document.getElementById('ctx-member');
    const ctxWall = document.getElementById('ctx-wall');
    const ctxText = document.getElementById('ctx-text');
    if (ctxMember) ctxMember.classList.toggle('visible', tool === 'beam' || tool === 'line' || tool === 'column' || tool === 'edge');
    if (ctxWall) ctxWall.classList.toggle('visible', tool === 'wall');
    if (ctxText) ctxText.classList.toggle('visible', tool === 'text');
}

// Patch setActiveTool for context updates
const _origSetActiveForCtx = setActiveTool;
setActiveTool = function(tool) {
    _origSetActiveForCtx(tool);
    updateToolbarContext(tool);
};
updateToolbarContext('select');

// Overflow dropdown
const moreBtn = document.getElementById('btn-annotate-more');
const overflowMenu = document.getElementById('annotate-overflow-menu');
if (moreBtn && overflowMenu) {
    moreBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        overflowMenu.classList.toggle('open');
    });
    document.addEventListener('click', () => overflowMenu.classList.remove('open'));
    overflowMenu.querySelectorAll('.rbtn-sm, .tbtn').forEach(btn => {
        btn.addEventListener('click', () => overflowMenu.classList.remove('open'));
    });
}

// Keyboard shortcuts: V=select, L=line, C=column, T=text
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'v') setActiveTool('select');
    if (e.key === 'l') setActiveTool('line');
    if (e.key === 'c') setActiveTool('column');
    if (e.key === 't') setActiveTool('callout');
    if (e.key === 'b') setActiveTool('bracingWall');
    if (e.key === 'e') setActiveTool('buildingEnvelope');
    if (e.key === 'Escape') {
        if (lineToolState.placing) {
            lineToolState.placing = false;
            lineToolState.startPoint = null;
            engine.requestRender();
        } else if (wallToolState.placing) {
            wallToolState.placing = false;
            wallToolState.startPoint = null;
            wallToolState.currentEnd = null;
            engine.requestRender();
        } else if (typeof bracingWallToolState !== 'undefined' && bracingWallToolState.placing) {
            bracingWallToolState.placing = false;
            bracingWallToolState.startPoint = null;
            bracingWallToolState.currentEnd = null;
            engine.requestRender();
        } else if (typeof envelopeToolState !== 'undefined' && envelopeToolState.points.length > 0) {
            envelopeToolState.points = [];
            envelopeToolState.currentEnd = null;
            engine.requestRender();
        } else if (typeof ridgeLineToolState !== 'undefined' && ridgeLineToolState.points.length > 0) {
            ridgeLineToolState.points = [];
            ridgeLineToolState.currentEnd = null;
            engine.requestRender();
        } else {
            setActiveTool('select');
        }
    }
});

// ── Element Type ─────────────────────────────────────────

const elementTypeSelect = document.getElementById('element-type');

function getActiveLayer() {
    return elementTypeSelect.value;
}

// ── Line Tool ────────────────────────────────────────────

const lineToolState = {
    placing: false,
    startPoint: null,  // { x, y } in sheet-mm
    currentEnd: null,  // { x, y } in sheet-mm (preview)
};

/**
 * Get the snapped + ortho-constrained cursor position for the line tool.
 */
function getLineToolPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    // Try snap first
    const snap = findSnap(sx, sy);
    let pos = snap
        ? { x: snap.x, y: snap.y }
        : engine.coords.screenToSheet(sx, sy);

    // Apply ortho constraint relative to start point
    if (lineToolState.placing && lineToolState.startPoint) {
        pos = applyOrtho(pos.x, pos.y, lineToolState.startPoint.x, lineToolState.startPoint.y);
    }

    return pos;
}

// Mouse move — update line preview
container.addEventListener('mousemove', (e) => {
    if (activeTool !== 'line') return;
    lineToolState.currentEnd = getLineToolPos(e);
    engine.requestRender();
});

// Mouse click — place line start/end
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'line') return;
    if (pdfState.calibrating) return;

    const pos = getLineToolPos(e);

    if (!lineToolState.placing) {
        // Start a new line
        lineToolState.placing = true;
        lineToolState.startPoint = pos;
    } else {
        // Finish the line
        const start = lineToolState.startPoint;
        const end = pos;

        // Minimum length check (at least 0.5mm on sheet)
        const dist = Math.sqrt(Math.pow(end.x - start.x, 2) + Math.pow(end.y - start.y, 2));
        if (dist < 0.5) return;

        // Convert to real-world coordinates
        const realStart = engine.coords.sheetToReal(start.x, start.y);
        const realEnd = engine.coords.sheetToReal(end.x, end.y);

        const layerId = getActiveLayer();
        const newLine = {
            id: generateId(),
            type: 'line',
            layer: layerId,
            x1: realStart.x, y1: realStart.y,
            x2: realEnd.x, y2: realEnd.y,
            typeRef: layerId === 'S-BEAM' ? (placementTypeRef.beam || 'SB1') : undefined,
        };

        history.execute({
            description: 'Draw line',
            execute() { project.elements.push(newLine); },
            undo() {
                const i = project.elements.indexOf(newLine);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });

        // Chain: start next line from this endpoint
        lineToolState.startPoint = { x: end.x, y: end.y };
        engine.requestRender();
    }
});

// Right-click or Escape finishes the chain
container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'line' && lineToolState.placing) {
        e.preventDefault();
        lineToolState.placing = false;
        lineToolState.startPoint = null;
        engine.requestRender();
    }
});

// ── Column Tool ──────────────────────────────────────────

let _colNextNum = 1;
const COL_DEFAULT_SIZE = 89; // mm — default SHS size for display

// ── Wall Tool ────────────────────────────────────────────

const wallToolState = {
    placing: false,
    startPoint: null,  // { x, y } in sheet-mm
    currentEnd: null,  // { x, y } in sheet-mm (preview)
};

let _wallBwNum = 1;
let _wallCwNum = 1;

/**
 * Parse wall-type string: "90-stud" → { wallType: 'stud', thickness: 90 }
 */
function parseWallType(typeStr) {
    const [thickStr, type] = typeStr.split('-');
    return { wallType: type, thickness: parseInt(thickStr) };
}

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'wall') return;
    if (pdfState.calibrating) return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    // Use snap
    const snap = findSnap(sx, sy);
    const sheetPos = snap
        ? { x: snap.x, y: snap.y }
        : engine.coords.screenToSheet(sx, sy);

    if (!wallToolState.placing) {
        // First click: start the wall
        wallToolState.placing = true;
        wallToolState.startPoint = { ...sheetPos };
        wallToolState.currentEnd = { ...sheetPos };
        engine.requestRender();
    } else {
        // Second click: finish the wall
        const typeSelect = document.getElementById('wall-type');
        const { wallType, thickness } = parseWallType(typeSelect.value);

        const realStart = engine.coords.sheetToReal(wallToolState.startPoint.x, wallToolState.startPoint.y);
        const realEnd = engine.coords.sheetToReal(wallToolState.currentEnd.x, wallToolState.currentEnd.y);

        // Use schedule type ref for tag
        const wallTypeRef = placementTypeRef.wall || 'BW1';
        const wallTypeData = project.scheduleTypes.wall[wallTypeRef] || {};
        let tag = wallTypeRef;

        const newWall = {
            id: generateId(),
            type: 'wall',
            layer: 'S-WALL',
            x1: realStart.x, y1: realStart.y,
            x2: realEnd.x, y2: realEnd.y,
            wallType: wallTypeData.wallType || wallType,
            thickness: wallTypeData.thickness || thickness,
            tag: tag,
            typeRef: wallTypeRef,
        };

        // Auto-link: create strip footing under wall if on Ground Floor
        let autoFooting = null;
        if (autoLinkSettings.enabled) {
            const sfTypeRef = placementTypeRef.stripfooting || 'SF1';
            const sfTypeData = project.scheduleTypes.stripfooting[sfTypeRef] || {};
            autoFooting = {
                id: generateId(),
                type: 'stripFooting',
                layer: 'S-FTNG',
                level: 'GF',
                x1: realStart.x, y1: realStart.y,
                x2: realEnd.x, y2: realEnd.y,
                tag: sfTypeRef,
                typeRef: sfTypeRef,
                footingWidth: parseInt(sfTypeData.width) || 300,
                footingDepth: parseInt(sfTypeData.depth) || 500,
                reinforcement: sfTypeData.reo || '',
                depthBelowFSL: parseInt(sfTypeData.setdown) || 200,
                wallRef: newWall.id,    // Link back to this wall
            };
        }

        history.execute({
            description: 'Draw wall: ' + (tag || 'Stud') + (autoFooting ? ' + auto strip footing' : ''),
            execute() {
                project.elements.push(newWall);
                if (autoFooting) project.elements.push(autoFooting);
            },
            undo() {
                if (autoFooting) {
                    const fi = project.elements.indexOf(autoFooting);
                    if (fi !== -1) project.elements.splice(fi, 1);
                }
                const i = project.elements.indexOf(newWall);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });

        // Chain mode: start next wall from where this one ended
        wallToolState.startPoint = { ...wallToolState.currentEnd };
        engine.requestRender();
    }
});

container.addEventListener('mousemove', (e) => {
    if (activeTool !== 'wall' || !wallToolState.placing || !wallToolState.startPoint) return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    let sheetPos = engine.coords.screenToSheet(sx, sy);

    // Check for snap
    const snap = findSnap(sx, sy);
    if (snap) sheetPos = { x: snap.x, y: snap.y };

    // Shift: constrain to 45° angles
    if (e.shiftKey && wallToolState.startPoint) {
        const dx = sheetPos.x - wallToolState.startPoint.x;
        const dy = sheetPos.y - wallToolState.startPoint.y;
        const angle = Math.atan2(dy, dx);
        const snapAngle = Math.round(angle / (Math.PI / 4)) * (Math.PI / 4);
        const dist = Math.sqrt(dx * dx + dy * dy);
        sheetPos = {
            x: wallToolState.startPoint.x + dist * Math.cos(snapAngle),
            y: wallToolState.startPoint.y + dist * Math.sin(snapAngle)
        };
    }

    wallToolState.currentEnd = { ...sheetPos };
    engine.requestRender();
});

container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'wall' && wallToolState.placing) {
        e.preventDefault();
        wallToolState.placing = false;
        wallToolState.startPoint = null;
        wallToolState.currentEnd = null;
        engine.requestRender();
    }
});

// ── Column tool state ──────────────────────────────────────
// Track last placed column position for Shift-ortho constraint
const _colToolState = {
    lastPlacedSheet: null,  // { x, y } in sheet-mm — position of last placed column
    previewPos: null,       // { x, y } in sheet-mm — current preview position
};

/** Get column tool position with snap + optional Shift-ortho from last column */
function getColToolPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    // Try snap first
    const snap = findSnap(sx, sy);
    let pos = snap
        ? { x: snap.x, y: snap.y }
        : engine.coords.screenToSheet(sx, sy);

    // Shift-ortho constraint from last placed column
    if (_shiftDown && _colToolState.lastPlacedSheet) {
        pos = applyOrtho(pos.x, pos.y, _colToolState.lastPlacedSheet.x, _colToolState.lastPlacedSheet.y);

        // After ortho constraint, try to re-snap along the constrained axis
        // This lets you lock to a row but still snap to grid intersections
        if (snap) {
            const ref = _colToolState.lastPlacedSheet;
            const dx = Math.abs(pos.x - ref.x);
            const dy = Math.abs(pos.y - ref.y);
            if (dy < 0.01) {
                // Horizontal lock — keep y fixed, try to snap x to grid
                pos.x = snap.x;
            } else if (dx < 0.01) {
                // Vertical lock — keep x fixed, try to snap y to grid
                pos.y = snap.y;
            }
        }
    }

    return pos;
}

// Column tool mousemove — show preview and ortho guide line
container.addEventListener('mousemove', (e) => {
    if (activeTool !== 'column') return;
    _colToolState.previewPos = getColToolPos(e);
    engine.requestRender();
});

// Column tool mousedown
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'column') return;
    if (pdfState.calibrating) return;

    const sheetPos = getColToolPos(e);
    const realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);
    const colTypeRef = placementTypeRef.column || 'SC1';
    const colTypeData = project.scheduleTypes.column[colTypeRef] || {};

    // Bottom floor columns extend above, others extend below by default
    const isBottomFloor = levelSystem.activeLevelIndex === 0;
    const newCol = {
        id: generateId(),
        type: 'column',
        layer: 'S-COLS',
        x: realPos.x, y: realPos.y,
        size: colTypeData.size || COL_DEFAULT_SIZE,
        tag: colTypeRef,
        typeRef: colTypeRef,
        extends: isBottomFloor ? 'above' : 'below',
    };

    // Auto-link: create pad footing under column
    let autoColFooting = null;
    if (autoLinkSettings.enabled) {
        const pfTypeRef = placementTypeRef.footing || 'PF1';
        const pfTypeData = project.scheduleTypes.padfooting[pfTypeRef] || {};
        autoColFooting = {
            id: generateId(),
            type: 'footing',
            layer: 'S-FTNG',
            level: 'GF',
            x: realPos.x,
            y: realPos.y,
            mark: pfTypeRef,
            typeRef: pfTypeRef,
            footingWidth: parseInt(pfTypeData.width) || 600,
            footingDepth: parseInt(pfTypeData.depth) || 300,
            reinforcement: pfTypeData.reo || '',
            depthBelowFSL: parseInt(pfTypeData.setdown) || 200,
            colRef: newCol.id,  // Link back to this column
        };
    }

    history.execute({
        description: 'Place column ' + colTypeRef + (autoColFooting ? ' + auto pad footing' : ''),
        execute() {
            project.elements.push(newCol);
            if (autoColFooting) project.elements.push(autoColFooting);
        },
        undo() {
            if (autoColFooting) {
                const fi = project.elements.indexOf(autoColFooting);
                if (fi !== -1) project.elements.splice(fi, 1);
            }
            const i = project.elements.indexOf(newCol);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    // Remember this column position for Shift-ortho
    _colToolState.lastPlacedSheet = { x: sheetPos.x, y: sheetPos.y };

    engine.requestRender();
});

// Column tool: Shift-ortho guide line render
engine.onRender((ctx, eng) => {
    if (activeTool !== 'column') return;
    if (!_shiftDown || !_colToolState.lastPlacedSheet || !_colToolState.previewPos) return;

    const coords = eng.coords;
    const ref = _colToolState.lastPlacedSheet;
    const preview = _colToolState.previewPos;

    // Draw dashed guide line from last placed column to current preview position
    const sp1 = coords.sheetToScreen(ref.x, ref.y);
    const sp2 = coords.sheetToScreen(preview.x, preview.y);

    ctx.save();
    ctx.setLineDash([6, 4]);
    ctx.strokeStyle = 'rgba(43, 124, 208, 0.5)';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(sp1.x, sp1.y);
    ctx.lineTo(sp2.x, sp2.y);
    ctx.stroke();
    ctx.setLineDash([]);

    // Draw small cross at preview position
    const r = 4;
    ctx.strokeStyle = '#2B7CD0';
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(sp2.x - r, sp2.y); ctx.lineTo(sp2.x + r, sp2.y);
    ctx.moveTo(sp2.x, sp2.y - r); ctx.lineTo(sp2.x, sp2.y + r);
    ctx.stroke();
    ctx.restore();
});

// Reset column tool state when switching tools
const _origSetActiveToolForCol = setActiveTool;
setActiveTool = function(tool) {
    if (tool !== 'column') {
        _colToolState.lastPlacedSheet = null;
        _colToolState.previewPos = null;
    }
    _origSetActiveToolForCol(tool);
};

// ── Text Tool (drag-box placement like Revit) ─────────────

const textToolState = {
    dragging: false,
    startScreen: null,   // { x, y } screen-px
    startSheet: null,    // { x, y } sheet-mm
    currentScreen: null, // { x, y } screen-px
};
let activeTextInput = null;

// Mousedown — begin drag box
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'text') return;
    if (pdfState.calibrating) return;

    // If there's already an input open, commit it first
    if (activeTextInput) commitTextInput();

    e.preventDefault();
    e.stopPropagation();

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    const snap = findSnap(sx, sy);
    const sheetPos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);

    textToolState.dragging = true;
    textToolState.startScreen = { x: sx, y: sy };
    textToolState.startSheet = sheetPos;
    textToolState.currentScreen = { x: sx, y: sy };
});

// Mousemove — update drag preview
window.addEventListener('mousemove', (e) => {
    if (!textToolState.dragging) return;
    const rect = container.getBoundingClientRect();
    textToolState.currentScreen = {
        x: e.clientX - rect.left,
        y: e.clientY - rect.top
    };
    engine.requestRender();
});

// Mouseup — finish drag, show text input
window.addEventListener('mouseup', (e) => {
    if (!textToolState.dragging) return;
    textToolState.dragging = false;

    const rect = container.getBoundingClientRect();
    const endX = e.clientX - rect.left;
    const endY = e.clientY - rect.top;

    // Calculate box in screen coords
    const x = Math.min(textToolState.startScreen.x, endX);
    const y = Math.min(textToolState.startScreen.y, endY);
    const w = Math.max(60, Math.abs(endX - textToolState.startScreen.x));
    const h = Math.max(24, Math.abs(endY - textToolState.startScreen.y));

    // Create the input at the drag box location
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'text-input-overlay';
    input.style.left = x + 'px';
    input.style.top = y + 'px';
    input.style.width = w + 'px';
    input.style.height = h + 'px';
    input.style.fontSize = Math.max(12, h * 0.6) + 'px';
    input.placeholder = 'Type text...';
    input._sheetPos = { x: textToolState.startSheet.x, y: textToolState.startSheet.y };
    input._boxHeight = h;

    container.appendChild(input);

    // Small delay to avoid the mouseup stealing focus
    setTimeout(() => {
        input.focus();
    }, 30);

    activeTextInput = input;

    input.addEventListener('keydown', (ev) => {
        ev.stopPropagation(); // prevent tool shortcuts while typing
        if (ev.key === 'Enter') {
            ev.preventDefault();
            commitTextInput();
        }
        if (ev.key === 'Escape') {
            ev.preventDefault();
            cancelTextInput();
        }
    });

    input.addEventListener('blur', () => {
        setTimeout(() => {
            if (activeTextInput === input) commitTextInput();
        }, 150);
    });

    engine.requestRender();
});

function commitTextInput() {
    if (!activeTextInput) return;
    const input = activeTextInput;
    const text = input.value.trim();
    const sheetPos = input._sheetPos;
    // Use the toolbar text size setting
    const textSizeSelect = document.getElementById('text-size');
    const fontSize = textSizeSelect ? parseFloat(textSizeSelect.value) : 3.5;

    activeTextInput = null;
    if (input.parentNode) input.parentNode.removeChild(input);

    if (!text) return;

    const realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);

    const newText = {
        id: generateId(),
        type: 'text',
        layer: 'S-ANNO',
        x: realPos.x, y: realPos.y,
        text: text,
        fontSize: fontSize,
    };

    history.execute({
        description: 'Add text: ' + text,
        execute() { project.elements.push(newText); },
        undo() {
            const i = project.elements.indexOf(newText);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    engine.requestRender();
}

function cancelTextInput() {
    if (!activeTextInput) return;
    const input = activeTextInput;
    activeTextInput = null;
    if (input.parentNode) input.parentNode.removeChild(input);
    engine.requestRender();
}

// Draw text drag preview box
function drawTextDragPreview(ctx, eng) {
    if (activeTool !== 'text' || !textToolState.dragging) return;

    const s = textToolState.startScreen;
    const c = textToolState.currentScreen;
    const x = Math.min(s.x, c.x);
    const y = Math.min(s.y, c.y);
    const w = Math.abs(c.x - s.x);
    const h = Math.abs(c.y - s.y);

    ctx.strokeStyle = '#2B7CD0';
    ctx.lineWidth = 1.5;
    ctx.setLineDash([5, 3]);
    ctx.strokeRect(x, y, w, h);
    ctx.setLineDash([]);

    // "T" indicator
    ctx.fillStyle = 'rgba(43,124,208,0.15)';
    ctx.fillRect(x, y, w, h);
    ctx.font = 'bold 14px "Segoe UI", Arial, sans-serif';
    ctx.fillStyle = '#2B7CD0';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    if (w > 20 && h > 15) {
        ctx.fillText('T', x + w / 2, y + h / 2);
    }
    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';
}

// Register the text drag preview renderer
engine.onRender(drawTextDragPreview);

// ── Element Rendering ────────────────────────────────────

function drawElements(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    for (const el of project.getVisibleElements()) {
        const layer = project.layers[el.layer];
        if (!layer) continue;
        const isSelected = (selectedElement === el);

        // ── Walls (double-line with mitred corners and hatch) ──
        if (el.type === 'wall') {
            const wdx = el.x2 - el.x1;
            const wdy = el.y2 - el.y1;
            const wlen = Math.sqrt(wdx * wdx + wdy * wdy);
            if (wlen < 0.1) continue;

            // Direction and perpendicular of this wall
            const ux = wdx / wlen, uy = wdy / wlen; // unit along wall
            const nx = -uy, ny = ux; // perpendicular
            const halfT = el.thickness / 2;

            // Base corner points (no extension)
            let s1a = { x: el.x1 + nx * halfT, y: el.y1 + ny * halfT }; // start, side A
            let s1b = { x: el.x1 - nx * halfT, y: el.y1 - ny * halfT }; // start, side B
            let s2a = { x: el.x2 + nx * halfT, y: el.y2 + ny * halfT }; // end, side A
            let s2b = { x: el.x2 - nx * halfT, y: el.y2 - ny * halfT }; // end, side B

            // Find connected walls at each endpoint and compute mitre extensions
            const joinTol = el.thickness * 0.6;
            let startJoined = null, endJoined = null;

            for (const other of project.getVisibleElements()) {
                if (other === el || other.type !== 'wall') continue;
                if (other.wallType !== el.wallType) continue;
                // Check start of this wall connects to either end of other
                const d1s = Math.sqrt(Math.pow(el.x1 - other.x1, 2) + Math.pow(el.y1 - other.y1, 2));
                const d1e = Math.sqrt(Math.pow(el.x1 - other.x2, 2) + Math.pow(el.y1 - other.y2, 2));
                if (d1s < joinTol || d1e < joinTol) startJoined = other;
                const d2s = Math.sqrt(Math.pow(el.x2 - other.x1, 2) + Math.pow(el.y2 - other.y1, 2));
                const d2e = Math.sqrt(Math.pow(el.x2 - other.x2, 2) + Math.pow(el.y2 - other.y2, 2));
                if (d2s < joinTol || d2e < joinTol) endJoined = other;
            }

            // Line-line intersection: returns point where line (p1→p2) meets line (p3→p4), or null
            function lineIntersect(p1x,p1y,p2x,p2y, p3x,p3y,p4x,p4y) {
                const d = (p1x-p2x)*(p3y-p4y) - (p1y-p2y)*(p3x-p4x);
                if (Math.abs(d) < 0.0001) return null; // parallel
                const t = ((p1x-p3x)*(p3y-p4y) - (p1y-p3y)*(p3x-p4x)) / d;
                return { x: p1x + t*(p2x-p1x), y: p1y + t*(p2y-p1y) };
            }

            // Compute mitred corner: extend each edge line of this wall to intersect the corresponding edge line of the other wall
            function mitreCorner(thisEnd, otherWall, sideA, sideB) {
                const odx = otherWall.x2 - otherWall.x1;
                const ody = otherWall.y2 - otherWall.y1;
                const olen = Math.sqrt(odx*odx + ody*ody);
                if (olen < 0.1) return;
                const oux = odx/olen, ouy = ody/olen;
                const onx = -ouy, ony = oux;
                const oHalfT = otherWall.thickness / 2;

                // Other wall's two edge lines in real coords
                const oA1 = { x: otherWall.x1 + onx*oHalfT, y: otherWall.y1 + ony*oHalfT };
                const oA2 = { x: otherWall.x2 + onx*oHalfT, y: otherWall.y2 + ony*oHalfT };
                const oB1 = { x: otherWall.x1 - onx*oHalfT, y: otherWall.y1 - ony*oHalfT };
                const oB2 = { x: otherWall.x2 - onx*oHalfT, y: otherWall.y2 - ony*oHalfT };

                // This wall's two edge lines: sideA and sideB (line along ux,uy direction)
                // We extend them by finding intersection with the outer edge of the other wall

                // For side A of this wall: intersect with both edges of other wall, pick the one that extends outward
                const intAoA = lineIntersect(sideA.x, sideA.y, sideA.x + ux*1000, sideA.y + uy*1000,
                                             oA1.x, oA1.y, oA2.x, oA2.y);
                const intAoB = lineIntersect(sideA.x, sideA.y, sideA.x + ux*1000, sideA.y + uy*1000,
                                             oB1.x, oB1.y, oB2.x, oB2.y);

                const intBoA = lineIntersect(sideB.x, sideB.y, sideB.x + ux*1000, sideB.y + uy*1000,
                                             oA1.x, oA1.y, oA2.x, oA2.y);
                const intBoB = lineIntersect(sideB.x, sideB.y, sideB.x + ux*1000, sideB.y + uy*1000,
                                             oB1.x, oB1.y, oB2.x, oB2.y);

                // Pick the intersections that are furthest from the centreline (outer corner)
                const junctionPt = thisEnd === 'start' ? { x: el.x1, y: el.y1 } : { x: el.x2, y: el.y2 };

                // For each side, pick the intersection that extends the wall outward
                function bestIntersection(base, int1, int2, dir) {
                    // dir: -1 for start (extending backward), +1 for end (extending forward)
                    const candidates = [int1, int2].filter(p => p !== null);
                    if (candidates.length === 0) return;
                    // Pick the one that's further in the extension direction
                    let best = null, bestProj = -Infinity;
                    for (const c of candidates) {
                        const proj = ((c.x - junctionPt.x) * ux + (c.y - junctionPt.y) * uy) * dir;
                        if (proj > bestProj) { bestProj = proj; best = c; }
                    }
                    if (best && bestProj > 0) {
                        base.x = best.x;
                        base.y = best.y;
                    }
                }

                const dir = thisEnd === 'start' ? -1 : 1;
                bestIntersection(sideA, intAoA, intAoB, dir);
                bestIntersection(sideB, intBoA, intBoB, dir);
            }

            if (startJoined) mitreCorner('start', startJoined, s1a, s1b);
            if (endJoined) mitreCorner('end', endJoined, s2a, s2b);

            // Convert to screen
            const p1a = coords.realToScreen(s1a.x, s1a.y);
            const p1b = coords.realToScreen(s1b.x, s1b.y);
            const p2a = coords.realToScreen(s2a.x, s2a.y);
            const p2b = coords.realToScreen(s2b.x, s2b.y);

            ctx.save();

            // Fill rectangle
            ctx.fillStyle = el.wallType === 'stud' ? '#F0F0F0' : '#FFFFFF';
            ctx.beginPath();
            ctx.moveTo(p1a.x, p1a.y); ctx.lineTo(p2a.x, p2a.y);
            ctx.lineTo(p2b.x, p2b.y); ctx.lineTo(p1b.x, p1b.y);
            ctx.closePath();
            ctx.fill();

            // Supported-wall indicator: if wall below aligns, add subtle grey tint
            if (typeof project.getGhostElements === 'function') {
                const ghostWalls = project.getGhostElements().filter(g => g.type === 'wall');
                const wallMidX = (el.x1 + el.x2) / 2, wallMidY = (el.y1 + el.y2) / 2;
                const alignTol = Math.max(el.thickness, 300); // tolerance for alignment check
                for (const gw of ghostWalls) {
                    // Check if ghost wall centreline is close to this wall centreline
                    const gwMidX = (gw.x1 + gw.x2) / 2, gwMidY = (gw.y1 + gw.y2) / 2;
                    const dist = Math.sqrt(Math.pow(wallMidX - gwMidX, 2) + Math.pow(wallMidY - gwMidY, 2));
                    if (dist < alignTol + wlen / 2) {
                        // Check angular alignment (walls should be roughly parallel)
                        const gwDx = gw.x2 - gw.x1, gwDy = gw.y2 - gw.y1;
                        const gwLen = Math.sqrt(gwDx * gwDx + gwDy * gwDy);
                        if (gwLen > 0.1) {
                            const dot = Math.abs((wdx * gwDx + wdy * gwDy) / (wlen * gwLen));
                            if (dot > 0.85) { // roughly parallel (within ~30°)
                                // Project ghost wall midpoint onto this wall line — check perpendicular distance
                                const t = ((gwMidX - el.x1) * wdx + (gwMidY - el.y1) * wdy) / (wlen * wlen);
                                const projX = el.x1 + t * wdx, projY = el.y1 + t * wdy;
                                const perpDist = Math.sqrt(Math.pow(gwMidX - projX, 2) + Math.pow(gwMidY - projY, 2));
                                if (perpDist < alignTol) {
                                    // Wall is supported — draw subtle grey tint
                                    ctx.globalAlpha = 0.12;
                                    ctx.fillStyle = '#8899AA';
                                    ctx.beginPath();
                                    ctx.moveTo(p1a.x, p1a.y); ctx.lineTo(p2a.x, p2a.y);
                                    ctx.lineTo(p2b.x, p2b.y); ctx.lineTo(p1b.x, p1b.y);
                                    ctx.closePath();
                                    ctx.fill();
                                    ctx.globalAlpha = 1.0;
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            // Hatch fill (clipped to wall polygon, using GLOBAL sheet coords for continuity at corners)
            if (el.wallType !== 'stud') {
                ctx.save();
                ctx.beginPath();
                ctx.moveTo(p1a.x, p1a.y); ctx.lineTo(p2a.x, p2a.y);
                ctx.lineTo(p2b.x, p2b.y); ctx.lineTo(p1b.x, p1b.y);
                ctx.closePath();
                ctx.clip();

                // Global bounds for hatch (use sheet origin so hatch aligns across walls)
                const sheetOrigin = coords.sheetToScreen(0, 0);
                const sheetEnd = coords.sheetToScreen(CONFIG.SHEET_WIDTH_MM, CONFIG.SHEET_HEIGHT_MM);
                const gBounds = {
                    x: Math.min(p1a.x, p1b.x, p2a.x, p2b.x) - 10,
                    y: Math.min(p1a.y, p1b.y, p2a.y, p2b.y) - 10,
                    w: Math.abs(Math.max(p1a.x, p1b.x, p2a.x, p2b.x) - Math.min(p1a.x, p1b.x, p2a.x, p2b.x)) + 20,
                    h: Math.abs(Math.max(p1a.y, p1b.y, p2a.y, p2b.y) - Math.min(p1a.y, p1b.y, p2a.y, p2b.y)) + 20,
                };

                if (el.wallType === 'block') {
                    // 45° hatch using global grid so corners align
                    ctx.strokeStyle = '#000000';
                    ctx.lineWidth = Math.max(0.5, 0.15 * zoom);
                    const spacing = Math.max(4, 3 * zoom);
                    // Use sheet-origin-based offset so all walls share the same hatch grid
                    const ox = sheetOrigin.x % spacing;
                    const oy = sheetOrigin.y % spacing;
                    const diag = gBounds.w + gBounds.h;
                    ctx.beginPath();
                    for (let d = -diag; d < diag; d += spacing) {
                        ctx.moveTo(gBounds.x + d + ox, gBounds.y + oy);
                        ctx.lineTo(gBounds.x + d + gBounds.h + ox, gBounds.y + gBounds.h + oy);
                    }
                    ctx.stroke();
                } else if (el.wallType === 'concrete') {
                    ctx.fillStyle = '#000000';
                    const spacing = Math.max(4, 2.5 * zoom);
                    // Global-aligned stipple
                    const startX = Math.floor(gBounds.x / spacing) * spacing;
                    const startY = Math.floor(gBounds.y / spacing) * spacing;
                    for (let x = startX; x < gBounds.x + gBounds.w; x += spacing) {
                        for (let y = startY; y < gBounds.y + gBounds.h; y += spacing) {
                            const pox = ((x * 7 + y * 13) % spacing) - spacing / 2;
                            const poy = ((x * 11 + y * 3) % spacing) - spacing / 2;
                            ctx.beginPath();
                            ctx.arc(x + pox * 0.5, y + poy * 0.5, Math.max(0.5, 0.3 * zoom), 0, Math.PI * 2);
                            ctx.fill();
                        }
                    }
                }
                ctx.restore();
            }

            // Stud interior lines
            if (el.wallType === 'stud') {
                ctx.strokeStyle = '#CCCCCC';
                ctx.lineWidth = Math.max(0.5, 0.1 * zoom);
                // Two lines near each face
                for (const sign of [0.35, -0.35]) {
                    const lx1 = el.x1 + nx * halfT * sign;
                    const ly1 = el.y1 + ny * halfT * sign;
                    const lx2 = el.x2 + nx * halfT * sign;
                    const ly2 = el.y2 + ny * halfT * sign;
                    const ls1 = coords.realToScreen(lx1, ly1);
                    const ls2 = coords.realToScreen(lx2, ly2);
                    ctx.beginPath(); ctx.moveTo(ls1.x, ls1.y); ctx.lineTo(ls2.x, ls2.y); ctx.stroke();
                }
            }

            // Edge lines (two parallel sides)
            const edgeColor = el.wallType === 'stud' ? '#888888' : '#000000';
            const edgeW = el.wallType === 'concrete' ? Math.max(1, 0.5 * zoom) :
                          el.wallType === 'block' ? Math.max(1, 0.35 * zoom) :
                          Math.max(0.5, 0.15 * zoom);
            ctx.strokeStyle = isSelected ? '#2B7CD0' : edgeColor;
            ctx.lineWidth = edgeW;
            ctx.beginPath(); ctx.moveTo(p1a.x, p1a.y); ctx.lineTo(p2a.x, p2a.y); ctx.stroke();
            ctx.beginPath(); ctx.moveTo(p1b.x, p1b.y); ctx.lineTo(p2b.x, p2b.y); ctx.stroke();

            // End caps (solid line closing each end, unless joined to another wall)
            if (!startJoined) {
                ctx.beginPath(); ctx.moveTo(p1a.x, p1a.y); ctx.lineTo(p1b.x, p1b.y); ctx.stroke();
            }
            if (!endJoined) {
                ctx.beginPath(); ctx.moveTo(p2a.x, p2a.y); ctx.lineTo(p2b.x, p2b.y); ctx.stroke();
            }

            // Wall tag rendering (rotated to match wall angle)
            if (el.tag && zoom > 0.3) {
                const midX = (el.x1 + el.x2) / 2;
                const midY = (el.y1 + el.y2) / 2;
                const midScreen = coords.realToScreen(midX, midY);

                // Calculate wall angle, keep text readable (flip if upside-down)
                let wallAngle = Math.atan2(wdy, wdx);
                // Convert to screen angle (Y is inverted in screen coords)
                const p1s = coords.realToScreen(el.x1, el.y1);
                const p2s = coords.realToScreen(el.x2, el.y2);
                let screenAngle = Math.atan2(p2s.y - p1s.y, p2s.x - p1s.x);
                if (screenAngle > Math.PI / 2) screenAngle -= Math.PI;
                if (screenAngle < -Math.PI / 2) screenAngle += Math.PI;

                const fontSize = Math.max(1, 2.5 * zoom);
                const baseOffset = (halfT / CONFIG.drawingScale) * zoom + fontSize * 0.6 + 2;
                // Apply custom tag offset if user has dragged it
                const tagDx = (el._tagOffsetX || 0) / CONFIG.drawingScale * zoom;
                const tagDy = (el._tagOffsetY || 0) / CONFIG.drawingScale * zoom;
                ctx.save();
                ctx.translate(midScreen.x + tagDx, midScreen.y + tagDy);
                ctx.rotate(screenAngle);
                ctx.font = `${fontSize}px "Segoe UI", Arial, sans-serif`;
                ctx.fillStyle = isSelected ? '#2B7CD0' : '#000000';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                ctx.fillText(el.tag, 0, -baseOffset);
                ctx.restore();
            }

            ctx.restore();
        }

        // ── Lines ──
        if (el.type === 'line') {
            const p1 = coords.realToScreen(el.x1, el.y1);
            const p2 = coords.realToScreen(el.x2, el.y2);

            ctx.strokeStyle = isSelected ? '#2B7CD0' : layer.color;
            ctx.lineWidth = Math.max(1, layer.lineWeight * zoom);
            const pattern = DASH_PATTERNS[layer.pattern] || [];
            ctx.setLineDash(pattern.length > 0 ? pattern.map(d => d * zoom) : []);

            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.stroke();

            if (isSelected) {
                const r = 4;
                ctx.fillStyle = '#2B7CD0';
                ctx.setLineDash([]);
                ctx.fillRect(p1.x - r, p1.y - r, r * 2, r * 2);
                ctx.fillRect(p2.x - r, p2.y - r, r * 2, r * 2);
            }
        }

        // ── Columns (filled square with X + tag) ──
        if (el.type === 'column') {
            const cp = coords.realToScreen(el.x, el.y);
            const halfSize = (el.size / CONFIG.drawingScale) * zoom / 2;
            const r = Math.max(3, halfSize);

            // Filled square
            ctx.fillStyle = isSelected ? 'rgba(43,124,208,0.3)' : 'rgba(0,0,0,0.15)';
            ctx.fillRect(cp.x - r, cp.y - r, r * 2, r * 2);

            // Square outline
            ctx.strokeStyle = isSelected ? '#2B7CD0' : '#000000';
            ctx.lineWidth = Math.max(1, 0.5 * zoom);
            ctx.setLineDash([]);
            ctx.strokeRect(cp.x - r, cp.y - r, r * 2, r * 2);

            // X cross inside
            ctx.beginPath();
            ctx.moveTo(cp.x - r, cp.y - r);
            ctx.lineTo(cp.x + r, cp.y + r);
            ctx.moveTo(cp.x + r, cp.y - r);
            ctx.lineTo(cp.x - r, cp.y + r);
            ctx.stroke();

            // Tag label — show member size if assigned, otherwise type ref (SC1)
            if (el.tag && zoom > 0.3) {
                const typeRef = el.typeRef || el.tag;
                const schedData = project.scheduleTypes.column[typeRef];
                const shortTag = (schedData && schedData.size) ? schedData.size : typeRef;
                const fontSize = Math.max(1, 2.5 * zoom);
                // Apply custom tag offset if user has dragged it
                const tagDx = (el._tagOffsetX || 0) / CONFIG.drawingScale * zoom;
                const tagDy = (el._tagOffsetY || 0) / CONFIG.drawingScale * zoom;
                ctx.save();
                ctx.translate(cp.x + r + 3 + tagDx, cp.y + r + 3 + tagDy);
                ctx.rotate(-Math.PI / 4); // 45 degrees
                ctx.font = `${fontSize}px "Segoe UI", Arial, sans-serif`;
                ctx.fillStyle = isSelected ? '#2B7CD0' : '#000000';
                ctx.textAlign = 'left';
                ctx.textBaseline = 'middle';
                ctx.fillText(shortTag, 0, 0);
                ctx.restore();
            }
        }

        // ── Text annotations ──
        if (el.type === 'text') {
            const tp = coords.realToScreen(el.x, el.y);
            const fontSize = Math.max(7, (el.fontSize || 3.5) * zoom);

            ctx.font = `${fontSize}px "Architects Daughter", cursive`;
            ctx.fillStyle = isSelected ? '#2B7CD0' : layer.color;
            ctx.textAlign = 'left';
            ctx.textBaseline = 'middle';
            ctx.fillText(el.text, tp.x, tp.y);

            // Store computed width for hit-testing
            el._renderWidth = ctx.measureText(el.text).width / zoom;
            el._renderFontSize = fontSize;

            if (isSelected) {
                const metrics = ctx.measureText(el.text);
                ctx.strokeStyle = '#2B7CD0';
                ctx.lineWidth = 1;
                ctx.setLineDash([3, 2]);
                ctx.strokeRect(tp.x - 2, tp.y - fontSize / 2 - 2, metrics.width + 4, fontSize + 4);
                ctx.setLineDash([]);
            }
        }
    }

    ctx.setLineDash([]);
    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';

    // Draw wall preview
    if (activeTool === 'wall' && wallToolState.placing && wallToolState.startPoint && wallToolState.currentEnd) {
        const typeSelect = document.getElementById('wall-type');
        const { wallType, thickness } = parseWallType(typeSelect.value);

        const realStart = coords.sheetToReal(wallToolState.startPoint.x, wallToolState.startPoint.y);
        const realEnd = coords.sheetToReal(wallToolState.currentEnd.x, wallToolState.currentEnd.y);

        const dx = realEnd.x - realStart.x;
        const dy = realEnd.y - realStart.y;
        const len = Math.sqrt(dx * dx + dy * dy);
        if (len > 0.1) {
            const nx = -dy / len;
            const ny = dx / len;
            const halfT = thickness / 2;

            const p1a = coords.realToScreen(realStart.x + nx * halfT, realStart.y + ny * halfT);
            const p1b = coords.realToScreen(realStart.x - nx * halfT, realStart.y - ny * halfT);
            const p2a = coords.realToScreen(realEnd.x + nx * halfT, realEnd.y + ny * halfT);
            const p2b = coords.realToScreen(realEnd.x - nx * halfT, realEnd.y - ny * halfT);

            ctx.globalAlpha = 0.4;
            ctx.save();

            // Preview fill
            if (wallType === 'stud') {
                ctx.fillStyle = '#F0F0F0';
            } else if (wallType === 'block') {
                ctx.fillStyle = '#FFFFDD';
            } else if (wallType === 'concrete') {
                ctx.fillStyle = '#E8E8E8';
            }

            ctx.beginPath();
            ctx.moveTo(p1a.x, p1a.y);
            ctx.lineTo(p2a.x, p2a.y);
            ctx.lineTo(p2b.x, p2b.y);
            ctx.lineTo(p1b.x, p1b.y);
            ctx.closePath();
            ctx.fill();

            // Dashed preview outline
            ctx.strokeStyle = '#2B7CD0';
            ctx.lineWidth = Math.max(1, 1 * zoom);
            ctx.setLineDash([5, 3]);
            ctx.stroke();
            ctx.setLineDash([]);

            ctx.restore();
            ctx.globalAlpha = 1.0;

            // Length label
            const midX = (p1a.x + p2a.x) / 2;
            const midY = (p1a.y + p2a.y) / 2;
            const lenText = len >= 1000 ? (len / 1000).toFixed(len >= 10000 ? 1 : 2) + ' m' : Math.round(len) + ' mm';

            ctx.font = 'bold 11px "Segoe UI", Arial, sans-serif';
            ctx.fillStyle = '#2B7CD0';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            ctx.fillText(lenText, midX, midY - 6);
            ctx.textAlign = 'left';
            ctx.textBaseline = 'top';

            // Start point indicator
            ctx.fillStyle = '#2B7CD0';
            ctx.beginPath();
            ctx.arc(p1a.x, p1a.y, 3, 0, Math.PI * 2);
            ctx.fill();
        }
    }

    // Draw line preview
    if (activeTool === 'line' && lineToolState.placing && lineToolState.startPoint && lineToolState.currentEnd) {
        const layerId = getActiveLayer();
        const layer = project.layers[layerId];

        const sp = coords.sheetToScreen(lineToolState.startPoint.x, lineToolState.startPoint.y);
        const ep = coords.sheetToScreen(lineToolState.currentEnd.x, lineToolState.currentEnd.y);

        ctx.strokeStyle = layer ? layer.color : '#000000';
        ctx.lineWidth = Math.max(1, (layer ? layer.lineWeight : 0.35) * zoom);
        ctx.globalAlpha = 0.6;

        const pattern = DASH_PATTERNS[layer ? layer.pattern : 'solid'] || [];
        ctx.setLineDash(pattern.length > 0 ? pattern.map(d => d * zoom) : []);

        ctx.beginPath();
        ctx.moveTo(sp.x, sp.y);
        ctx.lineTo(ep.x, ep.y);
        ctx.stroke();
        ctx.globalAlpha = 1.0;
        ctx.setLineDash([]);

        // Length label
        const realStart = coords.sheetToReal(lineToolState.startPoint.x, lineToolState.startPoint.y);
        const realEnd = coords.sheetToReal(lineToolState.currentEnd.x, lineToolState.currentEnd.y);
        const lenMM = Math.sqrt(Math.pow(realEnd.x - realStart.x, 2) + Math.pow(realEnd.y - realStart.y, 2));
        const midX = (sp.x + ep.x) / 2;
        const midY = (sp.y + ep.y) / 2;
        let lenText = lenMM >= 1000 ? (lenMM / 1000).toFixed(lenMM >= 10000 ? 1 : 2) + ' m' : Math.round(lenMM) + ' mm';

        ctx.font = 'bold 11px "Segoe UI", Arial, sans-serif';
        ctx.fillStyle = '#2B7CD0';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(lenText, midX, midY - 6);
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';

        ctx.fillStyle = '#2B7CD0';
        ctx.beginPath();
        ctx.arc(sp.x, sp.y, 3, 0, Math.PI * 2);
        ctx.fill();
    }

    // Column preview at cursor (always follows mouse, snaps when near node)
    if (activeTool === 'column') {
        const snap = snapState.activeSnap;
        let cp;
        if (snap) {
            cp = coords.sheetToScreen(snap.x, snap.y);
        } else {
            cp = { x: cursorPos.screenX, y: cursorPos.screenY };
        }
        const halfSize = (COL_DEFAULT_SIZE / CONFIG.drawingScale) * zoom / 2;
        const r = Math.max(3, halfSize);
        ctx.save();
        ctx.globalAlpha = snap ? 0.5 : 0.3;
        ctx.fillStyle = '#2B7CD0';
        ctx.fillRect(cp.x - r, cp.y - r, r * 2, r * 2);
        ctx.strokeStyle = '#1a5fa0';
        ctx.lineWidth = Math.max(1, 0.5 * zoom);
        ctx.strokeRect(cp.x - r, cp.y - r, r * 2, r * 2);
        // Cross-hatch lines
        ctx.beginPath();
        ctx.moveTo(cp.x - r, cp.y - r); ctx.lineTo(cp.x + r, cp.y + r);
        ctx.moveTo(cp.x + r, cp.y - r); ctx.lineTo(cp.x - r, cp.y + r);
        ctx.stroke();
        // Centre dot
        ctx.globalAlpha = 1.0;
        ctx.fillStyle = '#ff3333';
        ctx.beginPath();
        ctx.arc(cp.x, cp.y, 2.5, 0, Math.PI * 2);
        ctx.fill();
        ctx.restore();
    }

    // Footing preview at cursor — uses active placement type from schedule
    if (activeTool === 'footing') {
        const snap = snapState.activeSnap;
        let cp;
        if (snap) {
            cp = coords.sheetToScreen(snap.x, snap.y);
        } else {
            cp = { x: cursorPos.screenX, y: cursorPos.screenY };
        }
        // Get dimensions from active placement type
        const _pfRef = (typeof placementTypeRef !== 'undefined' ? placementTypeRef.footing : null) || 'PF1';
        const _pfData = (project.scheduleTypes.padfooting || {})[_pfRef] || {};
        const _pfW = _pfData.width || 1000;
        const _pfL = _pfData.length || _pfW;
        const halfW = (_pfW / 2 / CONFIG.drawingScale) * zoom;
        const halfL = (_pfL / 2 / CONFIG.drawingScale) * zoom;
        const rw = Math.max(6, halfW);
        const rl = Math.max(6, halfL);
        // Colour from schedule type
        const _pfColor = _pfData.color || '#1a5fa0';
        ctx.save();
        // Fill with type colour
        ctx.globalAlpha = snap ? 0.15 : 0.08;
        ctx.fillStyle = _pfColor;
        ctx.fillRect(cp.x - rw, cp.y - rl, rw * 2, rl * 2);
        // Outline
        ctx.globalAlpha = snap ? 0.6 : 0.3;
        ctx.strokeStyle = _pfColor;
        ctx.lineWidth = Math.max(1, 0.6 * zoom);
        ctx.strokeRect(cp.x - rw, cp.y - rl, rw * 2, rl * 2);
        // Inner cross
        ctx.setLineDash([3, 3]);
        ctx.beginPath();
        ctx.moveTo(cp.x - rw, cp.y); ctx.lineTo(cp.x + rw, cp.y);
        ctx.moveTo(cp.x, cp.y - rl); ctx.lineTo(cp.x, cp.y + rl);
        ctx.stroke();
        ctx.setLineDash([]);
        // Type label above the preview
        ctx.globalAlpha = 0.7;
        ctx.fillStyle = _pfColor;
        ctx.font = `bold ${Math.max(1, 10 * zoom)}px sans-serif`;
        ctx.textAlign = 'center';
        ctx.fillText(_pfRef, cp.x, cp.y - rl - 4 * zoom);
        // Centre dot
        ctx.globalAlpha = 1.0;
        ctx.fillStyle = '#ff3333';
        ctx.beginPath();
        ctx.arc(cp.x, cp.y, 2.5, 0, Math.PI * 2);
        ctx.fill();
        ctx.restore();
    }
}

// Register element renderer (after grids, before snap indicator)
engine._renderCallbacks.splice(3, 0, drawElements);

// ── Node Handle Rendering for Selected Elements ──────────
engine.onRender((ctx, eng) => {
    if (activeTool !== 'select') return;
    if (!selectedElement) return;
    const el = selectedElement;
    // Only show node handles for elements with endpoints
    if (el.type !== 'wall' && el.type !== 'line' && el.type !== 'stripFooting') return;

    const p1 = eng.coords.realToSheet(el.x1, el.y1);
    const p2 = eng.coords.realToSheet(el.x2, el.y2);
    const sp1 = eng.coords.sheetToScreen(p1.x, p1.y);
    const sp2 = eng.coords.sheetToScreen(p2.x, p2.y);
    const r = 5;

    ctx.save();
    // Start node handle
    ctx.fillStyle = '#ffffff';
    ctx.strokeStyle = '#2B7CD0';
    ctx.lineWidth = 2;
    ctx.fillRect(sp1.x - r, sp1.y - r, r * 2, r * 2);
    ctx.strokeRect(sp1.x - r, sp1.y - r, r * 2, r * 2);
    // End node handle
    ctx.fillRect(sp2.x - r, sp2.y - r, r * 2, r * 2);
    ctx.strokeRect(sp2.x - r, sp2.y - r, r * 2, r * 2);
    ctx.restore();
});

// ── Selection + Drag-to-Move ─────────────────────────────

let selectedElement = null;

const dragState = {
    dragging: false,
    el: null,
    startSheet: null,    // where the mouse went down (sheet-mm)
    origCoords: null,    // original element coords before drag
    nodeIndex: -1,       // -1 = whole element, 0 = start node, 1 = end node
};

function hitTestElement(sheetPos, tolerance) {
    for (let i = project.elements.length - 1; i >= 0; i--) {
        const el = project.elements[i];
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;

        // === CORE ELEMENTS ===
        if (el.type === 'line') {
            const p1 = engine.coords.realToSheet(el.x1, el.y1);
            const p2 = engine.coords.realToSheet(el.x2, el.y2);
            if (pointToSegmentDist(sheetPos.x, sheetPos.y, p1.x, p1.y, p2.x, p2.y) < tolerance)
                return el;
        }
        if (el.type === 'column') {
            const cp = engine.coords.realToSheet(el.x, el.y);
            const hs = (el.size / CONFIG.drawingScale) / 2 + tolerance;
            if (Math.abs(sheetPos.x - cp.x) < hs && Math.abs(sheetPos.y - cp.y) < hs) return el;
        }
        if (el.type === 'text') {
            const tp = engine.coords.realToSheet(el.x, el.y);
            const tw = (el._renderWidth || el.text.length * (el.fontSize || 3.5) * 0.7) + 2;
            const th = (el.fontSize || 3.5) * 1.6;
            if (sheetPos.x >= tp.x - 1 && sheetPos.x <= tp.x + tw &&
                sheetPos.y >= tp.y - th / 2 - 1 && sheetPos.y <= tp.y + th / 2 + 1) return el;
        }
        if (el.type === 'dimension') {
            const p1 = engine.coords.realToSheet(el.x1, el.y1);
            const p2 = engine.coords.realToSheet(el.x2, el.y2);
            if (pointToSegmentDist(sheetPos.x, sheetPos.y, p1.x, p1.y, p2.x, p2.y) < tolerance + (el.offset || 5))
                return el;
        }

        // === POLYLINE (check both edges and interior) ===
        if (el.type === 'polyline') {
            const pts = el.points;
            // Edge hit-test
            for (let j = 0; j < pts.length - 1; j++) {
                const p1 = engine.coords.realToSheet(pts[j].x, pts[j].y);
                const p2 = engine.coords.realToSheet(pts[j+1].x, pts[j+1].y);
                if (pointToSegmentDist(sheetPos.x, sheetPos.y, p1.x, p1.y, p2.x, p2.y) < tolerance)
                    return el;
            }
            // Closed polyline: check closing edge and interior
            if (el.closed && pts.length > 2) {
                const pLast = engine.coords.realToSheet(pts[pts.length-1].x, pts[pts.length-1].y);
                const pFirst = engine.coords.realToSheet(pts[0].x, pts[0].y);
                if (pointToSegmentDist(sheetPos.x, sheetPos.y, pLast.x, pLast.y, pFirst.x, pFirst.y) < tolerance)
                    return el;
                // Interior hit-test
                const sheetPts = pts.map(p => engine.coords.realToSheet(p.x, p.y));
                if (pointInPolygon(sheetPos.x, sheetPos.y, sheetPts))
                    return el;
            }
        }

        // === LEADER ===
        if (el.type === 'leader') {
            const p1 = engine.coords.realToSheet(el.x1, el.y1);
            const p2 = engine.coords.realToSheet(el.x2, el.y2);
            if (pointToSegmentDist(sheetPos.x, sheetPos.y, p1.x, p1.y, p2.x, p2.y) < tolerance + 2)
                return el;
        }

        // === CLOUD ===
        if (el.type === 'cloud') {
            // Check segments of cloud path
            for (let j = 0; j < el.points.length - 1; j++) {
                const p1 = engine.coords.realToSheet(el.points[j].x, el.points[j].y);
                const p2 = engine.coords.realToSheet(el.points[j+1].x, el.points[j+1].y);
                if (pointToSegmentDist(sheetPos.x, sheetPos.y, p1.x, p1.y, p2.x, p2.y) < tolerance + 3)
                    return el;
            }
        }

        // === CHAIN DIMENSION ===
        if (el.type === 'chaindim') {
            for (let j = 0; j < el.points.length - 1; j++) {
                const p1 = engine.coords.realToSheet(el.points[j].x, el.points[j].y);
                const p2 = engine.coords.realToSheet(el.points[j+1].x, el.points[j+1].y);
                if (pointToSegmentDist(sheetPos.x, sheetPos.y, p1.x, p1.y, p2.x, p2.y) < tolerance + 8)
                    return el;
            }
        }

        // === NOTES BOX ===
        if (el.type === 'notesbox') {
            const pos = engine.coords.realToSheet(el.x, el.y);
            if (sheetPos.x >= pos.x - tolerance && sheetPos.x <= pos.x + el.width + tolerance &&
                sheetPos.y >= pos.y - tolerance && sheetPos.y <= pos.y + el.height + tolerance)
                return el;
        }

        // === TABLE ===
        if (el.type === 'table') {
            const pos = engine.coords.realToSheet(el.x, el.y);
            const cols = el.rows[0]?.length || 1;
            const rows = el.rows?.length || 1;
            const totalW = cols * el.cellWidth;
            const totalH = rows * el.cellHeight;
            if (sheetPos.x >= pos.x - tolerance && sheetPos.x <= pos.x + totalW + tolerance &&
                sheetPos.y >= pos.y - tolerance && sheetPos.y <= pos.y + totalH + tolerance)
                return el;
        }

        // === SLAB CALLOUT ===
        if (el.type === 'slabcallout') {
            const pos = engine.coords.realToSheet(el.x, el.y);
            if (sheetPos.x >= pos.x - tolerance && sheetPos.x <= pos.x + el.width + tolerance &&
                sheetPos.y >= pos.y - tolerance && sheetPos.y <= pos.y + el.height + tolerance)
                return el;
        }

        // === PHASE B TOOLS (Edge, Fall Arrow, Step Line) ===
        if (el.type === 'edge' || el.type === 'fallarrow' || el.type === 'step') {
            const p1 = engine.coords.realToSheet(el.x1, el.y1);
            const p2 = engine.coords.realToSheet(el.x2, el.y2);
            if (pointToSegmentDist(sheetPos.x, sheetPos.y, p1.x, p1.y, p2.x, p2.y) < tolerance + 2)
                return el;
        }

        // === PHASE C TOOLS (RL Marker, Section Ref) ===
        if (el.type === 'rlmarker') {
            const cp = engine.coords.realToSheet(el.x, el.y);
            const hs = (6 / CONFIG.drawingScale) + tolerance;
            if (Math.abs(sheetPos.x - cp.x) < hs && Math.abs(sheetPos.y - cp.y) < hs) return el;
        }
        if (el.type === 'secref') {
            const cp = engine.coords.realToSheet(el.x, el.y);
            const hs = (8 / CONFIG.drawingScale) + tolerance;
            if (Math.abs(sheetPos.x - cp.x) < hs && Math.abs(sheetPos.y - cp.y) < hs) return el;
        }
        if (el.type === 'reftag') {
            const cp = engine.coords.realToSheet(el.x, el.y);
            const hs = (8 / CONFIG.drawingScale) + tolerance;
            if (Math.abs(sheetPos.x - cp.x) < hs && Math.abs(sheetPos.y - cp.y) < hs) return el;
        }

        // === PHASE D TOOLS (Footing) ===
        if (el.type === 'footing') {
            const cp = engine.coords.realToSheet(el.x, el.y);
            const fw = el.footingWidth || el.width || 1000;
            const hs = (fw / 2 / CONFIG.drawingScale) + tolerance;
            if (Math.abs(sheetPos.x - cp.x) < hs && Math.abs(sheetPos.y - cp.y) < hs) return el;
        }

        // === WALL ===
        if (el.type === 'wall') {
            const p1 = engine.coords.realToSheet(el.x1, el.y1);
            const p2 = engine.coords.realToSheet(el.x2, el.y2);
            const dx = p2.x - p1.x;
            const dy = p2.y - p1.y;
            const len = Math.sqrt(dx * dx + dy * dy);
            if (len < 0.1) continue;

            // Hit test perpendicular distance (thickness + tolerance)
            const nx = -dy / len;
            const ny = dx / len;
            const halfT = (el.thickness / CONFIG.drawingScale / 2) + tolerance;

            // Check point-to-wall distance
            const t = Math.max(0, Math.min(len, (sheetPos.x - p1.x) * dx + (sheetPos.y - p1.y) * dy)) / (len * len);
            const closestX = p1.x + t * (p2.x - p1.x);
            const closestY = p1.y + t * (p2.y - p1.y);
            const dist = Math.sqrt(Math.pow(sheetPos.x - closestX, 2) + Math.pow(sheetPos.y - closestY, 2));

            if (dist < halfT) return el;
        }

        // === BRACING WALL (line-based hit test) ===
        if (el.type === 'bracingWall') {
            const p1 = engine.coords.realToSheet(el.x1, el.y1);
            const p2 = engine.coords.realToSheet(el.x2, el.y2);
            const dx = p2.x - p1.x;
            const dy = p2.y - p1.y;
            const len = Math.sqrt(dx * dx + dy * dy);
            if (len < 0.1) continue;
            const t = Math.max(0, Math.min(len, (sheetPos.x - p1.x) * dx + (sheetPos.y - p1.y) * dy)) / (len * len);
            const closestX = p1.x + t * (p2.x - p1.x);
            const closestY = p1.y + t * (p2.y - p1.y);
            const dist = Math.sqrt(Math.pow(sheetPos.x - closestX, 2) + Math.pow(sheetPos.y - closestY, 2));
            if (dist < tolerance + 3 / engine.viewport.zoom) return el;
        }

        // === BUILDING ENVELOPE (polygon edge hit test) ===
        if (el.type === 'buildingEnvelope' && el.points && el.points.length >= 3) {
            for (let i = 0; i < el.points.length; i++) {
                const j = (i + 1) % el.points.length;
                const p1 = engine.coords.realToSheet(el.points[i].x, el.points[i].y);
                const p2 = engine.coords.realToSheet(el.points[j].x, el.points[j].y);
                const dx = p2.x - p1.x;
                const dy = p2.y - p1.y;
                const len = Math.sqrt(dx * dx + dy * dy);
                if (len < 0.1) continue;
                const t = Math.max(0, Math.min(len, (sheetPos.x - p1.x) * dx + (sheetPos.y - p1.y) * dy)) / (len * len);
                const closestX = p1.x + t * (p2.x - p1.x);
                const closestY = p1.y + t * (p2.y - p1.y);
                const dist = Math.sqrt(Math.pow(sheetPos.x - closestX, 2) + Math.pow(sheetPos.y - closestY, 2));
                if (dist < tolerance + 3 / engine.viewport.zoom) return el;
            }
        }

        // === RIDGE LINE (polyline edge hit test) ===
        if (el.type === 'ridgeLine' && el.points && el.points.length >= 2) {
            for (let i = 0; i < el.points.length - 1; i++) {
                const p1 = engine.coords.realToSheet(el.points[i].x, el.points[i].y);
                const p2 = engine.coords.realToSheet(el.points[i + 1].x, el.points[i + 1].y);
                const dx = p2.x - p1.x;
                const dy = p2.y - p1.y;
                const len = Math.sqrt(dx * dx + dy * dy);
                if (len < 0.1) continue;
                const t = Math.max(0, Math.min(len, (sheetPos.x - p1.x) * dx + (sheetPos.y - p1.y) * dy)) / (len * len);
                const closestX = p1.x + t * (p2.x - p1.x);
                const closestY = p1.y + t * (p2.y - p1.y);
                const dist = Math.sqrt(Math.pow(sheetPos.x - closestX, 2) + Math.pow(sheetPos.y - closestY, 2));
                if (dist < tolerance + 3 / engine.viewport.zoom) return el;
            }
        }
    }
    return null;
}

// Mousedown — select and start drag (all element types)
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (e.altKey) return; // Alt+drag is reserved for tag dragging
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'select') return;
    if (pdfState.calibrating) return;

    const sheetPos = engine.getSheetPos(e);
    const tolerance = 4 / engine.viewport.zoom;
    const hit = hitTestElement(sheetPos, tolerance);

    selectedElement = hit;

    if (hit) {
        dragState.dragging = true;
        dragState.el = hit;
        dragState.startSheet = { x: sheetPos.x, y: sheetPos.y };
        dragState.nodeIndex = -1; // default: drag whole element

        // Store original coords based on element type
        if (hit.type === 'line' || hit.type === 'dimension' || hit.type === 'leader' || hit.type === 'callout' || hit.type === 'wall' || hit.type === 'stripFooting' || hit.type === 'bracingWall') {
            dragState.origCoords = { x1: hit.x1, y1: hit.y1, x2: hit.x2, y2: hit.y2 };

            // Callout: check if near arrow tip (node 0) — otherwise drag text box only (node 1)
            // This gives Bluebeam-style behaviour: drag body = move text box, arrow stays put
            if (hit.type === 'callout') {
                const nodeTol = 6 / engine.viewport.zoom;
                const p1 = engine.coords.realToSheet(hit.x1, hit.y1);
                const d1 = Math.sqrt(Math.pow(sheetPos.x - p1.x, 2) + Math.pow(sheetPos.y - p1.y, 2));
                if (d1 < nodeTol) {
                    dragState.nodeIndex = 0; // dragging arrow tip
                } else {
                    dragState.nodeIndex = 1; // dragging text box (default for callout)
                }
            }

            // Check if near an endpoint — if so, drag just that node
            if (hit.type === 'wall' || hit.type === 'line' || hit.type === 'stripFooting' || hit.type === 'bracingWall') {
                const nodeTol = 6 / engine.viewport.zoom; // screen pixels converted to sheet-mm
                const p1 = engine.coords.realToSheet(hit.x1, hit.y1);
                const p2 = engine.coords.realToSheet(hit.x2, hit.y2);
                const d1 = Math.sqrt(Math.pow(sheetPos.x - p1.x, 2) + Math.pow(sheetPos.y - p1.y, 2));
                const d2 = Math.sqrt(Math.pow(sheetPos.x - p2.x, 2) + Math.pow(sheetPos.y - p2.y, 2));
                if (d1 < nodeTol && d1 <= d2) {
                    dragState.nodeIndex = 0; // dragging start node
                } else if (d2 < nodeTol) {
                    dragState.nodeIndex = 1; // dragging end node
                }
            }
        } else if (hit.type === 'polyline' || hit.type === 'cloud') {
            dragState.origCoords = { points: hit.points.map(p => ({ x: p.x, y: p.y })) };
        } else {
            dragState.origCoords = { x: hit.x, y: hit.y };
        }
        container.style.cursor = dragState.nodeIndex >= 0 ? 'crosshair' : 'move';

        // Capture linked footings for auto-link drag movement
        const linked = findLinkedFootings(hit);
        if (linked.length > 0) {
            dragState._linkedFootings = linked.map(f => ({
                footing: f,
                orig: f.x1 !== undefined
                    ? { x1: f.x1, y1: f.y1, x2: f.x2, y2: f.y2 }
                    : { x: f.x, y: f.y }
            }));
        } else {
            dragState._linkedFootings = null;
        }
    }

    engine.requestRender();
});

// Mousemove — drag element
window.addEventListener('mousemove', (e) => {
    if (!dragState.dragging) return;
    const sheetPos = engine.getSheetPos(e);
    const dx = sheetPos.x - dragState.startSheet.x;
    const dy = sheetPos.y - dragState.startSheet.y;
    const realDx = dx * CONFIG.drawingScale;
    const realDy = dy * CONFIG.drawingScale;

    const el = dragState.el;
    if (el.type === 'line' || el.type === 'dimension' || el.type === 'leader' || el.type === 'callout' || el.type === 'edge' || el.type === 'fallarrow' || el.type === 'step' || el.type === 'wall' || el.type === 'stripFooting' || el.type === 'bracingWall') {
        if (dragState.nodeIndex === 0) {
            // Drag start node only — use snap for precision
            const snap = findSnap(e.clientX - container.getBoundingClientRect().left, e.clientY - container.getBoundingClientRect().top);
            if (snap) {
                const realSnap = engine.coords.sheetToReal(snap.x, snap.y);
                el.x1 = realSnap.x;
                el.y1 = realSnap.y;
            } else {
                el.x1 = dragState.origCoords.x1 + realDx;
                el.y1 = dragState.origCoords.y1 + realDy;
            }
        } else if (dragState.nodeIndex === 1) {
            // Drag end node only — use snap for precision
            const snap = findSnap(e.clientX - container.getBoundingClientRect().left, e.clientY - container.getBoundingClientRect().top);
            if (snap) {
                const realSnap = engine.coords.sheetToReal(snap.x, snap.y);
                el.x2 = realSnap.x;
                el.y2 = realSnap.y;
            } else {
                el.x2 = dragState.origCoords.x2 + realDx;
                el.y2 = dragState.origCoords.y2 + realDy;
            }
        } else {
            // Drag whole element
            el.x1 = dragState.origCoords.x1 + realDx;
            el.y1 = dragState.origCoords.y1 + realDy;
            el.x2 = dragState.origCoords.x2 + realDx;
            el.y2 = dragState.origCoords.y2 + realDy;
        }

        // Linked movement: drag linked footings along with wall
        if (el.type === 'wall' && dragState._linkedFootings) {
            for (const lf of dragState._linkedFootings) {
                if (dragState.nodeIndex === 0) {
                    lf.footing.x1 = el.x1; lf.footing.y1 = el.y1;
                } else if (dragState.nodeIndex === 1) {
                    lf.footing.x2 = el.x2; lf.footing.y2 = el.y2;
                } else {
                    lf.footing.x1 = lf.orig.x1 + realDx; lf.footing.y1 = lf.orig.y1 + realDy;
                    lf.footing.x2 = lf.orig.x2 + realDx; lf.footing.y2 = lf.orig.y2 + realDy;
                }
            }
        }
    } else {
        el.x = dragState.origCoords.x + realDx;
        el.y = dragState.origCoords.y + realDy;

        // Linked movement: drag linked pad footing along with column
        if (el.type === 'column' && dragState._linkedFootings) {
            for (const lf of dragState._linkedFootings) {
                lf.footing.x = lf.orig.x + realDx;
                lf.footing.y = lf.orig.y + realDy;
            }
        }
    }

    engine.requestRender();
});

// Mouseup — finish drag
window.addEventListener('mouseup', (e) => {
    if (!dragState.dragging) return;

    const el = dragState.el;
    const orig = dragState.origCoords;
    const isLineType = (el.type === 'line' || el.type === 'dimension' || el.type === 'leader' || el.type === 'callout' || el.type === 'edge' || el.type === 'fallarrow' || el.type === 'step' || el.type === 'wall' || el.type === 'stripFooting');
    const moved = isLineType
        ? (Math.abs(el.x1 - orig.x1) > 0.1 || Math.abs(el.y1 - orig.y1) > 0.1 || Math.abs(el.x2 - orig.x2) > 0.1 || Math.abs(el.y2 - orig.y2) > 0.1)
        : (el.type === 'polyline' || el.type === 'cloud')
            ? true // polylines always count as moved if drag happened
            : (Math.abs(el.x - orig.x) > 0.1 || Math.abs(el.y - orig.y) > 0.1);

    // Snapshot linked footing positions for undo
    const linkedSnap = dragState._linkedFootings
        ? dragState._linkedFootings.map(lf => {
            const f = lf.footing;
            return f.x1 !== undefined
                ? { footing: f, newCoords: { x1: f.x1, y1: f.y1, x2: f.x2, y2: f.y2 }, origCoords: { ...lf.orig } }
                : { footing: f, newCoords: { x: f.x, y: f.y }, origCoords: { ...lf.orig } };
        })
        : [];

    if (moved) {
        if (isLineType) {
            const nx1 = el.x1, ny1 = el.y1, nx2 = el.x2, ny2 = el.y2;
            const desc = dragState.nodeIndex >= 0 ? 'Move node' : 'Move element';
            history.undoStack.push({
                description: desc + (linkedSnap.length ? ' + linked footing' : ''),
                execute() {
                    el.x1 = nx1; el.y1 = ny1; el.x2 = nx2; el.y2 = ny2;
                    for (const ls of linkedSnap) Object.assign(ls.footing, ls.newCoords);
                },
                undo() {
                    el.x1 = orig.x1; el.y1 = orig.y1; el.x2 = orig.x2; el.y2 = orig.y2;
                    for (const ls of linkedSnap) Object.assign(ls.footing, ls.origCoords);
                }
            });
        } else if (!(el.type === 'polyline' || el.type === 'cloud')) {
            const nx = el.x, ny = el.y;
            history.undoStack.push({
                description: 'Move element' + (linkedSnap.length ? ' + linked footing' : ''),
                execute() {
                    el.x = nx; el.y = ny;
                    for (const ls of linkedSnap) Object.assign(ls.footing, ls.newCoords);
                },
                undo() {
                    el.x = orig.x; el.y = orig.y;
                    for (const ls of linkedSnap) Object.assign(ls.footing, ls.origCoords);
                }
            });
        }
        history.redoStack.length = 0;
    }

    dragState.dragging = false;
    dragState.el = null;
    dragState._linkedFootings = null;
    container.style.cursor = '';
    engine.requestRender();
});

// ── Tag Drag System (Alt+click+drag to move tags) ─────────
const tagDrag = {
    active: false,
    el: null,
    startScreen: null,
    origOffset: null,
};

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || !e.altKey) return;
    if (activeTool !== 'select') return;

    const sheetPos = engine.getSheetPos(e);
    const tolerance = 6 / engine.viewport.zoom;
    const hit = hitTestElement(sheetPos, tolerance);
    if (!hit) return;

    // Only allow tag drag on elements that have tags
    if (!hit.tag && !hit.mark) return;

    e.preventDefault();
    e.stopPropagation();
    tagDrag.active = true;
    tagDrag.el = hit;
    tagDrag.startScreen = { x: e.clientX, y: e.clientY };
    tagDrag.origOffset = { x: hit._tagOffsetX || 0, y: hit._tagOffsetY || 0 };
    container.style.cursor = 'move';
});

window.addEventListener('mousemove', (e) => {
    if (!tagDrag.active) return;
    const dx = e.clientX - tagDrag.startScreen.x;
    const dy = e.clientY - tagDrag.startScreen.y;
    // Convert screen delta to real-mm delta
    const realDx = dx / engine.viewport.zoom * CONFIG.drawingScale;
    const realDy = dy / engine.viewport.zoom * CONFIG.drawingScale;
    tagDrag.el._tagOffsetX = tagDrag.origOffset.x + realDx;
    tagDrag.el._tagOffsetY = tagDrag.origOffset.y + realDy;
    engine.requestRender();
});

window.addEventListener('mouseup', (e) => {
    if (!tagDrag.active) return;
    const el = tagDrag.el;
    const orig = { ...tagDrag.origOffset };
    const newOff = { x: el._tagOffsetX || 0, y: el._tagOffsetY || 0 };
    if (Math.abs(newOff.x - orig.x) > 0.1 || Math.abs(newOff.y - orig.y) > 0.1) {
        history.undoStack.push({
            description: 'Move tag',
            execute() { el._tagOffsetX = newOff.x; el._tagOffsetY = newOff.y; },
            undo() { el._tagOffsetX = orig.x; el._tagOffsetY = orig.y; }
        });
        history.redoStack.length = 0;
    }
    tagDrag.active = false;
    tagDrag.el = null;
    container.style.cursor = '';
    engine.requestRender();
});

// Delete selected element
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if ((e.key === 'Delete' || e.key === 'Backspace') && selectedElement && activeTool === 'select') {
        const removed = selectedElement;
        history.execute({
            description: 'Delete element',
            execute() {
                const i = project.elements.indexOf(removed);
                if (i !== -1) project.elements.splice(i, 1);
            },
            undo() { project.elements.push(removed); }
        });
        selectedElement = null;
        engine.requestRender();
    }
});

/** Point-to-line-segment distance */
function pointToSegmentDist(px, py, x1, y1, x2, y2) {
    const dx = x2 - x1, dy = y2 - y1;
    const lenSq = dx * dx + dy * dy;
    if (lenSq === 0) return Math.sqrt((px - x1) ** 2 + (py - y1) ** 2);
    let t = ((px - x1) * dx + (py - y1) * dy) / lenSq;
    t = Math.max(0, Math.min(1, t));
    const projX = x1 + t * dx, projY = y1 + t * dy;
    return Math.sqrt((px - projX) ** 2 + (py - projY) ** 2);
}

// ── Layer Panel ──────────────────────────────────────────

const layerPanel = document.getElementById('layer-panel');
const layerList = document.getElementById('layer-list');

// Layers toggle button
document.getElementById('btn-layers').addEventListener('click', () => {
    layerPanel.classList.toggle('hidden');
});

// Close panel on header click
document.getElementById('layer-panel-hdr').addEventListener('click', () => {
    layerPanel.classList.add('hidden');
});

function buildLayerPanel() {
    layerList.innerHTML = '';
    for (const [id, layer] of Object.entries(project.layers)) {
        const row = document.createElement('div');
        row.className = 'layer-row' + (id === elementTypeSelect.value ? ' active' : '') +
            (!layer.visible ? ' hidden-layer' : '');

        // Colour swatch showing line style
        const swatch = document.createElement('div');
        swatch.className = 'layer-swatch';
        swatch.style.background = layer.color;
        swatch.style.height = Math.max(2, layer.lineWeight * 4) + 'px';

        const name = document.createElement('span');
        name.className = 'layer-name';
        name.textContent = layer.name;

        const vis = document.createElement('button');
        vis.className = 'layer-vis';
        vis.textContent = layer.visible ? '👁' : '—';
        vis.title = layer.visible ? 'Hide layer' : 'Show layer';

        vis.addEventListener('click', (e) => {
            e.stopPropagation();
            layer.visible = !layer.visible;
            buildLayerPanel();
            engine.requestRender();
        });

        row.addEventListener('click', () => {
            elementTypeSelect.value = id;
            buildLayerPanel();
        });

        row.appendChild(swatch);
        row.appendChild(name);
        row.appendChild(vis);
        layerList.appendChild(row);
    }
}

buildLayerPanel();

// Rebuild panel when element type changes
elementTypeSelect.addEventListener('change', () => {
    buildLayerPanel();
});

// ── Text Size: apply to selected text + sync on selection ──

const textSizeDropdown = document.getElementById('text-size');

// When dropdown changes, update the selected text element
textSizeDropdown.addEventListener('change', () => {
    if (selectedElement && selectedElement.type === 'text') {
        const newSize = parseFloat(textSizeDropdown.value);
        const oldSize = selectedElement.fontSize;
        if (Math.abs(newSize - oldSize) > 0.01) {
            const el = selectedElement;
            history.execute({
                description: 'Resize text',
                execute() { el.fontSize = newSize; },
                undo() { el.fontSize = oldSize; }
            });
            engine.requestRender();
        }
    }
});

// Sync dropdown when a text element is selected
// (patch into render to check after each frame)
let _lastSyncedElement = null;
engine.onRender(() => {
    if (selectedElement && selectedElement.type === 'text' && selectedElement !== _lastSyncedElement) {
        // Find closest matching option
        const fs = selectedElement.fontSize || 3.5;
        let bestOpt = '3.5';
        let bestDiff = Infinity;
        for (const opt of textSizeDropdown.options) {
            const diff = Math.abs(parseFloat(opt.value) - fs);
            if (diff < bestDiff) { bestDiff = diff; bestOpt = opt.value; }
        }
        textSizeDropdown.value = bestOpt;
        _lastSyncedElement = selectedElement;
    } else if (!selectedElement || selectedElement.type !== 'text') {
        _lastSyncedElement = null;
    }
});

// ══════════════════════════════════════════════════════════
