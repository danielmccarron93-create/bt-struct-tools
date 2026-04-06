// ── app.js ─────────────────────────────────────────────────

const CONFIG = {
    // A1 sheet dimensions in mm
    SHEET_WIDTH_MM: 841,
    SHEET_HEIGHT_MM: 594,

    // Margins (mm)
    MARGIN_LEFT: 20,
    MARGIN_RIGHT: 10,
    MARGIN_TOP: 10,
    MARGIN_BOTTOM: 10,

    // Title block
    TITLE_BLOCK_HEIGHT_MM: 40,

    // Drawing scale (1:100 default — 1mm on sheet = 100mm real)
    drawingScale: 100,

    // Background grid (sheet-mm spacing)
    GRID_MINOR_MM: 5,
    GRID_MAJOR_MM: 25,

    // Grid visibility (off by default)
    gridVisible: false,

    // Theme — Light / Revit-like
    CANVAS_BG: '#D6D6D6',
    SHEET_BG: '#FFFFFF',
    SHEET_BORDER: '#999999',
    GRID_MINOR_COLOR: '#E8E8E8',
    GRID_MAJOR_COLOR: '#D0D0D0',
};

const container = document.getElementById('canvas-container');
const engine = new CanvasEngine(container, CONFIG);
const project = new ProjectData();
const history = new History();

// ── Schedule Types & Colour Palette ────────────────────────

const SCHEDULE_COLORS = [
    '#93C5FD', // blue
    '#86EFAC', // green
    '#FCD34D', // amber
    '#FCA5A5', // rose
    '#67E8F9', // cyan
    '#C4B5FD', // purple
    '#FDBA74', // orange
    '#CBD5E1', // slate
];

// ── Australian Steel Section Database (OneSteel/Liberty) ──────
const STEEL_SECTIONS = {
    UB: [
        '150UB14.0','150UB18.0',
        '180UB16.1','180UB18.1','180UB22.2',
        '200UB18.2','200UB22.3','200UB25.4','200UB29.8',
        '250UB25.7','250UB31.4','250UB37.3',
        '310UB32.0','310UB40.4','310UB46.2',
        '360UB44.7','360UB50.7','360UB56.7',
        '410UB53.7','410UB59.7',
        '460UB67.1','460UB74.6','460UB82.1',
        '530UB82.0','530UB92.4',
        '610UB101','610UB113','610UB125'
    ],
    UC: [
        '100UC14.8',
        '150UC23.4','150UC30.0','150UC37.2',
        '200UC46.2','200UC52.2','200UC59.5',
        '250UC72.9','250UC89.5',
        '310UC96.8','310UC118','310UC137','310UC158'
    ],
    PFC: [
        '75PFC','100PFC','125PFC','150PFC','180PFC',
        '200PFC','230PFC','250PFC','300PFC','380PFC'
    ],
    SHS: [
        '20x1.6','25x1.6','25x2.0',
        '30x1.6','30x2.0',
        '35x1.6','35x2.0',
        '40x1.6','40x2.0','40x2.5','40x3.0',
        '50x1.6','50x2.0','50x2.5','50x3.0','50x4.0',
        '65x2.0','65x2.5','65x3.0','65x3.5','65x4.0','65x5.0',
        '75x2.0','75x2.5','75x3.0','75x3.5','75x4.0','75x5.0',
        '89x2.0','89x2.5','89x3.0','89x3.5','89x5.0','89x6.0',
        '100x2.0','100x2.5','100x3.0','100x4.0','100x5.0','100x6.0',
        '125x3.0','125x4.0','125x5.0','125x6.0',
        '150x4.0','150x5.0','150x6.0','150x8.0','150x9.0',
        '200x5.0','200x6.0','200x8.0','200x9.0','200x10.0',
        '250x6.0','250x8.0','250x9.0','250x10.0',
        '300x8.0','300x9.0','300x10.0','300x12.0',
        '350x8.0','350x10.0','350x12.0',
        '400x10.0','400x12.0','400x16.0'
    ],
    RHS: [
        '50x25x1.6','50x25x2.0','50x25x2.5',
        '65x35x2.0','65x35x2.5','65x35x3.0',
        '75x25x1.6','75x25x2.0','75x25x2.5',
        '75x50x1.6','75x50x2.0','75x50x2.5','75x50x3.0',
        '100x50x2.0','100x50x2.5','100x50x3.0','100x50x4.0',
        '125x75x2.0','125x75x2.5','125x75x3.0','125x75x4.0','125x75x5.0',
        '150x50x2.0','150x50x2.5','150x50x3.0','150x50x4.0',
        '150x100x3.0','150x100x4.0','150x100x5.0','150x100x6.0',
        '200x100x4.0','200x100x5.0','200x100x6.0',
        '250x150x5.0','250x150x6.0','250x150x8.0',
        '300x200x6.0','300x200x8.0','300x200x10.0',
        '350x250x8.0','350x250x10.0',
        '400x200x8.0','400x200x10.0','400x200x12.0',
        '400x300x8.0','400x300x10.0','400x300x12.0'
    ],
    CHS: [
        '21.3x2.0','26.9x2.0','33.7x2.0','33.7x2.6',
        '42.4x2.0','42.4x2.6',
        '48.3x2.0','48.3x2.6','48.3x3.2',
        '60.3x2.0','60.3x2.6','60.3x3.2',
        '76.1x2.3','76.1x3.2','76.1x3.6',
        '88.9x2.6','88.9x3.2','88.9x4.0','88.9x5.0',
        '101.6x2.6','101.6x3.2','101.6x4.0','101.6x5.0',
        '114.3x3.2','114.3x3.6','114.3x4.0','114.3x5.0','114.3x6.0',
        '139.7x3.5','139.7x4.0','139.7x5.0','139.7x6.0',
        '165.1x3.5','165.1x5.0','165.1x6.0',
        '219.1x4.8','219.1x6.4','219.1x8.0','219.1x10.0',
        '273.1x5.0','273.1x6.4','273.1x8.0','273.1x9.3','273.1x12.7',
        '323.9x5.0','323.9x6.4','323.9x8.0','323.9x9.5','323.9x12.7',
        '355.6x6.4','355.6x8.0','355.6x9.5','355.6x12.7',
        '406.4x6.4','406.4x9.5','406.4x12.7',
        '457.0x6.4','457.0x9.5','457.0x12.7',
        '508.0x6.4','508.0x9.5','508.0x12.7'
    ],
    EA: [
        '25x25x3','25x25x5',
        '30x30x3','30x30x5',
        '40x40x3','40x40x5','40x40x6',
        '45x45x3','45x45x5','45x45x6',
        '50x50x3','50x50x5','50x50x6','50x50x8',
        '55x55x5','55x55x6',
        '60x60x5','60x60x6','60x60x8',
        '65x65x5','65x65x6','65x65x8','65x65x10',
        '75x75x5','75x75x6','75x75x8','75x75x10',
        '90x90x6','90x90x8','90x90x10',
        '100x100x6','100x100x8','100x100x10','100x100x12',
        '125x125x8','125x125x10','125x125x12','125x125x16',
        '150x150x10','150x150x12','150x150x16','150x150x19',
        '200x200x13','200x200x16','200x200x18','200x200x20','200x200x26'
    ],
    UA: [
        '65x50x5','65x50x6','65x50x8',
        '75x50x5','75x50x6','75x50x8',
        '100x75x6','100x75x8','100x75x10',
        '125x75x6','125x75x8','125x75x10','125x75x12',
        '150x90x8','150x90x10','150x90x12','150x90x16',
        '150x100x10','150x100x12'
    ]
};

// Format display name for SHS/RHS/CHS sections
function formatSectionName(sectionType, rawSize) {
    if (sectionType === 'SHS') return rawSize + ' SHS';
    if (sectionType === 'RHS') return rawSize + ' RHS';
    if (sectionType === 'CHS') return rawSize + ' CHS';
    if (sectionType === 'EA') return rawSize + ' EA';
    if (sectionType === 'UA') return rawSize + ' UA';
    return rawSize; // UB, UC, PFC already include type in name
}

const STEEL_GRADES = ['300', '350', '450'];

// ── Wall Type Database ────────────────────────────────────────
const WALL_TYPES = {
    '190 Block':   { thickness: 190 },
    '140 Block':   { thickness: 140 },
    '110 Block':   { thickness: 110 },
    '90 Block':    { thickness: 90 },
    'RC Wall':     { thickness: 200 },
    'Timber Stud': { thickness: 90 },
};

// Initialize schedule types data model
project.scheduleTypes = {
    padfooting: {
        PF1: { width: '', length: '', depth: '', reo: '', setdown: 200, rect: false, color: '#93C5FD' },
        PF2: { width: '', length: '', depth: '', reo: '', setdown: 200, rect: false, color: '#86EFAC' },
        PF3: { width: '', length: '', depth: '', reo: '', setdown: 200, rect: false, color: '#FCD34D' },
        PF4: { width: '', length: '', depth: '', reo: '', setdown: 200, rect: false, color: '#FCA5A5' },
        PF5: { width: '', length: '', depth: '', reo: '', setdown: 200, rect: false, color: '#67E8F9' },
    },
    stripfooting: {
        SF1: { width: '', depth: '', reo: '', setdown: '', top: '', color: '#93C5FD' },
        SF2: { width: '', depth: '', reo: '', setdown: '', top: '', color: '#86EFAC' },
        SF3: { width: '', depth: '', reo: '', setdown: '', top: '', color: '#FCD34D' },
        SF4: { width: '', depth: '', reo: '', setdown: '', top: '', color: '#FCA5A5' },
        SF5: { width: '', depth: '', reo: '', setdown: '', top: '', color: '#67E8F9' },
    },
    beam: {
        SB1: { sectionType: '', size: '', description: '', grade: '300', color: '#93C5FD' },
        SB2: { sectionType: '', size: '', description: '', grade: '300', color: '#86EFAC' },
        SB3: { sectionType: '', size: '', description: '', grade: '300', color: '#FCD34D' },
        SB4: { sectionType: '', size: '', description: '', grade: '300', color: '#FCA5A5' },
        SB5: { sectionType: '', size: '', description: '', grade: '300', color: '#67E8F9' },
    },
    column: {
        SC1: { sectionType: '', size: '', description: '', grade: '300', color: '#93C5FD' },
        SC2: { sectionType: '', size: '', description: '', grade: '300', color: '#86EFAC' },
        SC3: { sectionType: '', size: '', description: '', grade: '300', color: '#FCD34D' },
        SC4: { sectionType: '', size: '', description: '', grade: '300', color: '#FCA5A5' },
        SC5: { sectionType: '', size: '', description: '', grade: '300', color: '#67E8F9' },
    },
    wall: {
        BW1: { wallType: '', thickness: '', description: '', color: '#93C5FD' },
        BW2: { wallType: '', thickness: '', description: '', color: '#86EFAC' },
        BW3: { wallType: '', thickness: '', description: '', color: '#FCD34D' },
        BW4: { wallType: '', thickness: '', description: '', color: '#FCA5A5' },
        BW5: { wallType: '', thickness: '', description: '', color: '#67E8F9' },
    },
    bracingWall: {
        BR1: { bracingType: 'h-A', capacity: 6.4, minLength: 600, description: 'Ply Method A (M12 rods)', color: '#000000' },
        BR2: { bracingType: 'h-B', capacity: 6.0, minLength: 900, description: 'Ply Method B (no rods)', color: '#000000' },
        BR3: { bracingType: 'g', capacity: 3.4, minLength: 900, description: 'Ply 7mm F11 std nailing', color: '#000000' },
        BR4: { bracingType: 'nominal-1', capacity: 0.45, minLength: 450, description: 'Plasterboard 1 side', color: '#000000' },
        BR5: { bracingType: 'nominal-2', capacity: 0.75, minLength: 450, description: 'Plasterboard 2 sides', color: '#000000' },
    }
};

// Wire project info to engine for title block rendering
engine._titleBlockData = project.projectInfo;

// Store globals for debugging / inter-module access
window._app = { engine, project, history, CONFIG };

// ── Status Bar ─────────────────────────────────────────────

const statusScale = document.getElementById('status-scale');
const statusZoom = document.getElementById('status-zoom');
const statusCursor = document.getElementById('status-cursor');
const statusTool = document.getElementById('status-tool');
const statusLevel = document.getElementById('status-level');
const statusElCount = document.getElementById('status-elcount');
const statusSnapType = document.getElementById('status-snap-type');
const statusAutoSave = document.getElementById('status-autosave');

function updateStatusBar(e) {
    if (statusZoom) {
        statusZoom.textContent = (engine.viewport.zoom * 100).toFixed(0) + '%';
    }
    if (statusScale) {
        statusScale.textContent = '1:' + CONFIG.drawingScale;
    }
    if (e && statusCursor) {
        const real = engine.getRealPos(e);
        statusCursor.textContent =
            (real.x / 1000).toFixed(3) + ', ' + (real.y / 1000).toFixed(3) + ' m';
    }
    // Element count on active level
    if (statusElCount) {
        const vis = project.getVisibleElements ? project.getVisibleElements() : project.elements;
        statusElCount.textContent = vis.length;
    }
    // Active level name
    if (statusLevel && typeof levelSystem !== 'undefined' && levelSystem.levels) {
        const active = levelSystem.levels.find(l => l.id === levelSystem.activeLevel);
        statusLevel.textContent = active ? active.name : '—';
    }
    // Active snap type indicator
    if (statusSnapType && typeof snapState !== 'undefined') {
        if (snapState.activeSnap && snapState.activeSnap.type) {
            statusSnapType.textContent = snapState.activeSnap.type.label;
            statusSnapType.style.color = snapState.activeSnap.type.color;
        } else {
            statusSnapType.textContent = '';
        }
    }
}

// Update cursor position on mouse move
container.addEventListener('mousemove', (e) => {
    updateStatusBar(e);
});

// ── Render hook: draw elements ─────────────────────────────

engine.onRender((ctx, eng) => {
    // Future phases will render elements here
    // For now, just update the status bar zoom
    updateStatusBar();
});

// ── Keyboard Shortcuts ─────────────────────────────────────

window.addEventListener('keydown', (e) => {
    // Ctrl+Z / Cmd+Z = Undo
    if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
        e.preventDefault();
        if (history.undo()) engine.requestRender();
    }
    // Ctrl+Shift+Z or Ctrl+Y = Redo
    if ((e.ctrlKey || e.metaKey) && (e.key === 'Z' || e.key === 'y')) {
        e.preventDefault();
        if (history.redo()) engine.requestRender();
    }
    // F = Fit to view
    if (e.key === 'f' && !e.ctrlKey && !e.metaKey && document.activeElement === document.body) {
        engine.fitToView();
    }
});

// ── Undo/Redo/Paste Buttons ──────────────────────────────

const undoBtn = document.getElementById('btn-undo');
const redoBtn = document.getElementById('btn-redo');
const pasteBtn = document.getElementById('btn-paste');

function updateUndoRedoButtons() {
    if (undoBtn) {
        undoBtn.disabled = !history.canUndo;
        undoBtn.style.opacity = history.canUndo ? '1' : '0.4';
    }
    if (redoBtn) {
        redoBtn.disabled = !history.canRedo;
        redoBtn.style.opacity = history.canRedo ? '1' : '0.4';
    }
}

function updatePasteButton() {
    const hasClip = clipboard && (Array.isArray(clipboard) ? clipboard.length > 0 : true);
    if (pasteBtn) {
        pasteBtn.disabled = !hasClip;
        pasteBtn.style.opacity = hasClip ? '1' : '0.4';
    }
}

history.onChange(updateUndoRedoButtons);

if (undoBtn) {
    undoBtn.addEventListener('click', () => {
        if (history.undo()) engine.requestRender();
    });
}
if (redoBtn) {
    redoBtn.addEventListener('click', () => {
        if (history.redo()) engine.requestRender();
    });
}
if (pasteBtn) {
    pasteBtn.addEventListener('click', () => {
        // Trigger the same paste logic as Ctrl+V
        const evt = new KeyboardEvent('keydown', { key: 'v', ctrlKey: true, bubbles: true });
        window.dispatchEvent(evt);
    });
}

// ── Toolbar: Scale selector ────────────────────────────────

const scaleSelect = document.getElementById('scale-select');
if (scaleSelect) {
    scaleSelect.addEventListener('change', function () {
        const val = parseInt(this.value, 10);
        if (val > 0) {
            CONFIG.drawingScale = val;
            project.drawingScale = val;
            engine.requestRender();
        }
    });
}

// ── Toolbar: Fit to View button ────────────────────────────

const fitBtn = document.getElementById('btn-fit');
if (fitBtn) {
    fitBtn.addEventListener('click', () => engine.fitToView());
}

// ── Toolbar: Zoom In / Out ─────────────────────────────────

document.getElementById('btn-zoom-in')?.addEventListener('click', () => {
    const cx = engine.width / 2;
    const cy = engine.height / 2;
    const oldZoom = engine.viewport.zoom;
    const newZoom = Math.min(20, oldZoom * 1.25);
    engine.viewport.panX = cx - (cx - engine.viewport.panX) * (newZoom / oldZoom);
    engine.viewport.panY = cy - (cy - engine.viewport.panY) * (newZoom / oldZoom);
    engine.viewport.zoom = newZoom;
    engine.requestRender();
});

document.getElementById('btn-zoom-out')?.addEventListener('click', () => {
    const cx = engine.width / 2;
    const cy = engine.height / 2;
    const oldZoom = engine.viewport.zoom;
    const newZoom = Math.max(0.05, oldZoom / 1.25);
    engine.viewport.panX = cx - (cx - engine.viewport.panX) * (newZoom / oldZoom);
    engine.viewport.panY = cy - (cy - engine.viewport.panY) * (newZoom / oldZoom);
    engine.viewport.zoom = newZoom;
    engine.requestRender();
});

// ── Toolbar: Grid toggle ─────────────────────────────────

const gridBtn = document.getElementById('btn-grid');
function updateGridBtn() {
    if (gridBtn) {
        if (CONFIG.gridVisible) {
            gridBtn.classList.add('active');
        } else {
            gridBtn.classList.remove('active');
        }
    }
}
updateGridBtn();

if (gridBtn) {
    gridBtn.addEventListener('click', () => {
        CONFIG.gridVisible = !CONFIG.gridVisible;
        updateGridBtn();
        engine.requestRender();
    });
}

// Add G key shortcut for grid toggle
window.addEventListener('keydown', (e) => {
    if (e.key === 'g' && !e.ctrlKey && !e.metaKey && document.activeElement === document.body) {
        CONFIG.gridVisible = !CONFIG.gridVisible;
        updateGridBtn();
        engine.requestRender();
    }
});

// ══════════════════════════════════════════════════════════
