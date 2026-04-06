// ============================================================
// BUNDLED APPLICATION — All modules concatenated inline
// ============================================================

// ── coordinate.js ──────────────────────────────────────────

class CoordinateSystem {
    /**
     * @param {object} config - Sheet configuration (SHEET_WIDTH_MM, margins, etc.)
     * @param {object} viewport - Viewport state (panX, panY, zoom)
     */
    constructor(config, viewport) {
        this.config = config;
        this.viewport = viewport;
    }

    /** Drawing area bounds in sheet-mm */
    get drawArea() {
        const c = this.config;
        return {
            left: c.MARGIN_LEFT,
            top: c.MARGIN_TOP,
            right: c.SHEET_WIDTH_MM - c.MARGIN_RIGHT,
            bottom: c.SHEET_HEIGHT_MM - c.MARGIN_BOTTOM - c.TITLE_BLOCK_HEIGHT_MM,
            get width() { return this.right - this.left; },
            get height() { return this.bottom - this.top; }
        };
    }

    /** Sheet-mm → screen-px */
    sheetToScreen(sx, sy) {
        const v = this.viewport;
        return {
            x: sx * v.zoom + v.panX,
            y: sy * v.zoom + v.panY
        };
    }

    /** Screen-px → sheet-mm */
    screenToSheet(px, py) {
        const v = this.viewport;
        return {
            x: (px - v.panX) / v.zoom,
            y: (py - v.panY) / v.zoom
        };
    }

    /** Screen-px → real-world mm (using drawing scale) */
    screenToReal(px, py) {
        const sheet = this.screenToSheet(px, py);
        return this.sheetToReal(sheet.x, sheet.y);
    }

    /** Real-world mm → screen-px */
    realToScreen(rx, ry) {
        const sheet = this.realToSheet(rx, ry);
        return this.sheetToScreen(sheet.x, sheet.y);
    }

    /** Sheet-mm → real-world mm */
    sheetToReal(sx, sy) {
        const da = this.drawArea;
        const scale = this.config.drawingScale;
        return {
            x: (sx - da.left) * scale,
            y: (sy - da.top) * scale
        };
    }

    /** Real-world mm → sheet-mm */
    realToSheet(rx, ry) {
        const da = this.drawArea;
        const scale = this.config.drawingScale;
        return {
            x: da.left + rx / scale,
            y: da.top + ry / scale
        };
    }

    /** Convert a length in sheet-mm to screen-px */
    sheetLengthToScreen(mm) {
        return mm * this.viewport.zoom;
    }

    /** Convert a length in screen-px to sheet-mm */
    screenLengthToSheet(px) {
        return px / this.viewport.zoom;
    }

    /** Check if a sheet-mm point is within the drawing area */
    isInDrawArea(sx, sy) {
        const da = this.drawArea;
        return sx >= da.left && sx <= da.right && sy >= da.top && sy <= da.bottom;
    }
}

// ── history.js ─────────────────────────────────────────────

class History {
    constructor(maxSize = 100) {
        this.undoStack = [];
        this.redoStack = [];
        this.maxSize = maxSize;
        this._listeners = [];
    }

    /** Register a callback for when history changes */
    onChange(fn) {
        this._listeners.push(fn);
    }

    _notify() {
        for (const fn of this._listeners) fn();
    }

    /**
     * Execute a command and push it onto the undo stack.
     * @param {object} cmd - { execute(), undo(), description? }
     */
    execute(cmd) {
        cmd.execute();
        this.undoStack.push(cmd);
        if (this.undoStack.length > this.maxSize) {
            this.undoStack.shift();
        }
        // Clear redo stack on new action
        this.redoStack.length = 0;
        this._notify();
    }

    /** Undo the last command */
    undo() {
        if (this.undoStack.length === 0) return false;
        const cmd = this.undoStack.pop();
        cmd.undo();
        this.redoStack.push(cmd);
        this._notify();
        return true;
    }

    /** Redo the last undone command */
    redo() {
        if (this.redoStack.length === 0) return false;
        const cmd = this.redoStack.pop();
        cmd.execute();
        this.undoStack.push(cmd);
        this._notify();
        return true;
    }

    get canUndo() { return this.undoStack.length > 0; }
    get canRedo() { return this.redoStack.length > 0; }

    /** Clear all history */
    clear() {
        this.undoStack.length = 0;
        this.redoStack.length = 0;
        this._notify();
    }
}

/**
 * Helper: create a command that adds/removes an element from a collection.
 */
function addElementCmd(collection, element) {
    return {
        description: `Add ${element.type || 'element'}`,
        execute() { collection.push(element); },
        undo() {
            const idx = collection.indexOf(element);
            if (idx !== -1) collection.splice(idx, 1);
        }
    };
}

function removeElementCmd(collection, element) {
    let index = -1;
    return {
        description: `Remove ${element.type || 'element'}`,
        execute() {
            index = collection.indexOf(element);
            if (index !== -1) collection.splice(index, 1);
        },
        undo() {
            if (index !== -1) collection.splice(index, 0, element);
        }
    };
}

/**
 * Helper: batch multiple commands into one undo step.
 */
function batchCmd(commands, description = 'Batch') {
    return {
        description,
        execute() { commands.forEach(c => c.execute()); },
        undo() {
            // Undo in reverse order
            for (let i = commands.length - 1; i >= 0; i--) {
                commands[i].undo();
            }
        }
    };
}

// ── elements.js ────────────────────────────────────────────

let _nextId = 1;

/** Generate a unique element ID */
function generateId() {
    return _nextId++;
}

/** Reset ID counter (e.g., on project load) */
function resetIdCounter(startFrom = 1) {
    _nextId = startFrom;
}

/**
 * Layer definitions — pre-defined structural layers matching Revit conventions.
 * Each layer has a name, default colour, line weight (mm), line pattern, and visibility.
 */
const LAYERS = {
    'S-GRID': {
        name: 'Grids',
        color: '#808080',
        lineWeight: 0.18,
        pattern: 'solid',
        visible: true,
        locked: false,
        printColor: '#808080'
    },
    'S-COLS': {
        name: 'Columns',
        color: '#000000',
        lineWeight: 0.50,
        pattern: 'solid',
        visible: true,
        locked: false,
        printColor: '#000000'
    },
    'S-BEAM': {
        name: 'Beams',
        color: '#000000',
        lineWeight: 0.35,
        pattern: 'solid',
        visible: true,
        locked: false,
        printColor: '#000000'
    },
    'S-WALL': {
        name: 'Walls',
        color: '#000000',
        lineWeight: 0.50,
        pattern: 'solid',
        visible: true,
        locked: false,
        printColor: '#000000'
    },
    'S-SLAB': {
        name: 'Slabs',
        color: '#000000',
        lineWeight: 0.18,
        pattern: 'solid',
        visible: true,
        locked: false,
        printColor: '#000000'
    },
    'S-FTNG': {
        name: 'Footings',
        color: '#000000',
        lineWeight: 0.35,
        pattern: 'hidden',
        visible: true,
        locked: false,
        printColor: '#000000'
    },
    'S-ANNO': {
        name: 'Annotations',
        color: '#000000',
        lineWeight: 0.25,
        pattern: 'solid',
        visible: true,
        locked: false,
        printColor: '#000000'
    },
    'S-DIMS': {
        name: 'Dimensions',
        color: '#000000',
        lineWeight: 0.18,
        pattern: 'solid',
        visible: true,
        locked: false,
        printColor: '#000000'
    },
    'S-BRAC': {
        name: 'Bracing',
        color: '#000000',
        lineWeight: 0.35,
        pattern: 'solid',
        visible: true,
        locked: false,
        printColor: '#000000'
    },
    'S-RIDGE': {
        name: 'Ridge',
        color: '#dc2626',
        lineWeight: 0.5,
        pattern: 'solid',
        visible: true,
        locked: false,
        printColor: '#dc2626'
    },
    'S-ENVL': {
        name: 'Envelope',
        color: '#6366f1',
        lineWeight: 0.25,
        pattern: 'dashed',
        visible: true,
        locked: false,
        printColor: '#6366f1'
    }
};

/**
 * Dash patterns in sheet-mm for each pattern type.
 */
const DASH_PATTERNS = {
    'solid': [],
    'dashed': [3, 1.5],
    'hidden': [1.5, 0.75],
    'centerline': [6, 1.5, 1.5, 1.5],
    'dotted': [0.3, 1.0]
};

/**
 * The project data store. All elements live here.
 */
class ProjectData {
    constructor() {
        this.elements = [];
        this.layers = structuredClone(LAYERS);
        this.activeLayerId = 'S-BEAM';
        this.drawingScale = 100;
        this.projectInfo = {
            projectName: 'PROJECT NAME',
            address: 'SITE ADDRESS',
            drawingTitle: 'STRUCTURAL PLAN',
            drawingNumber: 'S-001',
            projectNumber: '',
            revision: 'A',
            scale: '1:100',
            status: 'FOR CONSTRUCTION',
            drawnBy: '',
            designedBy: '',
            checkedBy: '',
            approvedBy: '',
            date: new Date().toLocaleDateString('en-AU', { day:'2-digit', month:'2-digit', year:'numeric' }),
            revDesc: '',
            company: '',
            companySubtitle: 'STRUCTURAL ENGINEERS'
        };
    }

    /** Get all elements on a given layer */
    getElementsByLayer(layerId) {
        return this.elements.filter(el => el.layer === layerId);
    }

    /** Get all visible elements */
    getVisibleElements() {
        return this.elements.filter(el => {
            const layer = this.layers[el.layer];
            return layer && layer.visible;
        });
    }

    /** Serialise for save */
    toJSON() {
        return {
            version: 1,
            elements: this.elements,
            layers: this.layers,
            activeLayerId: this.activeLayerId,
            drawingScale: this.drawingScale,
            projectInfo: this.projectInfo,
            nextId: _nextId
        };
    }

    /** Deserialise from save */
    static fromJSON(data) {
        const proj = new ProjectData();
        proj.elements = data.elements || [];
        if (data.layers) proj.layers = data.layers;
        if (data.activeLayerId) proj.activeLayerId = data.activeLayerId;
        if (data.drawingScale) proj.drawingScale = data.drawingScale;
        if (data.projectInfo) proj.projectInfo = { ...proj.projectInfo, ...data.projectInfo };
        if (data.nextId) resetIdCounter(data.nextId);
        return proj;
    }
}
