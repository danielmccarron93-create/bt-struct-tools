// ============================================================
// 23-bracing-design.js — AS 1684 Bracing Wall Design Module
// Phase A: Bracing wall tool, rendering, calculation engine,
//          and summary panel
// ============================================================

// ── Bracing Wall Tool State ──────────────────────────────────

const bracingWallToolState = {
    placing: false,
    startPoint: null,   // { x, y } in sheet-mm
    currentEnd: null,   // { x, y } in sheet-mm (preview)
};

let _bracingWallNum = 1;

// ── AS 1684 Bracing Capacity Engine ──────────────────────────

/**
 * Table 8.18 — Structural bracing types and base capacities (kN/m).
 * Types a–n as per AS 1684.2-2010 Clause 8.3.6.4.
 * Note: type f ("designed to engineering principles") has no fixed capacity.
 */
const BRACING_TYPES = {
    'a':    { capacity: 0.8,  desc: 'Diagonal brace — tension only',          group: 'diagonal' },
    'b':    { capacity: 1.1,  desc: 'Diagonal brace — tension + compression', group: 'diagonal' },
    'c':    { capacity: 1.5,  desc: 'K-brace',                                group: 'diagonal' },
    'd':    { capacity: 3.0,  desc: 'Tensioned strap with stud straps',       group: 'diagonal' },
    'e':    { capacity: 3.4,  desc: 'Metal clad — sheet fixed to each stud',  group: 'sheet' },
    'f':    { capacity: null, desc: 'Designed to engineering principles',      group: 'engineered' },
    'g':    { capacity: 3.4,  desc: 'Plywood 7mm F11 std nailing',            group: 'plywood', jointSensitive: true },
    'h-A':  { capacity: 6.4,  desc: 'Plywood Method A (M12 rods)',            group: 'plywood', jointSensitive: true },
    'h-B':  { capacity: 6.0,  desc: 'Plywood Method B (no rods)',             group: 'plywood', jointSensitive: true },
    'i':    { capacity: 3.4,  desc: 'Hardboard 4.8mm Type A',                 group: 'hardboard', jointSensitive: true },
    'j':    { capacity: 4.0,  desc: 'Hardboard 4.8mm Type B',                 group: 'hardboard', jointSensitive: true },
    'k':    { capacity: 5.0,  desc: 'Fibre cement 4.5mm',                     group: 'fibrecement', jointSensitive: true },
    'l':    { capacity: 6.0,  desc: 'Hardboard 5.5mm Type A',                 group: 'hardboard', jointSensitive: true },
    'm':    { capacity: 7.5,  desc: 'Hardboard 5.5mm Type B',                 group: 'hardboard', jointSensitive: true },
    'n':    { capacity: 9.0,  desc: 'Hardboard 5.5mm Type C',                 group: 'hardboard', jointSensitive: true },
    'nominal-1': { capacity: 0.45, desc: 'Plasterboard 1 side (nominal)',     group: 'nominal' },
    'nominal-2': { capacity: 0.75, desc: 'Plasterboard 2 sides (nominal)',    group: 'nominal' },
};

/**
 * Table 8.19 — Height modification factors.
 * Applies to walls with ceiling height > 2700mm.
 * Discrete multipliers (no interpolation per standard).
 */
const HEIGHT_FACTORS = {
    2400: 1.0,
    2700: 1.0,
    3000: 0.9,
    3300: 0.8,
    3600: 0.75,
    3900: 0.7,
    4200: 0.64,
};

/**
 * Get height modification factor for a given wall height (mm).
 * Clause 8.3.6.4 — uses discrete values from Table 8.19.
 * For heights between table values, uses the lower factor (conservative).
 */
function getHeightFactor(heightMm) {
    if (heightMm <= 2700) return 1.0;
    const heights = Object.keys(HEIGHT_FACTORS).map(Number).sort((a, b) => a - b);
    for (let i = heights.length - 1; i >= 0; i--) {
        if (heightMm >= heights[i]) return HEIGHT_FACTORS[heights[i]];
    }
    return 1.0;
}

/**
 * Joint group reduction — Clause 8.3.6.4.
 * JD5 timbers reduce capacity of certain bracing types.
 *   types g–k: multiply by 0.875 (−12.5%)
 *   types l–n: multiply by 0.84  (−16%)
 */
function getJointGroupFactor(bracingType, jointGroup) {
    if (jointGroup !== 'JD5') return 1.0;
    const jd5_mild = ['g', 'h-A', 'h-B', 'i', 'j', 'k'];
    const jd5_heavy = ['l', 'm', 'n'];
    if (jd5_mild.includes(bracingType)) return 0.875;
    if (jd5_heavy.includes(bracingType)) return 0.84;
    return 1.0;
}

/**
 * Plywood length rules — Clause 8.3.6.5.
 * Type g: min 900mm full capacity, 600mm = half capacity, linear between.
 * Type h Method A: min 600mm.
 * Type h Method B: min 900mm.
 */
function getPlywoodLengthFactor(bracingType, lengthMm) {
    if (bracingType === 'g') {
        if (lengthMm >= 900) return 1.0;
        if (lengthMm >= 600) return 0.5 + 0.5 * (lengthMm - 600) / 300;
        return 0; // below min
    }
    if (bracingType === 'h-A') {
        return lengthMm >= 600 ? 1.0 : 0;
    }
    if (bracingType === 'h-B') {
        return lengthMm >= 900 ? 1.0 : 0;
    }
    return 1.0;
}

/**
 * Nominal bracing 50% capacity cap — Clause 8.3.6.3.
 * Total nominal bracing capacity cannot exceed 50% of total
 * racking force demand in any direction.
 */

/**
 * Calculate effective capacity of a single bracing wall element.
 * Returns kN (total for this wall length).
 *
 * @param {object} el - The bracing wall element
 * @param {object} typeData - Schedule type data (from project.scheduleTypes.bracingWall)
 * @param {number} ceilingHeight - Ceiling height in mm (default 2700)
 * @param {string} jointGroup - Timber joint group (default 'JD6')
 * @returns {number} Effective capacity in kN
 */
function calcBracingCapacity(el, typeData, ceilingHeight = 2700, jointGroup = 'JD6') {
    const bracingType = typeData.bracingType || 'g';
    const baseCapacity = BRACING_TYPES[bracingType]?.capacity;
    if (baseCapacity === null || baseCapacity === undefined) return 0;

    // Wall length in metres
    const dx = el.x2 - el.x1;
    const dy = el.y2 - el.y1;
    const lengthMm = Math.sqrt(dx * dx + dy * dy);
    const lengthM = lengthMm / 1000;

    // Minimum length check
    const minLen = typeData.minLength || 450;
    if (lengthMm < minLen) return 0;

    // Capacity per metre
    let capPerM = baseCapacity;

    // Height factor (Table 8.19)
    capPerM *= getHeightFactor(ceilingHeight);

    // Joint group reduction
    capPerM *= getJointGroupFactor(bracingType, jointGroup);

    // Plywood length factor (Clause 8.3.6.5)
    const plyFactor = getPlywoodLengthFactor(bracingType, lengthMm);
    if (plyFactor === 0) return 0;
    capPerM *= plyFactor;

    return capPerM * lengthM;
}

/**
 * Determine the direction a bracing wall resists.
 * Walls within 45° of X-axis → resist X-direction (N-S) racking.
 * Walls within 45° of Y-axis → resist Y-direction (E-W) racking.
 * Returns 'X', 'Y', or 'BOTH' (for 45° diagonal — unusual).
 */
function getBracingDirection(el) {
    const dx = Math.abs(el.x2 - el.x1);
    const dy = Math.abs(el.y2 - el.y1);
    if (dx < 0.1 && dy < 0.1) return 'X'; // degenerate
    const angle = Math.atan2(dy, dx) * (180 / Math.PI); // 0° = horizontal
    if (angle <= 22.5 || angle >= 157.5) return 'Y'; // ~horizontal wall → resists Y-dir force
    if (angle >= 67.5 && angle <= 112.5) return 'X'; // ~vertical wall → resists X-dir force
    // Diagonal walls — conservative: assign to weaker direction at summary level
    return angle < 45 ? 'Y' : 'X';
}


// ── Bracing Summary Calculation ──────────────────────────────

/**
 * Calculate bracing summary for all walls on the current level.
 * Returns { xCapacity, yCapacity, xNominal, yNominal,
 *           xDemand, yDemand, walls: [...] }
 */
function calcBracingSummary() {
    const currentLevel = typeof levelSystem !== 'undefined' ? levelSystem.currentLevel : 'GF';
    const ceilingHeight = typeof bracingSettings !== 'undefined' ? bracingSettings.ceilingHeight : 2700;
    const jointGroup = typeof bracingSettings !== 'undefined' ? bracingSettings.jointGroup : 'JD6';

    const walls = project.elements.filter(el =>
        el.type === 'bracingWall' &&
        (!el.level || el.level === currentLevel)
    );

    let xCapacity = 0, yCapacity = 0;
    let xNominal = 0, yNominal = 0;
    const wallResults = [];

    for (const el of walls) {
        const typeRef = el.typeRef || 'BR1';
        const typeData = project.scheduleTypes.bracingWall?.[typeRef] || {};
        const bracingType = typeData.bracingType || 'g';
        const isNominal = bracingType.startsWith('nominal');
        const capacity = calcBracingCapacity(el, typeData, ceilingHeight, jointGroup);
        const dir = getBracingDirection(el);

        const result = {
            el,
            typeRef,
            bracingType,
            capacity,
            direction: dir,
            isNominal,
            lengthMm: Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2)),
        };
        wallResults.push(result);

        if (dir === 'X' || dir === 'BOTH') {
            if (isNominal) xNominal += capacity; else xCapacity += capacity;
        }
        if (dir === 'Y' || dir === 'BOTH') {
            if (isNominal) yNominal += capacity; else yCapacity += capacity;
        }
    }

    // Demand — auto or manual
    let xDemand = 0, yDemand = 0;
    let windDemandResult = null;

    if (typeof bracingSettings !== 'undefined' && bracingSettings.demandMode === 'auto') {
        const envelope = findEnvelopeElement();
        if (envelope) {
            const bbox = envelopeBBox(envelope);
            if (bbox) {
                windDemandResult = calcWindDemand(bbox, bracingSettings);
                xDemand = windDemandResult.xDemand;
                yDemand = windDemandResult.yDemand;
                // Store for display
                bracingSettings.autoXDemand = xDemand;
                bracingSettings.autoYDemand = yDemand;
                bracingSettings.autoXBreakdown = windDemandResult.xBreakdown;
                bracingSettings.autoYBreakdown = windDemandResult.yBreakdown;
            }
        }
        // Allow manual override even in auto mode
        if (bracingSettings.xDemand > 0) xDemand = bracingSettings.xDemand;
        if (bracingSettings.yDemand > 0) yDemand = bracingSettings.yDemand;
    } else {
        // Manual mode (Phase A)
        xDemand = typeof bracingSettings !== 'undefined' ? (bracingSettings.xDemand || 0) : 0;
        yDemand = typeof bracingSettings !== 'undefined' ? (bracingSettings.yDemand || 0) : 0;
    }

    // Nominal 50% cap (Clause 8.3.6.3)
    const xNominalCapped = Math.min(xNominal, xDemand * 0.5);
    const yNominalCapped = Math.min(yNominal, yDemand * 0.5);

    return {
        xCapacityStructural: xCapacity,
        yCapacityStructural: yCapacity,
        xNominalRaw: xNominal,
        yNominalRaw: yNominal,
        xNominalCapped,
        yNominalCapped,
        xCapacityTotal: xCapacity + xNominalCapped,
        yCapacityTotal: yCapacity + yNominalCapped,
        xDemand,
        yDemand,
        xRatio: xDemand > 0 ? (xCapacity + xNominalCapped) / xDemand : 0,
        yRatio: yDemand > 0 ? (yCapacity + yNominalCapped) / yDemand : 0,
        walls: wallResults,
        ceilingHeight,
        jointGroup,
        windDemandResult,
    };
}


// ── Bracing Settings (persisted on project) ──────────────────

// Default bracing settings — attached to project on first use
if (typeof bracingSettings === 'undefined') {
    var bracingSettings = {
        ceilingHeight: 2700,
        jointGroup: 'JD6',
        windClass: 'N2',
        xDemand: 0,    // kN — user enters manually in Phase A
        yDemand: 0,    // kN
        // Phase B additions
        demandMode: 'manual',       // 'manual' | 'auto'
        roofType: 'hip',            // 'hip' | 'gable' | 'skillion'
        roofPitch: 22.5,            // degrees
        storeyPosition: 'single',   // 'single' | 'upper' | 'lower'
        // Auto-calculated (read-only, populated by calcWindDemand)
        autoXDemand: 0,
        autoYDemand: 0,
        autoXBreakdown: '',
        autoYBreakdown: '',
    };
}


// ══════════════════════════════════════════════════════════════
// PHASE B: Wind Pressure Engine, Area of Elevation, Auto Demand
// ══════════════════════════════════════════════════════════════

// ── Table 8.1 — Flat Vertical Surfaces (all wind classes) ───

const WIND_PRESSURE_TABLE_8_1 = { N1: 0.67, N2: 0.92, N3: 1.40, N4: 2.10 };

// ── Tables 8.2–8.5 — Pressures by width & pitch ────────────
// Structure: { windClass: { width: { pitch: pressure } } }
// Widths in m: 4–16, Pitches in degrees: 0, 5, 10, 15, 20, 25, 30, 35
// All values from AS 1684.2-2010 Tables 8.2–8.5 (exact)

// Tables 8.2–8.5: Exact data from AS 1684.2-2010
// Widths 4–16m, Pitches 0–35°, Wind classes N1–N4

const WIND_TABLE_8_2 = {
    N1: {
        4: {0:0.61,5:0.53,10:0.48,15:0.44,20:0.44,25:0.52,30:0.56,35:0.55}, 5: {0:0.61,5:0.52,10:0.46,15:0.41,20:0.42,25:0.50,30:0.54,35:0.53},
        6: {0:0.61,5:0.50,10:0.44,15:0.39,20:0.42,25:0.50,30:0.53,35:0.54}, 7: {0:0.61,5:0.49,10:0.42,15:0.38,20:0.43,25:0.51,30:0.53,35:0.54},
        8: {0:0.61,5:0.47,10:0.40,15:0.37,20:0.43,25:0.51,30:0.52,35:0.54}, 9: {0:0.61,5:0.46,10:0.39,15:0.36,20:0.44,25:0.52,30:0.51,35:0.54},
        10:{0:0.61,5:0.45,10:0.38,15:0.35,20:0.44,25:0.52,30:0.51,35:0.54}, 11:{0:0.61,5:0.44,10:0.36,15:0.35,20:0.45,25:0.52,30:0.51,35:0.55},
        12:{0:0.61,5:0.42,10:0.34,15:0.35,20:0.45,25:0.52,30:0.51,35:0.55}, 13:{0:0.61,5:0.41,10:0.33,15:0.36,20:0.46,25:0.52,30:0.52,35:0.55},
        14:{0:0.61,5:0.40,10:0.31,15:0.36,20:0.46,25:0.53,30:0.52,35:0.56}, 15:{0:0.61,5:0.39,10:0.30,15:0.36,20:0.47,25:0.53,30:0.52,35:0.56},
        16:{0:0.61,5:0.39,10:0.29,15:0.37,20:0.47,25:0.53,30:0.52,35:0.56}
    },
    N2: {
        4: {0:0.84,5:0.74,10:0.67,15:0.61,20:0.61,25:0.72,30:0.77,35:0.76}, 5: {0:0.84,5:0.71,10:0.64,15:0.57,20:0.58,25:0.69,30:0.75,35:0.74},
        6: {0:0.84,5:0.69,10:0.61,15:0.55,20:0.59,25:0.70,30:0.74,35:0.74}, 7: {0:0.84,5:0.67,10:0.58,15:0.53,20:0.59,25:0.70,30:0.73,35:0.74},
        8: {0:0.84,5:0.65,10:0.56,15:0.51,20:0.60,25:0.71,30:0.72,35:0.75}, 9: {0:0.84,5:0.64,10:0.54,15:0.49,20:0.61,25:0.71,30:0.71,35:0.75},
        10:{0:0.84,5:0.62,10:0.52,15:0.48,20:0.61,25:0.72,30:0.70,35:0.75}, 11:{0:0.84,5:0.60,10:0.50,15:0.48,20:0.62,25:0.72,30:0.71,35:0.75},
        12:{0:0.84,5:0.59,10:0.47,15:0.49,20:0.63,25:0.72,30:0.71,35:0.76}, 13:{0:0.84,5:0.57,10:0.45,15:0.49,20:0.63,25:0.73,30:0.71,35:0.77},
        14:{0:0.84,5:0.56,10:0.43,15:0.50,20:0.64,25:0.73,30:0.72,35:0.77}, 15:{0:0.84,5:0.55,10:0.42,15:0.50,20:0.65,25:0.73,30:0.72,35:0.77},
        16:{0:0.84,5:0.53,10:0.40,15:0.51,20:0.65,25:0.73,30:0.72,35:0.78}
    },
    N3: {
        4: {0:1.3,5:1.2,10:1.0,15:0.95,20:0.96,25:1.1,30:1.2,35:1.2}, 5: {0:1.3,5:1.1,10:1.00,15:0.89,20:0.91,25:1.1,30:1.2,35:1.2},
        6: {0:1.3,5:1.1,10:0.95,15:0.85,20:0.91,25:1.1,30:1.2,35:1.2}, 7: {0:1.3,5:1.1,10:0.91,15:0.82,20:0.93,25:1.1,30:1.1,35:1.2},
        8: {0:1.3,5:1.0,10:0.88,15:0.79,20:0.94,25:1.1,30:1.1,35:1.2}, 9: {0:1.3,5:0.99,10:0.84,15:0.77,20:0.95,25:1.1,30:1.1,35:1.2},
        10:{0:1.3,5:0.97,10:0.81,15:0.75,20:0.95,25:1.1,30:1.1,35:1.2}, 11:{0:1.3,5:0.94,10:0.78,15:0.75,20:0.97,25:1.1,30:1.1,35:1.2},
        12:{0:1.3,5:0.92,10:0.74,15:0.76,20:0.98,25:1.1,30:1.1,35:1.2}, 13:{0:1.3,5:0.90,10:0.71,15:0.77,20:0.99,25:1.1,30:1.1,35:1.2},
        14:{0:1.3,5:0.87,10:0.68,15:0.78,20:1.0,25:1.1,30:1.1,35:1.2}, 15:{0:1.3,5:0.85,10:0.65,15:0.79,20:1.0,25:1.1,30:1.1,35:1.2},
        16:{0:1.3,5:0.83,10:0.62,15:0.79,20:1.0,25:1.1,30:1.1,35:1.2}
    },
    N4: {
        4: {0:2.0,5:1.7,10:1.6,15:1.4,20:1.4,25:1.7,30:1.8,35:1.8}, 5: {0:2.0,5:1.7,10:1.5,15:1.3,20:1.3,25:1.6,30:1.8,35:1.7},
        6: {0:2.0,5:1.6,10:1.4,15:1.3,20:1.4,25:1.6,30:1.7,35:1.7}, 7: {0:2.0,5:1.6,10:1.4,15:1.2,20:1.4,25:1.6,30:1.7,35:1.7},
        8: {0:2.0,5:1.5,10:1.3,15:1.2,20:1.4,25:1.6,30:1.7,35:1.7}, 9: {0:2.0,5:1.5,10:1.3,15:1.1,20:1.4,25:1.7,30:1.7,35:1.7},
        10:{0:2.0,5:1.4,10:1.2,15:1.1,20:1.4,25:1.7,30:1.6,35:1.7}, 11:{0:2.0,5:1.4,10:1.2,15:1.1,20:1.4,25:1.7,30:1.6,35:1.8},
        12:{0:2.0,5:1.4,10:1.1,15:1.1,20:1.5,25:1.7,30:1.7,35:1.8}, 13:{0:2.0,5:1.3,10:1.1,15:1.1,20:1.5,25:1.7,30:1.7,35:1.8},
        14:{0:2.0,5:1.3,10:1.0,15:1.2,20:1.5,25:1.7,30:1.7,35:1.8}, 15:{0:2.0,5:1.3,10:0.97,15:1.2,20:1.5,25:1.7,30:1.7,35:1.8},
        16:{0:2.0,5:1.2,10:0.93,15:1.2,20:1.5,25:1.7,30:1.7,35:1.8}
    }
};

const WIND_TABLE_8_3 = {
    N1: {
        4: {0:0.61,5:0.58,10:0.56,15:0.54,20:0.54,25:0.60,30:0.62,35:0.61}, 5: {0:0.61,5:0.58,10:0.55,15:0.53,20:0.53,25:0.59,30:0.61,35:0.60},
        6: {0:0.61,5:0.57,10:0.54,15:0.52,20:0.52,25:0.59,30:0.60,35:0.59}, 7: {0:0.61,5:0.57,10:0.53,15:0.51,20:0.52,25:0.59,30:0.59,35:0.59},
        8: {0:0.61,5:0.56,10:0.53,15:0.50,20:0.52,25:0.58,30:0.58,35:0.59}, 9: {0:0.61,5:0.55,10:0.52,15:0.49,20:0.52,25:0.58,30:0.58,35:0.59},
        10:{0:0.61,5:0.55,10:0.51,15:0.48,20:0.52,25:0.58,30:0.57,35:0.59}, 11:{0:0.61,5:0.54,10:0.50,15:0.48,20:0.52,25:0.58,30:0.57,35:0.59},
        12:{0:0.61,5:0.54,10:0.49,15:0.48,20:0.52,25:0.58,30:0.57,35:0.59}, 13:{0:0.61,5:0.53,10:0.48,15:0.48,20:0.52,25:0.58,30:0.57,35:0.59},
        14:{0:0.61,5:0.53,10:0.47,15:0.48,20:0.52,25:0.58,30:0.57,35:0.59}, 15:{0:0.61,5:0.52,10:0.46,15:0.48,20:0.53,25:0.58,30:0.57,35:0.59},
        16:{0:0.61,5:0.52,10:0.45,15:0.48,20:0.53,25:0.58,30:0.57,35:0.59}
    },
    N2: {
        4: {0:0.84,5:0.81,10:0.78,15:0.75,20:0.75,25:0.83,30:0.85,35:0.84}, 5: {0:0.84,5:0.80,10:0.77,15:0.73,20:0.73,25:0.82,30:0.84,35:0.83},
        6: {0:0.84,5:0.79,10:0.75,15:0.72,20:0.73,25:0.81,30:0.83,35:0.82}, 7: {0:0.84,5:0.78,10:0.74,15:0.70,20:0.72,25:0.81,30:0.82,35:0.82},
        8: {0:0.84,5:0.78,10:0.73,15:0.69,20:0.72,25:0.81,30:0.81,35:0.82}, 9: {0:0.84,5:0.77,10:0.71,15:0.68,20:0.72,25:0.81,30:0.80,35:0.81},
        10:{0:0.84,5:0.76,10:0.70,15:0.67,20:0.72,25:0.81,30:0.79,35:0.81}, 11:{0:0.84,5:0.75,10:0.69,15:0.66,20:0.72,25:0.80,30:0.79,35:0.81},
        12:{0:0.84,5:0.74,10:0.68,15:0.66,20:0.72,25:0.80,30:0.79,35:0.81}, 13:{0:0.84,5:0.74,10:0.66,15:0.66,20:0.72,25:0.80,30:0.79,35:0.82},
        14:{0:0.84,5:0.73,10:0.65,15:0.66,20:0.73,25:0.80,30:0.79,35:0.82}, 15:{0:0.84,5:0.72,10:0.64,15:0.66,20:0.73,25:0.80,30:0.79,35:0.82},
        16:{0:0.84,5:0.72,10:0.63,15:0.66,20:0.73,25:0.80,30:0.79,35:0.82}
    },
    N3: {
        4: {0:1.3,5:1.3,10:1.2,15:1.2,20:1.2,25:1.3,30:1.3,35:1.3}, 5: {0:1.3,5:1.2,10:1.2,15:1.1,20:1.1,25:1.3,30:1.3,35:1.3},
        6: {0:1.3,5:1.2,10:1.2,15:1.1,20:1.1,25:1.3,30:1.3,35:1.3}, 7: {0:1.3,5:1.2,10:1.2,15:1.1,20:1.1,25:1.3,30:1.3,35:1.3},
        8: {0:1.3,5:1.2,10:1.1,15:1.1,20:1.1,25:1.3,30:1.3,35:1.3}, 9: {0:1.3,5:1.2,10:1.1,15:1.1,20:1.1,25:1.3,30:1.2,35:1.3},
        10:{0:1.3,5:1.2,10:1.1,15:1.0,20:1.1,25:1.3,30:1.2,35:1.3}, 11:{0:1.3,5:1.2,10:1.1,15:1.0,20:1.1,25:1.3,30:1.2,35:1.3},
        12:{0:1.3,5:1.2,10:1.1,15:1.0,20:1.1,25:1.3,30:1.2,35:1.3}, 13:{0:1.3,5:1.2,10:1.0,15:1.0,20:1.1,25:1.3,30:1.2,35:1.3},
        14:{0:1.3,5:1.1,10:1.0,15:1.0,20:1.1,25:1.3,30:1.2,35:1.3}, 15:{0:1.3,5:1.1,10:1.0,15:1.0,20:1.1,25:1.2,30:1.2,35:1.3},
        16:{0:1.3,5:1.1,10:0.98,15:1.0,20:1.1,25:1.2,30:1.2,35:1.3}
    },
    N4: {
        4: {0:2.0,5:1.9,10:1.8,15:1.7,20:1.7,25:1.9,30:2.0,35:2.0}, 5: {0:2.0,5:1.9,10:1.8,15:1.7,20:1.7,25:1.9,30:2.0,35:1.9},
        6: {0:2.0,5:1.8,10:1.8,15:1.7,20:1.7,25:1.9,30:1.9,35:1.9}, 7: {0:2.0,5:1.8,10:1.7,15:1.6,20:1.7,25:1.9,30:1.9,35:1.9},
        8: {0:2.0,5:1.8,10:1.7,15:1.6,20:1.7,25:1.9,30:1.9,35:1.9}, 9: {0:2.0,5:1.8,10:1.7,15:1.6,20:1.7,25:1.9,30:1.9,35:1.9},
        10:{0:2.0,5:1.8,10:1.6,15:1.6,20:1.7,25:1.9,30:1.9,35:1.9}, 11:{0:2.0,5:1.7,10:1.6,15:1.5,20:1.7,25:1.9,30:1.9,35:1.9},
        12:{0:2.0,5:1.7,10:1.6,15:1.5,20:1.7,25:1.9,30:1.9,35:1.9}, 13:{0:2.0,5:1.7,10:1.5,15:1.5,20:1.7,25:1.9,30:1.9,35:1.9},
        14:{0:2.0,5:1.7,10:1.5,15:1.5,20:1.7,25:1.9,30:1.9,35:1.9}, 15:{0:2.0,5:1.7,10:1.5,15:1.5,20:1.7,25:1.9,30:1.9,35:1.9},
        16:{0:2.0,5:1.7,10:1.5,15:1.5,20:1.7,25:1.9,30:1.9,35:1.9}
    }
};

const WIND_TABLE_8_4 = {
    N1: {
        4: {0:0.67,5:0.62,10:0.59,15:0.55,20:0.55,25:0.57,30:0.59,35:0.58}, 5: {0:0.67,5:0.61,10:0.57,15:0.53,20:0.53,25:0.56,30:0.58,35:0.57},
        6: {0:0.67,5:0.60,10:0.56,15:0.52,20:0.53,25:0.56,30:0.57,35:0.57}, 7: {0:0.67,5:0.59,10:0.54,15:0.50,20:0.52,25:0.56,30:0.56,35:0.57},
        8: {0:0.67,5:0.58,10:0.53,15:0.49,20:0.52,25:0.56,30:0.56,35:0.57}, 9: {0:0.67,5:0.57,10:0.51,15:0.48,20:0.52,25:0.56,30:0.55,35:0.57},
        10:{0:0.67,5:0.56,10:0.50,15:0.47,20:0.52,25:0.56,30:0.54,35:0.57}, 11:{0:0.67,5:0.55,10:0.49,15:0.46,20:0.52,25:0.56,30:0.54,35:0.57},
        12:{0:0.67,5:0.55,10:0.47,15:0.46,20:0.52,25:0.56,30:0.54,35:0.57}, 13:{0:0.67,5:0.54,10:0.46,15:0.46,20:0.52,25:0.56,30:0.55,35:0.57},
        14:{0:0.67,5:0.53,10:0.45,15:0.46,20:0.53,25:0.56,30:0.55,35:0.57}, 15:{0:0.67,5:0.52,10:0.44,15:0.46,20:0.53,25:0.56,30:0.55,35:0.58},
        16:{0:0.67,5:0.52,10:0.43,15:0.46,20:0.53,25:0.56,30:0.55,35:0.58}
    },
    N2: {
        4: {0:0.92,5:0.86,10:0.81,15:0.77,20:0.76,25:0.79,30:0.82,35:0.81}, 5: {0:0.92,5:0.84,10:0.79,15:0.74,20:0.73,25:0.77,30:0.81,35:0.79},
        6: {0:0.92,5:0.83,10:0.77,15:0.72,20:0.73,25:0.77,30:0.79,35:0.79}, 7: {0:0.92,5:0.82,10:0.75,15:0.70,20:0.73,25:0.77,30:0.78,35:0.79},
        8: {0:0.92,5:0.80,10:0.73,15:0.68,20:0.72,25:0.77,30:0.77,35:0.79}, 9: {0:0.92,5:0.79,10:0.71,15:0.66,20:0.72,25:0.77,30:0.76,35:0.79},
        10:{0:0.92,5:0.78,10:0.69,15:0.65,20:0.72,25:0.77,30:0.75,35:0.78}, 11:{0:0.92,5:0.77,10:0.68,15:0.64,20:0.72,25:0.77,30:0.75,35:0.79},
        12:{0:0.92,5:0.76,10:0.66,15:0.64,20:0.72,25:0.77,30:0.75,35:0.79}, 13:{0:0.92,5:0.75,10:0.64,15:0.64,20:0.73,25:0.77,30:0.75,35:0.79},
        14:{0:0.92,5:0.73,10:0.62,15:0.64,20:0.73,25:0.77,30:0.76,35:0.79}, 15:{0:0.92,5:0.72,10:0.60,15:0.64,20:0.73,25:0.77,30:0.76,35:0.80},
        16:{0:0.92,5:0.71,10:0.59,15:0.64,20:0.73,25:0.77,30:0.76,35:0.80}
    },
    N3: {
        4: {0:1.4,5:1.3,10:1.3,15:1.2,20:1.2,25:1.2,30:1.3,35:1.3}, 5: {0:1.4,5:1.3,10:1.2,15:1.2,20:1.1,25:1.2,30:1.3,35:1.2},
        6: {0:1.4,5:1.3,10:1.2,15:1.1,20:1.1,25:1.2,30:1.2,35:1.2}, 7: {0:1.4,5:1.3,10:1.2,15:1.1,20:1.1,25:1.2,30:1.2,35:1.2},
        8: {0:1.4,5:1.3,10:1.1,15:1.1,20:1.1,25:1.2,30:1.2,35:1.2}, 9: {0:1.4,5:1.2,10:1.1,15:1.0,20:1.1,25:1.2,30:1.2,35:1.2},
        10:{0:1.4,5:1.2,10:1.1,15:1.0,20:1.1,25:1.2,30:1.2,35:1.2}, 11:{0:1.4,5:1.2,10:1.1,15:1.0,20:1.1,25:1.2,30:1.2,35:1.2},
        12:{0:1.4,5:1.2,10:1.0,15:1.0,20:1.1,25:1.2,30:1.2,35:1.2}, 13:{0:1.4,5:1.2,10:1.0,15:1.0,20:1.1,25:1.2,30:1.2,35:1.2},
        14:{0:1.4,5:1.1,10:0.97,15:1.0,20:1.1,25:1.2,30:1.2,35:1.2}, 15:{0:1.4,5:1.1,10:0.94,15:1.0,20:1.1,25:1.2,30:1.2,35:1.2},
        16:{0:1.4,5:1.1,10:0.92,15:1.0,20:1.1,25:1.2,30:1.2,35:1.2}
    },
    N4: {
        4: {0:2.1,5:2.0,10:1.9,15:1.8,20:1.8,25:1.8,30:1.9,35:1.9}, 5: {0:2.1,5:2.0,10:1.8,15:1.7,20:1.7,25:1.8,30:1.9,35:1.8},
        6: {0:2.1,5:1.9,10:1.8,15:1.7,20:1.7,25:1.8,30:1.8,35:1.8}, 7: {0:2.1,5:1.9,10:1.7,15:1.6,20:1.7,25:1.8,30:1.8,35:1.8},
        8: {0:2.1,5:1.9,10:1.7,15:1.6,20:1.7,25:1.8,30:1.8,35:1.8}, 9: {0:2.1,5:1.8,10:1.7,15:1.5,20:1.7,25:1.8,30:1.8,35:1.8},
        10:{0:2.1,5:1.8,10:1.6,15:1.5,20:1.7,25:1.8,30:1.8,35:1.8}, 11:{0:2.1,5:1.8,10:1.6,15:1.5,20:1.7,25:1.8,30:1.8,35:1.8},
        12:{0:2.1,5:1.8,10:1.5,15:1.5,20:1.7,25:1.8,30:1.8,35:1.8}, 13:{0:2.1,5:1.7,10:1.5,15:1.5,20:1.7,25:1.8,30:1.8,35:1.8},
        14:{0:2.1,5:1.7,10:1.4,15:1.5,20:1.7,25:1.8,30:1.8,35:1.8}, 15:{0:2.1,5:1.7,10:1.4,15:1.5,20:1.7,25:1.8,30:1.8,35:1.9},
        16:{0:2.1,5:1.7,10:1.4,15:1.5,20:1.7,25:1.8,30:1.8,35:1.9}
    }
};

const WIND_TABLE_8_5 = {
    N1: {
        4: {0:0.67,5:0.65,10:0.64,15:0.63,20:0.62,25:0.63,30:0.64,35:0.63}, 5: {0:0.67,5:0.65,10:0.63,15:0.62,20:0.61,25:0.62,30:0.63,35:0.63},
        6: {0:0.67,5:0.64,10:0.63,15:0.61,20:0.61,25:0.62,30:0.63,35:0.62}, 7: {0:0.67,5:0.64,10:0.62,15:0.60,20:0.61,25:0.62,30:0.62,35:0.62},
        8: {0:0.67,5:0.64,10:0.62,15:0.60,20:0.61,25:0.62,30:0.62,35:0.62}, 9: {0:0.67,5:0.63,10:0.61,15:0.59,20:0.60,25:0.61,30:0.61,35:0.62},
        10:{0:0.67,5:0.63,10:0.60,15:0.58,20:0.60,25:0.61,30:0.61,35:0.61}, 11:{0:0.67,5:0.63,10:0.60,15:0.58,20:0.60,25:0.61,30:0.60,35:0.61},
        12:{0:0.67,5:0.62,10:0.59,15:0.58,20:0.60,25:0.61,30:0.60,35:0.61}, 13:{0:0.67,5:0.62,10:0.58,15:0.58,20:0.60,25:0.61,30:0.60,35:0.61},
        14:{0:0.67,5:0.62,10:0.58,15:0.58,20:0.60,25:0.61,30:0.60,35:0.61}, 15:{0:0.67,5:0.61,10:0.57,15:0.57,20:0.60,25:0.61,30:0.60,35:0.61},
        16:{0:0.67,5:0.61,10:0.57,15:0.57,20:0.60,25:0.61,30:0.60,35:0.61}
    },
    N2: {
        4: {0:0.92,5:0.90,10:0.89,15:0.87,20:0.86,25:0.87,30:0.88,35:0.87}, 5: {0:0.92,5:0.90,10:0.88,15:0.85,20:0.85,25:0.86,30:0.87,35:0.87},
        6: {0:0.92,5:0.89,10:0.87,15:0.84,20:0.85,25:0.86,30:0.87,35:0.86}, 7: {0:0.92,5:0.89,10:0.86,15:0.84,20:0.84,25:0.86,30:0.86,35:0.86},
        8: {0:0.92,5:0.88,10:0.85,15:0.83,20:0.84,25:0.85,30:0.85,35:0.86}, 9: {0:0.92,5:0.88,10:0.84,15:0.82,20:0.84,25:0.85,30:0.84,35:0.85},
        10:{0:0.92,5:0.87,10:0.84,15:0.81,20:0.83,25:0.85,30:0.84,35:0.85}, 11:{0:0.92,5:0.87,10:0.83,15:0.80,20:0.83,25:0.85,30:0.84,35:0.85},
        12:{0:0.92,5:0.86,10:0.82,15:0.80,20:0.83,25:0.85,30:0.83,35:0.85}, 13:{0:0.92,5:0.86,10:0.81,15:0.80,20:0.83,25:0.84,30:0.83,35:0.85},
        14:{0:0.92,5:0.85,10:0.80,15:0.80,20:0.83,25:0.84,30:0.83,35:0.85}, 15:{0:0.92,5:0.85,10:0.79,15:0.79,20:0.83,25:0.84,30:0.83,35:0.85},
        16:{0:0.92,5:0.85,10:0.78,15:0.79,20:0.83,25:0.84,30:0.83,35:0.85}
    },
    N3: {
        4: {0:1.4,5:1.4,10:1.4,15:1.4,20:1.3,25:1.4,30:1.4,35:1.4}, 5: {0:1.4,5:1.4,10:1.4,15:1.3,20:1.3,25:1.3,30:1.4,35:1.3},
        6: {0:1.4,5:1.4,10:1.4,15:1.3,20:1.3,25:1.3,30:1.4,35:1.3}, 7: {0:1.4,5:1.4,10:1.3,15:1.3,20:1.3,25:1.3,30:1.3,35:1.3},
        8: {0:1.4,5:1.4,10:1.3,15:1.3,20:1.3,25:1.3,30:1.3,35:1.3}, 9: {0:1.4,5:1.4,10:1.3,15:1.3,20:1.3,25:1.3,30:1.3,35:1.3},
        10:{0:1.4,5:1.4,10:1.3,15:1.3,20:1.3,25:1.3,30:1.3,35:1.3}, 11:{0:1.4,5:1.4,10:1.3,15:1.3,20:1.3,25:1.3,30:1.3,35:1.3},
        12:{0:1.4,5:1.3,10:1.3,15:1.3,20:1.3,25:1.3,30:1.3,35:1.3}, 13:{0:1.4,5:1.3,10:1.3,15:1.2,20:1.3,25:1.3,30:1.3,35:1.3},
        14:{0:1.4,5:1.3,10:1.3,15:1.2,20:1.3,25:1.3,30:1.3,35:1.3}, 15:{0:1.4,5:1.3,10:1.2,15:1.2,20:1.3,25:1.3,30:1.3,35:1.3},
        16:{0:1.4,5:1.3,10:1.2,15:1.2,20:1.3,25:1.3,30:1.3,35:1.3}
    },
    N4: {
        4: {0:2.1,5:2.1,10:2.1,15:2.0,20:2.0,25:2.0,30:2.1,35:2.0}, 5: {0:2.1,5:2.1,10:2.0,15:2.0,20:2.0,25:2.0,30:2.0,35:2.0},
        6: {0:2.1,5:2.1,10:2.0,15:2.0,20:2.0,25:2.0,30:2.0,35:2.0}, 7: {0:2.1,5:2.1,10:2.0,15:1.9,20:2.0,25:2.0,30:2.0,35:2.0},
        8: {0:2.1,5:2.1,10:2.0,15:1.9,20:2.0,25:2.0,30:2.0,35:2.0}, 9: {0:2.1,5:2.0,10:2.0,15:1.9,20:1.9,25:2.0,30:2.0,35:2.0},
        10:{0:2.1,5:2.0,10:1.9,15:1.9,20:1.9,25:2.0,30:2.0,35:2.0}, 11:{0:2.1,5:2.0,10:1.9,15:1.9,20:1.9,25:2.0,30:1.9,35:2.0},
        12:{0:2.1,5:2.0,10:1.9,15:1.9,20:1.9,25:2.0,30:1.9,35:2.0}, 13:{0:2.1,5:2.0,10:1.9,15:1.9,20:1.9,25:2.0,30:1.9,35:2.0},
        14:{0:2.1,5:2.0,10:1.9,15:1.9,20:1.9,25:2.0,30:1.9,35:2.0}, 15:{0:2.1,5:2.0,10:1.8,15:1.8,20:1.9,25:2.0,30:1.9,35:2.0},
        16:{0:2.1,5:2.0,10:1.8,15:1.8,20:1.9,25:2.0,30:1.9,35:2.0}
    }
};


// ── Bilinear Interpolation ──────────────────────────────────

/**
 * Bilinear interpolation in a width × pitch table.
 * AS 1684 Appendix C explicitly permits linear interpolation.
 */
function bilinearInterpolate(table, width, pitch) {
    const widths = Object.keys(table).map(Number).sort((a, b) => a - b);
    const pitches = Object.keys(table[widths[0]]).map(Number).sort((a, b) => a - b);

    // Clamp to table bounds
    const w = Math.max(widths[0], Math.min(widths[widths.length - 1], width));
    const p = Math.max(pitches[0], Math.min(pitches[pitches.length - 1], pitch));

    // Find bounding indices
    let wi0 = 0, wi1 = 0;
    for (let i = 0; i < widths.length - 1; i++) {
        if (w >= widths[i] && w <= widths[i + 1]) { wi0 = i; wi1 = i + 1; break; }
    }
    if (w >= widths[widths.length - 1]) { wi0 = widths.length - 1; wi1 = wi0; }

    let pi0 = 0, pi1 = 0;
    for (let i = 0; i < pitches.length - 1; i++) {
        if (p >= pitches[i] && p <= pitches[i + 1]) { pi0 = i; pi1 = i + 1; break; }
    }
    if (p >= pitches[pitches.length - 1]) { pi0 = pitches.length - 1; pi1 = pi0; }

    const w0 = widths[wi0], w1 = widths[wi1];
    const p0 = pitches[pi0], p1 = pitches[pi1];

    const v00 = table[w0][p0];
    const v01 = table[w0][p1];
    const v10 = table[w1][p0];
    const v11 = table[w1][p1];

    // Interpolation weights
    const tw = (w1 !== w0) ? (w - w0) / (w1 - w0) : 0;
    const tp = (p1 !== p0) ? (p - p0) / (p1 - p0) : 0;

    const v0 = v00 + (v01 - v00) * tp;
    const v1 = v10 + (v11 - v10) * tp;
    return v0 + (v1 - v0) * tw;
}


// ── Wind Table Selection Logic ──────────────────────────────

/**
 * Select the correct wind pressure table per AS 1684 Clause 8.3.2.
 *
 * @param {string} roofType - 'hip' | 'gable' | 'skillion'
 * @param {string} storeyPos - 'single' | 'upper' | 'lower'
 * @param {string} windDir - 'longSide' | 'shortEnd'
 * @returns {{ tableRef: string, useTable81: boolean }}
 */
function selectWindTable(roofType, storeyPos, windDir) {
    // Skillion roofs and gable end faces → always Table 8.1
    if (roofType === 'skillion') {
        return { tableRef: '8.1', useTable81: true };
    }
    // Gable end: wind hitting the flat gable face → Table 8.1
    if (roofType === 'gable' && windDir === 'shortEnd') {
        return { tableRef: '8.1', useTable81: true };
    }
    // Otherwise select from Tables 8.2–8.5
    const isUpper = (storeyPos === 'single' || storeyPos === 'upper');
    if (windDir === 'longSide') {
        return { tableRef: isUpper ? '8.2' : '8.3', useTable81: false };
    } else {
        // shortEnd with hip roof
        return { tableRef: isUpper ? '8.4' : '8.5', useTable81: false };
    }
}

/**
 * Look up wind pressure for given parameters.
 *
 * @returns {{ pressure: number, tableRef: string, interpolated: boolean }}
 */
function lookupWindPressure(windClass, buildingWidthM, roofPitch, roofType, storeyPos, windDir) {
    const sel = selectWindTable(roofType, storeyPos, windDir);

    if (sel.useTable81) {
        return {
            pressure: WIND_PRESSURE_TABLE_8_1[windClass] || 0.92,
            tableRef: '8.1',
            interpolated: false,
        };
    }

    // Get the right table
    const tables = { '8.2': WIND_TABLE_8_2, '8.3': WIND_TABLE_8_3, '8.4': WIND_TABLE_8_4, '8.5': WIND_TABLE_8_5 };
    const table = tables[sel.tableRef]?.[windClass];
    if (!table) {
        // Fallback to Table 8.1 if table data missing
        return { pressure: WIND_PRESSURE_TABLE_8_1[windClass] || 0.92, tableRef: '8.1 (fallback)', interpolated: false };
    }

    const pressure = bilinearInterpolate(table, buildingWidthM, roofPitch);
    return {
        pressure: Math.round(pressure * 1000) / 1000,
        tableRef: sel.tableRef,
        interpolated: true,
    };
}


// ── Area of Elevation Calculator (Clause 8.3.3, Figure 8.2) ─

/**
 * Calculate area of elevation for one wind direction.
 *
 * @param {object} bbox - { minX, maxX, minY, maxY } in real-world mm
 * @param {string} direction - 'X' or 'Y' (which wall face the wind hits)
 * @param {object} settings - { ceilingHeight, roofPitch, roofType, storeyPosition }
 * @returns {{ area_m2, faceLength_m, wallHeight_m, roofRise_m, hBelow_m, breakdown }}
 */
function calcAreaOfElevation(bbox, direction, settings) {
    const ceilH = settings.ceilingHeight || 2700; // mm
    const pitch = settings.roofPitch || 22.5;     // degrees
    const roofType = settings.roofType || 'hip';
    const storeyPos = settings.storeyPosition || 'single';

    // Building dimensions from bounding box
    const lengthX = (bbox.maxX - bbox.minX);  // mm along X
    const lengthY = (bbox.maxY - bbox.minY);  // mm along Y

    // Wind on X-direction → hits the Y-face (wall running along Y-axis)
    // Face length = Y dimension, building width = X dimension
    let faceLength_mm, buildingWidth_mm;
    if (direction === 'X') {
        faceLength_mm = lengthY;
        buildingWidth_mm = lengthX;
    } else {
        faceLength_mm = lengthX;
        buildingWidth_mm = lengthY;
    }

    const faceLength_m = faceLength_mm / 1000;
    const wallHeight_m = ceilH / 1000;
    const buildingWidth_m = buildingWidth_mm / 1000;

    // Roof rise calculation
    const pitchRad = pitch * Math.PI / 180;
    let roofRise_m;
    if (roofType === 'hip') {
        // Hip roof: ridge doesn't extend to the full width
        // For area of elevation perpendicular to long side, use full half-width rise
        // For short end (hip face), the projected triangle is smaller — use 2/3 factor
        roofRise_m = (buildingWidth_m / 2) * Math.tan(pitchRad);
    } else if (roofType === 'gable') {
        roofRise_m = (buildingWidth_m / 2) * Math.tan(pitchRad);
    } else {
        // Skillion — full height difference across the width
        roofRise_m = buildingWidth_m * Math.tan(pitchRad);
    }

    // Height zone below floor level (h = half ceiling height)
    let hBelow_m = 0;
    if (storeyPos === 'upper') {
        hBelow_m = wallHeight_m / 2; // h below upper floor
    } else if (storeyPos === 'lower') {
        hBelow_m = wallHeight_m / 2; // h below lower floor
    }
    // Single storey on slab: hBelow = 0

    // Calculate area per Figure 8.2
    let totalHeight_m;
    if (storeyPos === 'single') {
        // Single storey: ground to ridge
        // Area = faceLength × (wallHeight + roofRise/2 for average)
        // Per Figure 8.2(a): wall height + roof contribution
        totalHeight_m = wallHeight_m + roofRise_m / 2;
    } else if (storeyPos === 'upper') {
        // Upper storey: h below floor + wall height + roof
        totalHeight_m = hBelow_m + wallHeight_m + roofRise_m / 2;
    } else {
        // Lower storey: h below + wall height + h above (no roof)
        const hAbove_m = wallHeight_m / 2;
        totalHeight_m = hBelow_m + wallHeight_m + hAbove_m;
    }

    const area_m2 = faceLength_m * totalHeight_m;

    // Build human-readable breakdown
    let breakdown = `Face: ${faceLength_m.toFixed(1)}m`;
    if (hBelow_m > 0) breakdown += ` | h below: ${hBelow_m.toFixed(2)}m`;
    breakdown += ` | Wall: ${wallHeight_m.toFixed(2)}m`;
    if (storeyPos === 'lower') {
        breakdown += ` | h above: ${(wallHeight_m / 2).toFixed(2)}m`;
    } else {
        breakdown += ` | Roof: ${(roofRise_m / 2).toFixed(2)}m`;
    }
    breakdown += ` | Total H: ${totalHeight_m.toFixed(2)}m`;
    breakdown += ` | Area: ${area_m2.toFixed(1)} m\u00B2`;

    return {
        area_m2: Math.round(area_m2 * 100) / 100,
        faceLength_m: Math.round(faceLength_m * 100) / 100,
        wallHeight_m: Math.round(wallHeight_m * 100) / 100,
        roofRise_m: Math.round(roofRise_m * 100) / 100,
        hBelow_m: Math.round(hBelow_m * 100) / 100,
        totalHeight_m: Math.round(totalHeight_m * 100) / 100,
        buildingWidth_m: Math.round(buildingWidth_m * 100) / 100,
        breakdown,
    };
}


// ── Racking Force Auto-Calculation ──────────────────────────

/**
 * Auto-calculate wind demand (racking force) from building envelope.
 *
 * @param {object} bbox - building bounding box { minX, maxX, minY, maxY } in mm
 * @param {object} settings - bracingSettings object
 * @returns {{ xDemand, yDemand, xBreakdown, yBreakdown, xPressureInfo, yPressureInfo }}
 */
function calcWindDemand(bbox, settings) {
    const windClass = settings.windClass || 'N2';
    const roofPitch = settings.roofPitch || 22.5;
    const roofType = settings.roofType || 'hip';
    const storeyPos = settings.storeyPosition || 'single';

    // Determine long side vs short end from building dimensions
    const dimX = (bbox.maxX - bbox.minX) / 1000; // metres
    const dimY = (bbox.maxY - bbox.minY) / 1000;

    // Determine best 3D geometry source: skeleton > ridgeLine > simplified
    const envelope = findEnvelopeElement();
    const skeleton = typeof findSkeletonElement === 'function' ? findSkeletonElement() : null;
    const ridge    = findRidgeElement();

    const rfLevel = typeof levelSystem !== 'undefined'
        ? levelSystem.levels.find(l => l.id === 'RF' || l.name.toLowerCase().includes('roof'))
        : null;
    const eavesH = rfLevel ? rfLevel.elevation : (settings.ceilingHeight || 2700);

    let xArea_m2, yArea_m2, areaMethod;

    const skeletonReady = !!(skeleton && skeleton.committed &&
        skeleton.faces && skeleton.faces.length > 0 &&
        skeleton.faces.every(f => f.pitch != null) &&
        typeof buildRoofModelFromSkeleton === 'function');

    if (skeletonReady) {
        // Best case: skeleton 3D model with per-face pitches
        const roofModel = buildRoofModelFromSkeleton(skeleton, eavesH);
        if (roofModel && typeof calcExactAreaOfElevation === 'function') {
            const xResult = calcExactAreaOfElevation(roofModel, 0);
            const yResult = calcExactAreaOfElevation(roofModel, Math.PI / 2);
            xArea_m2 = xResult.area_m2;
            yArea_m2 = yResult.area_m2;
            areaMethod = '3D skeleton';
        } else {
            areaMethod = null; // fall through
        }
    }

    if (!areaMethod && envelope && ridge && typeof buildRoofModel === 'function') {
        // Fallback: legacy ridgeLine 3D model
        const roofModel = buildRoofModel(envelope, ridge, eavesH);
        if (roofModel && typeof calcExactAreaOfElevation === 'function') {
            const xResult = calcExactAreaOfElevation(roofModel, 0);
            const yResult = calcExactAreaOfElevation(roofModel, Math.PI / 2);
            xArea_m2 = xResult.area_m2;
            yArea_m2 = yResult.area_m2;
            areaMethod = '3D ridge';
        }
    }

    if (!areaMethod) {
        // Simplified rectangular fallback
        const xArea = calcAreaOfElevation(bbox, 'X', settings);
        const yArea = calcAreaOfElevation(bbox, 'Y', settings);
        xArea_m2 = xArea.area_m2;
        yArea_m2 = yArea.area_m2;
        areaMethod = 'simplified';
    }

    // X-direction: wind acts on the Y-face
    const xIsLongSide = dimY >= dimX;
    const yIsLongSide = dimX >= dimY;

    const xWindDir = xIsLongSide ? 'longSide' : 'shortEnd';
    const xBuildingWidth = dimX;
    const xPressure = lookupWindPressure(windClass, xBuildingWidth, roofPitch, roofType, storeyPos, xWindDir);
    const xDemand = xArea_m2 * xPressure.pressure;

    const yWindDir = yIsLongSide ? 'longSide' : 'shortEnd';
    const yBuildingWidth = dimY;
    const yPressure = lookupWindPressure(windClass, yBuildingWidth, roofPitch, roofType, storeyPos, yWindDir);
    const yDemand = yArea_m2 * yPressure.pressure;

    // Build breakdowns
    const methodTag = areaMethod === 'simplified' ? '' : ` (${areaMethod})`;
    const xBreakdown = `${xArea_m2.toFixed(1)} m\u00B2${methodTag} \u00D7 ${xPressure.pressure.toFixed(3)} kPa (Tbl ${xPressure.tableRef}) = ${xDemand.toFixed(1)} kN`;
    const yBreakdown = `${yArea_m2.toFixed(1)} m\u00B2${methodTag} \u00D7 ${yPressure.pressure.toFixed(3)} kPa (Tbl ${yPressure.tableRef}) = ${yDemand.toFixed(1)} kN`;

    return {
        xDemand: Math.round(xDemand * 10) / 10,
        yDemand: Math.round(yDemand * 10) / 10,
        xBreakdown,
        yBreakdown,
        xArea_m2,
        yArea_m2,
        xPressureInfo: xPressure,
        yPressureInfo: yPressure,
        areaMethod,
    };
}

/**
 * Find the building envelope element on the current level.
 */
function findEnvelopeElement() {
    const currentLevel = typeof levelSystem !== 'undefined' ? levelSystem.currentLevel : 'GF';
    return project.elements.find(el =>
        el.type === 'buildingEnvelope' && (!el.level || el.level === currentLevel)
    );
}

/**
 * Get bounding box from an envelope element's points array.
 */
function envelopeBBox(envelope) {
    if (!envelope || !envelope.points || envelope.points.length < 3) return null;
    let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
    for (const pt of envelope.points) {
        minX = Math.min(minX, pt.x);
        maxX = Math.max(maxX, pt.x);
        minY = Math.min(minY, pt.y);
        maxY = Math.max(maxY, pt.y);
    }
    return { minX, maxX, minY, maxY };
}

/**
 * Polygon area via shoelace formula (mm² → m²).
 */
function envelopeArea(points) {
    if (!points || points.length < 3) return 0;
    let area = 0;
    const n = points.length;
    for (let i = 0; i < n; i++) {
        const j = (i + 1) % n;
        area += points[i].x * points[j].y - points[j].x * points[i].y;
    }
    return Math.abs(area) / 2 / 1e6; // convert mm² to m²
}


// ── Building Envelope Tool ──────────────────────────────────

const envelopeToolState = {
    points: [],        // [{x, y}] in sheet-mm
    currentEnd: null,  // preview point
};

let _envelopeNum = 1;

/**
 * Compute convex hull of a set of points (Andrew's monotone chain).
 * Returns points in CCW order.
 */
function convexHull(pts) {
    if (pts.length < 3) return pts.slice();
    const sorted = pts.slice().sort((a, b) => a.x - b.x || a.y - b.y);
    const cross = (O, A, B) => (A.x - O.x) * (B.y - O.y) - (A.y - O.y) * (B.x - O.x);
    const lower = [];
    for (const p of sorted) {
        while (lower.length >= 2 && cross(lower[lower.length - 2], lower[lower.length - 1], p) <= 0) lower.pop();
        lower.push(p);
    }
    const upper = [];
    for (let i = sorted.length - 1; i >= 0; i--) {
        const p = sorted[i];
        while (upper.length >= 2 && cross(upper[upper.length - 2], upper[upper.length - 1], p) <= 0) upper.pop();
        upper.push(p);
    }
    upper.pop();
    lower.pop();
    return lower.concat(upper);
}

/**
 * Auto-detect building envelope from wall elements on the current level.
 * Computes the convex hull of all wall endpoints.
 */
function autoDetectEnvelope() {
    const currentLevel = typeof levelSystem !== 'undefined' ? levelSystem.currentLevel : 'GF';
    const walls = project.elements.filter(el =>
        el.type === 'wall' && (!el.level || el.level === currentLevel)
    );
    if (walls.length < 2) return null;

    // Collect all wall endpoints in real-world mm
    const pts = [];
    for (const w of walls) {
        pts.push({ x: w.x1, y: w.y1 });
        pts.push({ x: w.x2, y: w.y2 });
    }

    const hull = convexHull(pts);
    if (hull.length < 3) return null;

    // Remove any existing envelope on this level
    const existing = findEnvelopeElement();
    if (existing) {
        const idx = project.elements.indexOf(existing);
        if (idx !== -1) project.elements.splice(idx, 1);
    }

    const envelope = {
        id: generateId(),
        type: 'buildingEnvelope',
        layer: 'S-ENVL',
        level: currentLevel,
        points: hull,
        closed: true,
    };

    history.execute({
        description: 'Auto-detect building envelope',
        execute() { project.elements.push(envelope); },
        undo() {
            const i = project.elements.indexOf(envelope);
            if (i !== -1) project.elements.splice(i, 1);
            // Restore previous envelope if one was removed
            if (existing) project.elements.push(existing);
        }
    });

    markComplianceDirty();
    engine.requestRender();
    return envelope;
}


// ── Envelope Drawing Tool Handlers ──────────────────────────

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'buildingEnvelope') return;
    if (pdfState.calibrating) return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    const snap = findSnap(sx, sy);
    let sheetPos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);

    // Shift: constrain to ortho (same logic as mousemove preview)
    if (e.shiftKey && envelopeToolState.points.length > 0) {
        const last = envelopeToolState.points[envelopeToolState.points.length - 1];
        const dx = sheetPos.x - last.x;
        const dy = sheetPos.y - last.y;
        if (Math.abs(dx) > Math.abs(dy)) {
            sheetPos = { x: sheetPos.x, y: last.y };
        } else {
            sheetPos = { x: last.x, y: sheetPos.y };
        }
    }

    // Check for closure — if near first point and we have ≥3 points
    if (envelopeToolState.points.length >= 3) {
        const first = envelopeToolState.points[0];
        const dist = Math.sqrt(Math.pow(sheetPos.x - first.x, 2) + Math.pow(sheetPos.y - first.y, 2));
        if (dist < 2 / (engine.viewport.zoom || 1)) {
            // Close the polygon
            _commitEnvelope();
            return;
        }
    }

    envelopeToolState.points.push({ ...sheetPos });
    envelopeToolState.currentEnd = { ...sheetPos };
    engine.requestRender();
});

container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'buildingEnvelope') return;
    if (envelopeToolState.points.length >= 3) {
        e.preventDefault();
        _commitEnvelope();
    }
});

container.addEventListener('mousemove', (e) => {
    if (activeTool !== 'buildingEnvelope' || envelopeToolState.points.length === 0) return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    let sheetPos = engine.coords.screenToSheet(sx, sy);
    const snap = findSnap(sx, sy);
    if (snap) sheetPos = { x: snap.x, y: snap.y };

    // Shift: constrain to ortho
    if (e.shiftKey && envelopeToolState.points.length > 0) {
        const last = envelopeToolState.points[envelopeToolState.points.length - 1];
        const dx = sheetPos.x - last.x;
        const dy = sheetPos.y - last.y;
        if (Math.abs(dx) > Math.abs(dy)) {
            sheetPos = { x: sheetPos.x, y: last.y };
        } else {
            sheetPos = { x: last.x, y: sheetPos.y };
        }
    }

    envelopeToolState.currentEnd = { ...sheetPos };
    engine.requestRender();
});

container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'buildingEnvelope' && envelopeToolState.points.length > 0) {
        e.preventDefault();
        if (envelopeToolState.points.length >= 3) {
            _commitEnvelope();
        } else {
            envelopeToolState.points = [];
            envelopeToolState.currentEnd = null;
            engine.requestRender();
        }
    }
});

// Enter key closes envelope polygon
document.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && activeTool === 'buildingEnvelope' && envelopeToolState.points.length >= 3) {
        e.preventDefault();
        _commitEnvelope();
    }
});

function _commitEnvelope() {
    const currentLevel = typeof levelSystem !== 'undefined' ? levelSystem.currentLevel : 'GF';

    // Convert sheet-mm points to real-world mm
    const realPoints = envelopeToolState.points.map(p => {
        const r = engine.coords.sheetToReal(p.x, p.y);
        return { x: r.x, y: r.y };
    });

    // Remove any existing envelope on this level
    const existing = findEnvelopeElement();

    const envelope = {
        id: generateId(),
        type: 'buildingEnvelope',
        layer: 'S-ENVL',
        level: currentLevel,
        points: realPoints,
        closed: true,
    };

    history.execute({
        description: 'Draw building envelope',
        execute() {
            if (existing) {
                const idx = project.elements.indexOf(existing);
                if (idx !== -1) project.elements.splice(idx, 1);
            }
            project.elements.push(envelope);
        },
        undo() {
            const i = project.elements.indexOf(envelope);
            if (i !== -1) project.elements.splice(i, 1);
            if (existing) project.elements.push(existing);
        }
    });

    // Reset tool state
    envelopeToolState.points = [];
    envelopeToolState.currentEnd = null;
    markComplianceDirty();
    engine.requestRender();
    updateBracingSummaryPanel();
}


// ── Envelope Rendering ──────────────────────────────────────

function drawBuildingEnvelopes(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    for (const el of project.getVisibleElements()) {
        if (el.type !== 'buildingEnvelope') continue;
        if (!el.points || el.points.length < 3) continue;

        const isSelected = (selectedElement === el);
        const color = isSelected ? '#2B7CD0' : '#6366f1';

        // Convert to screen coords
        const screenPts = el.points.map(p => coords.realToScreen(p.x, p.y));

        ctx.save();

        // Fill with semi-transparent
        ctx.fillStyle = isSelected ? 'rgba(43,124,208,0.06)' : 'rgba(99,102,241,0.04)';
        ctx.beginPath();
        ctx.moveTo(screenPts[0].x, screenPts[0].y);
        for (let i = 1; i < screenPts.length; i++) {
            ctx.lineTo(screenPts[i].x, screenPts[i].y);
        }
        ctx.closePath();
        ctx.fill();

        // Dashed outline
        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(0.5, 0.25 * zoom);
        ctx.setLineDash([6 * zoom, 3 * zoom]);
        ctx.globalAlpha = 0.6;
        ctx.beginPath();
        ctx.moveTo(screenPts[0].x, screenPts[0].y);
        for (let i = 1; i < screenPts.length; i++) {
            ctx.lineTo(screenPts[i].x, screenPts[i].y);
        }
        ctx.closePath();
        ctx.stroke();

        // Label at centroid
        if (zoom > 0.15) {
            const cx = screenPts.reduce((s, p) => s + p.x, 0) / screenPts.length;
            const cy = screenPts.reduce((s, p) => s + p.y, 0) / screenPts.length;
            const area = envelopeArea(el.points);
            const bbox = envelopeBBox(el);
            const dimX = bbox ? ((bbox.maxX - bbox.minX) / 1000).toFixed(1) : '?';
            const dimY = bbox ? ((bbox.maxY - bbox.minY) / 1000).toFixed(1) : '?';

            ctx.setLineDash([]);
            ctx.globalAlpha = 0.5;
            const fontSize = Math.max(7, 1.8 * zoom);
            ctx.font = `${fontSize}px "Segoe UI", Arial, sans-serif`;
            ctx.fillStyle = color;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText(`ENVELOPE ${dimX}m \u00D7 ${dimY}m`, cx, cy - fontSize * 0.7);
            ctx.fillText(`${area.toFixed(0)} m\u00B2`, cx, cy + fontSize * 0.7);
        }

        // Selection handles
        if (isSelected) {
            ctx.setLineDash([]);
            ctx.globalAlpha = 1;
            ctx.fillStyle = '#2B7CD0';
            for (const sp of screenPts) {
                ctx.fillRect(sp.x - 3, sp.y - 3, 6, 6);
            }
        }

        ctx.restore();
    }

    // Preview while drawing
    if (activeTool === 'buildingEnvelope' && envelopeToolState.points.length > 0) {
        const pts = envelopeToolState.points.map(p => coords.sheetToScreen(p.x, p.y));
        const cur = envelopeToolState.currentEnd ? coords.sheetToScreen(envelopeToolState.currentEnd.x, envelopeToolState.currentEnd.y) : null;

        ctx.save();
        ctx.strokeStyle = '#6366f1';
        ctx.lineWidth = Math.max(0.5, 0.25 * zoom);
        ctx.setLineDash([4, 3]);
        ctx.globalAlpha = 0.5;

        // Draw existing lines
        ctx.beginPath();
        ctx.moveTo(pts[0].x, pts[0].y);
        for (let i = 1; i < pts.length; i++) {
            ctx.lineTo(pts[i].x, pts[i].y);
        }
        if (cur) ctx.lineTo(cur.x, cur.y);
        // Close preview to first point
        if (pts.length >= 2) {
            ctx.setLineDash([2, 4]);
            ctx.lineTo(pts[0].x, pts[0].y);
        }
        ctx.stroke();

        // Semi-transparent fill
        if (pts.length >= 2 && cur) {
            ctx.fillStyle = 'rgba(99,102,241,0.05)';
            ctx.beginPath();
            ctx.moveTo(pts[0].x, pts[0].y);
            for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
            ctx.lineTo(cur.x, cur.y);
            ctx.closePath();
            ctx.fill();
        }

        // Vertex dots
        ctx.setLineDash([]);
        ctx.globalAlpha = 0.8;
        ctx.fillStyle = '#6366f1';
        for (const p of pts) {
            ctx.beginPath();
            ctx.arc(p.x, p.y, 3, 0, Math.PI * 2);
            ctx.fill();
        }

        ctx.restore();
    }
}

// Register envelope renderer
engine.onRender(drawBuildingEnvelopes);


// ══════════════════════════════════════════════════════════════
// RIDGE LINE — Drawing Tool, Rendering, Helpers
// ══════════════════════════════════════════════════════════════

const ridgeLineToolState = {
    points: [],        // [{x, y}] in sheet-mm
    currentEnd: null,
};

/**
 * Find the ridge line element on the roof level.
 */
function findRidgeElement() {
    // Search on RF level, or any level with a ridgeLine element
    return project.elements.find(el => el.type === 'ridgeLine');
}

/**
 * Switch to roof level and activate ridge line tool.
 * Called from the bracing panel "Draw ridge" button.
 */
function startDrawRidge() {
    // Find the RF (roof) level index
    const rfIndex = levelSystem.levels.findIndex(l => l.id === 'RF' || l.name.toLowerCase().includes('roof'));
    if (rfIndex >= 0) {
        switchToLevel(rfIndex);
    }
    setActiveTool('ridgeLine');
}

// ── Ridge Line Mousedown Handler ────────────────────────────

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'ridgeLine') return;
    if (pdfState.calibrating) return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    const snap = findSnap(sx, sy);
    const sheetPos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);

    ridgeLineToolState.points.push({ ...sheetPos });
    ridgeLineToolState.currentEnd = { ...sheetPos };
    engine.requestRender();
});

container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'ridgeLine') return;
    if (ridgeLineToolState.points.length >= 2) {
        e.preventDefault();
        _commitRidgeLine();
    }
});

container.addEventListener('mousemove', (e) => {
    if (activeTool !== 'ridgeLine' || ridgeLineToolState.points.length === 0) return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    let sheetPos = engine.coords.screenToSheet(sx, sy);
    const snap = findSnap(sx, sy);
    if (snap) sheetPos = { x: snap.x, y: snap.y };

    // Shift: constrain to ortho
    if (e.shiftKey && ridgeLineToolState.points.length > 0) {
        const last = ridgeLineToolState.points[ridgeLineToolState.points.length - 1];
        const dx = sheetPos.x - last.x;
        const dy = sheetPos.y - last.y;
        if (Math.abs(dx) > Math.abs(dy)) {
            sheetPos = { x: sheetPos.x, y: last.y };
        } else {
            sheetPos = { x: last.x, y: sheetPos.y };
        }
    }

    ridgeLineToolState.currentEnd = { ...sheetPos };
    engine.requestRender();
});

// Right-click or Enter to commit ridge
container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'ridgeLine' && ridgeLineToolState.points.length > 0) {
        e.preventDefault();
        if (ridgeLineToolState.points.length >= 2) {
            _commitRidgeLine();
        } else {
            ridgeLineToolState.points = [];
            ridgeLineToolState.currentEnd = null;
            engine.requestRender();
        }
    }
});

document.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && activeTool === 'ridgeLine' && ridgeLineToolState.points.length >= 2) {
        e.preventDefault();
        _commitRidgeLine();
    }
});

function _commitRidgeLine() {
    const currentLevel = typeof levelSystem !== 'undefined' ? levelSystem.currentLevel : 'RF';

    // Convert sheet-mm points to real-world mm
    const realPoints = ridgeLineToolState.points.map(p => {
        const r = engine.coords.sheetToReal(p.x, p.y);
        return { x: r.x, y: r.y };
    });

    // Default pitches: 22.5° both sides for each segment
    const pitches = [];
    for (let i = 0; i < realPoints.length - 1; i++) {
        pitches.push({ a: 22.5, b: 22.5 });
    }

    // Remove any existing ridge on any level (one ridge per project for now)
    const existing = findRidgeElement();

    const ridge = {
        id: generateId(),
        type: 'ridgeLine',
        layer: 'S-RIDGE',
        level: currentLevel,
        points: realPoints,
        closed: false,
        pitches: pitches,
    };

    history.execute({
        description: 'Draw ridge line',
        execute() {
            if (existing) {
                const idx = project.elements.indexOf(existing);
                if (idx !== -1) project.elements.splice(idx, 1);
            }
            project.elements.push(ridge);
        },
        undo() {
            const i = project.elements.indexOf(ridge);
            if (i !== -1) project.elements.splice(i, 1);
            if (existing) project.elements.push(existing);
        }
    });

    ridgeLineToolState.points = [];
    ridgeLineToolState.currentEnd = null;
    markComplianceDirty();
    engine.requestRender();
    updateBracingSummaryPanel();
}


// ── Ridge Line Rendering ────────────────────────────────────

function drawRidgeLines(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    // Draw committed ridge elements (show on ALL levels as ghost reference)
    for (const el of project.elements) {
        if (el.type !== 'ridgeLine') continue;
        if (!el.points || el.points.length < 2) continue;

        const layer = project.layers[el.layer];
        if (layer && !layer.visible) continue;

        const isSelected = (selectedElement === el);
        const isCurrentLevel = (el.level === (typeof levelSystem !== 'undefined' ? levelSystem.currentLevel : 'GF'));
        const color = isSelected ? '#2B7CD0' : '#dc2626';
        const alpha = isCurrentLevel ? 1.0 : 0.2; // Ghost on other levels

        const screenPts = el.points.map(p => coords.realToScreen(p.x, p.y));

        ctx.save();
        ctx.globalAlpha = alpha;

        // Draw ridge line — thick red
        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(1.5, 0.5 * zoom);
        ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(screenPts[0].x, screenPts[0].y);
        for (let i = 1; i < screenPts.length; i++) {
            ctx.lineTo(screenPts[i].x, screenPts[i].y);
        }
        ctx.stroke();

        // Draw pitch labels at segment midpoints
        if (zoom > 0.2 && el.pitches) {
            for (let i = 0; i < screenPts.length - 1 && i < el.pitches.length; i++) {
                const mx = (screenPts[i].x + screenPts[i + 1].x) / 2;
                const my = (screenPts[i].y + screenPts[i + 1].y) / 2;
                const p = el.pitches[i];
                const label = `${p.a}\u00B0 / ${p.b}\u00B0`;

                const fontSize = Math.max(7, 1.8 * zoom);
                ctx.font = `bold ${fontSize}px "Segoe UI", Arial, sans-serif`;
                ctx.fillStyle = color;
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                ctx.fillText(label, mx, my - 4);

                // Small perpendicular tick marks to show pitch direction
                const dx = screenPts[i + 1].x - screenPts[i].x;
                const dy = screenPts[i + 1].y - screenPts[i].y;
                const len = Math.sqrt(dx * dx + dy * dy);
                if (len > 5) {
                    const nx = -dy / len, ny = dx / len;
                    const tickLen = Math.max(4, 2 * zoom);
                    ctx.lineWidth = Math.max(0.5, 0.2 * zoom);
                    ctx.beginPath();
                    ctx.moveTo(mx - nx * tickLen, my - ny * tickLen);
                    ctx.lineTo(mx + nx * tickLen, my + ny * tickLen);
                    ctx.stroke();
                    // Label sides A/B
                    if (zoom > 0.4) {
                        ctx.font = `${Math.max(5, 1.2 * zoom)}px "Segoe UI", Arial, sans-serif`;
                        ctx.globalAlpha = alpha * 0.6;
                        ctx.fillText('A', mx - nx * (tickLen + fontSize * 0.4), my - ny * (tickLen + fontSize * 0.4));
                        ctx.fillText('B', mx + nx * (tickLen + fontSize * 0.4), my + ny * (tickLen + fontSize * 0.4));
                        ctx.globalAlpha = alpha;
                    }
                }
            }
        }

        // Ridge endpoint markers (small diamonds)
        ctx.fillStyle = color;
        for (const sp of screenPts) {
            const r = isSelected ? 4 : 3;
            ctx.beginPath();
            ctx.moveTo(sp.x, sp.y - r);
            ctx.lineTo(sp.x + r, sp.y);
            ctx.lineTo(sp.x, sp.y + r);
            ctx.lineTo(sp.x - r, sp.y);
            ctx.closePath();
            ctx.fill();
        }

        // Selection handles
        if (isSelected) {
            ctx.fillStyle = '#2B7CD0';
            ctx.globalAlpha = 1;
            for (const sp of screenPts) {
                ctx.fillRect(sp.x - 3, sp.y - 3, 6, 6);
            }
        }

        ctx.restore();
    }

    // Preview while drawing
    if (activeTool === 'ridgeLine' && ridgeLineToolState.points.length > 0) {
        const pts = ridgeLineToolState.points.map(p => coords.sheetToScreen(p.x, p.y));
        const cur = ridgeLineToolState.currentEnd ? coords.sheetToScreen(ridgeLineToolState.currentEnd.x, ridgeLineToolState.currentEnd.y) : null;

        ctx.save();
        ctx.strokeStyle = '#dc2626';
        ctx.lineWidth = Math.max(1.5, 0.5 * zoom);
        ctx.setLineDash([4, 3]);
        ctx.globalAlpha = 0.6;

        ctx.beginPath();
        ctx.moveTo(pts[0].x, pts[0].y);
        for (let i = 1; i < pts.length; i++) {
            ctx.lineTo(pts[i].x, pts[i].y);
        }
        if (cur) ctx.lineTo(cur.x, cur.y);
        ctx.stroke();

        // Vertex dots
        ctx.setLineDash([]);
        ctx.globalAlpha = 0.8;
        ctx.fillStyle = '#dc2626';
        for (const p of pts) {
            ctx.beginPath();
            ctx.arc(p.x, p.y, 3, 0, Math.PI * 2);
            ctx.fill();
        }

        // Show length preview
        if (cur && pts.length >= 1) {
            const lastPt = pts[pts.length - 1];
            const sheetLen = Math.sqrt(
                Math.pow(ridgeLineToolState.currentEnd.x - ridgeLineToolState.points[ridgeLineToolState.points.length - 1].x, 2) +
                Math.pow(ridgeLineToolState.currentEnd.y - ridgeLineToolState.points[ridgeLineToolState.points.length - 1].y, 2)
            );
            const realLen = sheetLen * CONFIG.drawingScale;
            if (realLen > 50) {
                const mx = (lastPt.x + cur.x) / 2, my = (lastPt.y + cur.y) / 2;
                ctx.font = `${Math.max(8, 2 * zoom)}px "Segoe UI", Arial, sans-serif`;
                ctx.fillStyle = '#dc2626';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                ctx.fillText(Math.round(realLen) + ' mm', mx, my - 6);
            }
        }

        ctx.restore();
    }
}

// Register ridge renderer
engine.onRender(drawRidgeLines);


// ── Bracing Wall Mousedown Handler ──────────────────────────

container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'bracingWall') return;
    if (pdfState.calibrating) return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    // Use snap
    const snap = findSnap(sx, sy);
    const sheetPos = snap
        ? { x: snap.x, y: snap.y }
        : engine.coords.screenToSheet(sx, sy);

    if (!bracingWallToolState.placing) {
        // First click: start
        bracingWallToolState.placing = true;
        bracingWallToolState.startPoint = { ...sheetPos };
        bracingWallToolState.currentEnd = { ...sheetPos };
        engine.requestRender();
    } else {
        // Second click: finish placement
        const realStart = engine.coords.sheetToReal(
            bracingWallToolState.startPoint.x, bracingWallToolState.startPoint.y
        );
        const realEnd = engine.coords.sheetToReal(
            bracingWallToolState.currentEnd.x, bracingWallToolState.currentEnd.y
        );

        const typeRef = placementTypeRef.bracingWall || 'BR1';
        const typeData = project.scheduleTypes.bracingWall?.[typeRef] || {};

        const tag = typeRef;
        const currentLevel = typeof levelSystem !== 'undefined' ? levelSystem.currentLevel : 'GF';

        const newBracingWall = {
            id: generateId(),
            type: 'bracingWall',
            layer: 'S-BRAC',
            level: currentLevel,
            x1: realStart.x, y1: realStart.y,
            x2: realEnd.x,   y2: realEnd.y,
            tag: tag,
            typeRef: typeRef,
        };

        history.execute({
            description: 'Draw bracing wall: ' + tag,
            execute() {
                project.elements.push(newBracingWall);
            },
            undo() {
                const i = project.elements.indexOf(newBracingWall);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });

        // Chain mode: start next from where this one ended
        bracingWallToolState.startPoint = { ...bracingWallToolState.currentEnd };
        engine.requestRender();

        // Update bracing summary panel if visible
        updateBracingSummaryPanel();
    }
});

// ── Bracing Wall Mousemove Handler ──────────────────────────

container.addEventListener('mousemove', (e) => {
    if (activeTool !== 'bracingWall' || !bracingWallToolState.placing || !bracingWallToolState.startPoint) return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    let sheetPos = engine.coords.screenToSheet(sx, sy);

    // Check for snap
    const snap = findSnap(sx, sy);
    if (snap) sheetPos = { x: snap.x, y: snap.y };

    // Shift: constrain to 45° angles
    if (e.shiftKey && bracingWallToolState.startPoint) {
        const dx = sheetPos.x - bracingWallToolState.startPoint.x;
        const dy = sheetPos.y - bracingWallToolState.startPoint.y;
        const angle = Math.atan2(dy, dx);
        const snapAngle = Math.round(angle / (Math.PI / 4)) * (Math.PI / 4);
        const dist = Math.sqrt(dx * dx + dy * dy);
        sheetPos = {
            x: bracingWallToolState.startPoint.x + dist * Math.cos(snapAngle),
            y: bracingWallToolState.startPoint.y + dist * Math.sin(snapAngle)
        };
    }

    bracingWallToolState.currentEnd = { ...sheetPos };
    engine.requestRender();
});

// ── Bracing Wall Right-Click Cancel ─────────────────────────

container.addEventListener('contextmenu', (e) => {
    if (activeTool === 'bracingWall' && bracingWallToolState.placing) {
        e.preventDefault();
        bracingWallToolState.placing = false;
        bracingWallToolState.startPoint = null;
        bracingWallToolState.currentEnd = null;
        engine.requestRender();
    }
});


// ── Bracing Wall Rendering ──────────────────────────────────

/**
 * Panel width in real-world mm — represents the drawn width of the
 * bracing panel rectangle on the plan.  Typical stud wall = 90mm,
 * but for clarity on bracing plans we draw slightly wider so the X
 * is legible at standard 1:100 print scale.
 */
const BRACE_PANEL_WIDTH_MM = 120; // real-world mm

/**
 * Draw bracing walls as thin-lined rectangular panels with a single
 * corner-to-corner X inside — matching standard Australian structural
 * drawing convention (ref: ACOR bracing plan style).
 *
 * Label format: "1600 (6.4)" = length mm (capacity kN/m)
 */
function drawBracingWalls(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    for (const el of project.getVisibleElements()) {
        if (el.type !== 'bracingWall') continue;

        const layer = project.layers[el.layer];
        if (!layer) continue;
        const isSelected = (selectedElement === el);

        const p1 = coords.realToScreen(el.x1, el.y1);
        const p2 = coords.realToScreen(el.x2, el.y2);
        const dx = p2.x - p1.x;
        const dy = p2.y - p1.y;
        const len = Math.sqrt(dx * dx + dy * dy);
        if (len < 1) continue;

        // Get type colour — default black for structural drawing convention
        const typeRef = el.typeRef || 'BR1';
        const typeData = project.scheduleTypes.bracingWall?.[typeRef] || {};
        const color = isSelected ? '#2B7CD0' : (typeData.color || '#000000');

        // ── Panel rectangle geometry ──
        // halfW = half the panel width in screen pixels
        const halfW = (BRACE_PANEL_WIDTH_MM / 2 / CONFIG.drawingScale) * zoom;
        const ux = dx / len, uy = dy / len;  // unit along wall
        const nx = -uy, ny = ux;             // perpendicular

        // Four corners of the panel rectangle
        const c1 = { x: p1.x + nx * halfW, y: p1.y + ny * halfW };
        const c2 = { x: p1.x - nx * halfW, y: p1.y - ny * halfW };
        const c3 = { x: p2.x - nx * halfW, y: p2.y - ny * halfW };
        const c4 = { x: p2.x + nx * halfW, y: p2.y + ny * halfW };

        ctx.save();
        const lineW = Math.max(0.5, 0.35 * zoom);

        // ── Panel outline (thin rectangle) ──
        ctx.strokeStyle = color;
        ctx.lineWidth = lineW;
        ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(c1.x, c1.y);
        ctx.lineTo(c4.x, c4.y);
        ctx.lineTo(c3.x, c3.y);
        ctx.lineTo(c2.x, c2.y);
        ctx.closePath();
        ctx.stroke();

        // ── Diagonal X (corner-to-corner) ──
        ctx.beginPath();
        ctx.moveTo(c1.x, c1.y);
        ctx.lineTo(c3.x, c3.y);
        ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(c2.x, c2.y);
        ctx.lineTo(c4.x, c4.y);
        ctx.stroke();

        // ── Label: "1600 (6.4)" — length mm then capacity ──
        if (zoom > 0.3) {
            const midX = (p1.x + p2.x) / 2;
            const midY = (p1.y + p2.y) / 2;

            let screenAngle = Math.atan2(dy, dx);
            if (screenAngle > Math.PI / 2) screenAngle -= Math.PI;
            if (screenAngle < -Math.PI / 2) screenAngle += Math.PI;

            // Real length in mm
            const realLenMm = Math.sqrt(
                Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2)
            );
            const baseCap = typeData.capacity || (BRACING_TYPES[typeData.bracingType]?.capacity) || 0;
            const labelText = `${Math.round(realLenMm)} (${baseCap !== null ? baseCap.toFixed(1) : '?'})`;

            const fontSize = Math.max(1, 2 * zoom);
            const baseOffset = halfW + fontSize * 0.6 + 2;
            const tagDx = (el._tagOffsetX || 0) / CONFIG.drawingScale * zoom;
            const tagDy = (el._tagOffsetY || 0) / CONFIG.drawingScale * zoom;

            ctx.translate(midX + tagDx, midY + tagDy);
            ctx.rotate(screenAngle);
            ctx.font = `${fontSize}px "Segoe UI", Arial, sans-serif`;
            ctx.fillStyle = isSelected ? '#2B7CD0' : '#000000';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            ctx.fillText(labelText, 0, -baseOffset);
            ctx.setTransform(1, 0, 0, 1, 0, 0);
        }

        // ── Selection handles ──
        if (isSelected) {
            const r = 4;
            ctx.fillStyle = '#2B7CD0';
            ctx.setLineDash([]);
            ctx.fillRect(p1.x - r, p1.y - r, r * 2, r * 2);
            ctx.fillRect(p2.x - r, p2.y - r, r * 2, r * 2);
        }

        ctx.restore();
    }

    // ── Preview while placing ──
    if (activeTool === 'bracingWall' && bracingWallToolState.placing && bracingWallToolState.startPoint) {
        const s1 = coords.sheetToScreen(bracingWallToolState.startPoint.x, bracingWallToolState.startPoint.y);
        const s2 = coords.sheetToScreen(bracingWallToolState.currentEnd.x, bracingWallToolState.currentEnd.y);
        const pdx = s2.x - s1.x, pdy = s2.y - s1.y;
        const plen = Math.sqrt(pdx * pdx + pdy * pdy);

        const typeRef = placementTypeRef.bracingWall || 'BR1';
        const typeData = project.scheduleTypes.bracingWall?.[typeRef] || {};
        const previewColor = typeData.color || '#000000';

        if (plen > 1) {
            const halfW = (BRACE_PANEL_WIDTH_MM / 2 / CONFIG.drawingScale) * zoom;
            const pux = pdx / plen, puy = pdy / plen;
            const pnx = -puy, pny = pux;

            const pc1 = { x: s1.x + pnx * halfW, y: s1.y + pny * halfW };
            const pc2 = { x: s1.x - pnx * halfW, y: s1.y - pny * halfW };
            const pc3 = { x: s2.x - pnx * halfW, y: s2.y - pny * halfW };
            const pc4 = { x: s2.x + pnx * halfW, y: s2.y + pny * halfW };

            ctx.save();
            ctx.strokeStyle = previewColor;
            ctx.lineWidth = Math.max(0.5, 0.35 * zoom);
            ctx.globalAlpha = 0.5;
            ctx.setLineDash([4, 3]);

            // Panel outline
            ctx.beginPath();
            ctx.moveTo(pc1.x, pc1.y);
            ctx.lineTo(pc4.x, pc4.y);
            ctx.lineTo(pc3.x, pc3.y);
            ctx.lineTo(pc2.x, pc2.y);
            ctx.closePath();
            ctx.stroke();

            // X diagonals
            ctx.setLineDash([]);
            ctx.beginPath();
            ctx.moveTo(pc1.x, pc1.y); ctx.lineTo(pc3.x, pc3.y); ctx.stroke();
            ctx.beginPath();
            ctx.moveTo(pc2.x, pc2.y); ctx.lineTo(pc4.x, pc4.y); ctx.stroke();

            // Preview length
            const sheetLen = Math.sqrt(
                Math.pow(bracingWallToolState.currentEnd.x - bracingWallToolState.startPoint.x, 2) +
                Math.pow(bracingWallToolState.currentEnd.y - bracingWallToolState.startPoint.y, 2)
            );
            const realLen = sheetLen * CONFIG.drawingScale;
            if (realLen > 50) {
                ctx.globalAlpha = 0.8;
                const mx = (s1.x + s2.x) / 2, my = (s1.y + s2.y) / 2;
                ctx.font = `${Math.max(8, 2 * zoom)}px "Segoe UI", Arial, sans-serif`;
                ctx.fillStyle = previewColor;
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                ctx.fillText(Math.round(realLen) + ' mm', mx, my - halfW - 4);
            }
            ctx.restore();
        }
    }
}

// Register renderer
engine.onRender(drawBracingWalls);


// ── Bracing Summary Panel UI ─────────────────────────────────

/**
 * Render the Phase D compliance section as HTML for the summary panel.
 */
function renderComplianceSection() {
    const compliance = getCachedCompliance();
    const spacingIssues = compliance.issues.filter(i => i.type === 'spacing');
    const cornerIssues = compliance.issues.filter(i => i.type === 'corner');
    const propIssues = compliance.issues.filter(i => i.type === 'proportioning');

    const statusColor = compliance.overallOk ? '#16a34a' : '#dc2626';
    const statusLabel = compliance.overallOk ? 'PASS' : 'FAIL';
    const statusIcon = compliance.overallOk ? '✓' : '✗';

    let html = `
        <div style="margin-top:10px;padding-top:8px;border-top:1px solid #e5e7eb;">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
                <span style="font-weight:700;font-size:11px;">Compliance Checks</span>
                <span style="font-weight:700;font-size:12px;color:${statusColor};">${statusIcon} ${statusLabel}</span>
            </div>`;

    // Spacing
    const xSpacing = compliance.spacing.x;
    const ySpacing = compliance.spacing.y;
    const spacingOk = spacingIssues.length === 0;
    html += `<div style="font-size:10px;margin-bottom:4px;">
        <span style="color:${spacingOk ? '#16a34a' : '#dc2626'};font-weight:600;">${spacingOk ? '✓' : '✗'} Spacing</span>
        <span style="color:#999;"> (max ${(getMaxSpacing(bracingSettings.windClass, 3.4) / 1000).toFixed(0)}m for ${bracingSettings.windClass})</span>
    </div>`;
    if (spacingIssues.length > 0) {
        for (const s of spacingIssues) {
            const aRef = s.wallA ? (s.wallA.typeRef || '?') : 'Edge';
            const bRef = s.wallB ? (s.wallB.typeRef || '?') : 'Edge';
            html += `<div style="font-size:9px;color:#dc2626;padding-left:12px;margin-bottom:2px;">
                ${s.dir}-dir: ${aRef} → ${bRef} = ${(s.gap/1000).toFixed(1)}m > ${(s.max/1000).toFixed(1)}m
            </div>`;
        }
    }

    // Corners
    const cornersOk = cornerIssues.length === 0;
    if (compliance.hasExtent) {
        html += `<div style="font-size:10px;margin-bottom:4px;margin-top:4px;">
            <span style="color:${cornersOk ? '#16a34a' : '#dc2626'};font-weight:600;">${cornersOk ? '✓' : '✗'} Corner bracing</span>
            <span style="color:#999;"> (within 4.5m)</span>
        </div>`;
        if (cornerIssues.length > 0) {
            for (const c of cornerIssues) {
                html += `<div style="font-size:9px;color:#dc2626;padding-left:12px;margin-bottom:2px;">
                    ${c.label}: nearest at ${(c.distance/1000).toFixed(1)}m
                </div>`;
            }
        }
    }

    // Proportioning
    const propOk = propIssues.length === 0;
    if (compliance.hasExtent) {
        html += `<div style="font-size:10px;margin-bottom:4px;margin-top:4px;">
            <span style="color:${propOk ? '#16a34a' : '#f59e0b'};font-weight:600;">${propOk ? '✓' : '⚠'} Distribution</span>
            <span style="color:#999;"> (min 33% each half)</span>
        </div>`;
        if (propIssues.length > 0) {
            for (const p of propIssues) {
                html += `<div style="font-size:9px;color:#f59e0b;padding-left:12px;margin-bottom:2px;">
                    ${p.dir}-dir: weaker half has ${(p.ratio * 100).toFixed(0)}% of capacity
                </div>`;
            }
        }

    }

    if (!compliance.hasExtent) {
        html += `<div style="font-size:9px;color:#999;margin-top:4px;font-style:italic;">
            Draw regular walls to define building envelope for corner &amp; distribution checks.
        </div>`;
    }

    html += '</div>';
    return html;
}

/**
 * Update the bracing summary panel (sidebar or overlay).
 * Creates/updates HTML in #bracing-summary-panel.
 */
function updateBracingSummaryPanel() {
    const panel = document.getElementById('bracing-summary-panel');
    if (!panel) return;
    if (panel.classList.contains('hidden')) return; // Don't update if closed
    markComplianceDirty();

    const summary = calcBracingSummary();

    const xOk = summary.xRatio >= 1.0;
    const yOk = summary.yRatio >= 1.0;
    const xPct = summary.xDemand > 0 ? (summary.xRatio * 100).toFixed(0) : '—';
    const yPct = summary.yDemand > 0 ? (summary.yRatio * 100).toFixed(0) : '—';

    const isAuto = bracingSettings.demandMode === 'auto';
    const envelope = findEnvelopeElement();
    const envBBox = envelope ? envelopeBBox(envelope) : null;
    const envDimX = envBBox ? ((envBBox.maxX - envBBox.minX) / 1000).toFixed(1) : '—';
    const envDimY = envBBox ? ((envBBox.maxY - envBBox.minY) / 1000).toFixed(1) : '—';
    const hasWalls = project.elements.some(el => el.type === 'wall');

    panel.innerHTML = `
        <div class="bracing-panel-header">
            <span style="font-weight:700;font-size:12px;">Bracing Summary — AS 1684</span>
            <button class="tbtn" onclick="document.getElementById('bracing-summary-panel').classList.add('hidden')" style="font-size:10px;padding:2px 6px;">\u00D7</button>
        </div>

        <div class="bracing-settings-row">
            <label>Ceiling Ht
                <select id="bracing-ceiling-ht" onchange="bracingSettings.ceilingHeight=parseInt(this.value);updateBracingSummaryPanel();engine.requestRender();">
                    ${[2400,2700,3000,3300,3600,3900,4200].map(h =>
                        `<option value="${h}"${h===summary.ceilingHeight?' selected':''}>${h}mm</option>`
                    ).join('')}
                </select>
            </label>
            <label>Joint Grp
                <select id="bracing-joint-grp" onchange="bracingSettings.jointGroup=this.value;updateBracingSummaryPanel();engine.requestRender();">
                    ${['JD6','JD5'].map(j =>
                        `<option value="${j}"${j===summary.jointGroup?' selected':''}>${j}</option>`
                    ).join('')}
                </select>
            </label>
            <label>Wind
                <select id="bracing-wind-class" onchange="bracingSettings.windClass=this.value;updateBracingSummaryPanel();engine.requestRender();">
                    ${['N1','N2','N3','N4'].map(w =>
                        `<option value="${w}"${w===bracingSettings.windClass?' selected':''}>${w}</option>`
                    ).join('')}
                </select>
            </label>
        </div>

        <!-- Building Envelope (always visible) -->
        <div style="padding:6px 10px;border-bottom:1px solid #e5e7eb;">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;">
                <span style="font-weight:700;font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#71717a;">Building Envelope</span>
                <span style="font-size:10px;">
                    ${envelope
                        ? `<span style="color:#16a34a;font-weight:600;">\u2713 ${envDimX}m \u00D7 ${envDimY}m</span>`
                        : `<span style="color:#999;">Not defined</span>`
                    }
                </span>
            </div>
            <div style="display:flex;gap:4px;margin-bottom:2px;">
                ${hasWalls ? `<button class="tbtn" style="font-size:9px;padding:2px 6px;" onclick="autoDetectEnvelope();updateBracingSummaryPanel();engine.requestRender();">Auto-detect from walls</button>` : ''}
                <button class="tbtn" style="font-size:9px;padding:2px 6px;" onclick="setActiveTool('buildingEnvelope');">Draw envelope</button>
            </div>
        </div>

        <!-- Roof Skeleton section -->
        <div style="padding:6px 10px;border-bottom:1px solid #e5e7eb;">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;">
                <span style="font-weight:700;font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#71717a;">Roof Skeleton</span>
                <span style="font-size:10px;">
                    ${(() => {
                        const sk = typeof findSkeletonElement === 'function' ? findSkeletonElement() : null;
                        if (!sk) return '<span style="color:#999;">Not defined</span>';
                        const drawnEdges = sk.edges ? sk.edges.filter(e => e.source === 'drawn').length : 0;
                        const faceCount = sk.faces ? sk.faces.length : 0;
                        const pitchedCount = sk.faces ? sk.faces.filter(f => f.pitch != null).length : 0;
                        const allPitched = faceCount > 0 && pitchedCount === faceCount;
                        const color = allPitched ? '#16a34a' : '#f59e0b';
                        return '<span style="color:' + color + ';font-weight:600;">' +
                            (allPitched ? '\u2713 ' : '\u26A0 ') +
                            drawnEdges + ' lines \u00B7 ' + pitchedCount + '/' + faceCount + ' faces pitched</span>';
                    })()}
                </span>
            </div>
            ${(() => {
                const sk = typeof findSkeletonElement === 'function' ? findSkeletonElement() : null;
                if (sk && sk.faces && sk.faces.length > 0) {
                    let html = '';
                    sk.faces.forEach((f, i) => {
                        const val = f.pitch != null ? f.pitch : '';
                        html += '<div style="display:flex;gap:4px;align-items:center;margin-bottom:3px;font-size:10px;">';
                        html += '<span style="color:#71717a;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' + (f.label || ('Face ' + (i+1))) + '</span>';
                        html += '<input type="number" value="' + val + '" min="0" max="60" step="0.5" placeholder="°" '
                             + 'style="width:44px;font-size:10px;padding:1px 3px;" '
                             + 'onchange="(function(){var sk=findSkeletonElement();if(sk&&sk.faces[' + i + ']){sk.faces[' + i + '].pitch=parseFloat(this.value)||null;markComplianceDirty();engine.requestRender();updateBracingSummaryPanel();}}).call(this)">\u00B0';
                        html += '</div>';
                    });
                    return html;
                }
                return '';
            })()}
            <div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:2px;">
                <button class="tbtn" style="font-size:9px;padding:2px 6px;" onclick="startDrawSkeleton();">Draw skeleton</button>
                <button class="tbtn" style="font-size:9px;padding:2px 6px;" onclick="generateHipRoofFromEnvelope();">Auto hip roof</button>
                ${(() => {
                    const sk = typeof findSkeletonElement === 'function' ? findSkeletonElement() : null;
                    if (sk) return '<button class="tbtn" style="font-size:9px;padding:2px 6px;color:#dc2626;" onclick="var s=findSkeletonElement();if(s){var idx=project.elements.indexOf(s);if(idx!==-1)project.elements.splice(idx,1);markComplianceDirty();engine.requestRender();updateBracingSummaryPanel();}">Clear</button>';
                    return '';
                })()}
            </div>
        </div>

        <!-- Phase B: Demand Mode -->
        <div style="padding:6px 10px;border-bottom:1px solid #e5e7eb;">
            <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
                <span style="font-weight:700;font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#71717a;">Wind Demand</span>
                <label style="font-size:10px;cursor:pointer;">
                    <input type="radio" name="demand-mode" value="auto" ${isAuto?'checked':''}
                        onchange="bracingSettings.demandMode='auto';bracingSettings.xDemand=0;bracingSettings.yDemand=0;updateBracingSummaryPanel();engine.requestRender();">
                    Auto
                </label>
                <label style="font-size:10px;cursor:pointer;">
                    <input type="radio" name="demand-mode" value="manual" ${!isAuto?'checked':''}
                        onchange="bracingSettings.demandMode='manual';updateBracingSummaryPanel();">
                    Manual
                </label>
            </div>

            ${isAuto ? `
                <!-- Auto demand controls -->
                <div style="display:flex;gap:6px;margin-bottom:6px;flex-wrap:wrap;">
                    <label style="font-size:10px;">Roof
                        <select style="font-size:10px;padding:1px 3px;" onchange="bracingSettings.roofType=this.value;updateBracingSummaryPanel();engine.requestRender();">
                            ${['hip','gable','skillion'].map(r =>
                                `<option value="${r}"${r===bracingSettings.roofType?' selected':''}>${r.charAt(0).toUpperCase()+r.slice(1)}</option>`
                            ).join('')}
                        </select>
                    </label>
                    <label style="font-size:10px;">Pitch
                        <input type="number" value="${bracingSettings.roofPitch}" min="0" max="45" step="0.5"
                            style="width:45px;font-size:10px;padding:1px 3px;"
                            onchange="bracingSettings.roofPitch=parseFloat(this.value)||22.5;updateBracingSummaryPanel();engine.requestRender();">\u00B0
                    </label>
                    <label style="font-size:10px;">Storey
                        <select style="font-size:10px;padding:1px 3px;" onchange="bracingSettings.storeyPosition=this.value;updateBracingSummaryPanel();engine.requestRender();">
                            ${[['single','Single'],['upper','Upper'],['lower','Lower']].map(([v,l]) =>
                                `<option value="${v}"${v===bracingSettings.storeyPosition?' selected':''}>${l}</option>`
                            ).join('')}
                        </select>
                    </label>
                </div>

                ${!envelope ? `
                    <div style="font-size:9px;color:#999;font-style:italic;margin-bottom:4px;">Define a building envelope above to auto-calculate demand.</div>
                ` : ''}

                ${summary.windDemandResult ? `
                    <!-- Auto-calculated breakdown -->
                    <div style="font-size:9px;color:#374151;background:#f9fafb;border:1px solid #e5e7eb;border-radius:4px;padding:4px 6px;margin-bottom:4px;">
                        <div style="margin-bottom:2px;"><strong>X-Dir:</strong> ${summary.windDemandResult.xBreakdown}</div>
                        <div><strong>Y-Dir:</strong> ${summary.windDemandResult.yBreakdown}</div>
                    </div>
                ` : ''}

                <!-- Manual override in auto mode -->
                <div style="font-size:9px;color:#71717a;margin-top:2px;">
                    Override: X
                    <input type="number" value="${bracingSettings.xDemand||''}" min="0" step="0.1" placeholder="auto"
                        style="width:40px;font-size:9px;padding:1px 2px;"
                        onchange="bracingSettings.xDemand=parseFloat(this.value)||0;updateBracingSummaryPanel();">
                    kN &nbsp; Y
                    <input type="number" value="${bracingSettings.yDemand||''}" min="0" step="0.1" placeholder="auto"
                        style="width:40px;font-size:9px;padding:1px 2px;"
                        onchange="bracingSettings.yDemand=parseFloat(this.value)||0;updateBracingSummaryPanel();">
                    kN
                </div>
            ` : `
                <!-- Manual demand inputs -->
                <div class="bracing-demand-row">
                    <label>X Demand (kN)
                        <input type="number" id="bracing-x-demand" value="${summary.xDemand}" min="0" step="0.1"
                            onchange="bracingSettings.xDemand=parseFloat(this.value)||0;updateBracingSummaryPanel();">
                    </label>
                    <label>Y Demand (kN)
                        <input type="number" id="bracing-y-demand" value="${summary.yDemand}" min="0" step="0.1"
                            onchange="bracingSettings.yDemand=parseFloat(this.value)||0;updateBracingSummaryPanel();">
                    </label>
                </div>
            `}
        </div>

        <table class="bracing-summary-table">
            <thead>
                <tr><th></th><th>X-Dir</th><th>Y-Dir</th></tr>
            </thead>
            <tbody>
                <tr>
                    <td>Demand</td>
                    <td>${summary.xDemand.toFixed(1)} kN</td>
                    <td>${summary.yDemand.toFixed(1)} kN</td>
                </tr>
                <tr>
                    <td>Structural</td>
                    <td>${summary.xCapacityStructural.toFixed(1)} kN</td>
                    <td>${summary.yCapacityStructural.toFixed(1)} kN</td>
                </tr>
                <tr>
                    <td>Nominal</td>
                    <td>${summary.xNominalCapped.toFixed(1)} kN <span style="color:#999;font-size:9px;">(raw ${summary.xNominalRaw.toFixed(1)})</span></td>
                    <td>${summary.yNominalCapped.toFixed(1)} kN <span style="color:#999;font-size:9px;">(raw ${summary.yNominalRaw.toFixed(1)})</span></td>
                </tr>
                <tr style="font-weight:700;">
                    <td>Total Capacity</td>
                    <td>${summary.xCapacityTotal.toFixed(1)} kN</td>
                    <td>${summary.yCapacityTotal.toFixed(1)} kN</td>
                </tr>
                <tr style="font-weight:700;font-size:13px;">
                    <td>Ratio</td>
                    <td style="color:${xOk ? '#16a34a' : '#dc2626'};">${xPct}%${xOk ? ' \u2713' : ' \u2717'}</td>
                    <td style="color:${yOk ? '#16a34a' : '#dc2626'};">${yPct}%${yOk ? ' \u2713' : ' \u2717'}</td>
                </tr>
            </tbody>
        </table>

        <div class="bracing-wall-list">
            <div style="font-weight:600;font-size:10px;margin-bottom:4px;color:#6b7280;">WALL BREAKDOWN (${summary.walls.length} walls)</div>
            ${summary.walls.map(w => {
                const dir = w.direction === 'X' ? '\u2195' : '\u2194';
                const nomTag = w.isNominal ? ' <span style="color:#999;font-size:9px;">NOM</span>' : '';
                return `<div class="bracing-wall-item" onclick="selectedElement=project.elements.find(e=>e.id===${w.el.id});engine.requestRender();showProperties();">
                    <span style="color:${(project.scheduleTypes.bracingWall?.[w.typeRef]||{}).color||'#000000'};font-weight:700;">${w.typeRef}</span>
                    <span>${dir} ${(w.lengthMm/1000).toFixed(2)}m</span>
                    <span style="font-weight:600;">${w.capacity.toFixed(1)} kN${nomTag}</span>
                </div>`;
            }).join('')}
        </div>

        ${renderComplianceSection()}

        <div style="margin-top:8px;font-size:9px;color:#999;text-align:center;">
            AS 1684.2-2010 Clause 8.3.6 | Phase A+B+D
        </div>
    `;
}

// Toggle bracing panel visibility
function toggleBracingPanel() {
    const panel = document.getElementById('bracing-summary-panel');
    if (!panel) return;
    panel.classList.toggle('hidden');
    if (!panel.classList.contains('hidden')) {
        updateBracingSummaryPanel();
    }
}

// ══════════════════════════════════════════════════════════════
// PHASE D: Compliance Checks — AS 1684.2-2010 Clause 8.3.6.6
// ══════════════════════════════════════════════════════════════

/**
 * Max spacing between bracing walls per wind class.
 * N1/N2: flat 9000mm (Clause 8.3.6.6).
 * N3: Table 8.20, N4: Table 8.21 — simplified lookup by capacity.
 */
function getMaxSpacing(windClass, capacityKnPerM) {
    if (windClass === 'N1' || windClass === 'N2') return 9000;
    if (windClass === 'N3') {
        if (capacityKnPerM >= 3.4) return 9000;
        if (capacityKnPerM >= 2.0) return 7200;
        if (capacityKnPerM >= 1.0) return 5400;
        return 3600;
    }
    if (windClass === 'N4') {
        if (capacityKnPerM >= 6.0) return 9000;
        if (capacityKnPerM >= 3.4) return 7200;
        if (capacityKnPerM >= 2.0) return 5400;
        return 3600;
    }
    return 9000;
}

/**
 * Minimum distance from point p to line segment (el.x1,y1)→(el.x2,y2).
 */
function pointToSegmentDist(p, el) {
    const dx = el.x2 - el.x1;
    const dy = el.y2 - el.y1;
    const lenSq = dx * dx + dy * dy;
    if (lenSq < 0.01) return Math.sqrt(Math.pow(p.x - el.x1, 2) + Math.pow(p.y - el.y1, 2));
    let t = ((p.x - el.x1) * dx + (p.y - el.y1) * dy) / lenSq;
    t = Math.max(0, Math.min(1, t));
    const cx = el.x1 + t * dx, cy = el.y1 + t * dy;
    return Math.sqrt(Math.pow(p.x - cx, 2) + Math.pow(p.y - cy, 2));
}

function wallMidpoint(el) {
    return { x: (el.x1 + el.x2) / 2, y: (el.y1 + el.y2) / 2 };
}

/**
 * Auto-detect building extent.
 * Priority: 1) Building envelope polygon, 2) Regular walls, 3) Bracing walls + margin.
 */
function detectBuildingExtent() {
    // 1. Prefer drawn/auto-detected building envelope
    const envelope = findEnvelopeElement();
    if (envelope) {
        const bbox = envelopeBBox(envelope);
        if (bbox) return bbox;
    }
    // 2. Regular wall elements
    const walls = project.elements.filter(el => el.type === 'wall');
    if (walls.length >= 2) {
        let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
        for (const w of walls) {
            minX = Math.min(minX, w.x1, w.x2);
            maxX = Math.max(maxX, w.x1, w.x2);
            minY = Math.min(minY, w.y1, w.y2);
            maxY = Math.max(maxY, w.y1, w.y2);
        }
        return { minX, maxX, minY, maxY };
    }
    // 3. Bracing walls fallback
    const bw = project.elements.filter(el => el.type === 'bracingWall');
    if (bw.length === 0) return null;
    let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
    for (const w of bw) {
        minX = Math.min(minX, w.x1, w.x2);
        maxX = Math.max(maxX, w.x1, w.x2);
        minY = Math.min(minY, w.y1, w.y2);
        maxY = Math.max(maxY, w.y1, w.y2);
    }
    const m = 500;
    return { minX: minX - m, maxX: maxX + m, minY: minY - m, maxY: maxY + m };
}

/**
 * Check spacing compliance for one direction.
 */
function checkSpacing(walls, direction, windClass, buildingExtent) {
    if (walls.length === 0) return [];
    const axis = direction === 'X' ? 'x' : 'y';
    const typeDataMap = project.scheduleTypes.bracingWall || {};

    const sorted = walls.map(el => {
        const mid = wallMidpoint(el);
        const td = typeDataMap[el.typeRef || 'BR1'] || {};
        const bt = td.bracingType || 'g';
        return { el, pos: mid[axis], capPerM: BRACING_TYPES[bt]?.capacity || 0 };
    }).sort((a, b) => a.pos - b.pos);

    const results = [];

    if (buildingExtent) {
        const edgeMin = direction === 'X' ? buildingExtent.minX : buildingExtent.minY;
        const gap = sorted[0].pos - edgeMin;
        const max = getMaxSpacing(windClass, sorted[0].capPerM);
        results.push({ wallA: null, wallB: sorted[0].el, gap: Math.round(gap), maxAllowed: max, ok: gap <= max });
    }
    for (let i = 0; i < sorted.length - 1; i++) {
        const gap = sorted[i + 1].pos - sorted[i].pos;
        const weaker = Math.min(sorted[i].capPerM, sorted[i + 1].capPerM);
        const max = getMaxSpacing(windClass, weaker);
        results.push({ wallA: sorted[i].el, wallB: sorted[i + 1].el, gap: Math.round(gap), maxAllowed: max, ok: gap <= max });
    }
    if (buildingExtent) {
        const edgeMax = direction === 'X' ? buildingExtent.maxX : buildingExtent.maxY;
        const last = sorted[sorted.length - 1];
        const gap = edgeMax - last.pos;
        const max = getMaxSpacing(windClass, last.capPerM);
        results.push({ wallA: last.el, wallB: null, gap: Math.round(gap), maxAllowed: max, ok: gap <= max });
    }
    return results;
}

/**
 * Check corner bracing — Clause 8.3.6.6.
 */
function checkCornerBracing(walls, buildingExtent, threshold) {
    if (!buildingExtent) return [];
    const { minX, maxX, minY, maxY } = buildingExtent;
    const corners = [
        { x: minX, y: minY, label: 'BL' },
        { x: maxX, y: minY, label: 'BR' },
        { x: maxX, y: maxY, label: 'TR' },
        { x: minX, y: maxY, label: 'TL' },
    ];
    return corners.map(corner => {
        let nd = Infinity, nw = null;
        for (const el of walls) {
            const d = pointToSegmentDist(corner, el);
            if (d < nd) { nd = d; nw = el; }
        }
        return { corner, nearestWall: nw, distance: Math.round(nd), threshold, ok: nd <= threshold };
    });
}

/**
 * Check proportioning — even distribution per Clause 8.3.6.6.
 */
function checkProportioning(walls, direction, buildingExtent) {
    const minRatio = 0.33;
    if (walls.length === 0 || !buildingExtent) return { halfACap: 0, halfBCap: 0, totalCap: 0, ratio: 0, ok: true };
    const axis = direction === 'X' ? 'x' : 'y';
    const midLine = direction === 'X'
        ? (buildingExtent.minX + buildingExtent.maxX) / 2
        : (buildingExtent.minY + buildingExtent.maxY) / 2;
    const typeDataMap = project.scheduleTypes.bracingWall || {};
    const ceilingHeight = typeof bracingSettings !== 'undefined' ? bracingSettings.ceilingHeight : 2700;
    const jointGroup = typeof bracingSettings !== 'undefined' ? bracingSettings.jointGroup : 'JD6';
    let halfACap = 0, halfBCap = 0;
    for (const el of walls) {
        const mid = wallMidpoint(el);
        const td = typeDataMap[el.typeRef || 'BR1'] || {};
        const cap = calcBracingCapacity(el, td, ceilingHeight, jointGroup);
        if (mid[axis] <= midLine) halfACap += cap; else halfBCap += cap;
    }
    const totalCap = halfACap + halfBCap;
    if (totalCap === 0) return { halfACap: 0, halfBCap: 0, totalCap: 0, ratio: 0, ok: true };
    const ratio = Math.min(halfACap, halfBCap) / totalCap;
    return {
        halfACap: Math.round(halfACap * 10) / 10,
        halfBCap: Math.round(halfBCap * 10) / 10,
        totalCap: Math.round(totalCap * 10) / 10,
        ratio: Math.round(ratio * 100) / 100,
        minRatio, ok: ratio >= minRatio
    };
}

// ── Compliance Cache ────────────────────────────────────────
let _complianceCache = null;
let _complianceDirty = true;

function markComplianceDirty() { _complianceDirty = true; }

function getCachedCompliance() {
    if (_complianceDirty || !_complianceCache) {
        _complianceCache = runComplianceChecks();
        _complianceDirty = false;
    }
    return _complianceCache;
}

/**
 * Run all compliance checks. Returns unified result.
 */
function runComplianceChecks() {
    const currentLevel = typeof levelSystem !== 'undefined' ? levelSystem.currentLevel : 'GF';
    const windClass = bracingSettings.windClass || 'N2';
    const bracingWalls = project.elements.filter(el =>
        el.type === 'bracingWall' && (!el.level || el.level === currentLevel)
    );
    const buildingExtent = detectBuildingExtent();

    const xWalls = bracingWalls.filter(el => getBracingDirection(el) === 'X');
    const yWalls = bracingWalls.filter(el => getBracingDirection(el) === 'Y');

    // Only run corner/proportioning checks when there are enough bracing walls to be meaningful
    const enoughWalls = bracingWalls.length >= 2;

    const result = {
        hasExtent: !!buildingExtent,
        buildingExtent,
        xWallCount: xWalls.length,
        yWallCount: yWalls.length,
        spacing: { x: checkSpacing(xWalls, 'X', windClass, buildingExtent), y: checkSpacing(yWalls, 'Y', windClass, buildingExtent) },
        corners: enoughWalls ? checkCornerBracing(bracingWalls, buildingExtent, 4500) : [],
        proportioning: {
            x: (buildingExtent && enoughWalls) ? checkProportioning(xWalls, 'X', buildingExtent) : null,
            y: (buildingExtent && enoughWalls) ? checkProportioning(yWalls, 'Y', buildingExtent) : null
        },
        overallOk: true,
        issues: [],
    };

    // Collect failures
    for (const s of [...result.spacing.x, ...result.spacing.y]) {
        if (!s.ok) {
            result.overallOk = false;
            const dir = result.spacing.x.includes(s) ? 'X' : 'Y';
            result.issues.push({ type: 'spacing', dir, gap: s.gap, max: s.maxAllowed, wallA: s.wallA, wallB: s.wallB });
        }
    }
    for (const c of result.corners) {
        if (!c.ok && isFinite(c.distance)) {
            result.overallOk = false;
            result.issues.push({ type: 'corner', label: c.corner.label, distance: c.distance, threshold: c.threshold, corner: c.corner });
        }
    }
    if (result.proportioning.x && !result.proportioning.x.ok) {
        result.overallOk = false;
        result.issues.push({ type: 'proportioning', dir: 'X', ratio: result.proportioning.x.ratio });
    }
    if (result.proportioning.y && !result.proportioning.y.ok) {
        result.overallOk = false;
        result.issues.push({ type: 'proportioning', dir: 'Y', ratio: result.proportioning.y.ratio });
    }

    return result;
}


// ── Compliance Overlay Rendering ─────────────────────────────

/**
 * Draw compliance indicators on the canvas:
 *   - Red dashed lines between walls with spacing violations
 *   - Red corner markers where corner bracing is missing
 *   - Amber midline indicator for proportioning failures
 */
function drawComplianceOverlay(ctx, eng) {
    // Only draw when bracing panel is visible
    const panel = document.getElementById('bracing-summary-panel');
    if (!panel || panel.classList.contains('hidden')) return;

    const compliance = getCachedCompliance();
    if (compliance.overallOk) return; // all good, no overlay needed

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    ctx.save();

    // ── Spacing violations: red dashed line between midpoints ──
    const spacingIssues = compliance.issues.filter(i => i.type === 'spacing');
    for (const issue of spacingIssues) {
        let pA, pB;
        if (issue.wallA && issue.wallB) {
            const midA = wallMidpoint(issue.wallA);
            const midB = wallMidpoint(issue.wallB);
            pA = coords.realToScreen(midA.x, midA.y);
            pB = coords.realToScreen(midB.x, midB.y);
        } else if (issue.wallA && compliance.buildingExtent) {
            // Wall to edge
            const mid = wallMidpoint(issue.wallA);
            const dir = issue.dir;
            const ext = compliance.buildingExtent;
            pA = coords.realToScreen(mid.x, mid.y);
            pB = coords.realToScreen(
                dir === 'X' ? ext.maxX : mid.x,
                dir === 'Y' ? ext.maxY : mid.y
            );
        } else if (issue.wallB && compliance.buildingExtent) {
            const mid = wallMidpoint(issue.wallB);
            const dir = issue.dir;
            const ext = compliance.buildingExtent;
            pA = coords.realToScreen(
                dir === 'X' ? ext.minX : mid.x,
                dir === 'Y' ? ext.minY : mid.y
            );
            pB = coords.realToScreen(mid.x, mid.y);
        }
        if (pA && pB) {
            ctx.strokeStyle = '#dc2626';
            ctx.lineWidth = Math.max(1, 1.5 * zoom);
            ctx.setLineDash([6, 4]);
            ctx.globalAlpha = 0.7;
            ctx.beginPath();
            ctx.moveTo(pA.x, pA.y);
            ctx.lineTo(pB.x, pB.y);
            ctx.stroke();

            // Gap label
            const mx = (pA.x + pB.x) / 2, my = (pA.y + pB.y) / 2;
            const fontSize = Math.max(8, 2 * zoom);
            ctx.setLineDash([]);
            ctx.globalAlpha = 0.85;
            ctx.font = `bold ${fontSize}px "Segoe UI", Arial, sans-serif`;
            ctx.fillStyle = '#dc2626';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            ctx.fillText(`${(issue.gap / 1000).toFixed(1)}m > ${(issue.max / 1000).toFixed(1)}m`, mx, my - 4);
        }
    }

    // ── Corner violations: red circle at corner ──
    const cornerIssues = compliance.issues.filter(i => i.type === 'corner');
    for (const issue of cornerIssues) {
        const sp = coords.realToScreen(issue.corner.x, issue.corner.y);
        const r = Math.max(6, 4 * zoom);

        ctx.globalAlpha = 0.6;
        ctx.fillStyle = '#dc2626';
        ctx.beginPath();
        ctx.arc(sp.x, sp.y, r, 0, Math.PI * 2);
        ctx.fill();

        // X mark
        ctx.strokeStyle = '#ffffff';
        ctx.lineWidth = Math.max(1, 0.3 * zoom);
        ctx.setLineDash([]);
        ctx.globalAlpha = 0.9;
        const hr = r * 0.5;
        ctx.beginPath();
        ctx.moveTo(sp.x - hr, sp.y - hr); ctx.lineTo(sp.x + hr, sp.y + hr); ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(sp.x + hr, sp.y - hr); ctx.lineTo(sp.x - hr, sp.y + hr); ctx.stroke();
    }

    // ── Proportioning violations: amber midline ──
    const propIssues = compliance.issues.filter(i => i.type === 'proportioning');
    if (propIssues.length > 0 && compliance.buildingExtent) {
        const ext = compliance.buildingExtent;
        for (const issue of propIssues) {
            ctx.strokeStyle = '#f59e0b';
            ctx.lineWidth = Math.max(1, 1 * zoom);
            ctx.setLineDash([8, 4]);
            ctx.globalAlpha = 0.5;

            if (issue.dir === 'X') {
                const midX = (ext.minX + ext.maxX) / 2;
                const pTop = coords.realToScreen(midX, ext.minY);
                const pBot = coords.realToScreen(midX, ext.maxY);
                ctx.beginPath(); ctx.moveTo(pTop.x, pTop.y); ctx.lineTo(pBot.x, pBot.y); ctx.stroke();
            } else {
                const midY = (ext.minY + ext.maxY) / 2;
                const pLeft = coords.realToScreen(ext.minX, midY);
                const pRight = coords.realToScreen(ext.maxX, midY);
                ctx.beginPath(); ctx.moveTo(pLeft.x, pLeft.y); ctx.lineTo(pRight.x, pRight.y); ctx.stroke();
            }

            // Label
            const fontSize = Math.max(7, 1.8 * zoom);
            ctx.globalAlpha = 0.7;
            ctx.setLineDash([]);
            ctx.font = `${fontSize}px "Segoe UI", Arial, sans-serif`;
            ctx.fillStyle = '#f59e0b';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            if (issue.dir === 'X') {
                const midX = (ext.minX + ext.maxX) / 2;
                const p = coords.realToScreen(midX, ext.minY);
                ctx.fillText(`${issue.dir}-dir: ${(issue.ratio * 100).toFixed(0)}% imbalance`, p.x, p.y - 4);
            } else {
                const midY = (ext.minY + ext.maxY) / 2;
                const p = coords.realToScreen(ext.minX, midY);
                ctx.fillText(`${issue.dir}-dir: ${(issue.ratio * 100).toFixed(0)}% imbalance`, p.x + 40, p.y - 4);
            }
        }
    }

    ctx.restore();
}

// Register compliance overlay renderer
engine.onRender(drawComplianceOverlay);


// Expose globally
window.updateBracingSummaryPanel = updateBracingSummaryPanel;
window.toggleBracingPanel = toggleBracingPanel;
window.calcBracingSummary = calcBracingSummary;
window.calcBracingCapacity = calcBracingCapacity;
window.runComplianceChecks = runComplianceChecks;
window.getCachedCompliance = getCachedCompliance;
window.markComplianceDirty = markComplianceDirty;
window.bracingSettings = bracingSettings;
// Phase B exports
window.lookupWindPressure = lookupWindPressure;
window.calcAreaOfElevation = calcAreaOfElevation;
window.calcWindDemand = calcWindDemand;
window.autoDetectEnvelope = autoDetectEnvelope;
window.findEnvelopeElement = findEnvelopeElement;
window.envelopeBBox = envelopeBBox;
window.envelopeArea = envelopeArea;
// Ridge line exports
window.findRidgeElement = findRidgeElement;
window.startDrawRidge = startDrawRidge;
