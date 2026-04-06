// ══════════════════════════════════════════════════════════
// ── 25 · ANALYSIS MODEL EXPORT (SpaceGass + DXF) ────────
// ══════════════════════════════════════════════════════════
//
// Extracts a structural stick model from the StructuralSketch
// project data (nodes, members, sections, materials, supports)
// and exports in formats suitable for structural analysis:
//
//   • SpaceGass text input (.txt)
//   • DXF stick model (.dxf) — universal, works with
//     SpaceGass / Microstran / Strand7 / ETABS / SAP2000
//
// The extraction logic mirrors rebuild3DScene() in 11-3d-engine.js
// but outputs analytical data instead of Three.js meshes.
//
// ══════════════════════════════════════════════════════════


// ──────────────────────────────────────────────────────────
// §1  CONFIGURATION & CONSTANTS
// ──────────────────────────────────────────────────────────

const ANALYSIS_EXPORT = {
    // Node merge tolerance (mm) — points within this distance
    // are treated as the same node. Grid snap means most will
    // be exact, but this catches rounding.
    MERGE_TOL: 5.0,

    // Material property database (AS-referenced values)
    MATERIALS: {
        STEEL: {
            id: 1,
            name: 'Steel - Grade 300',
            E: 200000,       // MPa  (AS 4100 Table 2.1)
            G: 80000,        // MPa
            density: 7850,   // kg/m³
            fy: 300,         // MPa  (AS 4100 Table 2.1, t ≤ 17mm)
            poisson: 0.3,
            alpha: 11.7e-6,  // /°C
        },
        TIMBER: {
            id: 2,
            name: 'Timber - GL17',
            E: 16300,        // MPa  (AS 1720.1 Table H3.1 — GL17)
            G: 1080,         // MPa
            density: 600,    // kg/m³ (typical GLT)
            poisson: 0.35,
            alpha: 5.0e-6,
        },
        TIMBER_MGP10: {
            id: 3,
            name: 'Timber - MGP10',
            E: 10000,        // MPa  (AS 1720.1)
            G: 660,
            density: 550,
            poisson: 0.35,
            alpha: 5.0e-6,
        },
        CONCRETE: {
            id: 4,
            name: 'Concrete - 40 MPa',
            E: 32800,        // MPa  (AS 3600 Cl 3.1.2 — Ecj)
            G: 13700,
            density: 2400,   // kg/m³
            poisson: 0.2,
            alpha: 10.0e-6,
        },
    },

    // Default member end release:
    //   'F' = Fixed, 'P' = Pinned (moment released)
    // For steel beams with simple connections (web cleats, fin plates),
    // default to pinned ends. Columns default to fixed-fixed.
    DEFAULT_BEAM_RELEASE: { startMz: 'P', endMz: 'P' },  // pin-pin about major axis
    DEFAULT_COL_RELEASE:  { startMz: 'F', endMz: 'F' },  // fixed-fixed
};


// ──────────────────────────────────────────────────────────
// §2  NODE EXTRACTION & MERGING
// ──────────────────────────────────────────────────────────

/**
 * Extracts all structural nodes from project elements.
 * Merges coincident nodes within MERGE_TOL.
 *
 * Returns: {
 *   nodes: [ { id, x, y, z, label, gridRef } ],
 *   nodeMap: Map<string, nodeId>  // "x,y,z" → id for lookups
 * }
 *
 * Coordinate convention (matches SpaceGass):
 *   X = plan horizontal (StructuralSketch X)
 *   Y = plan vertical   (StructuralSketch Y — note: screen Y, not elevation)
 *   Z = elevation (up)
 *
 * All coordinates in metres.
 */
function extractAnalysisNodes() {
    const tol = ANALYSIS_EXPORT.MERGE_TOL;
    const nodes = [];
    const nodeMap = new Map(); // "roundedX,roundedY,roundedZ" → nodeId

    // Helper: find or create a node at (x, y, z) in mm
    function getOrCreateNode(xMM, yMM, zMM) {
        // Round to tolerance grid for merging
        const rx = Math.round(xMM / tol) * tol;
        const ry = Math.round(yMM / tol) * tol;
        const rz = Math.round(zMM / tol) * tol;
        const key = `${rx},${ry},${rz}`;

        if (nodeMap.has(key)) {
            return nodeMap.get(key);
        }

        const id = nodes.length + 1; // 1-indexed for analysis software
        const node = {
            id: id,
            x: rx / 1000,  // Convert mm → m
            y: ry / 1000,
            z: rz / 1000,
            label: '',
            gridRef: '',
            isSupport: false,
        };
        nodes.push(node);
        nodeMap.set(key, id);
        return id;
    }

    // ── Walk all visible structural elements ──
    const visibleElements = project.elements.filter(el => {
        const layer = project.layers[el.layer];
        return layer && layer.visible;
    });

    for (const el of visibleElements) {
        const lv = levelSystem.levels.find(l => l.id === el.level) || getActiveLevel();
        const elev = lv.elevation; // mm

        // ── BEAMS → 2 nodes per beam ──
        if (el.type === 'line' && el.layer === 'S-BEAM') {
            getOrCreateNode(el.x1, el.y1, elev);
            getOrCreateNode(el.x2, el.y2, elev);
        }

        // ── COLUMNS → 2 nodes (top + bottom) ──
        if (el.type === 'column') {
            const lvIdx = levelSystem.levels.findIndex(l => l.id === el.level);
            const dir = el.extends || 'below';
            let bottomElev = elev, topElev = elev;

            if (dir === 'below' && lvIdx > 0) {
                bottomElev = levelSystem.levels[lvIdx - 1].elevation;
                topElev = elev;
            } else if (dir === 'above') {
                bottomElev = elev;
                const nextLv = lvIdx < levelSystem.levels.length - 1
                    ? levelSystem.levels[lvIdx + 1] : lv;
                topElev = nextLv.elevation || (elev + (lv.height || 2700));
            } else if (dir === 'both') {
                bottomElev = lvIdx > 0
                    ? levelSystem.levels[lvIdx - 1].elevation : elev;
                topElev = lvIdx < levelSystem.levels.length - 1
                    ? levelSystem.levels[lvIdx + 1].elevation
                    : elev + (lv.height || 2700);
            } else {
                topElev = elev + (lv.height || 2700);
            }

            getOrCreateNode(el.x, el.y, bottomElev);
            getOrCreateNode(el.x, el.y, topElev);
        }
    }

    // ── Label nodes from structural grids ──
    if (typeof structuralGrids !== 'undefined' && structuralGrids.length > 0) {
        const orthoGrids = structuralGrids.filter(g => typeof isOrthoGrid === 'function' ? isOrthoGrid(g) : !g.type || g.type === 'ortho');
        const angledGrids = structuralGrids.filter(g => typeof isAngledGrid === 'function' ? isAngledGrid(g) : g.type === 'angled');
        const vGrids = orthoGrids.filter(g => g.axis === 'V')
            .sort((a, b) => a.position - b.position);
        const hGrids = orthoGrids.filter(g => g.axis === 'H')
            .sort((a, b) => a.position - b.position);

        for (const node of nodes) {
            const xMM = node.x * 1000;
            const yMM = node.y * 1000;

            // Find nearest V grid (matches X position)
            let vLabel = '';
            for (const g of vGrids) {
                if (Math.abs(g.position - xMM) < tol) {
                    vLabel = g.label;
                    break;
                }
            }

            // Find nearest H grid (matches Y position)
            let hLabel = '';
            for (const g of hGrids) {
                if (Math.abs(g.position - yMM) < tol) {
                    hLabel = g.label;
                    break;
                }
            }

            // Check angled grids — find nearest by point-to-line distance
            if (!vLabel && !hLabel) {
                for (const g of angledGrids) {
                    // Use real-world mm coordinates for angled grids
                    const dx = g.x2 - g.x1, dy = g.y2 - g.y1;
                    const lenSq = dx * dx + dy * dy;
                    if (lenSq < 0.01) continue;
                    let t = ((xMM - g.x1) * dx + (yMM - g.y1) * dy) / lenSq;
                    t = Math.max(0, Math.min(1, t));
                    const nx = g.x1 + t * dx, ny = g.y1 + t * dy;
                    const dist = Math.sqrt((xMM - nx) ** 2 + (yMM - ny) ** 2);
                    if (dist < tol) {
                        vLabel = g.label; // use as primary label
                        break;
                    }
                }
            }

            if (vLabel && hLabel) {
                node.gridRef = vLabel + '/' + hLabel;
            } else if (vLabel) {
                node.gridRef = vLabel;
            } else if (hLabel) {
                node.gridRef = hLabel;
            }

            // Build human-readable label: "A/1-L01"
            // Find which level this elevation belongs to
            let lvName = '';
            for (const lv of levelSystem.levels) {
                if (Math.abs(lv.elevation - node.z * 1000) < tol) {
                    lvName = lv.name;
                    break;
                }
            }

            if (node.gridRef && lvName) {
                node.label = node.gridRef + '-' + lvName;
            } else if (node.gridRef) {
                node.label = node.gridRef + '-Z' + node.z.toFixed(1);
            } else {
                node.label = 'N' + node.id;
            }
        }
    } else {
        // No grids — simple numbering
        for (const node of nodes) {
            node.label = 'N' + node.id;
        }
    }

    return { nodes, nodeMap, getOrCreateNode };
}


// ──────────────────────────────────────────────────────────
// §3  MEMBER GENERATION
// ──────────────────────────────────────────────────────────

/**
 * Generates analysis members from project elements.
 * Each member references start/end node IDs and carries
 * section and material data.
 *
 * Returns: [ {
 *   id, type, nodeI, nodeJ,
 *   sectionName, sectionType, sectionData,
 *   materialKey, tag,
 *   releaseI, releaseJ
 * } ]
 */
function extractAnalysisMembers(nodeResult) {
    const { nodes, nodeMap, getOrCreateNode } = nodeResult;
    const tol = ANALYSIS_EXPORT.MERGE_TOL;
    const members = [];
    let memId = 1;

    const visibleElements = project.elements.filter(el => {
        const layer = project.layers[el.layer];
        return layer && layer.visible;
    });

    for (const el of visibleElements) {
        const lv = levelSystem.levels.find(l => l.id === el.level) || getActiveLevel();
        const elev = lv.elevation;

        // ── BEAMS ──
        if (el.type === 'line' && el.layer === 'S-BEAM') {
            const nodeI = getOrCreateNode(el.x1, el.y1, elev);
            const nodeJ = getOrCreateNode(el.x2, el.y2, elev);

            if (nodeI === nodeJ) continue; // zero-length member

            const parsed = parseMemberFor3D(el.memberSize, el.memberCategory);
            const matKey = inferMaterialKey(el, parsed);

            members.push({
                id: memId++,
                type: 'beam',
                nodeI: nodeI,
                nodeJ: nodeJ,
                sectionName: el.memberSize || 'Unknown',
                sectionType: parsed ? parsed.type : 'RECT',
                sectionData: parsed,
                materialKey: matKey,
                tag: el.tag || el.smartTag || ('B' + memId),
                releaseI: ANALYSIS_EXPORT.DEFAULT_BEAM_RELEASE.startMz,
                releaseJ: ANALYSIS_EXPORT.DEFAULT_BEAM_RELEASE.endMz,
                sourceElement: el,
            });
        }

        // ── COLUMNS ──
        if (el.type === 'column') {
            const lvIdx = levelSystem.levels.findIndex(l => l.id === el.level);
            const dir = el.extends || 'below';
            let bottomElev = elev, topElev = elev;

            if (dir === 'below' && lvIdx > 0) {
                bottomElev = levelSystem.levels[lvIdx - 1].elevation;
                topElev = elev;
            } else if (dir === 'above') {
                bottomElev = elev;
                const nextLv = lvIdx < levelSystem.levels.length - 1
                    ? levelSystem.levels[lvIdx + 1] : lv;
                topElev = nextLv.elevation || (elev + (lv.height || 2700));
            } else if (dir === 'both') {
                bottomElev = lvIdx > 0
                    ? levelSystem.levels[lvIdx - 1].elevation : elev;
                topElev = lvIdx < levelSystem.levels.length - 1
                    ? levelSystem.levels[lvIdx + 1].elevation
                    : elev + (lv.height || 2700);
            } else {
                topElev = elev + (lv.height || 2700);
            }

            const nodeI = getOrCreateNode(el.x, el.y, bottomElev);
            const nodeJ = getOrCreateNode(el.x, el.y, topElev);

            if (nodeI === nodeJ) continue;

            const parsed = parseMemberFor3D(el.memberSize, el.memberCategory);
            const matKey = inferMaterialKey(el, parsed);

            members.push({
                id: memId++,
                type: 'column',
                nodeI: nodeI,  // bottom node
                nodeJ: nodeJ,  // top node
                sectionName: el.memberSize || 'Unknown',
                sectionType: parsed ? parsed.type : 'RECT',
                sectionData: parsed,
                materialKey: matKey,
                tag: el.tag || el.smartTag || ('C' + memId),
                releaseI: ANALYSIS_EXPORT.DEFAULT_COL_RELEASE.startMz,
                releaseJ: ANALYSIS_EXPORT.DEFAULT_COL_RELEASE.endMz,
                sourceElement: el,
            });
        }
    }

    return members;
}


/**
 * Infer material key from element properties and parsed section data.
 */
function inferMaterialKey(el, parsed) {
    // Explicit material on element
    if (el.material === 'timber' || el.material === 'GLT') return 'TIMBER';
    if (el.material === 'concrete') return 'CONCRETE';
    if (el.material === 'steel') return 'STEEL';

    // Infer from section parsing
    if (parsed) {
        if (parsed.material === 'timber') return 'TIMBER';
        if (parsed.type === 'RECT' && parsed.material === 'timber') return 'TIMBER';
    }

    // Infer from member category
    if (el.memberCategory === 'TIMBER' || el.memberCategory === 'GLT') return 'TIMBER';
    if (/MGP|GL\d|LVL/i.test(el.memberSize || '')) return 'TIMBER';

    // Default: steel
    return 'STEEL';
}


// ──────────────────────────────────────────────────────────
// §4  SUPPORT CONDITIONS
// ──────────────────────────────────────────────────────────

/**
 * Identify support nodes — base of lowest-level columns.
 * Returns array of { nodeId, type, restraints }
 *
 * Default: pinned base (Fx, Fy, Fz restrained; Mx, My, Mz free)
 * The engineer can refine to fixed in SpaceGass if needed.
 *
 * Restraint code: 6-digit string, each digit = 1 (restrained) or 0 (free)
 *   Position: Fx, Fy, Fz, Mx, My, Mz
 *   Pinned = "111000"
 *   Fixed  = "111111"
 */
function identifySupports(nodes, members) {
    const supports = [];

    // Find the lowest elevation in the model
    let minElev = Infinity;
    for (const n of nodes) {
        if (n.z < minElev) minElev = n.z;
    }

    // Mark bottom nodes of columns at lowest elevation as supports
    for (const mem of members) {
        if (mem.type !== 'column') continue;
        const botNode = nodes.find(n => n.id === mem.nodeI);
        if (!botNode) continue;

        // Support if this is at or near the lowest elevation
        if (Math.abs(botNode.z - minElev) < 0.01) {
            // Check not already added
            if (!supports.find(s => s.nodeId === botNode.id)) {
                botNode.isSupport = true;
                supports.push({
                    nodeId: botNode.id,
                    type: 'pinned',       // default assumption
                    restraints: '111000',  // Fx, Fy, Fz fixed; Mx, My, Mz free
                });
            }
        }
    }

    return supports;
}


// ──────────────────────────────────────────────────────────
// §5  SECTION PROPERTY MAPPING
// ──────────────────────────────────────────────────────────

/**
 * Map StructuralSketch section definitions to SpaceGass-compatible
 * section property descriptions.
 *
 * For standard Australian sections (UB, UC, SHS, RHS, CHS, PFC),
 * the name maps directly to SpaceGass's built-in library.
 *
 * For timber rectangular sections, we compute and export:
 *   A, Ix, Iy, Zx, Zy, J
 */
function buildSectionTable(members) {
    const sections = new Map(); // sectionName → section definition

    for (const mem of members) {
        if (sections.has(mem.sectionName)) continue;

        const sec = {
            name: mem.sectionName,
            type: mem.sectionType,
            libraryMatch: false,
            properties: null,  // computed for non-library sections
        };

        if (mem.sectionData) {
            const sd = mem.sectionData;

            if (sd.type === 'UB' || sd.type === 'PFC') {
                // Standard Australian section — maps directly to SpaceGass library
                sec.libraryMatch = true;
                // Still compute properties for DXF and validation
                const d = sd.data;
                sec.properties = {
                    d: d.d, bf: d.bf, tf: d.tf, tw: d.tw,
                    A: approxIBeamArea(d),
                    Ix: approxIBeamIx(d),
                    Iy: approxIBeamIy(d),
                };
            } else if (sd.type === 'SHS') {
                sec.libraryMatch = true;
                const B = sd.B, t = sd.t;
                sec.properties = {
                    B: B, D: B, t: t,
                    A: (B * B) - (B - 2 * t) * (B - 2 * t),
                    Ix: (B ** 4 - (B - 2 * t) ** 4) / 12,
                    Iy: (B ** 4 - (B - 2 * t) ** 4) / 12,
                };
            } else if (sd.type === 'RHS') {
                sec.libraryMatch = true;
                const B = sd.B, D = sd.D, t = sd.t;
                sec.properties = {
                    B: B, D: D, t: t,
                    A: (B * D) - (B - 2 * t) * (D - 2 * t),
                    Ix: (B * D ** 3 - (B - 2 * t) * (D - 2 * t) ** 3) / 12,
                    Iy: (D * B ** 3 - (D - 2 * t) * (B - 2 * t) ** 3) / 12,
                };
            } else if (sd.type === 'CHS') {
                sec.libraryMatch = true;
                const D = sd.D, t = sd.t;
                const ro = D / 2, ri = ro - t;
                sec.properties = {
                    D: D, t: t,
                    A: Math.PI * (ro * ro - ri * ri),
                    Ix: Math.PI / 4 * (ro ** 4 - ri ** 4),
                    Iy: Math.PI / 4 * (ro ** 4 - ri ** 4),
                };
            } else if (sd.type === 'RECT') {
                // Timber / generic rectangular section — must define fully
                sec.libraryMatch = false;
                const b = sd.width, d = sd.depth;
                sec.properties = {
                    b: b, d: d,
                    A: b * d,
                    Ix: b * d ** 3 / 12,
                    Iy: d * b ** 3 / 12,
                    Zx: b * d ** 2 / 6,
                    Zy: d * b ** 2 / 6,
                    J: rectTorsionJ(b, d),
                };
            }
        }

        sections.set(mem.sectionName, sec);
    }

    return sections;
}

// ── Approximate section property helpers ──

function approxIBeamArea(d) {
    // A ≈ 2·bf·tf + (d - 2·tf)·tw
    return 2 * d.bf * d.tf + (d.d - 2 * d.tf) * d.tw;
}

function approxIBeamIx(d) {
    // Ix ≈ bf·d³/12 - (bf-tw)·(d-2tf)³/12
    const dTot = d.d, bfv = d.bf, tfv = d.tf, twv = d.tw;
    return (bfv * dTot ** 3 - (bfv - twv) * (dTot - 2 * tfv) ** 3) / 12;
}

function approxIBeamIy(d) {
    // Iy ≈ 2·(tf·bf³/12) + (d-2tf)·tw³/12
    const dTot = d.d, bfv = d.bf, tfv = d.tf, twv = d.tw;
    return 2 * (tfv * bfv ** 3 / 12) + (dTot - 2 * tfv) * twv ** 3 / 12;
}

function rectTorsionJ(b, d) {
    // Torsion constant for rectangle: J ≈ β·b·d³
    // where β depends on b/d ratio (Timoshenko)
    const ratio = Math.max(b, d) / Math.min(b, d);
    let beta;
    if (ratio <= 1.0) beta = 0.141;
    else if (ratio <= 1.5) beta = 0.196;
    else if (ratio <= 2.0) beta = 0.229;
    else if (ratio <= 3.0) beta = 0.263;
    else if (ratio <= 5.0) beta = 0.291;
    else beta = 0.312;

    const a = Math.max(b, d), bv = Math.min(b, d);
    return beta * bv * a ** 3;
}


// ──────────────────────────────────────────────────────────
// §6  SELF-WEIGHT LOAD CASE
// ──────────────────────────────────────────────────────────

/**
 * Generate self-weight (G1) load case data.
 * SpaceGass handles self-weight natively via a gravity multiplier,
 * but we document the case for completeness.
 */
function buildSelfWeightCase() {
    return {
        id: 1,
        name: 'G1 - Self Weight',
        type: 'dead',
        gravity: { x: 0, y: 0, z: -1.0 },  // -1g in Z direction
    };
}


// ──────────────────────────────────────────────────────────
// §7  COMPLETE MODEL EXTRACTION
// ──────────────────────────────────────────────────────────

/**
 * Extract the complete analysis model from StructuralSketch data.
 * Returns a clean data object that format writers consume.
 */
function extractAnalysisModel() {
    const nodeResult = extractAnalysisNodes();
    const members = extractAnalysisMembers(nodeResult);
    const supports = identifySupports(nodeResult.nodes, members);
    const sections = buildSectionTable(members);
    const loadCases = [buildSelfWeightCase()];

    // Collect unique materials actually used
    const usedMaterials = new Map();
    for (const mem of members) {
        const key = mem.materialKey;
        if (!usedMaterials.has(key) && ANALYSIS_EXPORT.MATERIALS[key]) {
            usedMaterials.set(key, ANALYSIS_EXPORT.MATERIALS[key]);
        }
    }

    // Summary stats
    const stats = {
        nodeCount: nodeResult.nodes.length,
        memberCount: members.length,
        beamCount: members.filter(m => m.type === 'beam').length,
        columnCount: members.filter(m => m.type === 'column').length,
        supportCount: supports.length,
        sectionCount: sections.size,
        materialCount: usedMaterials.size,
        levelCount: levelSystem.levels.length,
    };

    return {
        nodes: nodeResult.nodes,
        members: members,
        supports: supports,
        sections: sections,
        materials: usedMaterials,
        loadCases: loadCases,
        stats: stats,
        projectName: (project.projectInfo && project.projectInfo.name)
            || 'StructuralSketch Model',
        exportDate: new Date().toISOString().split('T')[0],
    };
}


// ──────────────────────────────────────────────────────────
// §8  SPACEGASS TEXT FORMAT WRITER
// ──────────────────────────────────────────────────────────

/**
 * Generate SpaceGass-compatible text input file.
 *
 * SpaceGass text format uses keyword blocks:
 *   TITLE, NODE, MEMBER, SECTION, MATERIAL, RESTRAINT, LOAD CASE, etc.
 *
 * Reference: SpaceGass User Manual — Text File Import
 */
function writeSpaceGassText(model) {
    const lines = [];
    const pad = (s, w) => String(s).padEnd(w);
    const padr = (s, w) => String(s).padStart(w);

    // ── Header ──
    lines.push('; ══════════════════════════════════════════════════════════');
    lines.push('; SpaceGass Analysis Model');
    lines.push('; Generated by StructuralSketch');
    lines.push('; Date: ' + model.exportDate);
    lines.push('; Project: ' + model.projectName);
    lines.push(';');
    lines.push('; Nodes: ' + model.stats.nodeCount);
    lines.push('; Members: ' + model.stats.memberCount +
        ' (' + model.stats.beamCount + ' beams, ' + model.stats.columnCount + ' columns)');
    lines.push('; Supports: ' + model.stats.supportCount);
    lines.push(';');
    lines.push('; ASSUMPTIONS:');
    lines.push(';   - Beam ends: Pin-Pin (simple connections) — review and adjust');
    lines.push(';   - Column ends: Fixed-Fixed — review and adjust');
    lines.push(';   - Supports: Pinned at column bases — review and adjust');
    lines.push(';   - Self-weight load case included (G1)');
    lines.push(';   - All coordinates in metres');
    lines.push('; ══════════════════════════════════════════════════════════');
    lines.push('');

    // ── Title ──
    lines.push('TITLE');
    lines.push(model.projectName + ' — exported from StructuralSketch');
    lines.push('');

    // ── Nodes ──
    lines.push('; Node    X(m)         Y(m)         Z(m)         Label');
    lines.push('NODE');
    for (const n of model.nodes) {
        lines.push(
            padr(n.id, 6) + '  ' +
            padr(n.x.toFixed(4), 12) + '  ' +
            padr(n.y.toFixed(4), 12) + '  ' +
            padr(n.z.toFixed(4), 12) +
            (n.label ? '  ; ' + n.label : '')
        );
    }
    lines.push('');

    // ── Materials ──
    lines.push('; MatID  E(MPa)       G(MPa)       Density(kg/m3) Poisson  Alpha      Name');
    lines.push('MATERIAL');
    for (const [key, mat] of model.materials) {
        lines.push(
            padr(mat.id, 6) + '  ' +
            padr(mat.E.toFixed(0), 12) + '  ' +
            padr(mat.G.toFixed(0), 12) + '  ' +
            padr(mat.density.toFixed(0), 12) + '  ' +
            padr(mat.poisson.toFixed(2), 8) + '  ' +
            padr(mat.alpha.toExponential(1), 10) + '  ' +
            '; ' + mat.name
        );
    }
    lines.push('');

    // ── Sections ──
    lines.push('; Section definitions');
    lines.push('; Standard sections (UB, SHS, RHS, CHS, PFC) reference SpaceGass library');
    lines.push('; Custom sections (timber rectangles) defined with explicit properties');
    lines.push('SECTION');

    let secId = 1;
    const sectionIdMap = new Map(); // sectionName → secId
    for (const [name, sec] of model.sections) {
        sectionIdMap.set(name, secId);

        if (sec.libraryMatch) {
            // Library section — reference by name
            lines.push(
                padr(secId, 6) + '  LIBRARY  ' + pad(name, 20) +
                '  ; ' + sec.type + ' — from AS section library'
            );
        } else if (sec.properties) {
            // Custom section — define properties
            const p = sec.properties;
            lines.push(
                padr(secId, 6) + '  GENERAL  ' +
                pad(name, 20) + '  ' +
                'A=' + p.A.toFixed(1) + '  ' +
                'Ix=' + p.Ix.toFixed(0) + '  ' +
                'Iy=' + p.Iy.toFixed(0) + '  ' +
                (p.J ? 'J=' + p.J.toFixed(0) + '  ' : '') +
                '; b=' + (p.b || 0) + ' d=' + (p.d || 0) + ' mm'
            );
        } else {
            lines.push(
                padr(secId, 6) + '  LIBRARY  ' + pad(name, 20) +
                '  ; UNRESOLVED — assign manually in SpaceGass'
            );
        }
        secId++;
    }
    lines.push('');

    // ── Members ──
    lines.push('; Mem   NodeI  NodeJ  Section  Material  ReleaseI  ReleaseJ  Tag');
    lines.push('MEMBER');
    for (const mem of model.members) {
        const secIdx = sectionIdMap.get(mem.sectionName) || 1;
        const matId = (ANALYSIS_EXPORT.MATERIALS[mem.materialKey] || {}).id || 1;

        lines.push(
            padr(mem.id, 6) + '  ' +
            padr(mem.nodeI, 6) + '  ' +
            padr(mem.nodeJ, 6) + '  ' +
            padr(secIdx, 8) + '  ' +
            padr(matId, 8) + '  ' +
            pad(mem.releaseI, 4) + '      ' +
            pad(mem.releaseJ, 4) + '      ' +
            '; ' + mem.tag + ' [' + mem.type + ']'
        );
    }
    lines.push('');

    // ── Restraints ──
    lines.push('; Node   Fx  Fy  Fz  Mx  My  Mz');
    lines.push('RESTRAINT');
    for (const sup of model.supports) {
        const r = sup.restraints;
        const node = model.nodes.find(n => n.id === sup.nodeId);
        lines.push(
            padr(sup.nodeId, 6) + '  ' +
            r[0] + '   ' + r[1] + '   ' + r[2] + '   ' +
            r[3] + '   ' + r[4] + '   ' + r[5] +
            '  ; ' + sup.type + (node ? ' @ ' + node.label : '')
        );
    }
    lines.push('');

    // ── Member End Releases ──
    lines.push('; Member releases (P = Pinned, F = Fixed about Mz)');
    lines.push('RELEASE');
    for (const mem of model.members) {
        if (mem.releaseI === 'P' || mem.releaseJ === 'P') {
            lines.push(
                padr(mem.id, 6) + '  ' +
                'I:' + (mem.releaseI === 'P' ? 'Mz' : '--') + '  ' +
                'J:' + (mem.releaseJ === 'P' ? 'Mz' : '--') +
                '  ; ' + mem.tag
            );
        }
    }
    lines.push('');

    // ── Self-Weight Load Case ──
    lines.push('LOAD CASE');
    lines.push('1  G1 - Self Weight  DEAD  GRAVITY  0  0  -9.81');
    lines.push('');

    lines.push('; ── End of SpaceGass Input File ──');
    lines.push('END');

    return lines.join('\n');
}


// ──────────────────────────────────────────────────────────
// §9  DXF STICK MODEL WRITER
// ──────────────────────────────────────────────────────────

/**
 * Generate a DXF file with LINE entities representing the stick model.
 *
 * Convention:
 *   - Each unique section gets its own layer (e.g., "310UB40.4", "89x89x5SHS")
 *   - Beams drawn as lines on section-named layers
 *   - Columns drawn as lines on section-named layers
 *   - Support nodes marked with POINT entities on layer "SUPPORTS"
 *
 * DXF coordinates:
 *   X = plan X (m)
 *   Y = plan Y (m) — note DXF Y = StructuralSketch screen Y
 *   Z = elevation (m)
 *
 * This format is importable by SpaceGass, Microstran, Strand7,
 * ETABS, SAP2000, and most structural analysis packages.
 */
function writeDXFStickModel(model) {
    const lines = [];

    // ── DXF Header ──
    lines.push('0', 'SECTION');
    lines.push('2', 'HEADER');
    lines.push('9', '$ACADVER');
    lines.push('1', 'AC1015');  // AutoCAD 2000 format — wide compatibility
    lines.push('9', '$INSUNITS');
    lines.push('70', '6');       // Units = metres
    lines.push('0', 'ENDSEC');

    // ── Tables Section (layers) ──
    lines.push('0', 'SECTION');
    lines.push('2', 'TABLES');
    lines.push('0', 'TABLE');
    lines.push('2', 'LAYER');

    // Collect unique layer names
    const layerNames = new Set();
    layerNames.add('SUPPORTS');
    layerNames.add('BEAMS');
    layerNames.add('COLUMNS');
    layerNames.add('NODE_LABELS');
    for (const mem of model.members) {
        layerNames.add(sanitiseDXFLayerName(mem.sectionName));
    }

    lines.push('70', String(layerNames.size));

    // Colour assignments
    const layerColours = {
        'SUPPORTS': 1,      // Red
        'BEAMS': 5,          // Blue
        'COLUMNS': 3,        // Green
        'NODE_LABELS': 7,    // White/black
    };
    let colourIdx = 30;

    for (const name of layerNames) {
        lines.push('0', 'LAYER');
        lines.push('2', name);
        lines.push('70', '0');
        lines.push('62', String(layerColours[name] || colourIdx++));
        lines.push('6', 'CONTINUOUS');
    }

    lines.push('0', 'ENDTAB');
    lines.push('0', 'ENDSEC');

    // ── Entities Section ──
    lines.push('0', 'SECTION');
    lines.push('2', 'ENTITIES');

    // Members as LINE entities
    for (const mem of model.members) {
        const ni = model.nodes.find(n => n.id === mem.nodeI);
        const nj = model.nodes.find(n => n.id === mem.nodeJ);
        if (!ni || !nj) continue;

        const layerName = sanitiseDXFLayerName(mem.sectionName);

        lines.push('0', 'LINE');
        lines.push('8', layerName);      // Layer
        // Start point
        lines.push('10', ni.x.toFixed(4));
        lines.push('20', ni.y.toFixed(4));
        lines.push('30', ni.z.toFixed(4));
        // End point
        lines.push('11', nj.x.toFixed(4));
        lines.push('21', nj.y.toFixed(4));
        lines.push('31', nj.z.toFixed(4));
    }

    // Support nodes as POINT entities
    for (const sup of model.supports) {
        const node = model.nodes.find(n => n.id === sup.nodeId);
        if (!node) continue;

        lines.push('0', 'POINT');
        lines.push('8', 'SUPPORTS');
        lines.push('10', node.x.toFixed(4));
        lines.push('20', node.y.toFixed(4));
        lines.push('30', node.z.toFixed(4));
    }

    // Node labels as TEXT entities
    for (const node of model.nodes) {
        if (!node.label) continue;
        lines.push('0', 'TEXT');
        lines.push('8', 'NODE_LABELS');
        lines.push('10', node.x.toFixed(4));
        lines.push('20', node.y.toFixed(4));
        lines.push('30', (node.z + 0.15).toFixed(4));  // Offset slightly above
        lines.push('40', '0.15');  // Text height (m)
        lines.push('1', node.label);
    }

    lines.push('0', 'ENDSEC');
    lines.push('0', 'EOF');

    return lines.join('\n');
}

/**
 * Sanitise a section name for use as a DXF layer name.
 * DXF layer names cannot contain certain characters.
 */
function sanitiseDXFLayerName(name) {
    return (name || 'UNKNOWN')
        .replace(/[<>\/\\\":;?*|=`]/g, '_')
        .replace(/\s+/g, '_')
        .substring(0, 255);
}


// ──────────────────────────────────────────────────────────
// §10  FILE DOWNLOAD HELPER
// ──────────────────────────────────────────────────────────

function downloadTextFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType || 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}


// ──────────────────────────────────────────────────────────
// §11  EXPORT DIALOG UI
// ──────────────────────────────────────────────────────────

/**
 * Show a modal export dialog with model summary, options, and
 * export buttons for SpaceGass and DXF formats.
 */
function showAnalysisExportDialog() {
    // Extract model first to show stats
    let model;
    try {
        model = extractAnalysisModel();
    } catch (e) {
        alert('Error extracting model: ' + e.message);
        console.error('[Export]', e);
        return;
    }

    if (model.stats.memberCount === 0) {
        alert('No structural members found to export.\n\n' +
            'Draw beams and columns on S-BEAM layer to create an analysis model.');
        return;
    }

    // ── Build Modal ──
    const overlay = document.createElement('div');
    overlay.id = 'analysis-export-overlay';
    overlay.style.cssText = `
        position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
        background: rgba(0,0,0,0.5); z-index: 20000;
        display: flex; align-items: center; justify-content: center;
        font-family: "Segoe UI", Arial, sans-serif;
    `;

    const dialog = document.createElement('div');
    dialog.style.cssText = `
        background: #fff; border-radius: 12px; padding: 28px 32px;
        max-width: 520px; width: 90%; box-shadow: 0 8px 40px rgba(0,0,0,0.3);
    `;

    // ── Title ──
    const title = document.createElement('h2');
    title.textContent = 'Export Analysis Model';
    title.style.cssText = 'margin: 0 0 6px 0; font-size: 18px; color: #1a1a1a;';
    dialog.appendChild(title);

    const subtitle = document.createElement('p');
    subtitle.textContent = 'Export stick model for structural analysis software';
    subtitle.style.cssText = 'margin: 0 0 18px 0; font-size: 12px; color: #888;';
    dialog.appendChild(subtitle);

    // ── Model Summary ──
    const summary = document.createElement('div');
    summary.style.cssText = `
        background: #f7f8fa; border-radius: 8px; padding: 14px 16px;
        margin-bottom: 18px; font-size: 13px; line-height: 1.7; color: #333;
    `;
    summary.innerHTML = `
        <div style="display:grid; grid-template-columns: 1fr 1fr; gap: 2px 20px;">
            <span><strong>Nodes:</strong> ${model.stats.nodeCount}</span>
            <span><strong>Members:</strong> ${model.stats.memberCount}</span>
            <span><strong>Beams:</strong> ${model.stats.beamCount}</span>
            <span><strong>Columns:</strong> ${model.stats.columnCount}</span>
            <span><strong>Supports:</strong> ${model.stats.supportCount}</span>
            <span><strong>Sections:</strong> ${model.stats.sectionCount}</span>
            <span><strong>Materials:</strong> ${model.stats.materialCount}</span>
            <span><strong>Levels:</strong> ${model.stats.levelCount}</span>
        </div>
    `;
    dialog.appendChild(summary);

    // ── Assumptions Panel ──
    const assumptions = document.createElement('div');
    assumptions.style.cssText = `
        background: #fffbeb; border: 1px solid #f0e4b8; border-radius: 8px;
        padding: 12px 14px; margin-bottom: 18px; font-size: 11.5px;
        line-height: 1.6; color: #7a6520;
    `;
    assumptions.innerHTML = `
        <strong style="font-size:12px;">Assumptions (review in analysis software):</strong><br>
        • Beam ends: <strong>Pinned</strong> (simple connections) — Mz released<br>
        • Column ends: <strong>Fixed–Fixed</strong><br>
        • Column bases: <strong>Pinned supports</strong> (Fx, Fy, Fz restrained)<br>
        • Self-weight load case (G1) included<br>
        • Steel: Grade 300, E = 200,000 MPa (AS 4100)<br>
        • Timber: GL17, E = 16,300 MPa (AS 1720.1)
    `;
    dialog.appendChild(assumptions);

    // ── Options ──
    const optionsDiv = document.createElement('div');
    optionsDiv.style.cssText = 'margin-bottom: 20px;';

    // Beam release option
    const releaseLabel = document.createElement('label');
    releaseLabel.style.cssText = 'display: flex; align-items: center; gap: 8px; font-size: 12.5px; color: #444; margin-bottom: 8px; cursor: pointer;';
    const releaseCheck = document.createElement('input');
    releaseCheck.type = 'checkbox';
    releaseCheck.checked = true;
    releaseCheck.id = 'export-pin-beams';
    releaseLabel.appendChild(releaseCheck);
    releaseLabel.appendChild(document.createTextNode('Pin beam ends (simple connections)'));
    optionsDiv.appendChild(releaseLabel);

    // Support type option
    const supportLabel = document.createElement('label');
    supportLabel.style.cssText = 'display: flex; align-items: center; gap: 8px; font-size: 12.5px; color: #444; cursor: pointer;';
    const supportCheck = document.createElement('input');
    supportCheck.type = 'checkbox';
    supportCheck.checked = false;
    supportCheck.id = 'export-fixed-supports';
    supportLabel.appendChild(supportCheck);
    supportLabel.appendChild(document.createTextNode('Fixed supports at column bases (default: pinned)'));
    optionsDiv.appendChild(supportLabel);

    dialog.appendChild(optionsDiv);

    // ── Export Buttons ──
    const btnRow = document.createElement('div');
    btnRow.style.cssText = 'display: flex; gap: 10px; margin-bottom: 12px;';

    const makeBtn = (text, primary, onClick) => {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.style.cssText = `
            flex: 1; padding: 10px 16px; border-radius: 7px; font-size: 13px;
            font-weight: 600; cursor: pointer; border: none;
            font-family: inherit; transition: background 0.15s;
            ${primary
                ? 'background: #2B7CD0; color: white;'
                : 'background: #e8eaed; color: #333;'
            }
        `;
        btn.addEventListener('mouseenter', () => {
            btn.style.background = primary ? '#1a5fa0' : '#d0d3d8';
        });
        btn.addEventListener('mouseleave', () => {
            btn.style.background = primary ? '#2B7CD0' : '#e8eaed';
        });
        btn.addEventListener('click', onClick);
        return btn;
    };

    // SpaceGass export
    btnRow.appendChild(makeBtn('Export SpaceGass (.txt)', true, () => {
        applyDialogOptions(model, releaseCheck.checked, supportCheck.checked);
        const content = writeSpaceGassText(model);
        const filename = (model.projectName.replace(/[^a-zA-Z0-9]/g, '_') || 'model')
            + '_SpaceGass.txt';
        downloadTextFile(content, filename, 'text/plain');
        showExportToast('SpaceGass file exported');
    }));

    // DXF export
    btnRow.appendChild(makeBtn('Export DXF (.dxf)', false, () => {
        applyDialogOptions(model, releaseCheck.checked, supportCheck.checked);
        const content = writeDXFStickModel(model);
        const filename = (model.projectName.replace(/[^a-zA-Z0-9]/g, '_') || 'model')
            + '_StickModel.dxf';
        downloadTextFile(content, filename, 'application/dxf');
        showExportToast('DXF stick model exported');
    }));

    dialog.appendChild(btnRow);

    // ── Close Button ──
    const closeRow = document.createElement('div');
    closeRow.style.cssText = 'text-align: right;';
    const closeBtn = document.createElement('button');
    closeBtn.textContent = 'Close';
    closeBtn.style.cssText = `
        background: none; border: none; color: #888; cursor: pointer;
        font-size: 12px; padding: 4px 12px; font-family: inherit;
    `;
    closeBtn.addEventListener('click', () => overlay.remove());
    closeRow.appendChild(closeBtn);
    dialog.appendChild(closeRow);

    overlay.appendChild(dialog);

    // Close on overlay click
    overlay.addEventListener('click', (e) => {
        if (e.target === overlay) overlay.remove();
    });

    // Close on Escape
    const escHandler = (e) => {
        if (e.key === 'Escape') {
            overlay.remove();
            window.removeEventListener('keydown', escHandler);
        }
    };
    window.addEventListener('keydown', escHandler);

    document.body.appendChild(overlay);
}

/**
 * Apply dialog options to the model before export.
 */
function applyDialogOptions(model, pinBeams, fixedSupports) {
    // Update beam releases
    for (const mem of model.members) {
        if (mem.type === 'beam') {
            mem.releaseI = pinBeams ? 'P' : 'F';
            mem.releaseJ = pinBeams ? 'P' : 'F';
        }
    }

    // Update support conditions
    for (const sup of model.supports) {
        if (fixedSupports) {
            sup.type = 'fixed';
            sup.restraints = '111111';
        } else {
            sup.type = 'pinned';
            sup.restraints = '111000';
        }
    }
}

/**
 * Show a brief toast notification on export success.
 */
function showExportToast(message) {
    const toast = document.createElement('div');
    toast.textContent = '✓ ' + message;
    toast.style.cssText = `
        position: fixed; bottom: 30px; left: 50%; transform: translateX(-50%);
        background: rgba(43,124,208,0.95); color: white; padding: 10px 24px;
        border-radius: 8px; font-size: 13px; font-family: "Segoe UI", Arial, sans-serif;
        z-index: 25000; pointer-events: none; box-shadow: 0 4px 16px rgba(0,0,0,0.2);
        animation: exportFadeUp 2.5s ease-out forwards;
    `;

    // Add animation if not present
    if (!document.getElementById('export-toast-style')) {
        const style = document.createElement('style');
        style.id = 'export-toast-style';
        style.textContent = `
            @keyframes exportFadeUp {
                0% { opacity: 1; transform: translateX(-50%) translateY(0); }
                70% { opacity: 1; }
                100% { opacity: 0; transform: translateX(-50%) translateY(-20px); }
            }
        `;
        document.head.appendChild(style);
    }

    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 2800);
}


// ──────────────────────────────────────────────────────────
// §12  UI INTEGRATION
// ──────────────────────────────────────────────────────────

// ── A) Hook into fullscreen 3D toolbar ──
// We patch enterFullscreen3D to add export buttons after the toolbar
// is built. This avoids modifying 11-3d-engine.js directly.

const _origEnterFullscreen3D = typeof enterFullscreen3D === 'function'
    ? enterFullscreen3D : null;

if (_origEnterFullscreen3D) {
    // We can't cleanly override enterFullscreen3D since it captures
    // local variables. Instead, we use a MutationObserver to detect
    // when the fullscreen overlay appears and inject our button.

    const _exportObserver = new MutationObserver((mutations) => {
        for (const mut of mutations) {
            for (const node of mut.addedNodes) {
                if (node.id === 'fullscreen-3d-overlay') {
                    // Find the toolbar (bottom-centre bar)
                    const toolbar = node.querySelector('div[style*="bottom: 24px"]');
                    if (toolbar) {
                        injectExportButtonIntoToolbar(toolbar);
                    }
                }
            }
        }
    });

    _exportObserver.observe(document.body, { childList: true });
}

/**
 * Inject an "Export Model" button into the fullscreen 3D toolbar.
 */
function injectExportButtonIntoToolbar(toolbar) {
    // Check if already injected
    if (toolbar.querySelector('#btn-export-analysis')) return;

    // Add a separator
    const sep = document.createElement('div');
    sep.style.cssText = 'width:1px; height:20px; background:rgba(255,255,255,0.15); margin:0 4px;';

    // Create export button
    const btn = document.createElement('button');
    btn.id = 'btn-export-analysis';
    btn.textContent = 'Export Model';
    btn.title = 'Export stick model for SpaceGass / analysis software';
    btn.style.cssText = `
        background: rgba(43, 124, 68, 0.85); border: none;
        color: #FFF; padding: 7px 16px;
        border-radius: 6px; cursor: pointer; font-size: 12px;
        font-family: inherit; font-weight: 600;
        transition: background 0.15s;
    `;
    btn.addEventListener('mouseenter', () => {
        btn.style.background = 'rgba(43, 124, 68, 1.0)';
    });
    btn.addEventListener('mouseleave', () => {
        btn.style.background = 'rgba(43, 124, 68, 0.85)';
    });
    btn.addEventListener('click', (e) => {
        e.stopPropagation();
        showAnalysisExportDialog();
    });

    // Insert before the Close button (last child)
    const closeBtn = toolbar.lastElementChild;
    const closeSep = closeBtn ? closeBtn.previousElementSibling : null;

    if (closeSep) {
        toolbar.insertBefore(sep, closeSep);
        toolbar.insertBefore(btn, closeSep);
    } else {
        toolbar.appendChild(sep);
        toolbar.appendChild(btn);
    }
}

// ── B) Add keyboard shortcut (Ctrl+Shift+E) ──
window.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.shiftKey && e.key === 'E') {
        e.preventDefault();
        showAnalysisExportDialog();
    }
});


// ── C) Expose API for external use / testing ──
window._analysisExport = {
    extractModel: extractAnalysisModel,
    writeSpaceGass: writeSpaceGassText,
    writeDXF: writeDXFStickModel,
    showDialog: showAnalysisExportDialog,
    config: ANALYSIS_EXPORT,
};

console.log('[StructuralSketch] Analysis Export module loaded (Ctrl+Shift+E or 3D → Export Model)');
