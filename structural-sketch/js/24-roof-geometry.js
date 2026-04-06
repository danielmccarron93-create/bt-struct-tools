// ============================================================
// 24-roof-geometry.js — 3D Roof Geometry Engine
// Computes roof surfaces from building envelope + ridge line,
// renders in 3D view, and calculates exact area of elevation.
// ============================================================

// ── Geometry Helpers ────────────────────────────────────────

/**
 * Signed perpendicular distance from point P to the line through A→B.
 * Positive = left side of A→B, negative = right side.
 */
function signedDistToLine(P, A, B) {
    const dx = B.x - A.x, dy = B.y - A.y;
    const len = Math.sqrt(dx * dx + dy * dy);
    if (len < 0.01) return 0;
    // Cross product gives signed area of parallelogram
    return ((B.x - A.x) * (P.y - A.y) - (B.y - A.y) * (P.x - A.x)) / len;
}

/**
 * Perpendicular distance from point P to line segment A→B (unsigned).
 */
function distToSegment(P, A, B) {
    const dx = B.x - A.x, dy = B.y - A.y;
    const lenSq = dx * dx + dy * dy;
    if (lenSq < 0.01) return Math.sqrt((P.x - A.x) ** 2 + (P.y - A.y) ** 2);
    let t = ((P.x - A.x) * dx + (P.y - A.y) * dy) / lenSq;
    t = Math.max(0, Math.min(1, t));
    const cx = A.x + t * dx, cy = A.y + t * dy;
    return Math.sqrt((P.x - cx) ** 2 + (P.y - cy) ** 2);
}

/**
 * Project point P onto the infinite line through A→B, return parameter t and closest point.
 */
function projectOntoLine(P, A, B) {
    const dx = B.x - A.x, dy = B.y - A.y;
    const lenSq = dx * dx + dy * dy;
    if (lenSq < 0.01) return { t: 0, x: A.x, y: A.y };
    const t = ((P.x - A.x) * dx + (P.y - A.y) * dy) / lenSq;
    return { t, x: A.x + t * dx, y: A.y + t * dy };
}

/**
 * Midpoint of a segment.
 */
function segMidpoint(A, B) {
    return { x: (A.x + B.x) / 2, y: (A.y + B.y) / 2 };
}


// ── Roof Model Builder ──────────────────────────────────────

/**
 * Build a 3D roof model from building envelope + ridge line.
 *
 * @param {object} envelope - building envelope element (points in real-world mm)
 * @param {object} ridge - ridge line element (points in real-world mm, pitches per segment)
 * @param {number} eavesHeightMM - eaves height above ground in mm
 * @returns {{ surfaces: Array, ridgeHeightMM: number, eavesHeightMM: number, wallFaces: Array }}
 */
function buildRoofModel(envelope, ridge, eavesHeightMM) {
    if (!envelope || !envelope.points || envelope.points.length < 3) return null;
    if (!ridge || !ridge.points || ridge.points.length < 2) return null;

    const footprint = envelope.points;
    const ridgePts = ridge.points;
    const pitches = ridge.pitches || [];
    const eavesH = eavesHeightMM;

    // For simple single-ridge roof, use the first segment's pitches
    // to determine ridge height from perpendicular distance to footprint edges
    const ridgeA = ridgePts[0];
    const ridgeB = ridgePts[ridgePts.length - 1];

    // Classify footprint edges as left (side A) or right (side B) of the ridge
    const classifiedEdges = classifyFootprintEdges(footprint, ridgeA, ridgeB);

    // Compute ridge height: use the minimum perpendicular distance
    // from the ridge line to the nearest footprint edge on each side
    let minDistA = Infinity, minDistB = Infinity;
    for (const ce of classifiedEdges) {
        const mid = segMidpoint(ce.p1, ce.p2);
        const d = Math.abs(signedDistToLine(mid, ridgeA, ridgeB));
        if (ce.side === 'A' && d < minDistA) minDistA = d;
        if (ce.side === 'B' && d < minDistB) minDistB = d;
    }

    // Use the pitch for each side to compute ridge height
    const pitchA = (pitches[0]?.a || 22.5) * Math.PI / 180;
    const pitchB = (pitches[0]?.b || 22.5) * Math.PI / 180;

    // Ridge height = perpendicular distance × tan(pitch)
    // Use the smaller of the two calculations (conservative — roof must meet at ridge)
    const ridgeHeightA = minDistA * Math.tan(pitchA);
    const ridgeHeightB = minDistB * Math.tan(pitchB);
    const ridgeHeightMM = Math.min(ridgeHeightA, ridgeHeightB);
    const ridgeAbsolute = eavesH + ridgeHeightMM;

    // ── Generate roof surfaces ──

    const surfaces = [];
    const wallFaces = [];

    // Wall faces: extrude each footprint edge from ground to eaves
    for (let i = 0; i < footprint.length; i++) {
        const j = (i + 1) % footprint.length;
        const p1 = footprint[i], p2 = footprint[j];
        wallFaces.push({
            type: 'wall',
            vertices: [
                { x: p1.x, y: p1.y, z: 0 },
                { x: p2.x, y: p2.y, z: 0 },
                { x: p2.x, y: p2.y, z: eavesH },
                { x: p1.x, y: p1.y, z: eavesH },
            ]
        });
    }

    // Determine hip vs gable at each ridge endpoint
    const ridgeStartDist = minDistToFootprint(ridgeA, footprint);
    const ridgeEndDist = minDistToFootprint(ridgeB, footprint);
    const hipThreshold = 500; // mm — if ridge end is within 500mm of footprint, it's a gable

    const startIsGable = ridgeStartDist < hipThreshold;
    const endIsGable = ridgeEndDist < hipThreshold;

    // ── Sloped roof surfaces (side A and side B) ──
    // Simplified approach: for each side, create a quad from the eaves edges to the ridge

    // Collect footprint vertices on each side
    const sideAVerts = [], sideBVerts = [];
    for (const pt of footprint) {
        const sd = signedDistToLine(pt, ridgeA, ridgeB);
        if (sd >= 0) sideAVerts.push(pt);
        else sideBVerts.push(pt);
    }

    // Side A slope surface: footprint verts at eaves height → ridge line at ridge height
    if (sideAVerts.length >= 2) {
        // Sort along ridge direction for consistent winding
        const ridgeDx = ridgeB.x - ridgeA.x, ridgeDy = ridgeB.y - ridgeA.y;
        sideAVerts.sort((a, b) => {
            return ((a.x - ridgeA.x) * ridgeDx + (a.y - ridgeA.y) * ridgeDy) -
                   ((b.x - ridgeA.x) * ridgeDx + (b.y - ridgeA.y) * ridgeDy);
        });

        const verts = [];
        for (const v of sideAVerts) {
            verts.push({ x: v.x, y: v.y, z: eavesH });
        }
        // Add ridge points in reverse
        verts.push({ x: ridgeB.x, y: ridgeB.y, z: ridgeAbsolute });
        verts.push({ x: ridgeA.x, y: ridgeA.y, z: ridgeAbsolute });
        surfaces.push({ type: 'slope', side: 'A', vertices: verts });
    }

    // Side B slope surface
    if (sideBVerts.length >= 2) {
        const ridgeDx = ridgeB.x - ridgeA.x, ridgeDy = ridgeB.y - ridgeA.y;
        sideBVerts.sort((a, b) => {
            return ((a.x - ridgeA.x) * ridgeDx + (a.y - ridgeA.y) * ridgeDy) -
                   ((b.x - ridgeA.x) * ridgeDx + (b.y - ridgeA.y) * ridgeDy);
        });

        const verts = [];
        for (const v of sideBVerts) {
            verts.push({ x: v.x, y: v.y, z: eavesH });
        }
        verts.push({ x: ridgeB.x, y: ridgeB.y, z: ridgeAbsolute });
        verts.push({ x: ridgeA.x, y: ridgeA.y, z: ridgeAbsolute });
        surfaces.push({ type: 'slope', side: 'B', vertices: verts });
    }

    // ── Hip ends (triangles from footprint corners to ridge endpoints) ──
    if (!startIsGable) {
        // Hip at ridge start: find nearest footprint corner
        const nearestCorner = findNearestPoint(ridgeA, footprint);
        if (nearestCorner) {
            surfaces.push({
                type: 'hip',
                vertices: [
                    { x: nearestCorner.x, y: nearestCorner.y, z: eavesH },
                    { x: ridgeA.x, y: ridgeA.y, z: ridgeAbsolute },
                ]
            });
        }
    }
    if (!endIsGable) {
        const nearestCorner = findNearestPoint(ridgeB, footprint);
        if (nearestCorner) {
            surfaces.push({
                type: 'hip',
                vertices: [
                    { x: nearestCorner.x, y: nearestCorner.y, z: eavesH },
                    { x: ridgeB.x, y: ridgeB.y, z: ridgeAbsolute },
                ]
            });
        }
    }

    // ── Gable ends (vertical triangles) ──
    if (startIsGable) {
        // Find the two footprint edges nearest to ridge start
        const nearEdge = findNearestFootprintEdge(ridgeA, footprint);
        if (nearEdge) {
            surfaces.push({
                type: 'gable',
                vertices: [
                    { x: nearEdge.p1.x, y: nearEdge.p1.y, z: eavesH },
                    { x: nearEdge.p2.x, y: nearEdge.p2.y, z: eavesH },
                    { x: ridgeA.x, y: ridgeA.y, z: ridgeAbsolute },
                ]
            });
        }
    }
    if (endIsGable) {
        const nearEdge = findNearestFootprintEdge(ridgeB, footprint);
        if (nearEdge) {
            surfaces.push({
                type: 'gable',
                vertices: [
                    { x: nearEdge.p1.x, y: nearEdge.p1.y, z: eavesH },
                    { x: nearEdge.p2.x, y: nearEdge.p2.y, z: eavesH },
                    { x: ridgeB.x, y: ridgeB.y, z: ridgeAbsolute },
                ]
            });
        }
    }

    return {
        surfaces,
        wallFaces,
        ridgeHeightMM: Math.round(ridgeHeightMM),
        ridgeAbsoluteMM: Math.round(ridgeAbsolute),
        eavesHeightMM: Math.round(eavesH),
        startIsGable,
        endIsGable,
    };
}


// ── Footprint Classification Helpers ────────────────────────

function classifyFootprintEdges(footprint, ridgeA, ridgeB) {
    const classified = [];
    for (let i = 0; i < footprint.length; i++) {
        const j = (i + 1) % footprint.length;
        const p1 = footprint[i], p2 = footprint[j];
        const mid = segMidpoint(p1, p2);
        const sd = signedDistToLine(mid, ridgeA, ridgeB);
        classified.push({ p1, p2, side: sd >= 0 ? 'A' : 'B', distance: Math.abs(sd) });
    }
    return classified;
}

function minDistToFootprint(point, footprint) {
    let min = Infinity;
    for (let i = 0; i < footprint.length; i++) {
        const j = (i + 1) % footprint.length;
        const d = distToSegment(point, footprint[i], footprint[j]);
        if (d < min) min = d;
    }
    return min;
}

function findNearestPoint(target, points) {
    let best = null, bestDist = Infinity;
    for (const p of points) {
        const d = Math.sqrt((p.x - target.x) ** 2 + (p.y - target.y) ** 2);
        if (d < bestDist) { bestDist = d; best = p; }
    }
    return best;
}

function findNearestFootprintEdge(point, footprint) {
    let best = null, bestDist = Infinity;
    for (let i = 0; i < footprint.length; i++) {
        const j = (i + 1) % footprint.length;
        const d = distToSegment(point, footprint[i], footprint[j]);
        if (d < bestDist) { bestDist = d; best = { p1: footprint[i], p2: footprint[j] }; }
    }
    return best;
}


// ── Exact Area of Elevation (3D Projection) ─────────────────

/**
 * Calculate the exact area of elevation by projecting all 3D faces
 * onto a vertical plane perpendicular to the wind direction.
 *
 * @param {object} roofModel - output from buildRoofModel()
 * @param {number} windAngleRad - wind direction in radians (0 = +X, π/2 = +Y)
 * @returns {{ area_m2: number, wallArea_m2: number, roofArea_m2: number }}
 */
function calcExactAreaOfElevation(roofModel, windAngleRad) {
    if (!roofModel) return { area_m2: 0, wallArea_m2: 0, roofArea_m2: 0 };

    // Wind direction unit vector (horizontal)
    const windX = Math.cos(windAngleRad);
    const windY = Math.sin(windAngleRad);

    let wallArea = 0;
    let roofArea = 0;

    // Project wall faces
    for (const face of roofModel.wallFaces) {
        const projected = projectFaceToElevation(face.vertices, windX, windY);
        if (projected > 0) wallArea += projected;
    }

    // Project roof surfaces
    for (const surface of roofModel.surfaces) {
        if (surface.vertices.length >= 3) {
            const projected = projectFaceToElevation(surface.vertices, windX, windY);
            if (projected > 0) roofArea += projected;
        }
    }

    const totalArea = wallArea + roofArea;
    return {
        area_m2: Math.round(totalArea / 1e6 * 100) / 100,     // mm² → m²
        wallArea_m2: Math.round(wallArea / 1e6 * 100) / 100,
        roofArea_m2: Math.round(roofArea / 1e6 * 100) / 100,
    };
}

/**
 * Project a 3D polygon face onto a vertical plane perpendicular to wind direction.
 * Returns the projected area in mm² (positive if facing the wind, 0 if facing away).
 */
function projectFaceToElevation(vertices, windX, windY) {
    if (vertices.length < 3) return 0;

    // Compute face normal (using first 3 vertices)
    const v0 = vertices[0], v1 = vertices[1], v2 = vertices[2];
    const ax = v1.x - v0.x, ay = v1.y - v0.y, az = (v1.z || 0) - (v0.z || 0);
    const bx = v2.x - v0.x, by = v2.y - v0.y, bz = (v2.z || 0) - (v0.z || 0);
    const nx = ay * bz - az * by;
    const ny = az * bx - ax * bz;
    // nz = ax * by - ay * bx; // not needed for facing check

    // Check if face is facing the wind (horizontal normal dot wind direction > 0)
    const horizLen = Math.sqrt(nx * nx + ny * ny);
    if (horizLen < 0.01) {
        // Horizontal face (like a flat roof) — projected area on vertical plane is ~0
        return 0;
    }
    const dotProduct = (nx / horizLen) * windX + (ny / horizLen) * windY;
    if (dotProduct <= 0) return 0; // Face is facing away from wind

    // Project vertices onto the elevation plane:
    // Elevation plane coordinates: (perpendicular to wind, z)
    // perpX = -windY, perpY = windX (rotation by 90°)
    const projected2D = vertices.map(v => ({
        u: -windY * v.x + windX * v.y,  // horizontal position on elevation plane
        v: v.z || 0,                      // vertical position (height)
    }));

    // Shoelace area of projected polygon
    let area = 0;
    for (let i = 0; i < projected2D.length; i++) {
        const j = (i + 1) % projected2D.length;
        area += projected2D[i].u * projected2D[j].v - projected2D[j].u * projected2D[i].v;
    }
    return Math.abs(area) / 2;
}


// ── Expose Globally ─────────────────────────────────────────

window.buildRoofModel = buildRoofModel;
window.calcExactAreaOfElevation = calcExactAreaOfElevation;
window.classifyFootprintEdges = classifyFootprintEdges;
