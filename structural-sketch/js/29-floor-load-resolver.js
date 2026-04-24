// ============================================================
// 29-floor-load-resolver.js — Polygonal floor load resolver
// ------------------------------------------------------------
// Pure geometry / load-lookup module. No DOM, no canvas, no
// schedule access. Consumers pass in an array of floorZone
// elements (already resolved from the schedule) and ask either:
//
//   FloorLoadResolver.getLoadAtPoint(x, y, zones, defaultG, defaultQ)
//     → { G_kPa, Q_kPa, spanDirection_deg, zoneId | null }
//
//   FloorLoadResolver.pointInPolygon(x, y, vertices)
//     → boolean  (vertices: [{x,y}, ...] in any consistent unit)
//
// Coordinate system: all inputs/outputs are in REAL-WORLD mm
// (not sheet-mm, not screen-px). Callers are responsible for
// converting if their data is in another frame.
//
// A "zone" here is the RESOLVED object:
//   { id, typeRef, vertices: [{x,y}, ...], G_kPa, Q_kPa,
//     spanDirection_deg, color }
//
// The element on project.elements has only { id, type:'floorZone',
// typeRef, points, layer, ... } — the resolver helper
// `resolveZonesFromElements()` joins the schedule entry in to
// produce the above.
//
// Zone ordering: later-drawn zones win (painter's algorithm),
// so we iterate in REVERSE and return the first hit. This
// matches the prototype behaviour and lets the engineer drop
// a heavy-load patch on top of a base floor load without
// editing the underlying zone.
//
// Later slices will add:
//   Slice 3 — buildPiecewiseLoadFns(beam, zones, swKNm)
//             for the AS 4100 beam engine
//   Slice 4 — tributary strip sampler for calc-sheet viz
// ============================================================

(function () {
    'use strict';

    // ---- Geometry ------------------------------------------------

    /**
     * Ray-casting point-in-polygon test.
     * @param {number} px  — test point X (real mm)
     * @param {number} py  — test point Y (real mm)
     * @param {Array<{x:number,y:number}>} verts — polygon vertices
     * @returns {boolean}
     */
    function pointInPolygon(px, py, verts) {
        if (!verts || verts.length < 3) return false;
        let inside = false;
        const n = verts.length;
        for (let i = 0, j = n - 1; i < n; j = i++) {
            const xi = verts[i].x, yi = verts[i].y;
            const xj = verts[j].x, yj = verts[j].y;
            // Classic ray-cast; strict inequality on yi/yj avoids
            // double-counting a vertex that sits exactly on the ray.
            const intersects = ((yi > py) !== (yj > py)) &&
                (px < ((xj - xi) * (py - yi)) / ((yj - yi) || 1e-12) + xi);
            if (intersects) inside = !inside;
        }
        return inside;
    }

    /**
     * Shoelace area for a closed polygon, in same units² as input.
     * Returned value is unsigned.
     */
    function polygonArea(verts) {
        if (!verts || verts.length < 3) return 0;
        let a = 0;
        const n = verts.length;
        for (let i = 0, j = n - 1; i < n; j = i++) {
            a += (verts[j].x + verts[i].x) * (verts[j].y - verts[i].y);
        }
        return Math.abs(a * 0.5);
    }

    // ---- Line-segment geometry utilities -------------------------

    /**
     * Compute intersection of two line segments (a1→a2) and (b1→b2).
     * Returns {x, y, t, u} where t is parametric position on segment A
     * and u on segment B (both in [0,1] if intersection is within segments).
     * Returns null if segments are parallel or don't intersect.
     */
    function segmentIntersection(a1, a2, b1, b2) {
        var dx1 = a2.x - a1.x, dy1 = a2.y - a1.y;
        var dx2 = b2.x - b1.x, dy2 = b2.y - b1.y;
        var denom = dx1 * dy2 - dy1 * dx2;
        if (Math.abs(denom) < 1e-10) return null; // parallel
        var t = ((b1.x - a1.x) * dy2 - (b1.y - a1.y) * dx2) / denom;
        var u = ((b1.x - a1.x) * dy1 - (b1.y - a1.y) * dx1) / denom;
        if (t < -1e-9 || t > 1 + 1e-9 || u < -1e-9 || u > 1 + 1e-9) return null;
        return { x: a1.x + t * dx1, y: a1.y + t * dy1, t: t, u: u };
    }

    /**
     * Clip a line segment (p1→p2) to a polygon (convex or concave).
     * Returns array of clipped segment pairs [{enter:{x,y}, exit:{x,y}}].
     * Uses parametric intersection: find all edge crossings, sort by t,
     * then pair up entry/exit points based on inside/outside state.
     */
    function clipSegmentToPolygon(p1, p2, polyVerts) {
        if (!polyVerts || polyVerts.length < 3) return [];
        var n = polyVerts.length;
        var tValues = [];

        // Find all intersection t-values along the line p1→p2
        for (var i = 0, j = n - 1; i < n; j = i++) {
            var hit = segmentIntersection(p1, p2, polyVerts[j], polyVerts[i]);
            if (hit) tValues.push(hit.t);
        }

        // Add t=0 and t=1 if inside polygon
        var p1Inside = pointInPolygon(p1.x, p1.y, polyVerts);
        var p2Inside = pointInPolygon(p2.x, p2.y, polyVerts);
        if (p1Inside) tValues.push(0);
        if (p2Inside) tValues.push(1);

        // Sort and deduplicate
        tValues.sort(function (a, b) { return a - b; });
        var unique = [];
        for (var k = 0; k < tValues.length; k++) {
            if (unique.length === 0 || tValues[k] - unique[unique.length - 1] > 1e-9) {
                unique.push(Math.max(0, Math.min(1, tValues[k])));
            }
        }

        // Pair into segments: test midpoint of each consecutive pair
        var dx = p2.x - p1.x, dy = p2.y - p1.y;
        var segments = [];
        for (var m = 0; m < unique.length - 1; m++) {
            var tMid = (unique[m] + unique[m + 1]) / 2;
            var mx = p1.x + tMid * dx, my = p1.y + tMid * dy;
            if (pointInPolygon(mx, my, polyVerts)) {
                segments.push({
                    enter: { x: p1.x + unique[m] * dx, y: p1.y + unique[m] * dy },
                    exit:  { x: p1.x + unique[m + 1] * dx, y: p1.y + unique[m + 1] * dy },
                });
            }
        }
        return segments;
    }

    /**
     * Project a point onto an axis defined by origin + unit direction.
     * Returns the scalar distance along the axis.
     */
    function projectOntoAxis(point, origin, dir) {
        return (point.x - origin.x) * dir.x + (point.y - origin.y) * dir.y;
    }

    // ---- Schedule join -------------------------------------------

    /**
     * Given project.elements and the floorLoad schedule, return
     * an array of resolved zone objects ready for load lookup.
     * Silently skips elements whose typeRef is missing from the
     * schedule (e.g. a legacy project with a deleted type).
     *
     * @param {Array} elements  — project.elements
     * @param {Object} floorLoadSchedule — project.scheduleTypes.floorLoad
     */
    function resolveZonesFromElements(elements, floorLoadSchedule) {
        if (!Array.isArray(elements) || !floorLoadSchedule) return [];
        const zones = [];
        for (let i = 0; i < elements.length; i++) {
            const el = elements[i];
            if (!el || el.type !== 'floorZone') continue;
            const t = floorLoadSchedule[el.typeRef];
            if (!t) continue; // dangling typeRef — skip
            const verts = Array.isArray(el.points) ? el.points : [];
            if (verts.length < 3) continue;
            // Per-element overrides take precedence over schedule entry
            zones.push({
                id: el.id,
                typeRef: el.typeRef,
                vertices: verts,
                G_kPa: (el.G_override !== undefined) ? Number(el.G_override) : Number(t.G) || 0,
                Q_kPa: (el.Q_override !== undefined) ? Number(el.Q_override) : Number(t.Q) || 0,
                spanDirection_deg: (el.spanDirection_override !== undefined) ? Number(el.spanDirection_override) : Number(t.spanDirection) || 0,
                color: t.color || '#A7F3D0',
            });
        }
        return zones;
    }

    // ---- Load lookup ---------------------------------------------

    /**
     * Return the floor G/Q at a given real-world point.
     * If no zone contains the point, returns the supplied
     * defaults (which may be 0 — the caller decides policy).
     *
     * @param {number} x_mm
     * @param {number} y_mm
     * @param {Array} zones — resolved zones (see resolveZonesFromElements)
     * @param {number} [defaultG=0]
     * @param {number} [defaultQ=0]
     * @returns {{G_kPa:number, Q_kPa:number, spanDirection_deg:number, zoneId:string|null}}
     */
    function getLoadAtPoint(x_mm, y_mm, zones, defaultG, defaultQ) {
        if (Array.isArray(zones)) {
            // Reverse iteration — later-drawn zones win
            for (let i = zones.length - 1; i >= 0; i--) {
                const z = zones[i];
                if (pointInPolygon(x_mm, y_mm, z.vertices)) {
                    return {
                        G_kPa: z.G_kPa,
                        Q_kPa: z.Q_kPa,
                        spanDirection_deg: z.spanDirection_deg,
                        zoneId: z.id,
                    };
                }
            }
        }
        return {
            G_kPa: (typeof defaultG === 'number') ? defaultG : 0,
            Q_kPa: (typeof defaultQ === 'number') ? defaultQ : 0,
            spanDirection_deg: 0,
            zoneId: null,
        };
    }

    /**
     * Axis-aligned bounding box of a polygon.
     * Useful for broad-phase culling in the beam engine (Slice 3)
     * before running the full point-in-polygon test on every
     * sample along a beam.
     */
    function polygonBBox(verts) {
        if (!verts || verts.length === 0) return null;
        let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
        for (let i = 0; i < verts.length; i++) {
            const v = verts[i];
            if (v.x < minX) minX = v.x;
            if (v.y < minY) minY = v.y;
            if (v.x > maxX) maxX = v.x;
            if (v.y > maxY) maxY = v.y;
        }
        return { minX, minY, maxX, maxY };
    }

    // ---- Piecewise load functions (Slice 3) ---------------------

    /**
     * Build w_G(x) and w_Q(x) closures for the piecewise beam solver.
     *
     * Given a beam's start/end in real-world mm, its tributary width,
     * and the resolved floor zones, return two functions that map a
     * position along the beam (x in metres from beam start) to the
     * dead / live UDL in kN/m at that point.
     *
     * The beam self-weight is NOT included here — the caller (or the
     * calc engine) adds it.  G/Q from the zones are area loads (kPa),
     * multiplied by the tributary width to give a line load (kN/m).
     *
     * If no zone covers a sample point, the supplied default G/Q are
     * used (typically the level's base floor load, or 0).
     *
     * @param {object} beam — { x1, y1, x2, y2 } in real mm
     * @param {number} tribWidth_mm — tributary width, mm
     * @param {Array}  zones — resolved zones (from resolveZonesFromElements)
     * @param {number} [defaultG=0] — default G kPa if no zone covers point
     * @param {number} [defaultQ=0] — default Q kPa if no zone covers point
     * @returns {{ w_G_of_x_kNm: function, w_Q_of_x_kNm: function }}
     */
    function buildPiecewiseLoadFns(beam, tribWidth_mm, zones, defaultG, defaultQ) {
        var x1 = beam.x1, y1 = beam.y1, x2 = beam.x2, y2 = beam.y2;
        var dx = x2 - x1, dy = y2 - y1;
        var beamLen_mm = Math.sqrt(dx * dx + dy * dy);
        if (beamLen_mm < 1) beamLen_mm = 1; // avoid div/0
        // Unit direction vector along beam
        var ux = dx / beamLen_mm, uy = dy / beamLen_mm;
        var tribW_m = tribWidth_mm / 1000;
        var dG = (typeof defaultG === 'number') ? defaultG : 0;
        var dQ = (typeof defaultQ === 'number') ? defaultQ : 0;

        // Sample at position x_m along beam (metres from start)
        function sampleAt(x_m) {
            var pos_mm = x_m * 1000; // convert to mm along beam
            var px = x1 + ux * pos_mm;
            var py = y1 + uy * pos_mm;
            return getLoadAtPoint(px, py, zones, dG, dQ);
        }

        return {
            w_G_of_x_kNm: function (x_m) {
                return sampleAt(x_m).G_kPa * tribW_m;
            },
            w_Q_of_x_kNm: function (x_m) {
                return sampleAt(x_m).Q_kPa * tribW_m;
            },
        };
    }

    // ---- Bay enumeration (Slice 5) --------------------------------

    /**
     * Detect rectangular bays from beams on the current level.
     *
     * A "bay" is defined by two parallel beams that share endpoints
     * on perpendicular beams. In the main app (unlike the prototype),
     * beams are free-form — not constrained to a regular grid. So we
     * use a different strategy:
     *
     * 1. Collect all beams on the given level
     * 2. Separate into X-ish (< 45° from horizontal) and Y-ish groups
     * 3. For each pair of adjacent parallel beams, form a bay bounded
     *    by the two beams and the perpendicular extent of the floor zones
     *
     * For the simpler prototype-like case (regular grid), this still
     * produces the correct rectangular bays.
     *
     * @param {Array} elements      — project.elements
     * @param {string} levelId      — level to scan
     * @param {string} spanDir      — 'X' or 'Y' (joist span direction)
     * @returns {Array<Object>}     — bay objects with { id, x0_mm, y0_mm, x1_mm, y1_mm, joistSpan_m, width_mm, height_mm, boundingBeams }
     */
    function enumerateBays(elements, levelId, spanDir) {
        if (!Array.isArray(elements)) return [];

        // Collect beams on this level
        var beams = [];
        for (var i = 0; i < elements.length; i++) {
            var el = elements[i];
            if (!el || el.type !== 'line' || el.layer !== 'S-BEAM') continue;
            if (el.level && el.level !== levelId) continue;
            beams.push(el);
        }

        // Classify beams as X-ish (< 45°) or Y-ish
        var xBeams = [], yBeams = [];
        for (var i = 0; i < beams.length; i++) {
            var b = beams[i];
            var dx = Math.abs(b.x2 - b.x1);
            var dy = Math.abs(b.y2 - b.y1);
            if (dx >= dy) {
                xBeams.push(b);
            } else {
                yBeams.push(b);
            }
        }

        // The "loaded" beams run perpendicular to the joist span direction.
        // If joists span in X, Y-beams carry load → bays are bounded by Y-beams, joists span between them.
        // If joists span in Y, X-beams carry load → bays bounded by X-beams.
        var loadedBeams, tieBeams;
        if (spanDir === 'X') {
            // Joists span X → loaded beams are Y-direction (perpendicular to span)
            loadedBeams = yBeams;
            tieBeams = xBeams;
        } else {
            // Joists span Y → loaded beams are X-direction
            loadedBeams = xBeams;
            tieBeams = yBeams;
        }

        // Sort loaded beams by their perpendicular-axis position
        // (midpoint along the span direction axis)
        // perpAxis = direction loaded beams RUN (for overlap check)
        // spanAxis = direction BETWEEN loaded beams (for sorting + joist span)
        var perpAxis = (spanDir === 'X') ? 'y' : 'x';
        var spanAxis = (spanDir === 'X') ? 'x' : 'y';
        loadedBeams.sort(function (a, b) {
            var aMid = ((a[spanAxis + '1'] || 0) + (a[spanAxis + '2'] || 0)) / 2;
            var bMid = ((b[spanAxis + '1'] || 0) + (b[spanAxis + '2'] || 0)) / 2;
            return aMid - bMid;
        });

        // Form bays between adjacent loaded beams
        var bays = [];
        for (var i = 0; i < loadedBeams.length - 1; i++) {
            var beamA = loadedBeams[i];
            var beamB = loadedBeams[i + 1];

            // Perpendicular extent (along span direction): overlap region
            var aMin = Math.min(beamA[perpAxis + '1'], beamA[perpAxis + '2']);
            var aMax = Math.max(beamA[perpAxis + '1'], beamA[perpAxis + '2']);
            var bMin = Math.min(beamB[perpAxis + '1'], beamB[perpAxis + '2']);
            var bMax = Math.max(beamB[perpAxis + '1'], beamB[perpAxis + '2']);
            var overlapMin = Math.max(aMin, bMin);
            var overlapMax = Math.min(aMax, bMax);
            if (overlapMax <= overlapMin) continue; // no overlap

            // Span-axis positions of the two beams
            var aMidSpan = ((beamA[spanAxis + '1'] || 0) + (beamA[spanAxis + '2'] || 0)) / 2;
            var bMidSpan = ((beamB[spanAxis + '1'] || 0) + (beamB[spanAxis + '2'] || 0)) / 2;
            var spanLo = Math.min(aMidSpan, bMidSpan);
            var spanHi = Math.max(aMidSpan, bMidSpan);
            var joistSpan_mm = spanHi - spanLo;
            if (joistSpan_mm < 100) continue; // degenerate

            var x0, y0, x1, y1;
            if (spanDir === 'X') {
                // Joists span in X between two Y-beams
                x0 = overlapMin; x1 = overlapMax; // along beam extent
                y0 = spanLo;     y1 = spanHi;     // between beams
                // Actually: joists span perpendicular to the loaded beams.
                // Loaded beams are Y-direction. Joist span is the X-distance between them? No.
                // Re-think: if spanDir = X, joists span in X. Loaded beams run in Y.
                // The joist clear span is the distance between the two Y-beams measured in X.
                // But Y-beams are vertical, so their X positions define the joist span.
                x0 = spanLo; x1 = spanHi;
                y0 = overlapMin; y1 = overlapMax;
                joistSpan_mm = x1 - x0;
            } else {
                // Joists span in Y between two X-beams
                x0 = overlapMin; x1 = overlapMax;
                y0 = spanLo;     y1 = spanHi;
                joistSpan_mm = y1 - y0;
            }

            bays.push({
                id: 'BAY-' + (bays.length + 1),
                x0_mm: x0,
                y0_mm: y0,
                x1_mm: x1,
                y1_mm: y1,
                width_mm: x1 - x0,
                height_mm: y1 - y0,
                joistSpan_mm: joistSpan_mm,
                joistSpan_m: joistSpan_mm / 1000,
                boundingBeams: [beamA.id, beamB.id],
            });
        }

        return bays;
    }

    /**
     * Compute worst-case (maximum) G and Q loads within a bay,
     * sampling across the resolved floor zones.
     *
     * The prototype used axis-aligned rectangle intersection because
     * zones were rectangles. Our zones are arbitrary polygons, so we
     * sample a grid of points within the bay and take the max.
     *
     * @param {Object} bay — { x0_mm, y0_mm, x1_mm, y1_mm }
     * @param {Array} zones — resolved zones
     * @param {number} defaultG — background G kPa
     * @param {number} defaultQ — background Q kPa
     * @returns {{ G_kPa, Q_kPa, zoneIds: string[], hasZoneOverlap: boolean }}
     */
    function bayWorstCaseLoad(bay, zones, defaultG, defaultQ) {
        var maxG = defaultG || 0;
        var maxQ = defaultQ || 0;
        var zoneIds = [];
        var hasOverlap = false;

        if (!zones || zones.length === 0) {
            return { G_kPa: maxG, Q_kPa: maxQ, zoneIds: zoneIds, hasZoneOverlap: false };
        }

        // Sample a grid of 5x5 points inside the bay
        var nSamples = 5;
        var dx = (bay.x1_mm - bay.x0_mm) / (nSamples + 1);
        var dy = (bay.y1_mm - bay.y0_mm) / (nSamples + 1);
        var seenZones = {};

        for (var ix = 1; ix <= nSamples; ix++) {
            for (var iy = 1; iy <= nSamples; iy++) {
                var px = bay.x0_mm + dx * ix;
                var py = bay.y0_mm + dy * iy;
                var hit = getLoadAtPoint(px, py, zones, defaultG, defaultQ);
                if (hit.G_kPa > maxG) maxG = hit.G_kPa;
                if (hit.Q_kPa > maxQ) maxQ = hit.Q_kPa;
                if (hit.zoneId && !seenZones[hit.zoneId]) {
                    seenZones[hit.zoneId] = true;
                    zoneIds.push(hit.zoneId);
                    hasOverlap = true;
                }
            }
        }

        return { G_kPa: maxG, Q_kPa: maxQ, zoneIds: zoneIds, hasZoneOverlap: hasOverlap };
    }

    // ---- Export --------------------------------------------------

    window.FloorLoadResolver = {
        version: '0.4-joistZone',
        pointInPolygon: pointInPolygon,
        polygonArea: polygonArea,
        polygonBBox: polygonBBox,
        resolveZonesFromElements: resolveZonesFromElements,
        getLoadAtPoint: getLoadAtPoint,
        buildPiecewiseLoadFns: buildPiecewiseLoadFns,
        enumerateBays: enumerateBays,
        bayWorstCaseLoad: bayWorstCaseLoad,
        // Geometry utilities for joist zone system
        segmentIntersection: segmentIntersection,
        clipSegmentToPolygon: clipSegmentToPolygon,
        projectOntoAxis: projectOntoAxis,
    };

    if (typeof console !== 'undefined') {
        console.log('[floor-load-resolver] 0.3-slice5 loaded');
    }
})();
