// ============================================================
// 19-raft-design.js — Raft Slab Design Mode (AS 2870)
// ============================================================
// Implements Option B: a dedicated design mode that transforms
// the StructuralSketch UI for raft slab design.
//
// Components:
//   1. Raft Design Panel (sidebar) — parameters, compliance, calcs
//   2. Geometry Engine — polygon ops, beam clipping (ported from standalone)
//   3. Beam Solver — optimised beam placement per site class
//   4. Canvas Renderer — beam pairs, annotations, dims in design mode
//   5. Commit-to-Plan — converts design into real SS elements
//
// Depends on: 01-core, 02-canvas-engine, 03-app-init, 13-schedules,
//             20-as3600-calcs, 21-as2870-calcs
// ============================================================

const RaftDesign = (() => {
    'use strict';

    // -----------------------------------------------------------------
    // MODULE STATE
    // -----------------------------------------------------------------

    const state = {
        active: false,                // Is raft design mode active?
        phase: 'draw',                // 'draw' | 'design' | 'committed'
        vertices: [],                 // Slab outline vertices [{x, y}] in real-world mm
        isClosed: false,
        layout: null,                 // Solved beam layout
        mouseReal: null,              // Current mouse in real-world mm
        snapMM: 500,                  // Snap to 500mm (0.5m) grid

        // Design parameters
        params: {
            siteClass: 'M',
            ys: null,                 // null = use default for site class
            Hs: 1.8,                  // m, default Brisbane
            location: 'Brisbane',
            constructionType: 'Articulated masonry veneer',
            concreteGrade: 'N25',
            beamWidth: 300,           // mm
            beamDepth: 400,           // mm
            slabThickness: 125,       // mm
            slabMesh: 'SL82',
            topCover: 30,             // mm
            botCover: 40,             // mm
            topReo: '3-L12TM',
            botReo: '3-L12TM',
            qu: 150,                  // kPa — ultimate bearing capacity
            finishes: 0.5,            // kPa
            partitions: 0.5,          // kPa
            Q: 1.5,                   // kPa imposed
            storeys: 1,
        },

        // Calculation results (populated after solve)
        calcResults: null,

        // Previous tool state (restored on exit)
        _prevTool: null,
        _prevLayer: null,
    };

    // -----------------------------------------------------------------
    // POLYGON GEOMETRY (ported from standalone tool)
    // -----------------------------------------------------------------

    function ensureCCW(verts) {
        let area = 0;
        const n = verts.length;
        for (let i = 0; i < n; i++) {
            const j = (i + 1) % n;
            area += verts[i].x * verts[j].y - verts[j].x * verts[i].y;
        }
        if (area < 0) return verts.slice().reverse();
        return verts.slice();
    }

    function pointInPolygon(px, py, verts) {
        let inside = false;
        const n = verts.length;
        for (let i = 0, j = n - 1; i < n; j = i++) {
            const xi = verts[i].x, yi = verts[i].y;
            const xj = verts[j].x, yj = verts[j].y;
            if ((yi > py) !== (yj > py) && px < (xj - xi) * (py - yi) / (yj - yi) + xi) {
                inside = !inside;
            }
        }
        return inside;
    }

    function pointOnBoundary(px, py, verts, tol = 20) {
        // tol in mm (real-world)
        const n = verts.length;
        for (let i = 0; i < n; i++) {
            const j = (i + 1) % n;
            const d = distPointToSegment(px, py, verts[i].x, verts[i].y, verts[j].x, verts[j].y);
            if (d < tol) return true;
        }
        return false;
    }

    function distPointToSegment(px, py, x1, y1, x2, y2) {
        const dx = x2 - x1, dy = y2 - y1;
        const lenSq = dx * dx + dy * dy;
        if (lenSq < 1) return Math.hypot(px - x1, py - y1);
        let t = ((px - x1) * dx + (py - y1) * dy) / lenSq;
        t = Math.max(0, Math.min(1, t));
        return Math.hypot(px - (x1 + t * dx), py - (y1 + t * dy));
    }

    function segSegIntersect(ax, ay, bx, by, cx, cy, dx, dy) {
        const denom = (bx - ax) * (dy - cy) - (by - ay) * (dx - cx);
        if (Math.abs(denom) < 1e-6) return null;
        const t = ((cx - ax) * (dy - cy) - (cy - ay) * (dx - cx)) / denom;
        const u = ((cx - ax) * (by - ay) - (cy - ay) * (bx - ax)) / denom;
        if (t >= -1e-6 && t <= 1 + 1e-6 && u >= -1e-6 && u <= 1 + 1e-6) {
            return {
                t: Math.max(0, Math.min(1, t)),
                x: ax + t * (bx - ax),
                y: ay + t * (by - ay),
            };
        }
        return null;
    }

    function clipBeamToInterior(x1, y1, x2, y2, verts, tol = 20) {
        const n = verts.length;
        let intPts = [];
        for (let i = 0; i < n; i++) {
            const j = (i + 1) % n;
            const hit = segSegIntersect(x1, y1, x2, y2,
                verts[i].x, verts[i].y, verts[j].x, verts[j].y);
            if (hit) intPts.push(hit);
        }
        if (intPts.length < 2) return [];

        const dx = x2 - x1, dy = y2 - y1;
        intPts.sort((a, b) => ((a.x - x1) * dx + (a.y - y1) * dy) - ((b.x - x1) * dx + (b.y - y1) * dy));

        // Remove near-duplicates
        let cleaned = [intPts[0]];
        for (let i = 1; i < intPts.length; i++) {
            if (Math.hypot(intPts[i].x - cleaned[cleaned.length - 1].x,
                           intPts[i].y - cleaned[cleaned.length - 1].y) > tol) {
                cleaned.push(intPts[i]);
            }
        }
        if (cleaned.length < 2) return [];

        let segs = [];
        for (let i = 0; i < cleaned.length - 1; i++) {
            const p1 = cleaned[i], p2 = cleaned[i + 1];
            const segLen = Math.hypot(p2.x - p1.x, p2.y - p1.y);
            if (segLen < tol) continue;
            const mx = (p1.x + p2.x) / 2, my = (p1.y + p2.y) / 2;
            if (pointInPolygon(mx, my, verts)) {
                segs.push({ x1: r(p1.x), y1: r(p1.y), x2: r(p2.x), y2: r(p2.y) });
            }
        }
        return segs;
    }

    function r(v) { return Math.round(v); }  // round to nearest mm

    function cross2d(v1, v2) { return v1[0] * v2[1] - v1[1] * v2[0]; }

    function polygonArea(verts) {
        let area = 0;
        const n = verts.length;
        for (let i = 0; i < n; i++) {
            const j = (i + 1) % n;
            area += verts[i].x * verts[j].y - verts[j].x * verts[i].y;
        }
        return Math.abs(area) / 2; // mm²
    }

    // -----------------------------------------------------------------
    // SLAB MODEL BUILDER
    // -----------------------------------------------------------------

    function buildSlabModel(rawVerts) {
        const verts = ensureCCW(rawVerts);
        const n = verts.length;
        const edges = [];
        for (let i = 0; i < n; i++) {
            edges.push({ p1: verts[i], p2: verts[(i + 1) % n] });
        }

        // Classify corners
        const corners = [];
        for (let i = 0; i < n; i++) {
            const prev = verts[(i - 1 + n) % n];
            const curr = verts[i];
            const next = verts[(i + 1) % n];
            const vIn = [curr.x - prev.x, curr.y - prev.y];
            const vOut = [next.x - curr.x, next.y - curr.y];
            const c = cross2d(vIn, vOut);

            const type = c > 0 ? 'EXTERNAL' : (c < 0 ? 'REENTRANT' : 'EXTERNAL');
            const edgeBefore = Math.hypot(vIn[0], vIn[1]);
            const edgeAfter = Math.hypot(vOut[0], vOut[1]);

            const corner = {
                point: { x: curr.x, y: curr.y },
                index: i,
                type: type,
                edgeBefore: edgeBefore,
                edgeAfter: edgeAfter,
                recessA: edgeBefore,
                recessB: edgeAfter,
                voidDx: 0, voidDy: 0, voidQuadrant: '',
            };

            if (type === 'REENTRANT') {
                const dirs = [[1, 1, 'NE'], [1, -1, 'SE'], [-1, 1, 'NW'], [-1, -1, 'SW']];
                for (const [dx, dy, name] of dirs) {
                    const tx = curr.x + dx * 500, ty = curr.y + dy * 500;
                    if (!pointInPolygon(tx, ty, verts) && !pointOnBoundary(tx, ty, verts, 50)) {
                        corner.voidDx = dx; corner.voidDy = dy; corner.voidQuadrant = name;
                        break;
                    }
                }
                if (!corner.voidQuadrant) {
                    for (const [dx, dy, name] of dirs) {
                        const tx = curr.x + dx * 1500, ty = curr.y + dy * 1500;
                        if (!pointInPolygon(tx, ty, verts) && !pointOnBoundary(tx, ty, verts, 50)) {
                            corner.voidDx = dx; corner.voidDy = dy; corner.voidQuadrant = name;
                            break;
                        }
                    }
                }
            }
            corners.push(corner);
        }

        const xs = verts.map(v => v.x), ys = verts.map(v => v.y);
        return {
            vertices: verts,
            corners: corners,
            edges: edges,
            bounds: { xMin: Math.min(...xs), yMin: Math.min(...ys), xMax: Math.max(...xs), yMax: Math.max(...ys) },
            get width() { return this.bounds.xMax - this.bounds.xMin; },
            get height() { return this.bounds.yMax - this.bounds.yMin; },
            get area() { return polygonArea(this.vertices); },
            get externalCorners() { return this.corners.filter(c => c.type === 'EXTERNAL'); },
            get reentrantCorners() { return this.corners.filter(c => c.type === 'REENTRANT'); },
        };
    }

    // -----------------------------------------------------------------
    // REENTRANT CORNER CLASSIFICATION
    // -----------------------------------------------------------------

    function classifyReentrant(corner, params) {
        const a = corner.recessA, b = corner.recessB;
        const oneSmall = (a < 1500 || b < 1500); // 1.5m in mm
        return {
            corner: corner.point,
            sideA: a, sideB: b,
            oneSmall: oneSmall,
            treatment: oneSmall ? 'FIGURE_5_4' : 'FULL_CONTINUITY',
            clause: oneSmall ? 'Cl 5.3.8 / Figure 5.4' : 'Cl 5.3.8 / Commentary C5.3.8 Note 2',
            note: oneSmall
                ? `One side < 1.5m (${(a/1000).toFixed(1)}m, ${(b/1000).toFixed(1)}m). Figure 5.4 details apply.`
                : `Both sides ≥ 1.5m (${(a/1000).toFixed(1)}m, ${(b/1000).toFixed(1)}m). Full continuity required.`,
        };
    }

    // -----------------------------------------------------------------
    // BEAM SOLVER
    // -----------------------------------------------------------------

    function optimisePositions(start, end, maxSp, mandatory = []) {
        const fixed = [...new Set([start, ...mandatory, end])].sort((a, b) => a - b);
        const all = new Set(mandatory);

        for (let i = 0; i < fixed.length - 1; i++) {
            const span = fixed[i + 1] - fixed[i];
            if (span <= maxSp) continue;
            const numBays = Math.ceil(span / maxSp);
            const baySize = span / numBays;
            for (let j = 1; j < numBays; j++) {
                all.add(r(fixed[i] + j * baySize));
            }
        }
        return [...all].filter(p => p > start + 10 && p < end - 10).sort((a, b) => a - b);
    }

    function solveBeamLayout(model, params) {
        const verts = model.vertices;
        const { xMin, yMin, xMax, yMax } = model.bounds;
        const maxSp = AS2870.MAX_SPACING[params.siteClass] * 1000;  // Convert m to mm
        const bw = params.beamWidth;
        const bd = params.beamDepth;
        const log = [];

        log.push(`BEAM SOLVER — Site Class ${params.siteClass}`);
        log.push(`Slab: ${(model.width/1000).toFixed(1)}m × ${(model.height/1000).toFixed(1)}m`);
        log.push(`Max spacing: ${maxSp/1000}m | Beam: ${bw}W × ${bd}D`);

        // Step 1: Edge beams (one per polygon edge)
        const edgeBeams = model.edges.map((edge, i) => ({
            x1: edge.p1.x, y1: edge.p1.y,
            x2: edge.p2.x, y2: edge.p2.y,
            type: 'EB1', category: 'edge',
            direction: Math.abs(edge.p2.x - edge.p1.x) < 50 ? 'NS' : 'EW',
            width: bw, depth: bd,
            get length() { return Math.hypot(this.x2 - this.x1, this.y2 - this.y1); },
            get midpoint() { return { x: (this.x1 + this.x2) / 2, y: (this.y1 + this.y2) / 2 }; },
        }));

        // Step 2: Reentrant analysis — mandatory beam positions
        const mandatoryXs = [], mandatoryYs = [];
        const treatments = {};

        for (const rc of model.reentrantCorners) {
            const t = classifyReentrant(rc, params);
            treatments[`${rc.point.x},${rc.point.y}`] = t;
            if (t.treatment === 'FULL_CONTINUITY') {
                mandatoryXs.push(rc.point.x);
                mandatoryYs.push(rc.point.y);
            } else {
                mandatoryXs.push(rc.point.x);
            }
        }

        // Step 3: Optimised grid positions
        const nsPositions = optimisePositions(xMin, xMax, maxSp, mandatoryXs);
        const ewPositions = optimisePositions(yMin, yMax, maxSp, mandatoryYs);

        // Step 4: Place & clip internal beams
        const internalBeams = [];
        for (const xPos of nsPositions) {
            const segs = clipBeamToInterior(xPos, yMin - 1000, xPos, yMax + 1000, verts);
            const isMandatory = mandatoryXs.some(mx => Math.abs(xPos - mx) < 50);
            for (const seg of segs) {
                internalBeams.push({
                    ...seg,
                    type: 'IB1', category: isMandatory ? 'continuity' : 'internal',
                    direction: 'NS', width: bw, depth: bd,
                    get length() { return Math.hypot(this.x2 - this.x1, this.y2 - this.y1); },
                    get midpoint() { return { x: (this.x1 + this.x2) / 2, y: (this.y1 + this.y2) / 2 }; },
                });
            }
        }
        for (const yPos of ewPositions) {
            const segs = clipBeamToInterior(xMin - 1000, yPos, xMax + 1000, yPos, verts);
            const isMandatory = mandatoryYs.some(my => Math.abs(yPos - my) < 50);
            for (const seg of segs) {
                internalBeams.push({
                    ...seg,
                    type: 'IB1', category: isMandatory ? 'continuity' : 'internal',
                    direction: 'EW', width: bw, depth: bd,
                    get length() { return Math.hypot(this.x2 - this.x1, this.y2 - this.y1); },
                    get midpoint() { return { x: (this.x1 + this.x2) / 2, y: (this.y1 + this.y2) / 2 }; },
                });
            }
        }

        // Build layout result
        const layout = {
            edgeBeams,
            internalBeams,
            deflectedBeams: [],
            antiCrackReos: [],
            annotations: [],
            params, model, nsPositions, ewPositions,
            get nsXs() { return [...new Set(this.internalBeams.filter(b => b.direction === 'NS').map(b => b.x1))].sort((a, b) => a - b); },
            get ewYs() { return [...new Set(this.internalBeams.filter(b => b.direction === 'EW').map(b => b.y1))].sort((a, b) => a - b); },
            get allBeams() { return [...this.edgeBeams, ...this.internalBeams]; },
        };

        // Step 5: Reentrant annotations & anti-crack reo
        for (const rc of model.reentrantCorners) {
            const t = treatments[`${rc.point.x},${rc.point.y}`];

            // Anti-crack reo (Cl 5.3.7) — diagonal bar at 45°, 2m min
            const diagOff = 2000 / Math.SQRT2; // 2m diagonal in mm
            const acr = {
                x1: rc.point.x, y1: rc.point.y,
                x2: r(rc.point.x - rc.voidDx * diagOff),
                y2: r(rc.point.y - rc.voidDy * diagOff),
                spec: '2× 3-L8TM (or 1× 3-L11TM or 3-N12)',
                minLength: 2000,
            };
            layout.antiCrackReos.push(acr);

            layout.annotations.push({
                corner: rc.point,
                treatment: t.treatment,
                clause: t.clause,
                description: t.note,
                antiCrack: acr,
                voidQuadrant: rc.voidQuadrant,
                voidDx: rc.voidDx, voidDy: rc.voidDy,
                twoD: 2 * bd,
            });
        }

        // Step 6: Spacing verification
        layout.report = verifyLayout(layout);
        layout.log = log;

        return layout;
    }

    function verifyLayout(layout) {
        const { xMin, yMin, xMax, yMax } = layout.model.bounds;
        const maxSp = AS2870.MAX_SPACING[layout.params.siteClass] * 1000;
        const nsXs = layout.nsXs;
        const ewYs = layout.ewYs;

        // Spacing checks
        const allX = [xMin, ...nsXs, xMax].sort((a, b) => a - b);
        const allY = [yMin, ...ewYs, yMax].sort((a, b) => a - b);
        const ewSpacings = []; for (let i = 0; i < allX.length - 1; i++) ewSpacings.push(allX[i + 1] - allX[i]);
        const nsSpacings = []; for (let i = 0; i < allY.length - 1; i++) nsSpacings.push(allY[i + 1] - allY[i]);

        const maxEW = ewSpacings.length ? Math.max(...ewSpacings) : 0;
        const maxNS = nsSpacings.length ? Math.max(...nsSpacings) : 0;
        const ewPass = maxEW <= maxSp + 10;
        const nsPass = maxNS <= maxSp + 10;

        // 4.0m corner rule
        const cornerChecks = [];
        for (const c of layout.model.externalCorners) {
            const minDx = nsXs.length ? Math.min(...nsXs.map(x => Math.abs(x - c.point.x))) : Infinity;
            const minDy = ewYs.length ? Math.min(...ewYs.map(y => Math.abs(y - c.point.y))) : Infinity;
            cornerChecks.push({
                corner: c.point,
                dx: minDx, dy: minDy,
                pass: minDx <= 4010 && minDy <= 4010,
            });
        }
        const cornersPass = cornerChecks.every(c => c.pass);
        const overall = ewPass && nsPass && cornersPass;

        return {
            overall, ewPass, nsPass, cornersPass,
            ewSpacings, nsSpacings, maxEW, maxNS, maxSp,
            cornerChecks,
            edgeBeamCount: layout.edgeBeams.length,
            internalBeamCount: layout.internalBeams.length,
            deflectedBeamCount: layout.deflectedBeams.length,
            antiCrackReoCount: layout.antiCrackReos.length,
        };
    }

    // -----------------------------------------------------------------
    // FULL DESIGN CALCULATION
    // -----------------------------------------------------------------

    /**
     * Run all AS 2870 + AS 3600 calculations after the beam layout is solved.
     */
    function runDesignCalculations() {
        if (!state.layout) return null;

        const p = state.params;
        const model = state.layout.model;
        const L = Math.max(model.width, model.height); // critical dimension in mm
        const ys = p.ys || AS2870.YS_DEFAULTS[p.siteClass];
        const fc = parseInt(p.concreteGrade.replace('N', ''));

        // 1. Design loads
        const loads = AS2870.designLoads({
            slabThickness: p.slabThickness,
            finishes: p.finishes,
            partitions: p.partitions,
            Q: p.Q,
            storeys: p.storeys,
        });

        // 2. Average beam spacing
        const nsXs = state.layout.nsXs;
        const ewYs = state.layout.ewYs;
        const avgSpacing = (() => {
            const spacings = [];
            const allX = [model.bounds.xMin, ...nsXs, model.bounds.xMax].sort((a, b) => a - b);
            for (let i = 0; i < allX.length - 1; i++) spacings.push(allX[i + 1] - allX[i]);
            return spacings.length > 0 ? spacings.reduce((a, b) => a + b) / spacings.length : p.beamWidth;
        })();

        // 3. Simplified method (Cl 4.5)
        const simplified = AS2870.simplifiedMethod({
            siteClass: p.siteClass,
            ys: ys,
            L: L,
            constructionType: p.constructionType,
            concreteGrade: p.concreteGrade,
            beamWidth: p.beamWidth,
            beamDepth: p.beamDepth,
            slabThickness: p.slabThickness,
            beamSpacing: avgSpacing,
        });

        // 4. Estimate design actions from differential movement
        const conc = AS2870.CONCRETE[p.concreteGrade] || AS2870.CONCRETE.N25;
        const flangeW = AS2870.effectiveFlangeWidth(p.beamWidth, L * 0.7, avgSpacing);
        const section = AS2870.tSectionProperties({
            beamWidth: p.beamWidth, beamDepth: p.beamDepth,
            slabThickness: p.slabThickness, flangeWidth: flangeW,
        });
        const EI_beam = conc.Ec * section.I; // N.mm²

        const govDelta = Math.max(
            simplified.edgeHeave.delta_mm,
            simplified.centreHeave.delta_mm
        );
        const Mstar_est = AS3600.estimateDesignMoment(EI_beam, govDelta, L);
        const Vstar_est = AS3600.estimateDesignShear(Mstar_est.Mstar, L);

        // 5. Beam design (AS 3600)
        const beamDesign = AS3600.raftBeamDesign({
            beamWidth: p.beamWidth,
            beamDepth: p.beamDepth,
            slabThickness: p.slabThickness,
            fc: fc,
            topReo: p.topReo,
            botReo: p.botReo,
            topCover: p.topCover,
            botCover: p.botCover,
            Mstar_hog: Mstar_est.Mstar_kNm,
            Mstar_sag: Mstar_est.Mstar_kNm,
            Vstar: Vstar_est.Vstar_kN,
            flangeWidth: flangeW,
        });

        // 6. Bearing check
        const bearing = AS2870.bearingCheck({
            G_kPa: loads.G,
            Q_kPa: loads.Q,
            slabArea_m2: model.area / 1e6,
            beamArea_m2: 0,
            qu_kPa: p.qu,
        });

        // 7. Suggested reinforcement if ductility fails
        let topSuggestion = null, botSuggestion = null;
        if (beamDesign.ductility_hog && !beamDesign.ductility_hog.pass) {
            topSuggestion = AS3600.suggestReinforcement({
                beamWidth: p.beamWidth, beamDepth: p.beamDepth,
                slabThickness: p.slabThickness, fc: fc,
                flangeWidth: flangeW, face: 'top',
            });
        }
        if (beamDesign.ductility_sag && !beamDesign.ductility_sag.pass) {
            botSuggestion = AS3600.suggestReinforcement({
                beamWidth: p.beamWidth, beamDepth: p.beamDepth,
                slabThickness: p.slabThickness, fc: fc,
                flangeWidth: flangeW, face: 'bottom',
            });
        }

        state.calcResults = {
            loads, simplified, beamDesign, bearing,
            Mstar_kNm: Mstar_est.Mstar_kNm,
            Vstar_kN: Vstar_est.Vstar_kN,
            flangeWidth: flangeW,
            avgSpacing: avgSpacing,
            ys: ys,
            L: L,
            topSuggestion, botSuggestion,
        };

        return state.calcResults;
    }

    // -----------------------------------------------------------------
    // CANVAS RENDERER (design mode overlay)
    // -----------------------------------------------------------------

    function renderDesignMode(ctx, eng) {
        if (!state.active) return;

        ctx.save();
        const coords = eng.coords;

        if (state.phase === 'draw') {
            renderDrawPhase(ctx, coords, eng);
        } else if (state.phase === 'design' || state.phase === 'committed') {
            renderDesignPhase(ctx, coords, eng);
        }

        ctx.restore();
    }

    function renderDrawPhase(ctx, coords, eng) {
        const verts = state.vertices;
        if (verts.length === 0 && !state.mouseReal) return;

        // Draw placed vertices and edges
        if (verts.length > 0) {
            ctx.strokeStyle = '#E94560';
            ctx.lineWidth = 2;
            ctx.setLineDash([]);
            ctx.beginPath();
            const s0 = coords.realToScreen(verts[0].x, verts[0].y);
            ctx.moveTo(s0.x, s0.y);
            for (let i = 1; i < verts.length; i++) {
                const s = coords.realToScreen(verts[i].x, verts[i].y);
                ctx.lineTo(s.x, s.y);
            }
            // Rubber band to mouse (with shift-constrained angle)
            if (state.mouseReal) {
                let mx = snapMM(state.mouseReal.x), my = snapMM(state.mouseReal.y);
                if (state.shiftHeld && verts.length > 0) {
                    const prev = verts[verts.length - 1];
                    const constrained = constrainAngle(prev, { x: mx, y: my });
                    mx = snapMM(constrained.x);
                    my = snapMM(constrained.y);
                }
                const sm = coords.realToScreen(mx, my);
                ctx.lineTo(sm.x, sm.y);
            }
            ctx.stroke();

            // Vertex dots
            for (const v of verts) {
                const s = coords.realToScreen(v.x, v.y);
                ctx.fillStyle = '#E94560';
                ctx.beginPath(); ctx.arc(s.x, s.y, 5, 0, Math.PI * 2); ctx.fill();
                // Label
                ctx.fillStyle = '#333';
                ctx.font = '10px monospace';
                ctx.fillText(`(${(v.x/1000).toFixed(1)}, ${(v.y/1000).toFixed(1)})`, s.x + 8, s.y - 8);
            }
        }

        // Snap crosshair at mouse (constrained when shift held)
        if (state.mouseReal) {
            let sx = snapMM(state.mouseReal.x), sy = snapMM(state.mouseReal.y);
            if (state.shiftHeld && verts.length > 0) {
                const prev = verts[verts.length - 1];
                const constrained = constrainAngle(prev, { x: sx, y: sy });
                sx = snapMM(constrained.x);
                sy = snapMM(constrained.y);
            }
            const sc = coords.realToScreen(sx, sy);
            ctx.strokeStyle = 'rgba(233, 69, 96, 0.3)';
            ctx.lineWidth = 1;
            ctx.setLineDash([4, 4]);
            ctx.beginPath(); ctx.moveTo(sc.x, 0); ctx.lineTo(sc.x, eng.height); ctx.stroke();
            ctx.beginPath(); ctx.moveTo(0, sc.y); ctx.lineTo(eng.width, sc.y); ctx.stroke();
            ctx.setLineDash([]);
        }
    }

    function renderDesignPhase(ctx, coords, eng) {
        if (!state.layout) return;
        const model = state.layout.model;

        // Slab outline fill
        ctx.fillStyle = 'rgba(200, 210, 230, 0.15)';
        ctx.beginPath();
        const sv0 = coords.realToScreen(model.vertices[0].x, model.vertices[0].y);
        ctx.moveTo(sv0.x, sv0.y);
        for (let i = 1; i < model.vertices.length; i++) {
            const sv = coords.realToScreen(model.vertices[i].x, model.vertices[i].y);
            ctx.lineTo(sv.x, sv.y);
        }
        ctx.closePath();
        ctx.fill();

        // Slab outline
        ctx.strokeStyle = '#333333';
        ctx.lineWidth = 2.5;
        ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(sv0.x, sv0.y);
        for (let i = 1; i < model.vertices.length; i++) {
            const sv = coords.realToScreen(model.vertices[i].x, model.vertices[i].y);
            ctx.lineTo(sv.x, sv.y);
        }
        ctx.closePath();
        ctx.stroke();

        // Edge beams (solid, grey)
        for (const b of state.layout.edgeBeams) {
            drawBeamPair(ctx, coords, b, '#666666', 1.5, []);
        }

        // Internal beams (dashed, blue)
        for (const b of state.layout.internalBeams) {
            const color = b.category === 'continuity' ? '#2563EB' : '#60A5FA';
            drawBeamPair(ctx, coords, b, color, 1.2, [6, 3]);
        }

        // Anti-crack reo (red chain-dash)
        for (const acr of state.layout.antiCrackReos) {
            const p1 = coords.realToScreen(acr.x1, acr.y1);
            const p2 = coords.realToScreen(acr.x2, acr.y2);
            ctx.strokeStyle = '#DC2626';
            ctx.lineWidth = 1.5;
            ctx.setLineDash([3, 2, 1, 2]);
            ctx.beginPath(); ctx.moveTo(p1.x, p1.y); ctx.lineTo(p2.x, p2.y); ctx.stroke();
            ctx.setLineDash([]);
        }

        // Reentrant corner annotations
        for (const ann of state.layout.annotations) {
            const pc = coords.realToScreen(ann.corner.x, ann.corner.y);
            const zoom = eng.viewport.zoom;
            const radius = (ann.twoD / CONFIG.drawingScale) * zoom;

            ctx.strokeStyle = '#DC2626';
            ctx.lineWidth = 1.5;
            ctx.setLineDash([5, 3]);
            ctx.beginPath(); ctx.arc(pc.x, pc.y, radius, 0, Math.PI * 2); ctx.stroke();
            ctx.setLineDash([]);

            ctx.fillStyle = '#DC2626';
            ctx.font = 'bold 10px sans-serif';
            const label = ann.treatment === 'FIGURE_5_4' ? 'Fig 5.4 Detail' : 'FULL CONTINUITY';
            ctx.fillText(label, pc.x - 30, pc.y + radius + 14);
            ctx.font = '9px sans-serif';
            ctx.fillText('Cl 5.3.7 / 5.3.8', pc.x - 30, pc.y + radius + 26);
        }

        // Pier markers (beam intersections)
        drawPiers(ctx, coords);

        // Beam type labels
        drawBeamLabels(ctx, coords);

        // Overall dimension chains
        drawRaftDimensions(ctx, coords, eng);

        // Corner vertices
        for (const c of model.corners) {
            const sc = coords.realToScreen(c.point.x, c.point.y);
            ctx.fillStyle = c.type === 'REENTRANT' ? '#DC2626' : '#333333';
            ctx.beginPath(); ctx.arc(sc.x, sc.y, 3, 0, Math.PI * 2); ctx.fill();
        }
    }

    function drawBeamPair(ctx, coords, beam, color, lw, dash) {
        const dx = beam.x2 - beam.x1, dy = beam.y2 - beam.y1;
        const len = Math.hypot(dx, dy);
        if (len < 1) return;
        const px = -dy / len, py = dx / len;
        const off = beam.width / 2;

        ctx.strokeStyle = color;
        ctx.lineWidth = lw;
        ctx.setLineDash(dash);

        for (const sign of [1, -1]) {
            const s1 = coords.realToScreen(beam.x1 + sign * px * off, beam.y1 + sign * py * off);
            const s2 = coords.realToScreen(beam.x2 + sign * px * off, beam.y2 + sign * py * off);
            ctx.beginPath(); ctx.moveTo(s1.x, s1.y); ctx.lineTo(s2.x, s2.y); ctx.stroke();
        }
        ctx.setLineDash([]);
    }

    function drawPiers(ctx, coords) {
        if (!state.layout) return;
        const model = state.layout.model;
        const verts = model.vertices;
        const piers = new Set();

        for (const c of model.corners) piers.add(`${c.point.x},${c.point.y}`);
        for (const x of state.layout.nsXs) {
            for (const y of [model.bounds.yMin, model.bounds.yMax, ...state.layout.ewYs]) {
                piers.add(`${r(x)},${r(y)}`);
            }
        }
        for (const y of state.layout.ewYs) {
            piers.add(`${r(model.bounds.xMin)},${r(y)}`);
            piers.add(`${r(model.bounds.xMax)},${r(y)}`);
        }

        ctx.fillStyle = '#FFFFFF';
        ctx.strokeStyle = '#666666';
        ctx.lineWidth = 1;
        for (const key of piers) {
            const [px, py] = key.split(',').map(Number);
            if (pointInPolygon(px, py, verts) || pointOnBoundary(px, py, verts, 50)) {
                const sc = coords.realToScreen(px, py);
                ctx.beginPath(); ctx.arc(sc.x, sc.y, 3.5, 0, Math.PI * 2);
                ctx.fill(); ctx.stroke();
            }
        }
    }

    function drawBeamLabels(ctx, coords) {
        if (!state.layout) return;
        ctx.font = 'bold 9px sans-serif';

        for (const beam of state.layout.allBeams) {
            if (beam.length < 1500) continue;
            const mid = beam.midpoint;
            const sm = coords.realToScreen(mid.x, mid.y);
            const dx = beam.x2 - beam.x1, dy = beam.y2 - beam.y1;
            const len = Math.hypot(dx, dy);
            if (len < 1) continue;
            const offX = (-dy / len) * 12, offY = (dx / len) * 12;
            ctx.fillStyle = beam.category === 'edge' ? '#666666' : '#2563EB';
            ctx.fillText(beam.type, sm.x + offX - 8, sm.y - offY + 4);
        }
    }

    function drawRaftDimensions(ctx, coords, eng) {
        if (!state.layout) return;
        const model = state.layout.model;
        const { xMin, yMin, xMax, yMax } = model.bounds;
        const dimOff = 1500 / CONFIG.drawingScale * eng.viewport.zoom; // offset in screen px

        ctx.strokeStyle = '#999999';
        ctx.fillStyle = '#666666';
        ctx.lineWidth = 0.8;
        ctx.font = '10px monospace';
        ctx.setLineDash([]);

        // Bottom overall dimension
        const s1 = coords.realToScreen(xMin, yMin);
        const s2 = coords.realToScreen(xMax, yMin);
        const yOff = s1.y + dimOff + 5;
        ctx.beginPath(); ctx.moveTo(s1.x, yOff); ctx.lineTo(s2.x, yOff); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(s1.x, yOff - 4); ctx.lineTo(s1.x, yOff + 4); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(s2.x, yOff - 4); ctx.lineTo(s2.x, yOff + 4); ctx.stroke();
        ctx.fillText(`${(model.width / 1000).toFixed(1)}m`, (s1.x + s2.x) / 2 - 15, yOff + 14);

        // Left overall dimension
        const s3 = coords.realToScreen(xMin, yMin);
        const s4 = coords.realToScreen(xMin, yMax);
        const xOff = s3.x - dimOff - 5;
        ctx.beginPath(); ctx.moveTo(xOff, s3.y); ctx.lineTo(xOff, s4.y); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(xOff - 4, s3.y); ctx.lineTo(xOff + 4, s3.y); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(xOff - 4, s4.y); ctx.lineTo(xOff + 4, s4.y); ctx.stroke();

        ctx.save();
        ctx.translate(xOff - 10, (s3.y + s4.y) / 2);
        ctx.rotate(-Math.PI / 2);
        ctx.fillText(`${(model.height / 1000).toFixed(1)}m`, -15, 0);
        ctx.restore();
    }

    function snapMM(val) {
        return Math.round(val / state.snapMM) * state.snapMM;
    }

    // -----------------------------------------------------------------
    // PANEL UI — builds the raft design sidebar content
    // -----------------------------------------------------------------

    function buildPanel() {
        const panel = document.getElementById('raft-design-panel');
        if (!panel) return;

        let html = '';

        // Header
        html += `<div class="raft-panel-header">
            <span class="raft-panel-title">RAFT SLAB DESIGN</span>
            <span class="raft-panel-subtitle">AS 2870-2011</span>
        </div>`;

        if (state.phase === 'draw') {
            html += buildDrawPhasePanel();
        } else {
            html += buildDesignPhasePanel();
        }

        panel.innerHTML = html;
        bindPanelEvents();
    }

    function buildDrawPhasePanel() {
        const p = state.params;
        return `
        <div class="raft-section">
            <div class="raft-section-title">DESIGN PARAMETERS</div>
            ${paramRow('Site Class', `<select id="raft-site-class" class="raft-input">
                ${['A','S','M','H1','H2','E'].map(c => `<option value="${c}" ${c===p.siteClass?'selected':''}>${c} (max ${AS2870.MAX_SPACING[c]}m)</option>`).join('')}
            </select>`)}
            ${paramRow('Construction', `<select id="raft-construction" class="raft-input">
                ${Object.keys(AS2870.DEFLECTION_LIMITS).map(c => `<option value="${c}" ${c===p.constructionType?'selected':''}>${c}</option>`).join('')}
            </select>`)}
            ${paramRow('Beam Width', `<input type="number" id="raft-bw" class="raft-input" value="${p.beamWidth}" step="50" min="200" max="600"> mm`)}
            ${paramRow('Beam Depth', `<input type="number" id="raft-bd" class="raft-input" value="${p.beamDepth}" step="50" min="300" max="800"> mm`)}
            ${paramRow('Concrete', `<select id="raft-concrete" class="raft-input">
                ${['N20','N25','N32','N40'].map(g => `<option value="${g}" ${g===p.concreteGrade?'selected':''}>${g}</option>`).join('')}
            </select>`)}
            ${paramRow('Slab Thick.', `<input type="number" id="raft-slab-t" class="raft-input" value="${p.slabThickness}" step="5" min="85" max="200"> mm`)}
        </div>

        <div class="raft-section">
            <div class="raft-section-title">SLAB OUTLINE</div>
            <p class="raft-hint">Click on the drawing sheet to place slab vertices. Right-click or press Enter to close and solve.</p>
            <div id="raft-vertex-list" class="raft-vertex-list">
                ${state.vertices.map((v, i) => `<div class="raft-vertex-item" data-idx="${i}">V${i}: (${(v.x/1000).toFixed(1)}, ${(v.y/1000).toFixed(1)})</div>`).join('')}
            </div>
            <div class="raft-btn-row">
                <button id="raft-undo-btn" class="raft-btn secondary" ${state.vertices.length===0?'disabled':''}>Undo</button>
                <button id="raft-clear-btn" class="raft-btn secondary" ${state.vertices.length===0?'disabled':''}>Clear</button>
            </div>
            <button id="raft-solve-btn" class="raft-btn primary" ${state.vertices.length<3?'disabled':''}>Close & Solve</button>
        </div>

        <div class="raft-section">
            <div class="raft-section-title">PRESET SLABS</div>
            <div class="raft-preset-grid">
                ${['rect','lshape','ushape','tshape'].map(p =>
                    `<button class="raft-btn secondary preset" data-preset="${p}">${p.charAt(0).toUpperCase()+p.slice(1)}</button>`
                ).join('')}
            </div>
        </div>`;
    }

    function buildDesignPhasePanel() {
        const p = state.params;
        const rep = state.layout ? state.layout.report : null;
        const calc = state.calcResults;

        let html = '';

        // Quick parameters (editable in design phase too)
        html += `<div class="raft-section">
            <div class="raft-section-title">DESIGN PARAMETERS</div>
            ${paramRow('Site Class', `<select id="raft-site-class" class="raft-input">
                ${['A','S','M','H1','H2','E'].map(c => `<option value="${c}" ${c===p.siteClass?'selected':''}>${c}</option>`).join('')}
            </select>`)}
            ${paramRow('Beam', `${p.beamWidth}W × ${p.beamDepth}D`)}
            ${paramRow('Concrete', p.concreteGrade)}
            ${paramRow('Slab', `${p.slabThickness}mm`)}
            <button id="raft-edit-params-btn" class="raft-btn secondary" style="margin-top:6px;">Edit Parameters</button>
        </div>`;

        // Layout compliance
        if (rep) {
            html += `<div class="raft-section">
                <div class="raft-section-title">LAYOUT COMPLIANCE</div>
                <div class="raft-check ${rep.overall?'pass':'fail'}">
                    ${rep.overall?'✓ ALL CHECKS PASS':'✗ ISSUES FOUND'}
                </div>
                <div class="raft-check-row ${rep.ewPass?'pass':'fail'}">
                    E-W Spacing: max ${(rep.maxEW/1000).toFixed(2)}m ${rep.ewPass?'≤':'>'} ${(rep.maxSp/1000).toFixed(1)}m
                </div>
                <div class="raft-check-row ${rep.nsPass?'pass':'fail'}">
                    N-S Spacing: max ${(rep.maxNS/1000).toFixed(2)}m ${rep.nsPass?'≤':'>'} ${(rep.maxSp/1000).toFixed(1)}m
                </div>
                <div class="raft-check-row ${rep.cornersPass?'pass':'fail'}">
                    4.0m Corner Rule: ${rep.cornersPass?'PASS':'FAIL'}
                </div>
                <div class="raft-summary">
                    ${rep.edgeBeamCount} edge + ${rep.internalBeamCount} internal beams
                    ${rep.antiCrackReoCount > 0 ? `| ${rep.antiCrackReoCount} anti-crack reos` : ''}
                </div>
            </div>`;
        }

        // Design calculations
        if (calc) {
            html += buildCalcSummaryPanel(calc);
        }

        // Actions
        html += `<div class="raft-section">
            <div class="raft-section-title">ACTIONS</div>
            <button id="raft-re-solve-btn" class="raft-btn secondary">Re-Solve Layout</button>
            <button id="raft-run-calcs-btn" class="raft-btn secondary" style="margin-top:6px;">Run Design Calcs</button>
            <button id="raft-commit-btn" class="raft-btn primary" style="margin-top:8px;"
                ${state.phase === 'committed' ? 'disabled' : ''}>
                ${state.phase === 'committed' ? '✓ Committed to Plan' : 'Commit to Plan'}
            </button>
            <button id="raft-reset-btn" class="raft-btn secondary" style="margin-top:6px;">Reset / New Slab</button>
        </div>`;

        return html;
    }

    function buildCalcSummaryPanel(calc) {
        let html = '';

        // Simplified method
        if (calc.simplified) {
            const s = calc.simplified;
            html += `<div class="raft-section collapsible">
                <div class="raft-section-title clickable" data-collapse="simp">
                    STIFFNESS CHECK (Cl 4.5) ${s.pass ? '<span class="pass-badge">PASS</span>' : '<span class="fail-badge">FAIL</span>'}
                </div>
                <div id="collapse-simp" class="raft-collapse-body">
                    <div class="raft-calc-row">ys = ${s.ys_mm.toFixed(1)} mm | L = ${(s.L_mm/1000).toFixed(1)} m</div>
                    <div class="raft-calc-row">EI = ${s.EI_kNm2_perM.toFixed(1)} kN.m²/m</div>
                    <div class="raft-calc-row">Φ = ${s.stiffnessParam.toExponential(3)}</div>
                    <div class="raft-check-row ${s.edgeHeave.pass?'pass':'fail'}">
                        Edge Heave: Δ = ${s.edgeHeave.delta_mm.toFixed(1)} mm ≤ ${s.deltaAllow_mm.toFixed(1)} mm
                    </div>
                    <div class="raft-check-row ${s.centreHeave.pass?'pass':'fail'}">
                        Centre Heave: Δ = ${s.centreHeave.delta_mm.toFixed(1)} mm ≤ ${s.deltaAllow_mm.toFixed(1)} mm
                    </div>
                </div>
            </div>`;
        }

        // Beam design
        if (calc.beamDesign) {
            const bd = calc.beamDesign;
            html += `<div class="raft-section collapsible">
                <div class="raft-section-title clickable" data-collapse="beam">
                    BEAM DESIGN (AS 3600) ${bd.overallPass ? '<span class="pass-badge">PASS</span>' : '<span class="fail-badge">FAIL</span>'}
                </div>
                <div id="collapse-beam" class="raft-collapse-body">
                    <div class="raft-calc-row">M* ≈ ${calc.Mstar_kNm.toFixed(1)} kN.m | V* ≈ ${calc.Vstar_kN.toFixed(1)} kN</div>
                    ${bd.sagging ? `<div class="raft-check-row ${bd.sagging.pass?'pass':'fail'}">
                        Sag: φMu = ${bd.sagging.phiMu_kNm.toFixed(1)} kN.m ${bd.sagging.pass?'≥':'<'} ${bd.sagging.Mstar.toFixed(1)}
                    </div>` : ''}
                    ${bd.hogging ? `<div class="raft-check-row ${bd.hogging.pass?'pass':'fail'}">
                        Hog: φMu = ${bd.hogging.phiMu_kNm.toFixed(1)} kN.m ${bd.hogging.pass?'≥':'<'} ${bd.hogging.Mstar.toFixed(1)}
                    </div>` : ''}
                    ${bd.ductility_sag ? `<div class="raft-check-row ${bd.ductility_sag.pass?'pass':'fail'}">
                        Ductility (sag): Mu/Mcr = ${bd.ductility_sag.ratio.toFixed(2)} ${bd.ductility_sag.pass?'≥':'<'} 1.2
                    </div>` : ''}
                    ${bd.ductility_hog ? `<div class="raft-check-row ${bd.ductility_hog.pass?'pass':'fail'}">
                        Ductility (hog): Mu/Mcr = ${bd.ductility_hog.ratio.toFixed(2)} ${bd.ductility_hog.pass?'≥':'<'} 1.2
                    </div>` : ''}
                    ${bd.shear ? `<div class="raft-check-row ${bd.shear.pass?'pass':'fail'}">
                        Shear: φVuc = ${bd.shear.phiVuc_kN.toFixed(1)} kN ${bd.shear.pass?'≥':'<'} ${bd.shear.Vstar.toFixed(1)} kN
                    </div>` : ''}
                </div>
            </div>`;
        }

        // Bearing
        if (calc.bearing) {
            const b = calc.bearing;
            html += `<div class="raft-section collapsible">
                <div class="raft-section-title clickable" data-collapse="bear">
                    BEARING (Cl 4.3) ${b.pass ? '<span class="pass-badge">PASS</span>' : '<span class="fail-badge">FAIL</span>'}
                </div>
                <div id="collapse-bear" class="raft-collapse-body">
                    <div class="raft-calc-row">q = ${b.q_applied.toFixed(1)} kPa ≤ qu/3 = ${b.q_allow.toFixed(1)} kPa</div>
                </div>
            </div>`;
        }

        return html;
    }

    function paramRow(label, value) {
        return `<div class="raft-param-row"><span class="raft-param-label">${label}</span><span class="raft-param-value">${value}</span></div>`;
    }

    // -----------------------------------------------------------------
    // PANEL EVENT BINDING
    // -----------------------------------------------------------------

    function bindPanelEvents() {
        // Parameter inputs
        const bind = (id, key, parse) => {
            const el = document.getElementById(id);
            if (!el) return;
            el.addEventListener('change', () => {
                state.params[key] = parse ? parse(el.value) : el.value;
                if (state.phase === 'design' && state.layout) {
                    // Re-solve on param change
                    solve();
                }
            });
        };
        bind('raft-site-class', 'siteClass');
        bind('raft-construction', 'constructionType');
        bind('raft-bw', 'beamWidth', Number);
        bind('raft-bd', 'beamDepth', Number);
        bind('raft-concrete', 'concreteGrade');
        bind('raft-slab-t', 'slabThickness', Number);

        // Draw phase buttons
        const undoBtn = document.getElementById('raft-undo-btn');
        if (undoBtn) undoBtn.addEventListener('click', () => { state.vertices.pop(); buildPanel(); engine.requestRender(); });
        const clearBtn = document.getElementById('raft-clear-btn');
        if (clearBtn) clearBtn.addEventListener('click', resetDraw);
        const solveBtn = document.getElementById('raft-solve-btn');
        if (solveBtn) solveBtn.addEventListener('click', () => { state.isClosed = true; solve(); });

        // Presets
        document.querySelectorAll('.preset[data-preset]').forEach(btn => {
            btn.addEventListener('click', () => loadPreset(btn.dataset.preset));
        });

        // Design phase buttons
        const reSolveBtn = document.getElementById('raft-re-solve-btn');
        if (reSolveBtn) reSolveBtn.addEventListener('click', solve);
        const runCalcsBtn = document.getElementById('raft-run-calcs-btn');
        if (runCalcsBtn) runCalcsBtn.addEventListener('click', () => { runDesignCalculations(); buildPanel(); });
        const commitBtn = document.getElementById('raft-commit-btn');
        if (commitBtn) commitBtn.addEventListener('click', commitToPlan);
        const resetBtn = document.getElementById('raft-reset-btn');
        if (resetBtn) resetBtn.addEventListener('click', resetDraw);
        const editParamsBtn = document.getElementById('raft-edit-params-btn');
        if (editParamsBtn) editParamsBtn.addEventListener('click', () => { state.phase = 'draw'; state.layout = null; state.calcResults = null; buildPanel(); engine.requestRender(); });

        // Collapsible sections
        document.querySelectorAll('.raft-section-title.clickable').forEach(el => {
            el.addEventListener('click', () => {
                const target = document.getElementById('collapse-' + el.dataset.collapse);
                if (target) target.classList.toggle('collapsed');
            });
        });
    }

    // -----------------------------------------------------------------
    // PRESETS (real-world mm coordinates)
    // -----------------------------------------------------------------

    const PRESETS = {
        rect:   [{x:0,y:0},{x:18000,y:0},{x:18000,y:12000},{x:0,y:12000}],
        lshape: [{x:0,y:0},{x:28000,y:0},{x:28000,y:8000},{x:20000,y:8000},{x:20000,y:14000},{x:0,y:14000}],
        ushape: [{x:0,y:0},{x:32000,y:0},{x:32000,y:18000},{x:26000,y:18000},{x:26000,y:8000},{x:6000,y:8000},{x:6000,y:18000},{x:0,y:18000}],
        tshape: [{x:0,y:10000},{x:14000,y:10000},{x:14000,y:0},{x:22000,y:0},{x:22000,y:10000},{x:36000,y:10000},{x:36000,y:18000},{x:0,y:18000}],
    };

    function loadPreset(name) {
        const preset = PRESETS[name];
        if (!preset) return;

        // Place presets at a sensible offset from the drawing origin
        const offset = { x: 5000, y: 5000 };
        state.vertices = preset.map(v => ({ x: v.x + offset.x, y: v.y + offset.y }));
        state.isClosed = true;
        solve();
    }

    // -----------------------------------------------------------------
    // SOLVE & RESET
    // -----------------------------------------------------------------

    function solve() {
        if (state.vertices.length < 3) return;

        const model = buildSlabModel(state.vertices);
        state.layout = solveBeamLayout(model, state.params);
        state.phase = 'design';

        // Auto-run design calcs
        runDesignCalculations();

        buildPanel();
        engine.requestRender();
    }

    function resetDraw() {
        state.vertices = [];
        state.isClosed = false;
        state.layout = null;
        state.calcResults = null;
        state.phase = 'draw';
        buildPanel();
        engine.requestRender();
    }

    // -----------------------------------------------------------------
    // COMMIT TO PLAN — converts design into real StructuralSketch elements
    // -----------------------------------------------------------------

    function commitToPlan() {
        if (!state.layout || state.phase === 'committed') return;

        const commands = [];
        const p = state.params;

        // Ensure schedule types exist for edge and internal beams
        ensureRaftScheduleTypes();

        // 1. Create slab outline as a closed polyline on S-SLAB layer
        const outlineEl = {
            id: generateId(),
            type: 'polyline',
            layer: 'S-SLAB',
            level: project.activeLevel || 'GF',
            points: state.vertices.map(v => ({ x: v.x, y: v.y })),
            closed: true,
            tag: 'RAFT SLAB',
            designData: {
                siteClass: p.siteClass,
                constructionType: p.constructionType,
                concreteGrade: p.concreteGrade,
                beamWidth: p.beamWidth,
                beamDepth: p.beamDepth,
                slabThickness: p.slabThickness,
                calcResults: state.calcResults ? {
                    simplified: state.calcResults.simplified ? {
                        pass: state.calcResults.simplified.pass,
                        edgeHeaveDelta: state.calcResults.simplified.edgeHeave.delta_mm,
                        centreHeaveDelta: state.calcResults.simplified.centreHeave.delta_mm,
                        deltaAllow: state.calcResults.simplified.deltaAllow_mm,
                    } : null,
                    beamDesign: state.calcResults.beamDesign ? {
                        overallPass: state.calcResults.beamDesign.overallPass,
                    } : null,
                    bearing: state.calcResults.bearing ? {
                        pass: state.calcResults.bearing.pass,
                    } : null,
                } : null,
            },
        };
        commands.push(addElementCmd(project.elements, outlineEl));

        // 2. Create edge beams as lines on S-FTNG layer
        for (const beam of state.layout.edgeBeams) {
            const el = {
                id: generateId(),
                type: 'line',
                layer: 'S-FTNG',
                level: project.activeLevel || 'GF',
                x1: beam.x1, y1: beam.y1,
                x2: beam.x2, y2: beam.y2,
                typeRef: 'EB1',
                tag: 'EB1',
                raftRef: outlineEl.id,
            };
            commands.push(addElementCmd(project.elements, el));
        }

        // 3. Create internal beams as lines on S-FTNG layer
        for (const beam of state.layout.internalBeams) {
            const el = {
                id: generateId(),
                type: 'line',
                layer: 'S-FTNG',
                level: project.activeLevel || 'GF',
                x1: beam.x1, y1: beam.y1,
                x2: beam.x2, y2: beam.y2,
                typeRef: 'IB1',
                tag: 'IB1',
                raftRef: outlineEl.id,
            };
            commands.push(addElementCmd(project.elements, el));
        }

        // Execute as a batch command (single undo step)
        history.execute(batchCmd(commands, 'Commit raft slab design'));

        state.phase = 'committed';
        buildPanel();
        engine.requestRender();
    }

    function ensureRaftScheduleTypes() {
        // Add raft beam types to the schedule if they don't exist
        if (!project.scheduleTypes.raftbeam) {
            project.scheduleTypes.raftbeam = {};
        }
        const p = state.params;
        if (!project.scheduleTypes.raftbeam.EB1) {
            project.scheduleTypes.raftbeam.EB1 = {
                description: 'Edge Beam',
                width: p.beamWidth,
                depth: p.beamDepth,
                topReo: p.topReo,
                botReo: p.botReo,
                concreteGrade: p.concreteGrade,
                color: '#666666',
            };
        }
        if (!project.scheduleTypes.raftbeam.IB1) {
            project.scheduleTypes.raftbeam.IB1 = {
                description: 'Internal Beam',
                width: p.beamWidth,
                depth: p.beamDepth,
                topReo: p.topReo,
                botReo: p.botReo,
                concreteGrade: p.concreteGrade,
                color: '#60A5FA',
            };
        }
    }

    // -----------------------------------------------------------------
    // MODE ENTRY / EXIT
    // -----------------------------------------------------------------

    function enterMode() {
        if (state.active) return;
        state.active = true;
        state._prevTool = window.activeTool || 'select';
        state._prevLayer = project.activeLayerId;
        window.activeTool = 'raft-design';

        // Show the panel
        const panel = document.getElementById('raft-design-panel');
        if (panel) panel.classList.remove('hidden');

        // Register render callback if not already
        if (!state._renderRegistered) {
            engine.onRender(renderDesignMode);
            state._renderRegistered = true;
        }

        buildPanel();
        engine.requestRender();

        // Update status bar
        const statusTool = document.getElementById('status-tool');
        if (statusTool) statusTool.textContent = 'Raft Slab Design';
    }

    function exitMode() {
        if (!state.active) return;
        state.active = false;
        window.activeTool = state._prevTool || 'select';
        project.activeLayerId = state._prevLayer || 'S-BEAM';

        const panel = document.getElementById('raft-design-panel');
        if (panel) panel.classList.add('hidden');

        engine.requestRender();

        const statusTool = document.getElementById('status-tool');
        if (statusTool) statusTool.textContent = state._prevTool || 'Select';
    }

    function toggleMode() {
        if (state.active) exitMode();
        else enterMode();
    }

    // -----------------------------------------------------------------
    // ANGLE CONSTRAINT HELPER (Shift = ortho / 45°)
    // -----------------------------------------------------------------

    /**
     * Constrain a point to the nearest 45° increment from a reference point.
     * Snaps to 0°, 45°, 90°, 135°, 180°, 225°, 270°, 315°.
     */
    function constrainAngle(ref, pt) {
        const dx = pt.x - ref.x;
        const dy = pt.y - ref.y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist === 0) return { x: pt.x, y: pt.y };

        const angle = Math.atan2(dy, dx);
        const snap45 = Math.round(angle / (Math.PI / 4)) * (Math.PI / 4);
        return {
            x: ref.x + dist * Math.cos(snap45),
            y: ref.y + dist * Math.sin(snap45)
        };
    }

    // -----------------------------------------------------------------
    // CANVAS EVENT HANDLERS (active only in raft design mode)
    // -----------------------------------------------------------------

    function handleMouseMove(e) {
        if (!state.active) return;
        const rect = engine.mainCanvas.getBoundingClientRect();
        const px = e.clientX - rect.left;
        const py = e.clientY - rect.top;
        state.mouseReal = engine.coords.screenToReal(px, py);
        state.shiftHeld = e.shiftKey;
        engine.requestRender();
    }

    function handleMouseDown(e) {
        if (!state.active || state.phase !== 'draw') return;
        if (e.button !== 0) return; // left click only
        if (e.ctrlKey || e.metaKey) return; // let pan/zoom through (shift used for constraint)

        const rect = engine.mainCanvas.getBoundingClientRect();
        const px = e.clientX - rect.left;
        const py = e.clientY - rect.top;
        const real = engine.coords.screenToReal(px, py);
        let v = { x: snapMM(real.x), y: snapMM(real.y) };

        // Apply 45° angle constraint when Shift held and we have a previous vertex
        if (e.shiftKey && state.vertices.length > 0) {
            const prev = state.vertices[state.vertices.length - 1];
            const constrained = constrainAngle(prev, v);
            v = { x: snapMM(constrained.x), y: snapMM(constrained.y) };
        }

        state.vertices.push(v);
        buildPanel();
        engine.requestRender();
        e.stopPropagation();
    }

    function handleDblClick(e) {
        if (!state.active || state.phase !== 'draw') return;
        // dblclick fires after two mousedowns, so the second click added a
        // duplicate vertex — remove it before closing
        if (state.vertices.length >= 4) {
            state.vertices.pop();
        }
        if (state.vertices.length >= 3) {
            state.isClosed = true;
            solve();
            e.preventDefault();
            e.stopPropagation();
        }
    }

    function handleContextMenu(e) {
        if (!state.active || state.phase !== 'draw') return;
        e.preventDefault();
        if (state.vertices.length >= 3) {
            state.isClosed = true;
            solve();
        }
    }

    function handleKeyDown(e) {
        if (!state.active) return;
        if (e.key === 'Enter' && state.phase === 'draw' && state.vertices.length >= 3) {
            state.isClosed = true;
            solve();
            e.preventDefault();
        }
        if (e.key === 'Escape') {
            if (state.phase === 'draw' && state.vertices.length > 0) {
                resetDraw();
            } else {
                exitMode();
            }
            e.preventDefault();
        }
    }

    // Wire up events on the canvas container
    function initEvents() {
        const c = engine.mainCanvas;
        c.addEventListener('mousemove', handleMouseMove);
        c.addEventListener('mousedown', handleMouseDown, true);
        c.addEventListener('dblclick', handleDblClick, true);
        c.addEventListener('contextmenu', handleContextMenu);
        document.addEventListener('keydown', handleKeyDown);
    }

    // -----------------------------------------------------------------
    // INIT
    // -----------------------------------------------------------------

    // Deferred init — called after DOM is ready
    function init() {
        initEvents();

        // Wire the ribbon button
        const btn = document.getElementById('btn-raft-design');
        if (btn) {
            btn.addEventListener('click', toggleMode);
        }
    }

    // Auto-init when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        // DOM already loaded (module loaded after other scripts)
        setTimeout(init, 0);
    }

    // -----------------------------------------------------------------
    // PUBLIC API
    // -----------------------------------------------------------------

    return {
        state,
        enterMode,
        exitMode,
        toggleMode,
        solve,
        resetDraw,
        commitToPlan,
        runDesignCalculations,
        buildPanel,
        // Expose for testing / external use
        buildSlabModel,
        solveBeamLayout,
        verifyLayout,
    };
})();
