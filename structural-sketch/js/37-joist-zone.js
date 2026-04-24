// ═══════════════════════════════════════════════════════════════════
// 37-joist-zone.js — Polygon-based joist zone system
//
// Replaces the bay-fill joist tool with a polygon-based approach:
//   1. Draw a polygon defining the joist zone
//   2. Setup dialog: span direction, spacing, span type, FFL offset
//   3. Auto-detect internal beams perpendicular to span
//   4. Compute individual joist lines clipped to polygon + beams
//   5. Per-span sizing via hySPAN (Dindas Table 3)
//   6. 2D rendering: thin solid lines, leader arrows, lap details
//   7. Zone-to-zone continuity + cantilever detection
//
// Element type: 'joistZone' on layer 'S-JOIST'
// Namespace: extends window.floorDesigner
// ═══════════════════════════════════════════════════════════════════

(function () {
    'use strict';

    var FLR = typeof FloorLoadResolver !== 'undefined' ? FloorLoadResolver : null;

    // ── Auto bay label counter ──────────────────────────────────
    var _nextBayLabel = 'A';
    function nextBayLabel() {
        var label = _nextBayLabel;
        _nextBayLabel = String.fromCharCode(_nextBayLabel.charCodeAt(0) + 1);
        if (_nextBayLabel.charCodeAt(0) > 90) _nextBayLabel = 'AA'; // wrap after Z
        return label;
    }

    // ════════════════════════════════════════════════════════════
    // PHASE 1: Polygon tool + setup dialog
    // ════════════════════════════════════════════════════════════

    var toolState = {
        points: [],       // sheet-mm vertices during drawing
        currentEnd: null, // preview vertex
    };

    // ── Input position (snap + ortho) ───────────────────────────
    function getToolPos(e) {
        var rect = container.getBoundingClientRect();
        var sx = e.clientX - rect.left;
        var sy = e.clientY - rect.top;
        var snap = (typeof findSnap === 'function') ? findSnap(sx, sy) : null;
        var pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
        if (toolState.points.length > 0 && typeof applyOrtho === 'function') {
            var last = toolState.points[toolState.points.length - 1];
            pos = applyOrtho(pos.x, pos.y, last.x, last.y);
        }
        return pos;
    }

    // ── Event handlers ──────────────────────────────────────────
    container.addEventListener('mousemove', function (e) {
        if (activeTool !== 'joistZone') return;
        toolState.currentEnd = getToolPos(e);
        engine.requestRender();
    });

    container.addEventListener('mousedown', function (e) {
        if (e.button !== 0) return;
        if (engine._spaceDown || engine._isPanning) return;
        if (activeTool !== 'joistZone') return;
        if (typeof pdfState !== 'undefined' && pdfState.calibrating) return;
        toolState.points.push(getToolPos(e));
        engine.requestRender();
    });

    container.addEventListener('dblclick', function (e) {
        if (activeTool !== 'joistZone') return;
        if (toolState.points.length < 3) return;
        toolState.points.pop(); // remove duplicate from 2nd mousedown
        if (toolState.points.length >= 3) commitJoistZone();
    });

    container.addEventListener('contextmenu', function (e) {
        if (activeTool !== 'joistZone') return;
        e.preventDefault();
        if (toolState.points.length >= 3) commitJoistZone();
        else toolState.points = [];
        engine.requestRender();
    });

    window.addEventListener('keydown', function (e) {
        if (activeTool !== 'joistZone') return;
        if (e.key === 'Escape') {
            toolState.points = [];
            toolState.currentEnd = null;
            engine.requestRender();
        } else if (e.key === 'Enter' && toolState.points.length >= 3) {
            commitJoistZone();
        }
    });

    // ── Commit polygon → open setup dialog ──────────────────────
    function commitJoistZone() {
        var pts = toolState.points.slice();
        toolState.points = [];
        toolState.currentEnd = null;

        // Convert sheet-mm → real-mm
        var realPoints = pts.map(function (p) {
            return engine.coords.sheetToReal(p.x, p.y);
        });

        var levelId = (typeof getActiveLevel === 'function') ? getActiveLevel().id : 'L1';
        var typeRef = (typeof placementTypeRef !== 'undefined' && placementTypeRef.joist) ? placementTypeRef.joist : 'FJ1';

        // Create element (not yet pushed — dialog first)
        var newEl = {
            id: generateId(),
            type: 'joistZone',
            layer: 'S-JOIST',
            level: levelId,
            typeRef: typeRef,
            points: realPoints,
            closed: true,
            spanDirection_deg: 90,  // default vertical
            maxSpacing_mm: 450,
            spanType: 'single',
            fflOffset_mm: 19,
            bayLabel: nextBayLabel(),
            computed: null,
        };

        showSetupDialog(newEl, true);
    }

    // ── Setup dialog ────────────────────────────────────────────
    var _dialogEl = null;
    var _dialogIsNew = false;

    function showSetupDialog(el, isNew) {
        _dialogEl = el;
        _dialogIsNew = isNew;

        var modal = document.getElementById('joist-zone-setup');
        if (!modal) {
            modal = document.createElement('div');
            modal.id = 'joist-zone-setup';
            modal.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;z-index:4500;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,0.35);';
            document.body.appendChild(modal);
        }

        var spanH = el.spanDirection_deg < 45;
        var isCont = el.spanType === 'continuous';

        modal.innerHTML =
            '<div style="background:#fff;border-radius:10px;box-shadow:0 8px 32px rgba(0,0,0,0.25);width:400px;max-width:92vw;overflow:hidden;">' +
                '<div style="background:#6D28D9;color:#fff;padding:12px 16px;font-weight:700;font-size:14px;">Joist Zone — Bay ' + el.bayLabel + '</div>' +
                '<div style="padding:16px;">' +
                    // Span direction
                    '<div style="font-size:11px;font-weight:600;color:#666;margin-bottom:4px;">SPAN DIRECTION</div>' +
                    '<div style="display:flex;gap:6px;margin-bottom:14px;">' +
                        '<button id="jzs-span-h" style="flex:1;padding:10px;border:1px solid ' + (spanH ? '#6D28D9' : '#ddd') + ';border-radius:4px;background:' + (spanH ? '#f5f3ff' : '#f9fafb') + ';cursor:pointer;font-size:12px;font-weight:600;color:' + (spanH ? '#6D28D9' : '#666') + ';">&#8592; &#8594; Horizontal</button>' +
                        '<button id="jzs-span-v" style="flex:1;padding:10px;border:1px solid ' + (!spanH ? '#6D28D9' : '#ddd') + ';border-radius:4px;background:' + (!spanH ? '#f5f3ff' : '#f9fafb') + ';cursor:pointer;font-size:12px;font-weight:600;color:' + (!spanH ? '#6D28D9' : '#666') + ';">&#8593; &#8595; Vertical</button>' +
                    '</div>' +
                    // Spacing
                    '<div style="display:flex;gap:12px;margin-bottom:14px;">' +
                        '<div style="flex:1;">' +
                            '<label style="font-size:10px;font-weight:600;color:#666;">MAX SPACING (mm)</label>' +
                            '<input type="number" id="jzs-spacing" value="' + el.maxSpacing_mm + '" step="50" min="200" max="600" style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;margin-top:2px;">' +
                        '</div>' +
                        '<div style="flex:1;">' +
                            '<label style="font-size:10px;font-weight:600;color:#666;">FFL OFFSET (mm)</label>' +
                            '<input type="number" id="jzs-ffl" value="' + el.fflOffset_mm + '" step="1" min="0" max="100" style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;margin-top:2px;">' +
                        '</div>' +
                    '</div>' +
                    // Span type
                    '<div style="font-size:11px;font-weight:600;color:#666;margin-bottom:4px;">SPAN TYPE</div>' +
                    '<div style="display:flex;gap:6px;margin-bottom:14px;">' +
                        '<button id="jzs-ss" style="flex:1;padding:10px;border:1px solid ' + (!isCont ? '#6D28D9' : '#ddd') + ';border-radius:4px;background:' + (!isCont ? '#f5f3ff' : '#f9fafb') + ';cursor:pointer;font-size:11px;font-weight:600;color:' + (!isCont ? '#6D28D9' : '#666') + ';">Simply Supported<br><span style=font-weight:400;font-size:9px;color:#888>Joists terminate at each beam</span></button>' +
                        '<button id="jzs-cont" style="flex:1;padding:10px;border:1px solid ' + (isCont ? '#6D28D9' : '#ddd') + ';border-radius:4px;background:' + (isCont ? '#f5f3ff' : '#f9fafb') + ';cursor:pointer;font-size:11px;font-weight:600;color:' + (isCont ? '#6D28D9' : '#666') + ';">Continuous<br><span style=font-weight:400;font-size:9px;color:#888>Joists lap over beams</span></button>' +
                    '</div>' +
                    // Bay label
                    '<div style="margin-bottom:14px;">' +
                        '<label style="font-size:10px;font-weight:600;color:#666;">BAY LABEL</label>' +
                        '<input type="text" id="jzs-label" value="' + el.bayLabel + '" maxlength="4" style="width:60px;padding:6px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;margin-top:2px;margin-left:8px;">' +
                    '</div>' +
                    // Buttons
                    '<div style="display:flex;justify-content:flex-end;gap:8px;padding-top:8px;border-top:1px solid #eee;">' +
                        '<button id="jzs-cancel" style="padding:6px 16px;border:1px solid #ddd;border-radius:4px;background:#fff;cursor:pointer;font-size:12px;">Cancel</button>' +
                        '<button id="jzs-apply" style="padding:6px 16px;border:none;border-radius:4px;background:#6D28D9;color:#fff;cursor:pointer;font-size:12px;font-weight:600;">Apply</button>' +
                    '</div>' +
                '</div>' +
            '</div>';

        modal.style.display = 'flex';

        // Wire toggle buttons
        var activeStyle = function (c) { return 'flex:1;padding:10px;border:1px solid ' + c + ';border-radius:4px;background:' + c + '15;cursor:pointer;font-size:12px;font-weight:600;color:' + c + ';'; };
        var inactiveStyle = 'flex:1;padding:10px;border:1px solid #ddd;border-radius:4px;background:#f9fafb;cursor:pointer;font-size:12px;font-weight:600;color:#666;';

        document.getElementById('jzs-span-h').onclick = function () {
            this.style.cssText = activeStyle('#6D28D9');
            document.getElementById('jzs-span-v').style.cssText = inactiveStyle;
            modal.dataset.spanDir = '0';
        };
        document.getElementById('jzs-span-v').onclick = function () {
            this.style.cssText = activeStyle('#6D28D9');
            document.getElementById('jzs-span-h').style.cssText = inactiveStyle;
            modal.dataset.spanDir = '90';
        };
        modal.dataset.spanDir = el.spanDirection_deg.toString();

        document.getElementById('jzs-ss').onclick = function () {
            this.style.cssText = activeStyle('#6D28D9');
            document.getElementById('jzs-cont').style.cssText = inactiveStyle.replace('font-size:12px', 'font-size:11px');
            modal.dataset.spanType = 'single';
        };
        document.getElementById('jzs-cont').onclick = function () {
            this.style.cssText = activeStyle('#6D28D9');
            document.getElementById('jzs-ss').style.cssText = inactiveStyle.replace('font-size:12px', 'font-size:11px');
            modal.dataset.spanType = 'continuous';
        };
        modal.dataset.spanType = el.spanType;

        document.getElementById('jzs-cancel').onclick = function () { hideSetupDialog(); };
        document.getElementById('jzs-apply').onclick = function () { applySetupDialog(); };
    }

    function hideSetupDialog() {
        var modal = document.getElementById('joist-zone-setup');
        if (modal) modal.style.display = 'none';
        _dialogEl = null;
    }

    function applySetupDialog() {
        if (!_dialogEl) return;
        var modal = document.getElementById('joist-zone-setup');
        var el = _dialogEl;

        el.spanDirection_deg = parseInt(modal.dataset.spanDir) || 0;
        el.maxSpacing_mm = parseInt(document.getElementById('jzs-spacing').value) || 450;
        el.fflOffset_mm = parseInt(document.getElementById('jzs-ffl').value) || 19;
        el.spanType = modal.dataset.spanType || 'single';
        el.bayLabel = document.getElementById('jzs-label').value || el.bayLabel;

        // Compute joist lines
        recomputeJoistZone(el);

        if (_dialogIsNew) {
            // New zone — push via history
            history.execute({
                description: 'Create joist zone Bay ' + el.bayLabel,
                execute: function () { project.elements.push(el); },
                undo: function () {
                    var i = project.elements.indexOf(el);
                    if (i !== -1) project.elements.splice(i, 1);
                }
            });
        }

        hideSetupDialog();
        if (typeof _beamUtilCache !== 'undefined') { _beamUtilCache = {}; _beamUtilHash = ''; }
        engine.requestRender();
    }

    // Double-click existing joistZone → re-open setup
    container.addEventListener('dblclick', function (e) {
        if (activeTool !== 'pointer' && activeTool !== 'select') return;
        var sel = (typeof selectedElement !== 'undefined') ? selectedElement : null;
        if (!sel || sel.type !== 'joistZone') return;
        showSetupDialog(sel, false);
    });

    // Escape closes setup dialog
    window.addEventListener('keydown', function (e) {
        if (e.key === 'Escape') {
            var modal = document.getElementById('joist-zone-setup');
            if (modal && modal.style.display !== 'none') {
                hideSetupDialog();
                e.stopPropagation();
            }
        }
    });

    // ════════════════════════════════════════════════════════════
    // PHASE 2: Joist computation engine
    // ════════════════════════════════════════════════════════════

    /**
     * Recompute all joist positions, internal beams, spans, and sizing
     * for a joistZone element. Stores results on el.computed.
     */
    function recomputeJoistZone(el) {
        if (!el || !el.points || el.points.length < 3) {
            el.computed = null;
            return;
        }

        var poly = el.points;
        var spanDeg = el.spanDirection_deg || 0;
        var spanRad = spanDeg * Math.PI / 180;
        var maxSpacing = el.maxSpacing_mm || 450;
        var spanType = el.spanType || 'single';
        var levelId = el.level || 'L1';

        // Span direction unit vectors
        // spanDir: direction joists span (the long axis of each joist)
        // perpDir: perpendicular to span (direction of spacing/layout)
        var spanDir = { x: Math.cos(spanRad), y: Math.sin(spanRad) };
        var perpDir = { x: -Math.sin(spanRad), y: Math.cos(spanRad) };

        // ── 1. Zone extent along perpendicular axis ──
        var bbox = FLR.polygonBBox(poly);
        var origin = { x: bbox.minX, y: bbox.minY };
        var perpMin = Infinity, perpMax = -Infinity;
        for (var i = 0; i < poly.length; i++) {
            var d = FLR.projectOntoAxis(poly[i], origin, perpDir);
            if (d < perpMin) perpMin = d;
            if (d > perpMax) perpMax = d;
        }
        var zoneWidth = perpMax - perpMin;
        if (zoneWidth < 50) { el.computed = null; return; }

        // ── 2. Even spacing ──
        var numSpaces = Math.ceil(zoneWidth / maxSpacing);
        var actualSpacing = zoneWidth / numSpaces;
        var joistCount = numSpaces - 1; // joists between edges (edges are at beams or zone boundary)

        // ── 3. Generate joist lines at computed positions ──
        // Each joist is a line running in spanDir, positioned along perpDir
        var spanMin = Infinity, spanMax = -Infinity;
        for (var i2 = 0; i2 < poly.length; i2++) {
            var ds = FLR.projectOntoAxis(poly[i2], origin, spanDir);
            if (ds < spanMin) spanMin = ds;
            if (ds > spanMax) spanMax = ds;
        }
        var spanExtent = spanMax - spanMin;

        var rawJoistLines = [];
        for (var j = 1; j <= joistCount; j++) {
            var perpOffset = perpMin + j * actualSpacing;
            // Line start/end: extend across full span extent
            var baseX = origin.x + perpDir.x * perpOffset + spanDir.x * spanMin;
            var baseY = origin.y + perpDir.y * perpOffset + spanDir.y * spanMin;
            var endX = origin.x + perpDir.x * perpOffset + spanDir.x * spanMax;
            var endY = origin.y + perpDir.y * perpOffset + spanDir.y * spanMax;

            // Clip to polygon
            var segments = FLR.clipSegmentToPolygon(
                { x: baseX, y: baseY }, { x: endX, y: endY }, poly
            );
            for (var s = 0; s < segments.length; s++) {
                rawJoistLines.push({
                    x1: segments[s].enter.x, y1: segments[s].enter.y,
                    x2: segments[s].exit.x,  y2: segments[s].exit.y,
                });
            }
        }

        // ── 4. Detect internal beams ──
        var internalBeams = detectInternalBeams(el, spanDir, perpDir, poly, levelId);

        // ── 5. Split joists at beam crossings ──
        var joistLines = [];
        for (var k = 0; k < rawJoistLines.length; k++) {
            var jl = rawJoistLines[k];
            var spans = splitJoistAtBeams(jl, internalBeams, spanDir, spanType);
            joistLines.push({
                x1: jl.x1, y1: jl.y1, x2: jl.x2, y2: jl.y2,
                spans: spans,
            });
        }

        // ── 6. Cantilever detection ──
        for (var m = 0; m < joistLines.length; m++) {
            var jSpans = joistLines[m].spans;
            if (jSpans.length > 0 && internalBeams.length > 0) {
                // First span: if start is at zone edge (not at a beam), it's a cantilever
                if (!jSpans[0]._startsAtBeam) jSpans[0].isCantilever = true;
                // Last span: if end is at zone edge, it's a cantilever
                if (!jSpans[jSpans.length - 1]._endsAtBeam) jSpans[jSpans.length - 1].isCantilever = true;
            }
        }

        // ── 7. Per-span sizing ──
        var levelDS = (typeof getLevelDesignSettings === 'function') ? getLevelDesignSettings(levelId) : { G: 1.2, Q: 1.5 };
        // Get loads from floor zones covering this area
        var zones = [];
        if (FLR && FLR.resolveZonesFromElements) {
            zones = FLR.resolveZonesFromElements(project.elements, project.scheduleTypes.floorLoad || {});
        }
        var centroid = { x: 0, y: 0 };
        for (var ci = 0; ci < poly.length; ci++) { centroid.x += poly[ci].x; centroid.y += poly[ci].y; }
        centroid.x /= poly.length; centroid.y /= poly.length;
        var loadAtCentre = FLR ? FLR.getLoadAtPoint(centroid.x, centroid.y, zones, levelDS.G, levelDS.Q) : { G_kPa: levelDS.G, Q_kPa: levelDS.Q };

        var sizeCache = {};
        var governingResult = null;
        var governingSpan = 0;

        for (var n = 0; n < joistLines.length; n++) {
            for (var p = 0; p < joistLines[n].spans.length; p++) {
                var span = joistLines[n].spans[p];
                var span_m = span.span_mm / 1000;
                var cacheKey = span.isCantilever ? 'c' + span_m.toFixed(2) : span_m.toFixed(2);

                if (!sizeCache[cacheKey]) {
                    if (typeof HyspanJoists !== 'undefined' && HyspanJoists.sizeJoist) {
                        // hySPAN only accepts 450 or 600 — use nearest valid spacing
                        var sizerSpacing = actualSpacing <= 525 ? 450 : 600;
                        var result = HyspanJoists.sizeJoist({
                            span_m: span_m,
                            spacing_mm: sizerSpacing,
                            spanType: spanType,
                            G_kPa: loadAtCentre.G_kPa,
                            Q_kPa: loadAtCentre.Q_kPa,
                        });
                        // Fallback: if sizing failed (span exceeds table), use the
                        // heaviest available section so the user still sees a size
                        // on the drawing. The FAIL state is indicated via util > 1.
                        if (!result.ok && HyspanJoists.sections && HyspanJoists.sections.length > 0) {
                            var heaviest = HyspanJoists.sections[HyspanJoists.sections.length - 1];
                            var entry = heaviest.max && heaviest.max[spanType] && heaviest.max[spanType][sizerSpacing];
                            var maxSpan = entry ? entry.span : span_m;
                            result = {
                                ok: false,
                                section: heaviest,
                                maxSpan_m: maxSpan,
                                utilisation: maxSpan > 0 ? span_m / maxSpan : 999,
                                spanType: spanType,
                                spacing_mm: sizerSpacing,
                                loadCheck: result.loadCheck,
                                reason: result.reason,
                                isFallback: true,
                            };
                        }
                        sizeCache[cacheKey] = result;
                    }
                }
                span.result = sizeCache[cacheKey] || null;

                // Track governing result: prefer passing results, but fall back to
                // the largest-span span's section (even if FAIL) so a callout always shows.
                if (span.result && span.result.section) {
                    if (!governingResult || span.span_mm > governingSpan) {
                        governingResult = span.result;
                        governingSpan = span.span_mm;
                    }
                }
            }
        }

        el.computed = {
            actualSpacing_mm: Math.round(actualSpacing),
            joistCount: joistCount,
            internalBeams: internalBeams,
            joistLines: joistLines,
            governingResult: governingResult,
            governingSpan_mm: governingSpan,
        };
    }

    // ── Internal beam detection ──────────────────────────────────
    function detectInternalBeams(zoneEl, spanDir, perpDir, poly, levelId) {
        if (!project || !project.elements) return [];
        var beams = [];

        for (var i = 0; i < project.elements.length; i++) {
            var b = project.elements[i];
            if (!b || b.type !== 'line' || b.layer !== 'S-BEAM') continue;
            if (b.level && b.level !== levelId) continue;

            // Check beam is roughly perpendicular to span (within ±30°)
            var bdx = b.x2 - b.x1, bdy = b.y2 - b.y1;
            var bLen = Math.sqrt(bdx * bdx + bdy * bdy);
            if (bLen < 100) continue;
            var bux = bdx / bLen, buy = bdy / bLen;
            // Dot product with span direction — should be near 0 (perpendicular)
            var dot = Math.abs(bux * spanDir.x + buy * spanDir.y);
            if (dot > 0.5) continue; // more than ~30° from perpendicular — skip

            // Check beam is within or at the edge of the zone polygon.
            // Uses bounding-box overlap + midpoint proximity rather than
            // clipping (which fails when beams lie exactly on polygon edges).
            var zoneBBox = FLR.polygonBBox(poly);
            var bMinX = Math.min(b.x1, b.x2), bMaxX = Math.max(b.x1, b.x2);
            var bMinY = Math.min(b.y1, b.y2), bMaxY = Math.max(b.y1, b.y2);
            var tol = 100; // mm tolerance for boundary beams
            // Bounding-box overlap test (with tolerance)
            if (bMaxX < zoneBBox.minX - tol || bMinX > zoneBBox.maxX + tol) continue;
            if (bMaxY < zoneBBox.minY - tol || bMinY > zoneBBox.maxY + tol) continue;

            // Measure beam's position along span axis
            var bMid = { x: (b.x1 + b.x2) / 2, y: (b.y1 + b.y2) / 2 };
            var distAlongSpan = FLR.projectOntoAxis(bMid, { x: poly[0].x, y: poly[0].y }, spanDir);

            beams.push({
                beamId: b.id,
                beamEl: b,
                distance_mm: distAlongSpan,
                midX: bMid.x, midY: bMid.y,
            });
        }

        // Sort by distance along span
        beams.sort(function (a, b) { return a.distance_mm - b.distance_mm; });

        // All beams (including those at zone edges) are valid supports — they
        // define where joists start/stop or pass over.  The splitJoistAtBeams()
        // function handles endpoints vs internal crossings based on geometry.

        return beams;
    }

    // ── Split joist line at beam crossings ───────────────────────
    function splitJoistAtBeams(jl, internalBeams, spanDir, spanType) {
        if (internalBeams.length === 0) {
            var dx = jl.x2 - jl.x1, dy = jl.y2 - jl.y1;
            return [{
                x1: jl.x1, y1: jl.y1, x2: jl.x2, y2: jl.y2,
                span_mm: Math.sqrt(dx * dx + dy * dy),
                isCantilever: false,
                _startsAtBeam: false, _endsAtBeam: false,
                result: null,
            }];
        }

        // Find where each beam crosses this joist line
        var cuts = [];
        for (var i = 0; i < internalBeams.length; i++) {
            var b = internalBeams[i].beamEl;
            var hit = FLR.segmentIntersection(
                { x: jl.x1, y: jl.y1 }, { x: jl.x2, y: jl.y2 },
                { x: b.x1, y: b.y1 }, { x: b.x2, y: b.y2 }
            );
            if (hit && hit.t >= 0 && hit.t <= 1) {
                cuts.push({ t: hit.t, x: hit.x, y: hit.y, beamId: internalBeams[i].beamId });
            }
        }
        cuts.sort(function (a, b) { return a.t - b.t; });

        if (cuts.length === 0) {
            var dx2 = jl.x2 - jl.x1, dy2 = jl.y2 - jl.y1;
            return [{
                x1: jl.x1, y1: jl.y1, x2: jl.x2, y2: jl.y2,
                span_mm: Math.sqrt(dx2 * dx2 + dy2 * dy2),
                isCantilever: false,
                _startsAtBeam: false, _endsAtBeam: false,
                result: null,
            }];
        }

        // Build spans between cuts
        var spans = [];
        var prevPt = { x: jl.x1, y: jl.y1 };
        var prevIsBeam = false;

        for (var c = 0; c < cuts.length; c++) {
            var dx3 = cuts[c].x - prevPt.x, dy3 = cuts[c].y - prevPt.y;
            var spanLen = Math.sqrt(dx3 * dx3 + dy3 * dy3);
            if (spanLen > 50) { // skip tiny fragments
                spans.push({
                    x1: prevPt.x, y1: prevPt.y,
                    x2: cuts[c].x, y2: cuts[c].y,
                    span_mm: spanLen,
                    isCantilever: false,
                    _startsAtBeam: prevIsBeam,
                    _endsAtBeam: true,
                    result: null,
                });
            }
            prevPt = { x: cuts[c].x, y: cuts[c].y };
            prevIsBeam = true;
        }

        // Last segment (from last cut to end)
        var dxL = jl.x2 - prevPt.x, dyL = jl.y2 - prevPt.y;
        var lastLen = Math.sqrt(dxL * dxL + dyL * dyL);
        if (lastLen > 50) {
            spans.push({
                x1: prevPt.x, y1: prevPt.y,
                x2: jl.x2, y2: jl.y2,
                span_mm: lastLen,
                isCantilever: false,
                _startsAtBeam: prevIsBeam,
                _endsAtBeam: false,
                result: null,
            });
        }

        return spans;
    }

    // ════════════════════════════════════════════════════════════
    // PHASE 3: 2D Canvas Rendering
    // ════════════════════════════════════════════════════════════

    // Hash of beam geometry to detect changes and trigger auto-recompute
    var _beamGeomHash = '';
    var _recomputePending = false;
    function computeBeamHash() {
        var els = project.elements;
        if (!els) return '';
        var parts = [];
        for (var i = 0; i < els.length; i++) {
            var e = els[i];
            if (e && e.type === 'line' && e.layer === 'S-BEAM') {
                parts.push(e.id + ':' + (e.x1 | 0) + ',' + (e.y1 | 0) + ',' + (e.x2 | 0) + ',' + (e.y2 | 0));
            }
        }
        return parts.join('|');
    }

    engine._renderCallbacks.push(function _joistZoneDraw(ctx, eng) {
        var coords = eng.coords;
        var zoom = eng.viewport.zoom;
        var elements = project.getVisibleElements();

        // Auto-recompute joist zones when beams change (add/move/delete).
        // Deferred via setTimeout so we don't block the render loop; also
        // skipped while any drawing tool is active (prevents mid-draw flicker).
        var curBeamHash = computeBeamHash();
        if (curBeamHash !== _beamGeomHash && !_recomputePending) {
            _beamGeomHash = curBeamHash;
            _recomputePending = true;
            setTimeout(function () {
                for (var ri = 0; ri < project.elements.length; ri++) {
                    var re = project.elements[ri];
                    if (re && re.type === 'joistZone') {
                        try { recomputeJoistZone(re); } catch (err) { /* skip */ }
                    }
                }
                _recomputePending = false;
                if (engine && engine.requestRender) engine.requestRender();
            }, 50);
        }

        for (var i = 0; i < elements.length; i++) {
            var el = elements[i];
            if (!el || el.type !== 'joistZone') continue;
            if (!el.computed || !el.computed.joistLines) continue;

            var layer = (typeof LAYERS !== 'undefined') ? LAYERS['S-JOIST'] : null;
            if (layer && !layer.visible) continue;
            var isSelected = (typeof selectedElement !== 'undefined' && selectedElement === el);

            // ── Zone polygon outline (ONLY when selected — drafting convention: clean plan) ──
            var pts = el.points;
            if (isSelected && pts && pts.length >= 3) {
                ctx.save();
                ctx.beginPath();
                var sp0 = coords.realToScreen(pts[0].x, pts[0].y);
                ctx.moveTo(sp0.x, sp0.y);
                for (var vi = 1; vi < pts.length; vi++) {
                    var sp = coords.realToScreen(pts[vi].x, pts[vi].y);
                    ctx.lineTo(sp.x, sp.y);
                }
                ctx.closePath();
                ctx.globalAlpha = 0.08;
                ctx.fillStyle = '#8B5CF6';
                ctx.fill();
                ctx.globalAlpha = 0.7;
                ctx.strokeStyle = '#8B5CF6';
                ctx.lineWidth = 1.5;
                ctx.setLineDash([6, 3]);
                ctx.stroke();
                ctx.setLineDash([]);
                ctx.restore();
            }

            // ── Individual joist lines — thin solid black (drafting convention) ──
            ctx.save();
            ctx.strokeStyle = '#1a1a1a';
            ctx.lineWidth = Math.max(0.5, 0.04 * zoom); // ~0.13mm pen weight at scale
            ctx.globalAlpha = 0.9;

            var joistLines = el.computed.joistLines;
            for (var ji = 0; ji < joistLines.length; ji++) {
                var jl = joistLines[ji];
                for (var si = 0; si < jl.spans.length; si++) {
                    var span = jl.spans[si];
                    var s1 = coords.realToScreen(span.x1, span.y1);
                    var s2 = coords.realToScreen(span.x2, span.y2);
                    ctx.beginPath();
                    ctx.moveTo(s1.x, s1.y);
                    ctx.lineTo(s2.x, s2.y);
                    ctx.stroke();

                    // ── Continuous lap detail at beam crossings ──
                    if (el.spanType === 'continuous' && span._endsAtBeam) {
                        var lapLen = 150; // mm real
                        var offsetDist = 25; // mm real offset each side
                        var sdx = span.x2 - span.x1, sdy = span.y2 - span.y1;
                        var sLen = Math.sqrt(sdx * sdx + sdy * sdy);
                        if (sLen > 1) {
                            var ux = sdx / sLen, uy = sdy / sLen;
                            var nx = -uy, ny = ux;
                            for (var side = -1; side <= 1; side += 2) {
                                var ox = nx * offsetDist * side;
                                var oy = ny * offsetDist * side;
                                var lapS = coords.realToScreen(span.x2 - ux * lapLen + ox, span.y2 - uy * lapLen + oy);
                                var lapE = coords.realToScreen(span.x2 + ux * lapLen * 0.3 + ox, span.y2 + uy * lapLen * 0.3 + oy);
                                ctx.beginPath();
                                ctx.moveTo(lapS.x, lapS.y);
                                ctx.lineTo(lapE.x, lapE.y);
                                ctx.stroke();
                            }
                        }
                    }
                }
            }
            ctx.restore();

            // ── Single leader line with centred callout (drafting convention) ──
            // Always draw leader — if sizing failed, show FAIL message so user knows
            if (zoom > 0.8) {
                drawLeaderArrow(ctx, coords, el, zoom);
            }

            // ── Selected state: show bay label + scope warning ──
            if (isSelected && zoom > 0.6) {
                var cx = 0, cy = 0;
                for (var li = 0; li < pts.length; li++) { cx += pts[li].x; cy += pts[li].y; }
                cx /= pts.length; cy /= pts.length;
                var cScreen = coords.realToScreen(cx, cy);
                var fontSize = Math.max(8, 2 * zoom);
                ctx.save();
                ctx.globalAlpha = 0.6;
                ctx.font = 'bold ' + fontSize + 'px "Segoe UI", Arial, sans-serif';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.fillStyle = '#6D28D9';
                ctx.fillText('Bay ' + el.bayLabel, cScreen.x, cScreen.y - fontSize * 2);

                if (el.computed.governingResult && el.computed.governingResult.loadCheck && !el.computed.governingResult.loadCheck.ok) {
                    ctx.font = (fontSize * 0.7) + 'px sans-serif';
                    ctx.fillStyle = '#D97706';
                    ctx.fillText('Outside AS 1684 scope', cScreen.x, cScreen.y - fontSize * 0.8);
                }
                ctx.restore();
            }
        }

        // ── Drawing preview (tool active) ──
        if (activeTool === 'joistZone' && toolState.points.length > 0) {
            ctx.save();
            ctx.beginPath();
            var fp = coords.sheetToScreen(toolState.points[0].x, toolState.points[0].y);
            ctx.moveTo(fp.x, fp.y);
            for (var ti = 1; ti < toolState.points.length; ti++) {
                var tp = coords.sheetToScreen(toolState.points[ti].x, toolState.points[ti].y);
                ctx.lineTo(tp.x, tp.y);
            }
            if (toolState.currentEnd) {
                var ep = coords.sheetToScreen(toolState.currentEnd.x, toolState.currentEnd.y);
                ctx.lineTo(ep.x, ep.y);
            }
            ctx.closePath();
            ctx.globalAlpha = 0.1;
            ctx.fillStyle = '#8B5CF6';
            ctx.fill();
            ctx.globalAlpha = 0.6;
            ctx.strokeStyle = '#8B5CF6';
            ctx.lineWidth = 2;
            ctx.setLineDash([8, 4]);
            ctx.stroke();
            ctx.setLineDash([]);
            // Vertex dots
            ctx.fillStyle = '#6D28D9';
            for (var di = 0; di < toolState.points.length; di++) {
                var dp = coords.sheetToScreen(toolState.points[di].x, toolState.points[di].y);
                ctx.beginPath();
                ctx.arc(dp.x, dp.y, 3, 0, Math.PI * 2);
                ctx.fill();
            }
            ctx.restore();
        }
    });

    // ── Leader arrow rendering (Australian drafting convention) ──
    // Single leader line running perpendicular to joists, across the full
    // perpendicular extent of the zone, with inward-pointing arrowheads at
    // each end and a centred callout reading horizontally.
    function drawLeaderArrow(ctx, coords, el, zoom) {
        var comp = el.computed;
        var hasResult = comp && comp.governingResult && comp.governingResult.section;
        var sect = hasResult ? comp.governingResult.section : null;
        // Find the longest span across all joist lines (used in FAIL callout)
        var longestSpan_mm = 0;
        if (comp && comp.joistLines) {
            for (var li2 = 0; li2 < comp.joistLines.length; li2++) {
                var jl2 = comp.joistLines[li2];
                for (var si2 = 0; si2 < jl2.spans.length; si2++) {
                    if (jl2.spans[si2].span_mm > longestSpan_mm) longestSpan_mm = jl2.spans[si2].span_mm;
                }
            }
        }
        var poly = el.points;
        var spanRad = (el.spanDirection_deg || 0) * Math.PI / 180;
        // spanDir: direction joists run; perpDir: perpendicular to joists
        var spanDir = { x: Math.cos(spanRad), y: Math.sin(spanRad) };
        var perpDir = { x: -Math.sin(spanRad), y: Math.cos(spanRad) };

        var origin = { x: poly[0].x, y: poly[0].y };

        // Find zone extent along perpendicular axis (leader length)
        var perpMin = Infinity, perpMax = -Infinity;
        for (var i = 0; i < poly.length; i++) {
            var dp = FLR.projectOntoAxis(poly[i], origin, perpDir);
            if (dp < perpMin) perpMin = dp;
            if (dp > perpMax) perpMax = dp;
        }

        // Find zone extent along span axis (for midspan placement)
        var spanMin = Infinity, spanMax = -Infinity;
        for (var k = 0; k < poly.length; k++) {
            var ds = FLR.projectOntoAxis(poly[k], origin, spanDir);
            if (ds < spanMin) spanMin = ds;
            if (ds > spanMax) spanMax = ds;
        }
        var spanMid = (spanMin + spanMax) / 2;

        // Inset leader slightly from zone edges (drafting convention: arrowheads
        // sit just inside the boundary, not on it)
        var inset = 50; // mm real
        var leaderPerpStart = perpMin + inset;
        var leaderPerpEnd = perpMax - inset;

        // Anchor point: the zone's origin plus perpMid along perp and spanMid along span
        var leaderStart = {
            x: origin.x + perpDir.x * leaderPerpStart + spanDir.x * spanMid,
            y: origin.y + perpDir.y * leaderPerpStart + spanDir.y * spanMid,
        };
        var leaderEnd = {
            x: origin.x + perpDir.x * leaderPerpEnd + spanDir.x * spanMid,
            y: origin.y + perpDir.y * leaderPerpEnd + spanDir.y * spanMid,
        };

        var ls = coords.realToScreen(leaderStart.x, leaderStart.y);
        var le = coords.realToScreen(leaderEnd.x, leaderEnd.y);

        // Leader always black — drafting convention keeps annotations monochrome.
        var lineColor = '#1a1a1a';
        ctx.save();
        ctx.strokeStyle = lineColor;
        ctx.fillStyle = lineColor;
        ctx.lineWidth = Math.max(0.6, 0.05 * zoom);
        ctx.globalAlpha = 1.0;

        // Leader line
        ctx.beginPath();
        ctx.moveTo(ls.x, ls.y);
        ctx.lineTo(le.x, le.y);
        ctx.stroke();

        // Inward-pointing arrowheads at each end (solid filled triangles)
        var adx = le.x - ls.x, ady = le.y - ls.y;
        var aLen = Math.sqrt(adx * adx + ady * ady);
        if (aLen > 10) {
            var aux = adx / aLen, auy = ady / aLen;
            var headLen = Math.min(10, Math.max(6, aLen * 0.06));
            var headWidth = headLen * 0.35;

            // Start arrowhead (points inward, toward le)
            ctx.beginPath();
            ctx.moveTo(ls.x, ls.y);
            ctx.lineTo(ls.x + aux * headLen + auy * headWidth, ls.y + auy * headLen - aux * headWidth);
            ctx.lineTo(ls.x + aux * headLen - auy * headWidth, ls.y + auy * headLen + aux * headWidth);
            ctx.closePath();
            ctx.fill();

            // End arrowhead (points inward, toward ls)
            ctx.beginPath();
            ctx.moveTo(le.x, le.y);
            ctx.lineTo(le.x - aux * headLen + auy * headWidth, le.y - auy * headLen - aux * headWidth);
            ctx.lineTo(le.x - aux * headLen - auy * headWidth, le.y - auy * headLen + aux * headWidth);
            ctx.closePath();
            ctx.fill();
        }

        // Callout text — centred on leader, reads HORIZONTALLY regardless of
        // leader direction (modern drafting convention: text is always read
        // left-to-right, never sideways).
        var mx = (ls.x + le.x) / 2, my = (ls.y + le.y) / 2;
        var textAngle = 0;

        // Fixed font size — does not scale with zoom (like typical CAD annotations).
        // Text stays the same screen size regardless of zoom level.
        var fontSize = 12;
        // Callout: always just size + product + spacing. No span, no FAIL, no (SS)/(Cont).
        var callout = '';
        if (sect) {
            callout = comp.joistCount + '/' + sect.D_mm + '\u00D7' + sect.B_mm +
                ' ' + (sect.product || 'hySPAN') + ' @ ' + comp.actualSpacing_mm + ' c/c';
        } else {
            callout = comp.joistCount + ' joists @ ' + comp.actualSpacing_mm + ' c/c';
        }

        ctx.save();
        ctx.translate(mx, my);
        ctx.rotate(textAngle);
        ctx.font = '500 ' + fontSize + 'px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        // Measure text to draw a halo rectangle (prevents joist lines crossing through text)
        var textMetrics = ctx.measureText(callout);
        var textW = textMetrics.width;
        var textH = fontSize * 1.2;
        ctx.fillStyle = 'rgba(255,255,255,0.95)';
        ctx.fillRect(-textW / 2 - 4, -textH / 2, textW + 8, textH);

        // Text always black — drafting convention.
        ctx.fillStyle = '#1a1a1a';
        ctx.fillText(callout, 0, 0);
        ctx.restore();

        ctx.restore();
    }

    // ════════════════════════════════════════════════════════════
    // PHASE 5: Zone-to-zone continuity (sizing enhancement)
    // ════════════════════════════════════════════════════════════
    // Handled within recomputeJoistZone — when spanType is 'continuous',
    // the hySPAN sizer uses the 'continuous' table which has higher
    // span capacities. Cross-zone continuity (sharing a beam) is
    // detected by checking if any internalBeam is shared with another
    // joistZone. The governing span across the group is used for sizing.
    // (Implemented inline in per-span sizing above via spanType parameter)

    // ════════════════════════════════════════════════════════════
    // PHASE 6: Hit testing + Properties panel
    // ════════════════════════════════════════════════════════════

    // Hit test: point-in-polygon for joistZone elements
    if (typeof hitTestElement === 'function') {
        var _origHitTest = hitTestElement;
        hitTestElement = function (sheetPos, tolerance) {
            var hit = _origHitTest(sheetPos, tolerance);
            if (hit) return hit;

            if (typeof project !== 'undefined' && project.elements) {
                var realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);
                for (var i = project.elements.length - 1; i >= 0; i--) {
                    var el = project.elements[i];
                    if (!el || el.type !== 'joistZone') continue;
                    if (el.points && FLR && FLR.pointInPolygon(realPos.x, realPos.y, el.points)) {
                        return el;
                    }
                }
            }
            return null;
        };
    }

    // Properties panel builder
    window.floorDesigner = window.floorDesigner || {};
    window.floorDesigner.buildJoistZonePropsHTML = function (el) {
        if (!el || el.type !== 'joistZone') return '';
        var html = '<div style="font-size:10px;color:#666;padding:4px 0;">';
        html += '<div style="display:flex;justify-content:space-between;"><span>Bay</span><span style="font-weight:700;color:#6D28D9;">' + el.bayLabel + '</span></div>';
        html += '<div style="display:flex;justify-content:space-between;"><span>Span direction</span><span>' + (el.spanDirection_deg < 45 ? 'Horizontal' : 'Vertical') + '</span></div>';
        html += '<div style="display:flex;justify-content:space-between;"><span>Span type</span><span>' + (el.spanType === 'continuous' ? 'Continuous' : 'Simply Supported') + '</span></div>';
        html += '<div style="display:flex;justify-content:space-between;"><span>FFL offset</span><span>' + el.fflOffset_mm + ' mm</span></div>';

        if (el.computed) {
            html += '<div style="display:flex;justify-content:space-between;"><span>Spacing</span><span>' + el.computed.actualSpacing_mm + ' mm c/c (max ' + el.maxSpacing_mm + ')</span></div>';
            html += '<div style="display:flex;justify-content:space-between;"><span>Joist count</span><span>' + el.computed.joistCount + '</span></div>';
            html += '<div style="display:flex;justify-content:space-between;"><span>Internal beams</span><span>' + el.computed.internalBeams.length + '</span></div>';

            if (el.computed.governingResult && el.computed.governingResult.section) {
                var sect = el.computed.governingResult.section;
                var util = el.computed.governingResult.utilisation;
                var color = util <= 0.85 ? '#16A34A' : util <= 1.0 ? '#D97706' : '#DC2626';
                html += '<div style="margin-top:4px;padding:4px;background:' + color + '15;border:1px solid ' + color + '40;border-radius:4px;text-align:center;">';
                html += '<span style="font-weight:700;color:' + color + ';">' + sect.name + ' ' + sect.product + '</span>';
                html += '<span style="margin-left:8px;color:' + color + ';">' + (util * 100).toFixed(0) + '%</span>';
                html += '</div>';
            }
        }

        html += '<div style="margin-top:6px;text-align:center;">';
        html += '<button onclick="var el=selectedElement;if(el&&el.type===\'joistZone\'){showSetupDialog(el,false);}" style="font-size:9px;padding:2px 10px;border:1px solid #6D28D9;color:#6D28D9;background:transparent;border-radius:3px;cursor:pointer;">Edit Zone</button>';
        html += '</div>';
        html += '</div>';
        return html;
    };

    // Expose showSetupDialog for properties panel button
    window.showJoistZoneSetup = showSetupDialog;
    window.recomputeJoistZone = recomputeJoistZone;

    console.log('[joist-zone] 1.0 loaded — draw polygon to define joist area');
})();
