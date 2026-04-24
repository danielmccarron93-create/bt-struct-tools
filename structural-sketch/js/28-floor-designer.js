// ============================================================
// 28-floor-designer.js — Floor Designer feature namespace
// ------------------------------------------------------------
// Slice 1: scaffolding only. Feature flag, namespace, schedule
// tab wiring.
//
// Slice 2 (this slice): Floor Load polygon drawing tool.
//   - New element type:  'floorZone'
//   - New layer:         'S-FLOORZONE'
//   - Tool state:        floorZoneToolState
//   - Active tool key:   'floorZone'
//   - Schedule category: 'floorLoad' (FL1..FLn)
//
// Slice 3 → AS 4100 engine replacement.
// Slice 4 → Calc sheet + trib viz + SB→FB promotion.
// Slice 5 → Joists (hySPAN residential scope).
// Slice 6 → Tonnage widget.
// Slice 7 → SpaceGass export.
//
// Schedule categories added by this feature:
//   floorLoad  — FL1..FLn — G/Q floor zones drawn on plan
//   floorBeam  — FB1..FBn — floor bearers (SB auto-promoted
//                           when joists span onto them)
//   joist      — FJ1..FJn — floor joists (hySPAN only for now)
//
// Existing SB (beam) schedule is untouched. An SB that directly
// carries a joist set will be re-tagged FB on promotion (Slice 4).
// An SB that only carries other beams (transfer beam) stays SB.
// ============================================================

(function () {
    // ------------------------------------------------------------
    // Namespace + feature flag
    // ------------------------------------------------------------
    window.floorDesigner = window.floorDesigner || {
        version: '0.2-slice2',
        enabled: true,
        categories: ['floorLoad', 'floorBeam', 'joist'],
        tools: {},
    };
    // Re-bump version if re-loaded against an earlier namespace
    window.floorDesigner.version = '0.2-slice2';

    if (typeof project !== 'undefined' && project) {
        project.features = project.features || {};
        if (project.features.floorDesigner === undefined) {
            project.features.floorDesigner = true;
        }
    }

    // Register placement type ref for the Floor Load tool so it
    // slots into the existing schedule-driven placement pattern
    // used by footings/walls/beams. 14-tools-extended.js creates
    // `placementTypeRef` — we just add our key.
    if (typeof placementTypeRef !== 'undefined' && placementTypeRef) {
        if (!placementTypeRef.floorLoad) placementTypeRef.floorLoad = 'FL1';
    }

    // ------------------------------------------------------------
    // Floor Load polygon tool state
    // ------------------------------------------------------------
    // Mirrors slabToolState — vertices kept in sheet-mm while drawing,
    // converted to real-mm on commit so they live in the model frame
    // alongside everything else. Rendering uses coords.realToScreen.
    const floorZoneToolState = {
        points: [],        // sheet-mm vertices, draw-time only
        currentEnd: null,  // preview next vertex, sheet-mm
    };
    window.floorDesigner.floorZoneToolState = floorZoneToolState;

    // ------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------
    function getActiveFloorLoadTypeRef() {
        if (typeof placementTypeRef !== 'undefined' && placementTypeRef && placementTypeRef.floorLoad) {
            return placementTypeRef.floorLoad;
        }
        return 'FL1';
    }

    function getFloorLoadTypeData(typeRef) {
        if (typeof project !== 'undefined'
            && project.scheduleTypes
            && project.scheduleTypes.floorLoad
            && project.scheduleTypes.floorLoad[typeRef]) {
            return project.scheduleTypes.floorLoad[typeRef];
        }
        return { G: 0, Q: 0, spanDirection: 0, color: '#A7F3D0', description: '' };
    }

    // Shoelace area in real mm² (closed polygon)
    function polygonArea(realPts) {
        if (!realPts || realPts.length < 3) return 0;
        let a = 0;
        const n = realPts.length;
        for (let i = 0, j = n - 1; i < n; j = i++) {
            a += (realPts[j].x + realPts[i].x) * (realPts[j].y - realPts[i].y);
        }
        return Math.abs(a * 0.5);
    }

    // Ray-cast point-in-polygon (works in any consistent frame).
    // Kept local so this tool still works even if the load
    // resolver module hasn't loaded yet (load order safety).
    function pointInPolyGeneric(px, py, verts) {
        if (!verts || verts.length < 3) return false;
        let inside = false;
        const n = verts.length;
        for (let i = 0, j = n - 1; i < n; j = i++) {
            const xi = verts[i].x, yi = verts[i].y;
            const xj = verts[j].x, yj = verts[j].y;
            const intersects = ((yi > py) !== (yj > py)) &&
                (px < ((xj - xi) * (py - yi)) / ((yj - yi) || 1e-12) + xi);
            if (intersects) inside = !inside;
        }
        return inside;
    }

    // ------------------------------------------------------------
    // Input position helper (with snap + ortho, matches slab tool)
    // ------------------------------------------------------------
    function getFloorZoneToolPos(e) {
        const rect = container.getBoundingClientRect();
        const sx = e.clientX - rect.left;
        const sy = e.clientY - rect.top;
        const snap = (typeof findSnap === 'function') ? findSnap(sx, sy) : null;
        let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
        if (floorZoneToolState.points.length > 0 && typeof applyOrtho === 'function') {
            const last = floorZoneToolState.points[floorZoneToolState.points.length - 1];
            pos = applyOrtho(pos.x, pos.y, last.x, last.y);
        }
        return pos;
    }

    // ------------------------------------------------------------
    // Input event handlers (scoped to activeTool === 'floorZone')
    // ------------------------------------------------------------
    container.addEventListener('mousemove', (e) => {
        if (activeTool !== 'floorZone') return;
        floorZoneToolState.currentEnd = getFloorZoneToolPos(e);
        engine.requestRender();
    });

    container.addEventListener('mousedown', (e) => {
        if (e.button !== 0) return;
        if (engine._spaceDown || engine._isPanning) return;
        if (activeTool !== 'floorZone') return;
        if (typeof pdfState !== 'undefined' && pdfState.calibrating) return;
        const pos = getFloorZoneToolPos(e);
        floorZoneToolState.points.push(pos);
        engine.requestRender();
    });

    container.addEventListener('dblclick', (e) => {
        if (activeTool !== 'floorZone') return;
        if (floorZoneToolState.points.length < 3) return;
        e.preventDefault();
        e.stopPropagation();
        // The 2nd mousedown of the dblclick already pushed a duplicate
        // vertex — drop it, then commit if we still have ≥3 points.
        floorZoneToolState.points.pop();
        if (floorZoneToolState.points.length >= 3) {
            commitFloorZone();
        }
    });

    container.addEventListener('contextmenu', (e) => {
        if (activeTool !== 'floorZone') return;
        if (floorZoneToolState.points.length >= 3) {
            e.preventDefault();
            commitFloorZone();
        } else {
            e.preventDefault();
            floorZoneToolState.points = [];
            floorZoneToolState.currentEnd = null;
            engine.requestRender();
        }
    });

    window.addEventListener('keydown', (e) => {
        if (activeTool !== 'floorZone') return;
        if (e.key === 'Enter' && floorZoneToolState.points.length >= 3) {
            commitFloorZone();
        }
        if (e.key === 'Escape') {
            floorZoneToolState.points = [];
            floorZoneToolState.currentEnd = null;
            engine.requestRender();
        }
    });

    // ------------------------------------------------------------
    // Commit — push a new floorZone element into project.elements
    // via the command pattern so undo/redo just works.
    // ------------------------------------------------------------
    function commitFloorZone() {
        const pts = floorZoneToolState.points;
        if (pts.length < 3) return;

        const typeRef = getActiveFloorLoadTypeRef();

        // Convert sheet → real, matching slab tool behaviour
        const realPoints = pts.map(p => engine.coords.sheetToReal(p.x, p.y));

        const newZone = {
            id: generateId(),
            type: 'floorZone',
            layer: 'S-FLOORZONE',
            level: (typeof getActiveLevel === 'function' && getActiveLevel())
                ? getActiveLevel().id
                : undefined,
            typeRef: typeRef,
            points: realPoints, // real-world mm vertices, NOT closed-duplicated
            closed: true,
        };

        history.execute({
            description: 'Draw floor load zone: ' + typeRef,
            execute() { project.elements.push(newZone); },
            undo() {
                const i = project.elements.indexOf(newZone);
                if (i !== -1) project.elements.splice(i, 1);
            }
        });

        floorZoneToolState.points = [];
        floorZoneToolState.currentEnd = null;
        engine.requestRender();
    }
    window.floorDesigner.commitFloorZone = commitFloorZone;

    // ------------------------------------------------------------
    // setActiveTool cleanup — clear our state when switching away
    // ------------------------------------------------------------
    // 06-drawing-tools.js clears every other tool's state inside
    // setActiveTool(). We hook the same way by wrapping it.
    if (typeof setActiveTool === 'function') {
        const _origSetActiveTool_FD = setActiveTool;
        setActiveTool = function (tool) {
            if (tool !== 'floorZone') {
                floorZoneToolState.points = [];
                floorZoneToolState.currentEnd = null;
            }
            _origSetActiveTool_FD(tool);

            // Toolbar button active state (button added in index.html)
            const btn = document.getElementById('btn-floor-load');
            if (btn) btn.classList.toggle('active', tool === 'floorZone');

            // Cursor — matches other polygon tools
            if (tool === 'floorZone') {
                container.style.cursor = 'crosshair';
            }

            // Status bar
            const statusTool = document.getElementById('status-tool');
            if (statusTool && tool === 'floorZone') {
                statusTool.textContent = 'Floor Load';
            }
        };
    }

    // ------------------------------------------------------------
    // Canvas rendering — append a new render callback.
    //
    // We deliberately don't try to splice into the existing
    // wrappedDrawElements chain: by the time this module loads,
    // 08-shapes.js has already replaced the slot with its own
    // closure, so indexOf(wrappedDrawElements) won't find it.
    // Appending instead means floor zones render AFTER all other
    // elements — the translucent fill sits over the plan, which
    // is the correct Z-order for a loads overlay (matches how
    // Bluebeam renders cloud/area markups).
    // ------------------------------------------------------------
    if (typeof engine !== 'undefined' && Array.isArray(engine._renderCallbacks)) {

        const fdDrawElements = function (ctx, eng) {
            const coords = eng.coords;
            const zoom = eng.viewport.zoom;

            // Committed floor zones
            const elements = (typeof project !== 'undefined' && project.getVisibleElements)
                ? project.getVisibleElements()
                : [];

            for (const el of elements) {
                if (!el || el.type !== 'floorZone') continue;
                const pts = el.points;
                if (!pts || pts.length < 3) continue;
                const layer = project.layers[el.layer] || project.layers['S-FLOORZONE'];
                if (!layer) continue;

                const typeData = getFloorLoadTypeData(el.typeRef);
                const zoneColor = typeData.color || '#A7F3D0';
                const isSelected = (typeof selectedElement !== 'undefined' && selectedElement === el);

                // Build the polygon path in screen coords
                ctx.beginPath();
                const p0 = coords.realToScreen(pts[0].x, pts[0].y);
                ctx.moveTo(p0.x, p0.y);
                for (let i = 1; i < pts.length; i++) {
                    const p = coords.realToScreen(pts[i].x, pts[i].y);
                    ctx.lineTo(p.x, p.y);
                }
                ctx.closePath();

                // Fill — tinted by schedule colour, translucent so
                // drawing underneath stays legible
                ctx.fillStyle = zoneColor;
                ctx.globalAlpha = isSelected ? 0.35 : 0.20;
                ctx.fill();
                ctx.globalAlpha = 1.0;

                // Dashed outline in layer colour
                ctx.strokeStyle = isSelected ? '#2B7CD0' : layer.color;
                ctx.lineWidth = Math.max(1, layer.lineWeight * zoom);
                const pattern = DASH_PATTERNS[layer.pattern] || [];
                ctx.setLineDash(pattern.length > 0 ? pattern.map(d => d * zoom) : []);
                ctx.stroke();
                ctx.setLineDash([]);

                // Label — "FL1  G=2.0  Q=3.0" at centroid
                // (simple average of vertices — good enough for a
                // convex-ish polygon, and we're labelling, not
                // computing first moment)
                let cx = 0, cy = 0;
                for (const v of pts) { cx += v.x; cy += v.y; }
                cx /= pts.length; cy /= pts.length;
                const sp = coords.realToScreen(cx, cy);

                // Per-zone overrides take precedence over schedule values
                const zoneG = (el.G_override !== undefined) ? Number(el.G_override) : Number(typeData.G) || 0;
                const zoneQ = (el.Q_override !== undefined) ? Number(el.Q_override) : Number(typeData.Q) || 0;
                const label = (el.typeRef || 'FL?')
                    + '  G=' + zoneG.toFixed(1)
                    + '  Q=' + zoneQ.toFixed(1);

                ctx.font = 'bold 11px "Segoe UI", Arial, sans-serif';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                // white halo for legibility over fills
                ctx.lineWidth = 3;
                ctx.strokeStyle = 'rgba(255,255,255,0.9)';
                ctx.strokeText(label, sp.x, sp.y);
                ctx.fillStyle = '#1f2937';
                ctx.fillText(label, sp.x, sp.y);

                // ── Span direction arrows ──
                // Draw small arrows at centroid showing joist span direction
                {
                    const spanDeg = (el.spanDirection_override !== undefined) ? Number(el.spanDirection_override) : Number(typeData.spanDirection) || 0;
                    const spanRad = spanDeg * Math.PI / 180;
                    // Arrow direction unit vector (in real-world frame, Y-down on screen)
                    const adx = Math.cos(spanRad);
                    const ady = Math.sin(spanRad);
                    // Convert to screen delta (real→screen may flip Y)
                    const refPt = coords.realToScreen(cx, cy);
                    const tipPt = coords.realToScreen(cx + adx * 500, cy + ady * 500);
                    const sdx = tipPt.x - refPt.x;
                    const sdy = tipPt.y - refPt.y;
                    const sLen = Math.sqrt(sdx * sdx + sdy * sdy);
                    if (sLen > 2) {
                        const ux = sdx / sLen, uy = sdy / sLen;
                        const arrowLen = Math.min(20, sLen * 0.8);
                        const headLen = arrowLen * 0.35;
                        ctx.save();
                        ctx.globalAlpha = 0.35;
                        ctx.strokeStyle = '#6B7280';
                        ctx.fillStyle = '#6B7280';
                        ctx.lineWidth = 1.5;
                        ctx.setLineDash([]);
                        // Draw two arrows offset above and below the label
                        for (const offsetPx of [-16, 16]) {
                            const ox = sp.x + uy * offsetPx;
                            const oy = sp.y - ux * offsetPx;
                            // Arrow shaft
                            ctx.beginPath();
                            ctx.moveTo(ox - ux * arrowLen / 2, oy - uy * arrowLen / 2);
                            ctx.lineTo(ox + ux * arrowLen / 2, oy + uy * arrowLen / 2);
                            ctx.stroke();
                            // Arrow head
                            const hx = ox + ux * arrowLen / 2;
                            const hy = oy + uy * arrowLen / 2;
                            ctx.beginPath();
                            ctx.moveTo(hx, hy);
                            ctx.lineTo(hx - ux * headLen + uy * headLen * 0.4, hy - uy * headLen - ux * headLen * 0.4);
                            ctx.lineTo(hx - ux * headLen - uy * headLen * 0.4, hy - uy * headLen + ux * headLen * 0.4);
                            ctx.closePath();
                            ctx.fill();
                        }
                        ctx.restore();
                    }
                }

                // Selection vertices
                if (isSelected) {
                    ctx.fillStyle = '#2B7CD0';
                    for (const pt of pts) {
                        const vp = coords.realToScreen(pt.x, pt.y);
                        ctx.fillRect(vp.x - 3, vp.y - 3, 6, 6);
                    }
                }
            }

            // Draw-time preview
            if (activeTool === 'floorZone' && floorZoneToolState.points.length > 0) {
                const layer = project.layers['S-FLOORZONE'];
                const typeRef = getActiveFloorLoadTypeRef();
                const typeData = getFloorLoadTypeData(typeRef);

                // Tinted preview fill — only if we have ≥3 points
                if (floorZoneToolState.points.length >= 3) {
                    ctx.beginPath();
                    const fp = coords.sheetToScreen(
                        floorZoneToolState.points[0].x,
                        floorZoneToolState.points[0].y
                    );
                    ctx.moveTo(fp.x, fp.y);
                    for (let i = 1; i < floorZoneToolState.points.length; i++) {
                        const p = coords.sheetToScreen(
                            floorZoneToolState.points[i].x,
                            floorZoneToolState.points[i].y
                        );
                        ctx.lineTo(p.x, p.y);
                    }
                    ctx.closePath();
                    ctx.fillStyle = typeData.color || '#A7F3D0';
                    ctx.globalAlpha = 0.15;
                    ctx.fill();
                    ctx.globalAlpha = 1.0;
                }

                // Outline preview (open path through vertices + cursor)
                ctx.strokeStyle = layer ? layer.color : '#059669';
                ctx.lineWidth = Math.max(1, (layer ? layer.lineWeight : 0.18) * zoom);
                ctx.globalAlpha = 0.8;
                const pattern = DASH_PATTERNS[layer ? layer.pattern : 'dashed'] || [];
                ctx.setLineDash(pattern.length > 0 ? pattern.map(d => d * zoom) : []);

                ctx.beginPath();
                const fp2 = coords.sheetToScreen(
                    floorZoneToolState.points[0].x,
                    floorZoneToolState.points[0].y
                );
                ctx.moveTo(fp2.x, fp2.y);
                for (let i = 1; i < floorZoneToolState.points.length; i++) {
                    const p = coords.sheetToScreen(
                        floorZoneToolState.points[i].x,
                        floorZoneToolState.points[i].y
                    );
                    ctx.lineTo(p.x, p.y);
                }
                if (floorZoneToolState.currentEnd) {
                    const ep = coords.sheetToScreen(
                        floorZoneToolState.currentEnd.x,
                        floorZoneToolState.currentEnd.y
                    );
                    ctx.lineTo(ep.x, ep.y);
                    // Dashed closing line back to first vertex
                    ctx.stroke();
                    ctx.beginPath();
                    ctx.moveTo(ep.x, ep.y);
                    ctx.lineTo(fp2.x, fp2.y);
                    ctx.setLineDash([5, 5]);
                    ctx.stroke();
                } else {
                    ctx.stroke();
                }
                ctx.setLineDash([]);
                ctx.globalAlpha = 1.0;

                // Vertex dots
                ctx.fillStyle = '#2B7CD0';
                for (const pt of floorZoneToolState.points) {
                    const vp = coords.sheetToScreen(pt.x, pt.y);
                    ctx.beginPath();
                    ctx.arc(vp.x, vp.y, 3, 0, Math.PI * 2);
                    ctx.fill();
                }
            }
        };

        engine._renderCallbacks.push(fdDrawElements);
    }

    // ------------------------------------------------------------
    // Hit testing — patch hitTestElement so selecting a filled
    // zone by clicking inside it works like the slab tool.
    // ------------------------------------------------------------
    if (typeof hitTestElement === 'function') {
        const _origHitTest_FD = hitTestElement;
        hitTestElement = function (sheetPos, tolerance) {
            // Try the normal hit-test first — things drawn on top
            // of a zone should still win.
            const hit = _origHitTest_FD(sheetPos, tolerance);
            if (hit) return hit;

            // Convert sheet → real, then point-in-polygon check
            // against any floorZone. Later-drawn zones win.
            if (typeof project !== 'undefined' && project.getVisibleElements) {
                const realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);
                const els = project.getVisibleElements();
                for (let i = els.length - 1; i >= 0; i--) {
                    const el = els[i];
                    if (!el || el.type !== 'floorZone') continue;
                    if (pointInPolyGeneric(realPos.x, realPos.y, el.points)) {
                        return el;
                    }
                }
            }
            return null;
        };
    }

    // ------------------------------------------------------------
    // Toolbar button wiring (button injected by index.html edit)
    // ------------------------------------------------------------
    const btnFloorLoad = document.getElementById('btn-floor-load');
    if (btnFloorLoad) {
        btnFloorLoad.addEventListener('click', () => setActiveTool('floorZone'));
    }

    // Keyboard shortcut: 'f' for Floor Load
    // (guarded against conflicts — see window keydown handler in
    // 06-drawing-tools.js for the list of shortcuts already in use)
    window.addEventListener('keydown', (e) => {
        if (document.activeElement !== document.body) return;
        if (e.ctrlKey || e.metaKey || e.altKey) return;
        if (e.key === 'f') setActiveTool('floorZone');
    });

    // ============================================================
    // Double-click beam → Calc Sheet modal (Slice 4)
    // ============================================================
    // When the pointer tool is active and the user double-clicks on
    // a beam (S-BEAM layer), open the full A4 calc sheet.
    if (typeof container !== 'undefined') {
        container.addEventListener('dblclick', function (e) {
            // Only fire when in pointer mode (no active drawing tool)
            if (typeof currentTool !== 'undefined' && currentTool && currentTool !== 'pointer') return;
            if (typeof activeTool !== 'undefined' && activeTool && activeTool !== 'pointer') return;

            // Must have calc sheet module
            if (typeof FloorCalcSheet === 'undefined' || !FloorCalcSheet.showCalcSheetModal) return;

            // Check if we have a selected beam element
            var beamEl = (typeof selectedElement !== 'undefined') ? selectedElement : null;
            if (!beamEl) return;

            // Must be a beam (on S-BEAM layer)
            if (beamEl.type !== 'line' || beamEl.layer !== 'S-BEAM') return;

            // Must have a section assigned
            var typeRef = beamEl.typeRef || beamEl.tag;
            if (!typeRef) return;
            var schedData = (project.scheduleTypes.beam && project.scheduleTypes.beam[typeRef]) ||
                            (project.scheduleTypes.floorBeam && project.scheduleTypes.floorBeam[typeRef]) || {};
            if (!schedData.size) return;

            FloorCalcSheet.showCalcSheetModal(beamEl);
        });
    }

    // ============================================================
    // Double-click floor zone → Zone Editor popup
    // ============================================================
    // Lets the user edit G, Q, and span direction per-zone with
    // common load case presets. Per-zone values override the schedule.

    var _zoneEditorEl = null; // currently-editing zone element

    function showZoneEditor(zoneEl) {
        _zoneEditorEl = zoneEl;
        var typeRef = zoneEl.typeRef || 'FL1';
        var typeData = getFloorLoadTypeData(typeRef);
        var curG = (zoneEl.G_override !== undefined) ? zoneEl.G_override : Number(typeData.G) || 0;
        var curQ = (zoneEl.Q_override !== undefined) ? zoneEl.Q_override : Number(typeData.Q) || 0;
        var curSpan = (zoneEl.spanDirection_override !== undefined) ? zoneEl.spanDirection_override : Number(typeData.spanDirection) || 0;

        // Create or reuse modal
        var modal = document.getElementById('zone-editor-modal');
        if (!modal) {
            modal = document.createElement('div');
            modal.id = 'zone-editor-modal';
            modal.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;z-index:4500;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,0.35);';
            document.body.appendChild(modal);
        }

        modal.innerHTML =
            '<div style="background:#fff;border-radius:10px;box-shadow:0 8px 32px rgba(0,0,0,0.25);width:380px;max-width:90vw;overflow:hidden;">' +
                '<div style="background:#1e40af;color:#fff;padding:12px 16px;font-weight:700;font-size:14px;">Floor Load — ' + typeRef + '</div>' +
                '<div style="padding:16px;">' +
                    // Presets
                    '<div style="font-size:11px;font-weight:600;color:#666;margin-bottom:6px;">LOAD PRESETS</div>' +
                    '<div style="display:flex;flex-wrap:wrap;gap:4px;margin-bottom:14px;" id="ze-presets">' +
                        '<button class="ze-preset" onclick="applyZonePreset(0.4,1.5)" style="font-size:10px;padding:4px 8px;border:1px solid #ddd;border-radius:4px;background:#f9fafb;cursor:pointer;">Residential<br><span style=color:#888>G=0.4 Q=1.5</span></button>' +
                        '<button class="ze-preset" onclick="applyZonePreset(1.0,3.0)" style="font-size:10px;padding:4px 8px;border:1px solid #ddd;border-radius:4px;background:#f9fafb;cursor:pointer;">Office<br><span style=color:#888>G=1.0 Q=3.0</span></button>' +
                        '<button class="ze-preset" onclick="applyZonePreset(1.0,4.0)" style="font-size:10px;padding:4px 8px;border:1px solid #ddd;border-radius:4px;background:#f9fafb;cursor:pointer;">Retail<br><span style=color:#888>G=1.0 Q=4.0</span></button>' +
                        '<button class="ze-preset" onclick="applyZonePreset(2.5,5.0)" style="font-size:10px;padding:4px 8px;border:1px solid #ddd;border-radius:4px;background:#f9fafb;cursor:pointer;">Assembly<br><span style=color:#888>G=2.5 Q=5.0</span></button>' +
                        '<button class="ze-preset" onclick="applyZonePreset(4.0,5.0)" style="font-size:10px;padding:4px 8px;border:1px solid #ddd;border-radius:4px;background:#f9fafb;cursor:pointer;">Plant/Storage<br><span style=color:#888>G=4.0 Q=5.0</span></button>' +
                        '<button class="ze-preset" onclick="applyZonePreset(0.5,0.25)" style="font-size:10px;padding:4px 8px;border:1px solid #ddd;border-radius:4px;background:#f9fafb;cursor:pointer;">Roof<br><span style=color:#888>G=0.5 Q=0.25</span></button>' +
                    '</div>' +
                    // Custom G/Q inputs
                    '<div style="display:flex;gap:12px;margin-bottom:14px;">' +
                        '<div style="flex:1;">' +
                            '<label style="font-size:10px;font-weight:600;color:#666;">G (dead) kPa</label>' +
                            '<input type="number" id="ze-G" value="' + curG.toFixed(1) + '" step="0.1" min="0" style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;margin-top:2px;">' +
                        '</div>' +
                        '<div style="flex:1;">' +
                            '<label style="font-size:10px;font-weight:600;color:#666;">Q (live) kPa</label>' +
                            '<input type="number" id="ze-Q" value="' + curQ.toFixed(1) + '" step="0.1" min="0" style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;margin-top:2px;">' +
                        '</div>' +
                    '</div>' +
                    // Span direction
                    '<div style="margin-bottom:14px;">' +
                        '<label style="font-size:10px;font-weight:600;color:#666;">JOIST SPAN DIRECTION</label>' +
                        '<div style="display:flex;gap:6px;margin-top:4px;">' +
                            '<button id="ze-span-0" onclick="setZoneSpanDir(0)" style="flex:1;padding:8px;border:1px solid ' + (curSpan < 45 ? '#1e40af' : '#ddd') + ';border-radius:4px;background:' + (curSpan < 45 ? '#eff6ff' : '#f9fafb') + ';cursor:pointer;font-size:11px;font-weight:600;color:' + (curSpan < 45 ? '#1e40af' : '#666') + ';">← → Horizontal</button>' +
                            '<button id="ze-span-90" onclick="setZoneSpanDir(90)" style="flex:1;padding:8px;border:1px solid ' + (curSpan >= 45 ? '#1e40af' : '#ddd') + ';border-radius:4px;background:' + (curSpan >= 45 ? '#eff6ff' : '#f9fafb') + ';cursor:pointer;font-size:11px;font-weight:600;color:' + (curSpan >= 45 ? '#1e40af' : '#666') + ';">↑ ↓ Vertical</button>' +
                        '</div>' +
                    '</div>' +
                    // Buttons
                    '<div style="display:flex;justify-content:flex-end;gap:8px;padding-top:8px;border-top:1px solid #eee;">' +
                        '<button onclick="hideZoneEditor()" style="padding:6px 16px;border:1px solid #ddd;border-radius:4px;background:#fff;cursor:pointer;font-size:12px;">Cancel</button>' +
                        '<button onclick="applyZoneEditor()" style="padding:6px 16px;border:none;border-radius:4px;background:#1e40af;color:#fff;cursor:pointer;font-size:12px;font-weight:600;">Apply</button>' +
                    '</div>' +
                '</div>' +
            '</div>';
        modal.style.display = 'flex';
    }

    function hideZoneEditor() {
        var modal = document.getElementById('zone-editor-modal');
        if (modal) modal.style.display = 'none';
        _zoneEditorEl = null;
    }

    // Preset buttons populate the G/Q fields
    window.applyZonePreset = function (g, q) {
        var gInput = document.getElementById('ze-G');
        var qInput = document.getElementById('ze-Q');
        if (gInput) gInput.value = g.toFixed(1);
        if (qInput) qInput.value = q.toFixed(1);
    };

    // Span direction buttons
    window.setZoneSpanDir = function (deg) {
        var btn0 = document.getElementById('ze-span-0');
        var btn90 = document.getElementById('ze-span-90');
        if (btn0 && btn90) {
            var active = 'flex:1;padding:8px;border:1px solid #1e40af;border-radius:4px;background:#eff6ff;cursor:pointer;font-size:11px;font-weight:600;color:#1e40af;';
            var inactive = 'flex:1;padding:8px;border:1px solid #ddd;border-radius:4px;background:#f9fafb;cursor:pointer;font-size:11px;font-weight:600;color:#666;';
            btn0.style.cssText = deg === 0 ? active : inactive;
            btn90.style.cssText = deg === 90 ? active : inactive;
        }
        // Store on a data attribute for retrieval on apply
        var modal = document.getElementById('zone-editor-modal');
        if (modal) modal.dataset.spanDir = deg;
    };

    // Apply button — saves per-zone overrides
    window.applyZoneEditor = function () {
        if (!_zoneEditorEl) return;
        var el = _zoneEditorEl;
        var gVal = parseFloat(document.getElementById('ze-G').value);
        var qVal = parseFloat(document.getElementById('ze-Q').value);
        var modal = document.getElementById('zone-editor-modal');
        var spanVal = modal && modal.dataset.spanDir !== undefined ? parseInt(modal.dataset.spanDir) : undefined;

        if (isNaN(gVal) || isNaN(qVal)) { hideZoneEditor(); return; }

        // Store old values for undo
        var oldG = el.G_override;
        var oldQ = el.Q_override;
        var oldSpan = el.spanDirection_override;

        // Apply via history for undo support
        if (typeof history !== 'undefined' && history.execute) {
            history.execute({
                description: 'Edit floor zone loads',
                execute: function () {
                    el.G_override = gVal;
                    el.Q_override = qVal;
                    if (spanVal !== undefined) el.spanDirection_override = spanVal;
                },
                undo: function () {
                    if (oldG !== undefined) el.G_override = oldG; else delete el.G_override;
                    if (oldQ !== undefined) el.Q_override = oldQ; else delete el.Q_override;
                    if (oldSpan !== undefined) el.spanDirection_override = oldSpan; else delete el.spanDirection_override;
                }
            });
        } else {
            el.G_override = gVal;
            el.Q_override = qVal;
            if (spanVal !== undefined) el.spanDirection_override = spanVal;
        }

        // Invalidate caches
        if (typeof _beamUtilCache !== 'undefined') { _beamUtilCache = {}; _beamUtilHash = ''; }
        if (typeof invalidateJoistBayCache === 'function') invalidateJoistBayCache();

        hideZoneEditor();
        if (typeof engine !== 'undefined' && engine.requestRender) engine.requestRender();
    };

    // Wire dblclick on floor zone → open editor
    if (typeof container !== 'undefined') {
        container.addEventListener('dblclick', function (e) {
            if (typeof activeTool !== 'undefined' && activeTool && activeTool !== 'pointer' && activeTool !== 'select') return;
            var zoneEl = (typeof selectedElement !== 'undefined') ? selectedElement : null;
            if (!zoneEl || zoneEl.type !== 'floorZone') return;
            showZoneEditor(zoneEl);
        });
    }

    // Escape closes zone editor
    window.addEventListener('keydown', function (e) {
        if (e.key === 'Escape') {
            var modal = document.getElementById('zone-editor-modal');
            if (modal && modal.style.display !== 'none') {
                hideZoneEditor();
                e.stopPropagation();
            }
        }
    });

    // Keyboard shortcut: Escape closes calc sheet modal
    window.addEventListener('keydown', function (e) {
        if (e.key === 'Escape' && typeof FloorCalcSheet !== 'undefined') {
            var modal = document.getElementById('cs-modal');
            if (modal && !modal.hasAttribute('hidden')) {
                FloorCalcSheet.hideCalcSheetModal();
                e.stopPropagation();
            }
        }
    });

    // ============================================================
    // Tributary strip visualisation (Slice 4)
    // ============================================================
    // When a beam is selected, draw a translucent strip showing its
    // tributary width on the canvas. This runs as part of the render
    // callback chain (already patched above with fdDrawElements).
    // We enhance fdDrawElements' parent scope by storing trib data
    // and rendering it during the draw pass.

    // Store computed trib data for the selected beam
    var _tribStripCache = { elId: null, strip: null };

    function _updateTribStrip() {
        var el = (typeof selectedElement !== 'undefined') ? selectedElement : null;
        if (!el || el.type !== 'line' || el.layer !== 'S-BEAM') {
            _tribStripCache.elId = null;
            _tribStripCache.strip = null;
            return;
        }
        if (_tribStripCache.elId === el.id) return; // already cached

        if (typeof calculateTributaryWidth !== 'function') {
            _tribStripCache.elId = el.id;
            _tribStripCache.strip = null;
            return;
        }

        var tribResult = calculateTributaryWidth(el);
        if (!tribResult || tribResult.tribWidth <= 0) {
            _tribStripCache.elId = el.id;
            _tribStripCache.strip = null;
            return;
        }

        // Build strip corners in real-world mm
        var dx = el.x2 - el.x1;
        var dy = el.y2 - el.y1;
        var beamLen = Math.sqrt(dx * dx + dy * dy);
        if (beamLen < 1) {
            _tribStripCache.elId = el.id;
            _tribStripCache.strip = null;
            return;
        }
        var ux = dx / beamLen, uy = dy / beamLen;
        var nx = -uy, ny = ux; // perpendicular

        var halfLeft  = tribResult.tribLeft  || tribResult.tribWidth / 2;
        var halfRight = tribResult.tribRight || tribResult.tribWidth / 2;

        // 4 corners of the trib strip in real-world mm
        _tribStripCache.elId = el.id;
        _tribStripCache.strip = {
            // corners: go along the beam, offset perp by trib distances
            // Note: "left" is in the +n direction, "right" is in -n direction
            p1: { x: el.x1 + nx * halfLeft,  y: el.y1 + ny * halfLeft },
            p2: { x: el.x2 + nx * halfLeft,  y: el.y2 + ny * halfLeft },
            p3: { x: el.x2 - nx * halfRight, y: el.y2 - ny * halfRight },
            p4: { x: el.x1 - nx * halfRight, y: el.y1 - ny * halfRight },
            tribWidth_mm: tribResult.tribWidth,
        };
    }

    // Patch the existing render callback to also draw the trib strip.
    // We wrap the existing fdDrawElements approach by adding another
    // callback that runs after it.
    if (typeof engine !== 'undefined' && engine._renderCallbacks) {
        engine._renderCallbacks.push(function _tribStripDraw(ctx, pass) {
            if (pass !== 'main') return;
            _updateTribStrip();
            var strip = _tribStripCache.strip;
            if (!strip) return;

            var coords = engine.coords;
            if (!coords || !coords.realToSheet || !coords.sheetToScreen) return;

            // Convert real → sheet → screen for each corner
            var corners = [strip.p1, strip.p2, strip.p3, strip.p4];
            var screenPts = [];
            for (var i = 0; i < 4; i++) {
                var sh = coords.realToSheet(corners[i].x, corners[i].y);
                var sc = coords.sheetToScreen(sh.x, sh.y);
                screenPts.push(sc);
            }

            // Draw translucent fill
            ctx.save();
            ctx.globalAlpha = 0.10;
            ctx.fillStyle = '#2B7CD0'; // main app accent
            ctx.beginPath();
            ctx.moveTo(screenPts[0].x, screenPts[0].y);
            for (var i = 1; i < 4; i++) {
                ctx.lineTo(screenPts[i].x, screenPts[i].y);
            }
            ctx.closePath();
            ctx.fill();

            // Dashed outline
            ctx.globalAlpha = 0.35;
            ctx.strokeStyle = '#2B7CD0';
            ctx.lineWidth = 1;
            ctx.setLineDash([6, 4]);
            ctx.stroke();

            // Label at midpoint
            ctx.globalAlpha = 0.7;
            ctx.setLineDash([]);
            var midX = (screenPts[0].x + screenPts[2].x) / 2;
            var midY = (screenPts[0].y + screenPts[2].y) / 2;
            ctx.font = '10px sans-serif';
            ctx.fillStyle = '#2B7CD0';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText('trib = ' + (strip.tribWidth_mm / 1000).toFixed(2) + ' m', midX, midY - 8);

            ctx.restore();
        });
    }

    // Clear trib strip cache when selection changes
    if (typeof document !== 'undefined') {
        // Listen for selection changes via the existing system
        var _origSelectForTrib = (typeof selectElement === 'function') ? selectElement : null;
        if (_origSelectForTrib) {
            // Note: we don't replace selectElement — the render callback
            // already re-checks selectedElement each frame via _updateTribStrip.
            // The cache auto-invalidates when selectedElement.id changes.
        }
    }

    // ============================================================
    // SB → FB silent auto-promotion (Slice 4)
    // ============================================================
    // When a beam on the S-BEAM layer (SB schedule) has floor zones
    // overlapping its tributary strip, it's functioning as a floor
    // bearer. Auto-promote it to the floorBeam (FB) schedule so the
    // enhanced engine runs correctly.
    //
    // This runs on-demand (when checking a beam) rather than as a
    // background process, to keep it predictable and undo-able.

    /**
     * Check if a beam should be promoted from SB to FB.
     * Criteria: beam has floor zones overlapping it, and it's currently
     * tagged with an SB-prefixed typeRef.
     *
     * @param {Object} beamEl — beam element
     * @returns {boolean} true if promoted
     */
    function maybePromoteToFloorBeam(beamEl) {
        if (!beamEl || beamEl.type !== 'line' || beamEl.layer !== 'S-BEAM') return false;

        var typeRef = beamEl.typeRef || beamEl.tag;
        if (!typeRef) return false;

        // Only promote SB-prefixed beams (not already FB)
        if (typeRef.indexOf('FB') === 0) return false;
        if (typeRef.indexOf('SB') !== 0) return false;

        // Check for floor zones
        if (typeof FloorLoadResolver === 'undefined') return false;
        var floorLoadSchedule = (typeof project !== 'undefined' && project.scheduleTypes)
            ? project.scheduleTypes.floorLoad || {}
            : {};
        var zones = FloorLoadResolver.resolveZonesFromElements(
            (typeof project !== 'undefined') ? project.elements : [], floorLoadSchedule
        );
        if (zones.length === 0) return false;

        // Check if any zone overlaps the beam's midpoint (quick heuristic)
        var mx = (beamEl.x1 + beamEl.x2) / 2;
        var my = (beamEl.y1 + beamEl.y2) / 2;
        var hit = FloorLoadResolver.getLoadAtPoint(mx, my, zones, 0, 0);
        if (!hit.zoneId) return false;

        // Beam has floor zones → promote to FB
        // Find or create a matching FB entry
        if (!project.scheduleTypes.floorBeam) project.scheduleTypes.floorBeam = {};
        var fbSchedule = project.scheduleTypes.floorBeam;

        // Get the SB schedule data
        var sbSchedule = project.scheduleTypes.beam || {};
        var sbData = sbSchedule[typeRef] || {};

        // Find an FB entry with the same size, or create one
        var targetFB = null;
        for (var key in fbSchedule) {
            if (fbSchedule[key].size === sbData.size && fbSchedule[key].grade === (sbData.grade || '300')) {
                targetFB = key;
                break;
            }
        }
        if (!targetFB) {
            // Create a new FB entry
            var fbCount = Object.keys(fbSchedule).length + 1;
            targetFB = 'FB' + fbCount;
            fbSchedule[targetFB] = {
                size: sbData.size || '',
                grade: sbData.grade || '300',
                notes: 'Auto-promoted from ' + typeRef,
            };
        }

        // Execute via command pattern for undo support
        if (typeof history !== 'undefined' && history.execute) {
            var oldTypeRef = typeRef;
            var newTypeRef = targetFB;
            history.execute({
                type: 'modifyElement',
                elementId: beamEl.id,
                description: 'Promote ' + oldTypeRef + ' → ' + newTypeRef + ' (floor bearer)',
                execute: function () {
                    beamEl.typeRef = newTypeRef;
                    if (beamEl.tag) beamEl.tag = newTypeRef;
                },
                undo: function () {
                    beamEl.typeRef = oldTypeRef;
                    if (beamEl.tag) beamEl.tag = oldTypeRef;
                },
            });
        } else {
            // No undo system — just set it directly
            beamEl.typeRef = targetFB;
            if (beamEl.tag) beamEl.tag = targetFB;
        }

        console.log('[floor-designer] Auto-promoted beam ' + beamEl.id + ' from ' + typeRef + ' → ' + targetFB);
        return true;
    }

    // Expose for external use
    window.floorDesigner.maybePromoteToFloorBeam = maybePromoteToFloorBeam;

    // Sanity log
    if (typeof console !== 'undefined') {
        console.log('[floor-designer] ' + window.floorDesigner.version + ' loaded — floor-zone tool ready');
    }
})();
