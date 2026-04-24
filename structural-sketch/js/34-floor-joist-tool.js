// ═══════════════════════════════════════════════════════════════════
// 34-floor-joist-tool.js — Fill-bay joist placement tool (Slice 5)
//
// Tool: user activates 'fillBayJoist' tool, clicks inside a bay
// (region between two parallel beams), and joists auto-populate at
// the selected spacing (450 or 600 c/c, per bay).
//
// This module handles:
//   1. Bay detection from beams on the current level
//   2. Fill-bay click interaction — place joists
//   3. hySPAN joist sizing per bay (warning-only scope gate)
//   4. Canvas rendering of placed joists + size label
//   5. Double-click bay → joist calc sheet
//   6. Properties panel info for joist bays
//
// Element model for a placed joist bay:
//   type:      'joistBay'
//   layer:     'S-JOIST'        (new layer)
//   typeRef:   'FJ1'            (joist schedule)
//   x0, y0, x1, y1:            bay bounds in real-world mm
//   joistSpan_mm:               clear span between bounding beams
//   spanDir:                    'X' or 'Y'
//   spacing_mm:                 450 or 600 (per bay)
//   spanType:                   'single' or 'continuous'
//   boundingBeams: [id, id]:    beam element IDs
//   joistResult:                cached HyspanJoists.sizeJoist() result
//
// Namespace: extends window.floorDesigner
// ═══════════════════════════════════════════════════════════════════

(function () {
    'use strict';

    // ── Layer registration ───────────────────────────────────────
    var joistLayerDef = {
        name: 'Floor Joists',
        color: '#8B5CF6',       // purple — distinct from beam (blue) and floor load (green)
        lineWeight: 0.10,
        pattern: 'solid',
        visible: true,
        locked: false,
        printColor: '#8B5CF6'
    };
    if (typeof LAYERS !== 'undefined' && LAYERS) {
        if (!LAYERS['S-JOIST']) LAYERS['S-JOIST'] = joistLayerDef;
    }
    // Also register in project.layers so getVisibleElements() includes joistBay elements
    if (typeof project !== 'undefined' && project.layers) {
        if (!project.layers['S-JOIST']) project.layers['S-JOIST'] = joistLayerDef;
    }

    // ── Joist bay state ──────────────────────────────────────────
    // Pre-computed bays for hover highlighting
    var _cachedBays = [];
    var _cachedBaysLevelId = null;
    var _hoveredBay = null;

    // Public invalidation — called when zone loads/span direction change
    window.invalidateJoistBayCache = function () {
        _cachedBays = [];
        _cachedBaysLevelId = null;
    };

    /**
     * Refresh cached bays from current elements + level.
     */
    function refreshBayCache() {
        if (typeof project === 'undefined' || !project.elements) return;
        var levelId = (typeof getActiveLevel === 'function') ? getActiveLevel().id : 'L1';
        if (levelId === _cachedBaysLevelId && _cachedBays.length > 0) return;

        // Resolve floor zones first — span direction comes from the zone, not level settings
        var zones = [];
        if (typeof FloorLoadResolver !== 'undefined' && typeof FloorLoadResolver.resolveZonesFromElements === 'function') {
            zones = FloorLoadResolver.resolveZonesFromElements(
                project.elements,
                project.scheduleTypes.floorLoad || {}
            );
        }

        // Determine span direction: prefer the first floor zone's spanDirection,
        // fall back to level design settings
        var ds = (typeof getLevelDesignSettings === 'function') ? getLevelDesignSettings(levelId) : { spanDirection: 0 };
        var spanDirDeg = ds.spanDirection || 0;
        if (zones.length > 0) {
            spanDirDeg = zones[0].spanDirection_deg || 0;
        }
        // 0 = joists span in X. 90 = joists span in Y.
        var spanDir = (spanDirDeg >= 45 && spanDirDeg < 135) ? 'Y' : 'X';

        _cachedBays = FloorLoadResolver.enumerateBays(project.elements, levelId, spanDir);
        _cachedBaysLevelId = levelId;

        var defaultG = ds.G || 1.2;
        var defaultQ = ds.Q || 1.5;

        for (var i = 0; i < _cachedBays.length; i++) {
            var bay = _cachedBays[i];
            bay.load = FloorLoadResolver.bayWorstCaseLoad(bay, zones, defaultG, defaultQ);
        }
    }

    /**
     * Find which cached bay contains a real-world point.
     */
    function bayAtPoint(rx, ry) {
        for (var i = _cachedBays.length - 1; i >= 0; i--) {
            var b = _cachedBays[i];
            if (rx >= b.x0_mm && rx <= b.x1_mm && ry >= b.y0_mm && ry <= b.y1_mm) {
                return b;
            }
        }
        return null;
    }

    /**
     * Check if a bay already has joists placed.
     */
    function bayHasJoists(bay) {
        if (!bay || typeof project === 'undefined') return false;
        for (var i = 0; i < project.elements.length; i++) {
            var el = project.elements[i];
            if (!el || el.type !== 'joistBay') continue;
            // Match by bounding box overlap (within tolerance)
            if (Math.abs(el.x0 - bay.x0_mm) < 50 &&
                Math.abs(el.y0 - bay.y0_mm) < 50 &&
                Math.abs(el.x1 - bay.x1_mm) < 50 &&
                Math.abs(el.y1 - bay.y1_mm) < 50) {
                return true;
            }
        }
        return false;
    }

    /**
     * Place joists in a bay. Creates a joistBay element via
     * command pattern for undo support.
     */
    function fillBayWithJoists(bay) {
        if (!bay) return;
        if (bayHasJoists(bay)) {
            console.log('[joist-tool] Bay already has joists');
            return;
        }

        // Determine spacing from joist schedule
        var joistSchedule = project.scheduleTypes.joist || {};
        var typeRef = (typeof placementTypeRef !== 'undefined' && placementTypeRef.joist)
            ? placementTypeRef.joist : 'FJ1';

        // Ensure FJ1 exists in schedule
        if (!joistSchedule[typeRef]) {
            joistSchedule[typeRef] = {
                size: '',
                spacing: 450,
                spanType: 'single',
                notes: 'Auto-created',
            };
            if (!project.scheduleTypes.joist) project.scheduleTypes.joist = joistSchedule;
        }

        var schedEntry = joistSchedule[typeRef];
        var spacing_mm = Number(schedEntry.spacing) || 450;
        var spanType = schedEntry.spanType || 'single';

        // Level design settings for defaults
        var levelId = (typeof getActiveLevel === 'function') ? getActiveLevel().id : 'L1';
        var ds = (typeof getLevelDesignSettings === 'function') ? getLevelDesignSettings(levelId) : { spanDirection: 0, G: 1.2, Q: 1.5 };
        var spanDir = (ds.spanDirection >= 45 && ds.spanDirection < 135) ? 'Y' : 'X';

        // Size the joist
        var joistResult = null;
        if (typeof HyspanJoists !== 'undefined' && HyspanJoists.sizeJoist) {
            joistResult = HyspanJoists.sizeJoist({
                span_m: bay.joistSpan_m,
                spacing_mm: spacing_mm,
                spanType: spanType,
                G_kPa: bay.load ? bay.load.G_kPa : (ds.G || 1.2),
                Q_kPa: bay.load ? bay.load.Q_kPa : (ds.Q || 1.5),
            });
        }

        // Create element
        var newId = 'joist-' + Date.now() + '-' + Math.random().toString(36).substr(2, 4);
        var el = {
            id: newId,
            type: 'joistBay',
            layer: 'S-JOIST',
            level: levelId,
            typeRef: typeRef,
            x0: bay.x0_mm,
            y0: bay.y0_mm,
            x1: bay.x1_mm,
            y1: bay.y1_mm,
            joistSpan_mm: bay.joistSpan_mm,
            spanDir: spanDir,
            spacing_mm: spacing_mm,
            spanType: spanType,
            boundingBeams: bay.boundingBeams || [],
            joistResult: joistResult,
        };

        // Execute via command pattern
        if (typeof history !== 'undefined' && history.execute) {
            history.execute({
                type: 'addElement',
                elementId: newId,
                description: 'Fill bay with joists (' + typeRef + ')',
                execute: function () {
                    project.elements.push(el);
                    _cachedBaysLevelId = null; // invalidate cache
                },
                undo: function () {
                    var idx = project.elements.indexOf(el);
                    if (idx >= 0) project.elements.splice(idx, 1);
                    _cachedBaysLevelId = null;
                },
            });
        } else {
            project.elements.push(el);
        }

        // Log scope warning if outside AS 1684 envelope
        if (joistResult && joistResult.loadCheck && !joistResult.loadCheck.ok) {
            console.warn('[joist-tool] Bay loads outside AS 1684 scope: ' + joistResult.loadCheck.reason);
        }

        // Update canvas
        if (typeof engine !== 'undefined' && engine.requestRender) {
            engine.requestRender();
        }

        console.log('[joist-tool] Filled bay ' + bay.id + ' with ' + typeRef +
            (joistResult && joistResult.ok ? ' → ' + joistResult.section.name + ' ' + joistResult.section.product : ''));
    }

    // ── Input event handlers ─────────────────────────────────────
    if (typeof container !== 'undefined') {

        // Click handler for fill-bay tool
        container.addEventListener('mousedown', function (e) {
            if (typeof activeTool === 'undefined' || activeTool !== 'fillBayJoist') return;
            if (e.button !== 0) return;

            refreshBayCache();
            var sheetPos = engine.coords.screenToSheet(e.offsetX, e.offsetY);
            var realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);
            var bay = bayAtPoint(realPos.x, realPos.y);

            if (bay && !bayHasJoists(bay)) {
                fillBayWithJoists(bay);
            }
        });

        // Hover handler for bay highlight
        container.addEventListener('mousemove', function (e) {
            if (typeof activeTool === 'undefined' || activeTool !== 'fillBayJoist') return;

            refreshBayCache();
            var sheetPos = engine.coords.screenToSheet(e.offsetX, e.offsetY);
            var realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);
            _hoveredBay = bayAtPoint(realPos.x, realPos.y);

            if (typeof engine !== 'undefined' && engine.requestRender) {
                engine.requestRender();
            }
        });
    }

    // ── setActiveTool wrapper ────────────────────────────────────
    if (typeof setActiveTool === 'function') {
        var _origSetActiveTool_JT = setActiveTool;
        setActiveTool = function (tool) {
            if (tool !== 'fillBayJoist') {
                _hoveredBay = null;
            }
            if (tool === 'fillBayJoist') {
                _cachedBaysLevelId = null; // force refresh
                refreshBayCache();
            }
            _origSetActiveTool_JT(tool);

            // Toolbar button active state
            var btn = document.getElementById('btn-floor-joist');
            if (btn) btn.classList.toggle('active', tool === 'fillBayJoist');

            // Cursor
            if (tool === 'fillBayJoist') {
                container.style.cursor = 'crosshair';
            }

            // Status bar
            var statusTool = document.getElementById('status-tool');
            if (statusTool && tool === 'fillBayJoist') {
                statusTool.textContent = 'Floor Joists';
            }
        };
    }

    // ── Keyboard shortcut ────────────────────────────────────────
    window.addEventListener('keydown', function (e) {
        if (document.activeElement !== document.body) return;
        if (e.ctrlKey || e.metaKey || e.altKey) return;
        if (e.key === 'j') setActiveTool('fillBayJoist');
    });

    // ── Canvas rendering ─────────────────────────────────────────
    if (typeof engine !== 'undefined' && Array.isArray(engine._renderCallbacks)) {

        engine._renderCallbacks.push(function _joistDraw(ctx, eng) {
            var coords = eng.coords;
            var zoom = eng.viewport.zoom;

            // ── Draw placed joist bays ──
            var elements = (typeof project !== 'undefined' && project.getVisibleElements)
                ? project.getVisibleElements()
                : (typeof project !== 'undefined' ? project.elements : []);

            for (var i = 0; i < elements.length; i++) {
                var el = elements[i];
                if (!el || el.type !== 'joistBay') continue;

                var layer = LAYERS['S-JOIST'];
                if (layer && !layer.visible) continue;

                // Convert bay corners to screen
                var sh0 = coords.realToSheet(el.x0, el.y0);
                var sh1 = coords.realToSheet(el.x1, el.y1);
                var sc0 = coords.sheetToScreen(sh0.x, sh0.y);
                var sc1 = coords.sheetToScreen(sh1.x, sh1.y);
                var x = Math.min(sc0.x, sc1.x);
                var y = Math.min(sc0.y, sc1.y);
                var w = Math.abs(sc1.x - sc0.x);
                var h = Math.abs(sc1.y - sc0.y);

                // Fill
                var isSelected = (typeof selectedElement !== 'undefined' && selectedElement === el);
                ctx.save();
                ctx.globalAlpha = isSelected ? 0.15 : 0.08;
                ctx.fillStyle = '#8B5CF6';
                ctx.fillRect(x, y, w, h);

                // Joist lines
                ctx.globalAlpha = 0.4;
                ctx.strokeStyle = '#8B5CF6';
                ctx.lineWidth = Math.max(0.5, 0.08 * zoom);
                ctx.setLineDash([4, 3]);

                var spacing_px;
                if (el.spanDir === 'X') {
                    // Joists span in X → draw horizontal lines at spacing in Y
                    var spacing_real = el.spacing_mm || 450;
                    var sh_top = coords.realToSheet(el.x0, el.y0);
                    var sh_top2 = coords.realToSheet(el.x0, el.y0 + spacing_real);
                    var sc_top = coords.sheetToScreen(sh_top.x, sh_top.y);
                    var sc_top2 = coords.sheetToScreen(sh_top2.x, sh_top2.y);
                    spacing_px = Math.abs(sc_top2.y - sc_top.y);
                    if (spacing_px > 3) {
                        for (var jy = y + spacing_px; jy < y + h - 1; jy += spacing_px) {
                            ctx.beginPath();
                            ctx.moveTo(x, jy);
                            ctx.lineTo(x + w, jy);
                            ctx.stroke();
                        }
                    }
                } else {
                    // Joists span in Y → draw vertical lines at spacing in X
                    var spacing_real = el.spacing_mm || 450;
                    var sh_left = coords.realToSheet(el.x0, el.y0);
                    var sh_left2 = coords.realToSheet(el.x0 + spacing_real, el.y0);
                    var sc_left = coords.sheetToScreen(sh_left.x, sh_left.y);
                    var sc_left2 = coords.sheetToScreen(sh_left2.x, sh_left2.y);
                    spacing_px = Math.abs(sc_left2.x - sc_left.x);
                    if (spacing_px > 3) {
                        for (var jx = x + spacing_px; jx < x + w - 1; jx += spacing_px) {
                            ctx.beginPath();
                            ctx.moveTo(jx, y);
                            ctx.lineTo(jx, y + h);
                            ctx.stroke();
                        }
                    }
                }

                ctx.setLineDash([]);

                // Label
                ctx.globalAlpha = 0.9;
                ctx.fillStyle = '#6D28D9';
                ctx.font = Math.max(9, 10 * Math.sqrt(zoom)) + 'px sans-serif';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                var label = el.typeRef || 'FJ';
                if (el.joistResult && el.joistResult.ok) {
                    label += ' ' + el.joistResult.section.name;
                    label += ' @ ' + (el.spacing_mm || 450);
                } else if (el.joistResult && !el.joistResult.ok) {
                    label = '⚠ ' + (el.joistResult.reason || 'No size').substring(0, 30);
                }
                ctx.fillText(label, x + w / 2, y + h / 2);

                // Scope warning badge
                if (el.joistResult && el.joistResult.loadCheck && !el.joistResult.loadCheck.ok) {
                    ctx.fillStyle = '#D97706';
                    ctx.font = Math.max(7, 8 * Math.sqrt(zoom)) + 'px sans-serif';
                    ctx.fillText('⚠ Outside AS 1684 scope', x + w / 2, y + h / 2 + 14);
                }

                // Selection outline
                if (isSelected) {
                    ctx.globalAlpha = 0.6;
                    ctx.strokeStyle = '#6D28D9';
                    ctx.lineWidth = 2;
                    ctx.setLineDash([]);
                    ctx.strokeRect(x, y, w, h);
                }

                ctx.restore();
            }

            // ── All detected bay outlines (fill-bay tool active) ──
            if (typeof activeTool !== 'undefined' && activeTool === 'fillBayJoist' && _cachedBays.length > 0) {
                for (var bi = 0; bi < _cachedBays.length; bi++) {
                    var hintBay = _cachedBays[bi];
                    if (_hoveredBay && hintBay.id === _hoveredBay.id) continue; // skip hovered (drawn separately)
                    var filled = bayHasJoists(hintBay);
                    var hsh0 = coords.realToSheet(hintBay.x0_mm, hintBay.y0_mm);
                    var hsh1 = coords.realToSheet(hintBay.x1_mm, hintBay.y1_mm);
                    var hsc0 = coords.sheetToScreen(hsh0.x, hsh0.y);
                    var hsc1 = coords.sheetToScreen(hsh1.x, hsh1.y);
                    var hbx = Math.min(hsc0.x, hsc1.x);
                    var hby = Math.min(hsc0.y, hsc1.y);
                    var hbw = Math.abs(hsc1.x - hsc0.x);
                    var hbh = Math.abs(hsc1.y - hsc0.y);

                    ctx.save();
                    ctx.globalAlpha = filled ? 0.03 : 0.06;
                    ctx.fillStyle = filled ? '#999' : '#8B5CF6';
                    ctx.fillRect(hbx, hby, hbw, hbh);
                    ctx.globalAlpha = filled ? 0.15 : 0.3;
                    ctx.strokeStyle = filled ? '#999' : '#8B5CF6';
                    ctx.lineWidth = 1;
                    ctx.setLineDash([6, 4]);
                    ctx.strokeRect(hbx, hby, hbw, hbh);
                    ctx.setLineDash([]);

                    if (!filled) {
                        ctx.globalAlpha = 0.4;
                        ctx.fillStyle = '#6D28D9';
                        ctx.font = '9px sans-serif';
                        ctx.textAlign = 'center';
                        ctx.textBaseline = 'middle';
                        ctx.fillText('click to fill', hbx + hbw / 2, hby + hbh / 2);
                    }
                    ctx.restore();
                }
            }

            // ── Hovered bay highlight (fill-bay tool active) ──
            if (_hoveredBay && typeof activeTool !== 'undefined' && activeTool === 'fillBayJoist') {
                var bay = _hoveredBay;
                var filled = bayHasJoists(bay);
                var sh0 = coords.realToSheet(bay.x0_mm, bay.y0_mm);
                var sh1 = coords.realToSheet(bay.x1_mm, bay.y1_mm);
                var sc0 = coords.sheetToScreen(sh0.x, sh0.y);
                var sc1 = coords.sheetToScreen(sh1.x, sh1.y);
                var bx = Math.min(sc0.x, sc1.x);
                var by = Math.min(sc0.y, sc1.y);
                var bw = Math.abs(sc1.x - sc0.x);
                var bh = Math.abs(sc1.y - sc0.y);

                ctx.save();
                ctx.globalAlpha = filled ? 0.05 : 0.12;
                ctx.fillStyle = filled ? '#999' : '#8B5CF6';
                ctx.fillRect(bx, by, bw, bh);
                ctx.globalAlpha = filled ? 0.2 : 0.5;
                ctx.strokeStyle = filled ? '#999' : '#8B5CF6';
                ctx.lineWidth = 2;
                ctx.setLineDash([8, 4]);
                ctx.strokeRect(bx, by, bw, bh);
                ctx.setLineDash([]);

                // Label
                ctx.globalAlpha = 0.8;
                ctx.fillStyle = filled ? '#666' : '#6D28D9';
                ctx.font = '11px sans-serif';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.fillText(
                    filled ? 'Joists placed' : 'Click to fill (' + bay.joistSpan_m.toFixed(2) + ' m span)',
                    bx + bw / 2, by + bh / 2
                );
                ctx.restore();
            }
        });
    }

    // ── Hit testing — select joist bay elements ──────────────────
    if (typeof hitTestElement === 'function') {
        var _origHitTest_JT = hitTestElement;
        hitTestElement = function (sheetPos, tolerance) {
            var hit = _origHitTest_JT(sheetPos, tolerance);
            if (hit) return hit;

            // Check joistBay elements (point-in-rect in real coordinates)
            if (typeof project !== 'undefined' && project.elements) {
                var realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);
                for (var i = project.elements.length - 1; i >= 0; i--) {
                    var el = project.elements[i];
                    if (!el || el.type !== 'joistBay') continue;
                    if (realPos.x >= el.x0 && realPos.x <= el.x1 &&
                        realPos.y >= el.y0 && realPos.y <= el.y1) {
                        return el;
                    }
                }
            }
            return null;
        };
    }

    // ── Double-click joist bay → calc sheet ──────────────────────
    if (typeof container !== 'undefined') {
        container.addEventListener('dblclick', function (e) {
            if (typeof currentTool !== 'undefined' && currentTool && currentTool !== 'pointer') return;
            if (typeof activeTool !== 'undefined' && activeTool && activeTool !== 'pointer') return;
            if (typeof FloorCalcSheet === 'undefined' || !FloorCalcSheet.showBayCalcSheetModal) return;

            var el = (typeof selectedElement !== 'undefined') ? selectedElement : null;
            if (!el || el.type !== 'joistBay') return;

            FloorCalcSheet.showBayCalcSheetModal(el);
        });
    }

    // ── Properties panel branch for joistBay ─────────────────────
    // This will be picked up by 12-properties.js's extension point.
    // We expose a function that generates the HTML.
    window.floorDesigner = window.floorDesigner || {};
    window.floorDesigner.buildJoistBayPropsHTML = function (el) {
        if (!el || el.type !== 'joistBay') return '';

        var html = '';
        var jr = el.joistResult;

        // Type badge
        html += '<div style="margin-bottom:6px;">';
        html += '<span class="prop-type-badge" style="background:#8B5CF6; color:white; padding:2px 8px; border-radius:3px; font-size:10px; font-weight:600;">' + (el.typeRef || 'FJ') + '</span>';
        html += '</div>';

        // Bay info
        html += '<div style="font-size:10px; color:#666;">';
        html += '<div style="display:flex; justify-content:space-between;"><span>Span</span><span>' + (el.joistSpan_mm / 1000).toFixed(2) + ' m (' + el.spanDir + '-dir)</span></div>';
        html += '<div style="display:flex; justify-content:space-between;"><span>Spacing</span><span>' + (el.spacing_mm || 450) + ' mm c/c</span></div>';
        html += '<div style="display:flex; justify-content:space-between;"><span>Span type</span><span>' + (el.spanType || 'single') + '</span></div>';
        html += '<div style="display:flex; justify-content:space-between;"><span>Bay size</span><span>' + ((el.x1 - el.x0) / 1000).toFixed(2) + ' × ' + ((el.y1 - el.y0) / 1000).toFixed(2) + ' m</span></div>';
        html += '</div>';

        // Joist result
        if (jr) {
            if (jr.ok) {
                var utilColor = jr.utilisation <= 0.9 ? '#16A34A' : jr.utilisation <= 1.0 ? '#D97706' : '#DC2626';
                html += '<div style="margin-top:6px; padding:4px; background:' + utilColor + '08; border-left:3px solid ' + utilColor + '; border-radius:0 3px 3px 0; font-size:10px;">';
                html += '<div style="font-weight:600; color:' + utilColor + ';">PASS — ' + (jr.utilisation * 100).toFixed(0) + '%</div>';
                html += '<div style="color:#666;">' + jr.section.name + ' ' + jr.section.product + '</div>';
                html += '<div style="color:#666;">Max span: ' + jr.maxSpan_m.toFixed(1) + ' m</div>';
                html += '</div>';
            } else {
                html += '<div style="margin-top:6px; padding:4px; background:#DC262608; border-left:3px solid #DC2626; border-radius:0 3px 3px 0; font-size:10px;">';
                html += '<div style="font-weight:600; color:#DC2626;">FAIL</div>';
                html += '<div style="color:#666;">' + (jr.reason || 'No size available') + '</div>';
                html += '</div>';
            }

            // Scope warning
            if (jr.loadCheck && !jr.loadCheck.ok) {
                html += '<div style="font-size:9px; color:#D97706; margin-top:4px;">⚠ ' + jr.loadCheck.reason + '</div>';
                html += '<div style="font-size:9px; color:#D97706;">Engineered design to AS 1720.1 required.</div>';
            }
        }

        // Calc sheet button
        if (typeof FloorCalcSheet !== 'undefined' && FloorCalcSheet.showBayCalcSheetModal) {
            html += '<div style="margin-top:6px; text-align:center;">';
            html += '<button onclick="FloorCalcSheet.showBayCalcSheetModal(selectedElement)" style="font-size:10px; padding:3px 14px; border:1px solid #8B5CF6; color:#8B5CF6; background:transparent; border-radius:3px; cursor:pointer;">View Calc Sheet</button>';
            html += '</div>';
        }

        return html;
    };

    // ── Invalidation hooks ───────────────────────────────────────
    // When elements change, invalidate the bay cache so it
    // recalculates next time the tool activates.
    if (typeof project !== 'undefined') {
        var _origPush = Array.prototype.push;
        // Simple approach: clear cache when any new element is added.
        // More robust approach would use MutationObserver on the array,
        // but for now this is sufficient.
        window.floorDesigner._invalidateBayCache = function () {
            _cachedBaysLevelId = null;
        };
    }

    // ── Toolbar button wiring ────────────────────────────────────
    // Main ribbon button now activates the new joistZone tool (see 37-joist-zone.js).
    // The old fillBayJoist tool is kept for backwards-compat with saved joistBay elements,
    // but not exposed via the ribbon. Users can still activate it via activeTool='fillBayJoist'
    // in the console if needed.

    // ── Register placement type ref ──────────────────────────────
    if (typeof placementTypeRef !== 'undefined' && placementTypeRef) {
        if (!placementTypeRef.joist) placementTypeRef.joist = 'FJ1';
    }

    // ── Console log ──────────────────────────────────────────────
    if (typeof console !== 'undefined') {
        console.log('[joist-tool] 1.0-slice5 loaded — fill-bay tool ready');
    }
})();
