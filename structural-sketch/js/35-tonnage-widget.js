// ═══════════════════════════════════════════════════════════════════
// 35-tonnage-widget.js — Material tonnage summary widget (Slice 6)
//
// Floating panel showing live material quantities:
//   - Steel beams: total mass from section catalogue × length
//   - Timber joists: total mass from hySPAN sections × count × span
//   - Category totals (steel tonnes, timber kg/m³)
//
// Recalculates on every render frame (debounced to avoid perf hit).
// Panel is toggled from a toolbar button.
//
// Namespace: window.tonnageWidget
// ═══════════════════════════════════════════════════════════════════

(function () {
    'use strict';

    // ── Aggregation engine ───────────────────────────────────────

    /**
     * Compute material quantities for the entire project.
     *
     * @returns {Object} {
     *   steel: { count, totalLength_m, totalMass_kg, items: [...] },
     *   timber: { count, totalMass_kg, totalVolume_m3, bayCount, items: [...] },
     *   total_kg: number,
     *   timestamp: number
     * }
     */
    function computeTonnage() {
        var result = {
            steel: { count: 0, totalLength_m: 0, totalMass_kg: 0, items: [] },
            timber: { count: 0, totalMass_kg: 0, totalVolume_m3: 0, bayCount: 0, items: [] },
            total_kg: 0,
            timestamp: Date.now(),
        };

        if (typeof project === 'undefined' || !project.elements) return result;

        var elements = project.elements;
        var beamSchedule = project.scheduleTypes.beam || {};
        var floorBeamSchedule = project.scheduleTypes.floorBeam || {};

        // ── Steel beams ──
        for (var i = 0; i < elements.length; i++) {
            var el = elements[i];
            if (!el) continue;

            // Beam elements (S-BEAM layer, type 'line')
            if (el.type === 'line' && el.layer === 'S-BEAM') {
                var typeRef = el.typeRef || el.tag;
                if (!typeRef) continue;

                var schedData = beamSchedule[typeRef] || floorBeamSchedule[typeRef] || {};
                var sizeStr = schedData.size;
                if (!sizeStr) continue;

                // Get mass per metre from section properties
                var mass_kg_m = 0;
                if (typeof getSectionProperties === 'function') {
                    var props = getSectionProperties(sizeStr);
                    if (props) mass_kg_m = props.mass || props.m || 0;
                }

                // Beam length
                var dx = (el.x2 || 0) - (el.x1 || 0);
                var dy = (el.y2 || 0) - (el.y1 || 0);
                var length_mm = Math.sqrt(dx * dx + dy * dy);
                var length_m = length_mm / 1000;

                if (length_m > 0 && mass_kg_m > 0) {
                    var mass_kg = mass_kg_m * length_m;
                    result.steel.count++;
                    result.steel.totalLength_m += length_m;
                    result.steel.totalMass_kg += mass_kg;
                    result.steel.items.push({
                        id: el.id,
                        typeRef: typeRef,
                        size: sizeStr,
                        length_m: length_m,
                        mass_kg_m: mass_kg_m,
                        mass_kg: mass_kg,
                    });
                }
            }

            // Column elements
            if (el.type === 'column') {
                var typeRef = el.typeRef || el.tag;
                if (!typeRef) continue;

                var colSchedule = project.scheduleTypes.column || {};
                var schedData = colSchedule[typeRef] || {};
                var sizeStr = schedData.size;
                if (!sizeStr) continue;

                var mass_kg_m = 0;
                if (typeof getSectionProperties === 'function') {
                    var props = getSectionProperties(sizeStr);
                    if (props) mass_kg_m = props.mass || props.m || 0;
                }

                // Column height — derive from level system
                var height_m = 2.7; // default 2700mm
                if (typeof levelSystem !== 'undefined' && levelSystem.levels) {
                    var lvIdx = levelSystem.levels.findIndex(function (l) { return l.id === el.level; });
                    if (lvIdx >= 0) {
                        var lv = levelSystem.levels[lvIdx];
                        height_m = (lv.height || 2700) / 1000;
                    }
                }

                if (mass_kg_m > 0) {
                    var mass_kg = mass_kg_m * height_m;
                    result.steel.count++;
                    result.steel.totalLength_m += height_m;
                    result.steel.totalMass_kg += mass_kg;
                    result.steel.items.push({
                        id: el.id,
                        typeRef: typeRef,
                        size: sizeStr,
                        length_m: height_m,
                        mass_kg_m: mass_kg_m,
                        mass_kg: mass_kg,
                        isColumn: true,
                    });
                }
            }

            // Joist bay elements
            if (el.type === 'joistBay') {
                var jr = el.joistResult;
                if (!jr || !jr.ok || !jr.section) continue;

                var joistSpan_m = (el.joistSpan_mm || 0) / 1000;
                var spacing_mm = el.spacing_mm || 450;
                var mass_kg_m = jr.section.mass_kg_m || 0;

                // Number of joists in the bay
                var bayWidth_mm;
                if (el.spanDir === 'X') {
                    // Joists span X, distributed along Y
                    bayWidth_mm = Math.abs((el.y1 || 0) - (el.y0 || 0));
                } else {
                    // Joists span Y, distributed along X
                    bayWidth_mm = Math.abs((el.x1 || 0) - (el.x0 || 0));
                }
                var numJoists = Math.max(1, Math.floor(bayWidth_mm / spacing_mm));

                var totalLength_m = numJoists * joistSpan_m;
                var totalMass_kg = totalLength_m * mass_kg_m;

                // Volume (rectangular cross-section)
                var D_mm = jr.section.D_mm || 0;
                var B_mm = jr.section.B_mm || 0;
                var volume_m3 = (D_mm * B_mm / 1e6) * totalLength_m;

                result.timber.bayCount++;
                result.timber.count += numJoists;
                result.timber.totalMass_kg += totalMass_kg;
                result.timber.totalVolume_m3 += volume_m3;
                result.timber.items.push({
                    id: el.id,
                    typeRef: el.typeRef || 'FJ',
                    section: jr.section.name,
                    product: jr.section.product,
                    numJoists: numJoists,
                    span_m: joistSpan_m,
                    totalLength_m: totalLength_m,
                    mass_kg: totalMass_kg,
                    volume_m3: volume_m3,
                });
            }

            // Joist zone elements (new polygon-based system)
            if (el.type === 'joistZone' && el.computed && el.computed.joistLines) {
                var gov = el.computed.governingResult;
                if (gov && gov.section) {
                    var jzMass = gov.section.mass_kg_m || 0;
                    var jzD = gov.section.D_mm || 0;
                    var jzB = gov.section.B_mm || 0;
                    var jzTotalLen = 0;
                    var jzCount = 0;
                    for (var jzi = 0; jzi < el.computed.joistLines.length; jzi++) {
                        var jzSpans = el.computed.joistLines[jzi].spans;
                        for (var jzs = 0; jzs < jzSpans.length; jzs++) {
                            jzTotalLen += jzSpans[jzs].span_mm / 1000;
                            jzCount++;
                        }
                    }
                    var jzTotalMass = jzTotalLen * jzMass;
                    var jzVolume = (jzD * jzB / 1e6) * jzTotalLen;

                    result.timber.bayCount++;
                    result.timber.count += jzCount;
                    result.timber.totalMass_kg += jzTotalMass;
                    result.timber.totalVolume_m3 += jzVolume;
                    result.timber.items.push({
                        id: el.id,
                        typeRef: el.typeRef || 'FJ',
                        section: gov.section.name,
                        product: gov.section.product,
                        numJoists: jzCount,
                        span_m: (el.computed.governingSpan_mm || 0) / 1000,
                        totalLength_m: jzTotalLen,
                        mass_kg: jzTotalMass,
                        volume_m3: jzVolume,
                        bayLabel: el.bayLabel,
                    });
                }
            }
        }

        result.total_kg = result.steel.totalMass_kg + result.timber.totalMass_kg;
        return result;
    }

    // ── Debounced update ─────────────────────────────────────────
    var _lastTonnage = null;
    var _debounceTimer = null;

    function requestUpdate() {
        if (_debounceTimer) return; // already scheduled
        _debounceTimer = setTimeout(function () {
            _debounceTimer = null;
            _lastTonnage = computeTonnage();
            renderPanel(_lastTonnage);
        }, 500); // 500ms debounce
    }

    // ── Panel rendering ──────────────────────────────────────────

    function renderPanel(data) {
        var body = document.getElementById('tonnage-panel-body');
        if (!body) return;

        if (!data) {
            body.innerHTML = '<div style="color:#999; font-size:10px; padding:8px;">No data</div>';
            return;
        }

        var steelT = (data.steel.totalMass_kg / 1000).toFixed(2);
        var timberKg = data.timber.totalMass_kg.toFixed(0);
        var timberM3 = data.timber.totalVolume_m3.toFixed(3);
        var totalT = (data.total_kg / 1000).toFixed(2);

        var html = '';

        // Steel section
        html += '<div style="margin-bottom:8px;">';
        html += '<div style="font-weight:600; font-size:10px; color:#1a5276; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">Steel</div>';
        if (data.steel.count > 0) {
            html += '<div style="font-size:22px; font-weight:700; color:#2c3e50; font-family:\'SF Mono\',Consolas,monospace;">' + steelT + ' <span style="font-size:11px; font-weight:400; color:#7f8c8d;">tonnes</span></div>';
            html += '<div style="font-size:9px; color:#7f8c8d;">' + data.steel.count + ' members · ' + data.steel.totalLength_m.toFixed(1) + ' m total</div>';
        } else {
            html += '<div style="font-size:10px; color:#999;">No steel elements</div>';
        }
        html += '</div>';

        // Timber section
        html += '<div style="margin-bottom:8px; padding-top:6px; border-top:1px solid #e0e3e6;">';
        html += '<div style="font-weight:600; font-size:10px; color:#1a5276; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">Timber (hySPAN)</div>';
        if (data.timber.bayCount > 0) {
            html += '<div style="font-size:22px; font-weight:700; color:#2c3e50; font-family:\'SF Mono\',Consolas,monospace;">' + timberM3 + ' <span style="font-size:11px; font-weight:400; color:#7f8c8d;">m³</span></div>';
            html += '<div style="font-size:9px; color:#7f8c8d;">' + data.timber.count + ' joists in ' + data.timber.bayCount + ' bays · ' + timberKg + ' kg</div>';
        } else {
            html += '<div style="font-size:10px; color:#999;">No joist bays</div>';
        }
        html += '</div>';

        // Total
        html += '<div style="padding-top:6px; border-top:2px solid #1a5276;">';
        html += '<div style="display:flex; justify-content:space-between; align-items:baseline;">';
        html += '<span style="font-weight:700; font-size:10px; color:#1a5276; text-transform:uppercase;">Total</span>';
        html += '<span style="font-size:16px; font-weight:700; color:#2c3e50; font-family:\'SF Mono\',Consolas,monospace;">' + totalT + ' t</span>';
        html += '</div>';
        html += '</div>';

        body.innerHTML = html;
    }

    // ── Toggle panel visibility ──────────────────────────────────

    function togglePanel() {
        var panel = document.getElementById('tonnage-panel');
        if (!panel) return;

        if (panel.classList.contains('hidden')) {
            panel.classList.remove('hidden');
            // Compute synchronously so the first frame shows data, not a placeholder
            _lastTonnage = computeTonnage();
            renderPanel(_lastTonnage);
        } else {
            panel.classList.add('hidden');
        }
    }

    function showPanel() {
        var panel = document.getElementById('tonnage-panel');
        if (panel) {
            panel.classList.remove('hidden');
            _lastTonnage = computeTonnage();
            renderPanel(_lastTonnage);
        }
    }

    function hidePanel() {
        var panel = document.getElementById('tonnage-panel');
        if (panel) panel.classList.add('hidden');
    }

    // ── Hook into render loop for live updates ───────────────────
    // When the panel is visible, recalculate on each render.
    if (typeof engine !== 'undefined' && Array.isArray(engine._renderCallbacks)) {
        engine._renderCallbacks.push(function _tonnageUpdate() {
            var panel = document.getElementById('tonnage-panel');
            if (!panel || panel.classList.contains('hidden')) return;
            requestUpdate();
        });
    }

    // ── Toolbar button wiring ────────────────────────────────────
    var btn = document.getElementById('btn-tonnage');
    if (btn) {
        btn.addEventListener('click', function () {
            togglePanel();
            btn.classList.toggle('active', !document.getElementById('tonnage-panel').classList.contains('hidden'));
        });
    }

    // ── Export ────────────────────────────────────────────────────
    window.tonnageWidget = {
        version: '1.0-slice6',
        computeTonnage: computeTonnage,
        togglePanel: togglePanel,
        showPanel: showPanel,
        hidePanel: hidePanel,
        requestUpdate: requestUpdate,
    };

    if (typeof console !== 'undefined') {
        console.log('[tonnage-widget] 1.0-slice6 loaded');
    }
})();
