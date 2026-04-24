// ═══════════════════════════════════════════════════════════════════
// 32-calc-sheet.js — A4 calc-sheet renderer for steel beams
//
// Ported from Floor Designer prototype calc-sheet.js (Phase 5).
// Pure JS, no DOM at module load. Namespace: window.FloorCalcSheet
//
// Changes from prototype:
//   - Namespace: FloorCalcSheet (not FloorDesigner.calcSheet)
//   - Section data comes from getSectionProperties() / FloorCalcEngine
//   - Constants from FloorCalcEngine.constants
//   - Bay/joist calc sheet deferred to Slice 5 — only beam sheet here
//   - Header pulls project.projectInfo (main app metadata)
//   - Diagram SVGs identical to prototype (inline, print-clean)
//   - renderBeamCalcSheet(ctx) is the main export
//   - showCalcSheetModal(beamEl) opens the overlay
//
// Reference: AS 4100-1998, AS/NZS 1170.0:2002, AS/NZS 3679.1.
// ═══════════════════════════════════════════════════════════════════

(function () {
    'use strict';

    // ── Number formatting ─────────────────────────────────────────
    var fmt0 = function (x) { return (x == null || isNaN(x)) ? '—' : Number(x).toFixed(0); };
    var fmt1 = function (x) { return (x == null || isNaN(x)) ? '—' : Number(x).toFixed(1); };
    var fmt2 = function (x) { return (x == null || isNaN(x)) ? '—' : Number(x).toFixed(2); };
    var fmt3 = function (x) { return (x == null || isNaN(x)) ? '—' : Number(x).toFixed(3); };
    var pct  = function (x) { return (x == null || isNaN(x)) ? '—' : (x * 100).toFixed(1) + '%'; };
    var esc  = function (s) {
        return String(s).replace(/[&<>]/g, function (c) {
            return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c];
        });
    };

    // ── Header block ──────────────────────────────────────────────
    function headerBlock(projectName, engineer, jobNo, sheetNo) {
        var today = new Date();
        var dateStr = today.toLocaleDateString('en-AU', { year: 'numeric', month: 'short', day: 'numeric' });
        return '<div class="cs-header">' +
            '<div>' +
                '<span class="cs-label">Project</span>' +
                '<span class="cs-value">' + esc(projectName || 'StructuralSketch') + '</span>' +
            '</div>' +
            '<div>' +
                '<span class="cs-label">Engineer</span>' +
                '<span class="cs-value">' + esc(engineer || '—') + '</span>' +
            '</div>' +
            '<div>' +
                '<span class="cs-label">Job no.</span>' +
                '<span class="cs-value">' + esc(jobNo || '—') + '</span>' +
            '</div>' +
            '<div>' +
                '<span class="cs-label">Date / Sheet</span>' +
                '<span class="cs-value">' + dateStr + ' · ' + esc(sheetNo || 'S-01') + '</span>' +
            '</div>' +
        '</div>';
    }

    // ── kv-row helper ─────────────────────────────────────────────
    function kvRows(rows) {
        var html = '<table class="cs-kv"><tbody>';
        for (var i = 0; i < rows.length; i++) {
            var r = rows[i];
            html += '<tr>' +
                '<td class="cs-k">' + r.k + '</td>' +
                '<td class="cs-v">' + r.v + '</td>' +
                '<td class="cs-clause">' + (r.clause || '') + '</td>' +
            '</tr>';
        }
        html += '</tbody></table>';
        return html;
    }

    // ── SVG diagram (M or V) ─────────────────────────────────────
    function svgDiagram(opts) {
        var x = opts.x, y = opts.y, title = opts.title, units = opts.units;
        var width  = opts.width  || 360;
        var height = opts.height || 120;
        var color  = opts.color  || '#2980b9';

        if (!x || !y || x.length !== y.length || x.length < 2) {
            return '<div class="cs-diagram-card"><div class="cs-diagram-title">' + title +
                   '</div><div style="font-size:8.5pt;color:#888;">No data</div></div>';
        }

        var xMin = x[0], xMax = x[x.length - 1];
        var yMax = 0, yMin = 0;
        for (var i = 0; i < y.length; i++) {
            if (y[i] > yMax) yMax = y[i];
            if (y[i] < yMin) yMin = y[i];
        }
        var yRange = Math.max(Math.abs(yMax), Math.abs(yMin), 1e-9);

        var padL = 38, padR = 10, padT = 14, padB = 22;
        var W = width, H = height;
        var plotW = W - padL - padR;
        var plotH = H - padT - padB;

        var xScale = function (xv) { return padL + ((xv - xMin) / (xMax - xMin)) * plotW; };
        var yScale = function (yv) {
            var zero_y = padT + plotH / 2;
            return zero_y - (yv / yRange) * (plotH / 2);
        };

        // Polyline points
        var curvePts = '';
        for (var i = 0; i < x.length; i++) {
            if (i > 0) curvePts += ' ';
            curvePts += xScale(x[i]).toFixed(1) + ',' + yScale(y[i]).toFixed(1);
        }
        var areaPts = xScale(xMin).toFixed(1) + ',' + yScale(0).toFixed(1) + ' ' +
                      curvePts + ' ' +
                      xScale(xMax).toFixed(1) + ',' + yScale(0).toFixed(1);

        // Peak marker
        var peakI = 0, peakAbs = 0;
        for (var i = 0; i < y.length; i++) {
            if (Math.abs(y[i]) > peakAbs) { peakAbs = Math.abs(y[i]); peakI = i; }
        }
        var peakX = xScale(x[peakI]);
        var peakY = yScale(y[peakI]);
        var peakLabel = y[peakI].toFixed(1) + ' ' + units;
        var zeroLineY = yScale(0);

        return '<div class="cs-diagram-card">' +
            '<div class="cs-diagram-title">' + title + '</div>' +
            '<svg viewBox="0 0 ' + W + ' ' + H + '" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg">' +
              '<line x1="' + padL + '" y1="' + padT + '" x2="' + padL + '" y2="' + (padT + plotH) + '" stroke="#aaa" stroke-width="0.8" />' +
              '<line x1="' + padL + '" y1="' + zeroLineY + '" x2="' + (W - padR) + '" y2="' + zeroLineY + '" stroke="#aaa" stroke-width="0.8" />' +
              '<text x="' + (padL - 4) + '" y="' + (padT + 4) + '" font-size="8" fill="#666" text-anchor="end">' + yMax.toFixed(0) + '</text>' +
              '<text x="' + (padL - 4) + '" y="' + (padT + plotH) + '" font-size="8" fill="#666" text-anchor="end">' + yMin.toFixed(0) + '</text>' +
              '<text x="' + padL + '" y="' + (H - 6) + '" font-size="8" fill="#666" text-anchor="start">0</text>' +
              '<text x="' + (W - padR) + '" y="' + (H - 6) + '" font-size="8" fill="#666" text-anchor="end">' + xMax.toFixed(2) + ' m</text>' +
              '<polygon points="' + areaPts + '" fill="' + color + '" fill-opacity="0.18" />' +
              '<polyline points="' + curvePts + '" fill="none" stroke="' + color + '" stroke-width="1.5" stroke-linejoin="round" />' +
              '<circle cx="' + peakX + '" cy="' + peakY + '" r="2.5" fill="' + color + '" />' +
              '<text x="' + peakX + '" y="' + (peakY - 5) + '" font-size="8.5" font-weight="600" fill="' + color + '" text-anchor="middle">' + peakLabel + '</text>' +
            '</svg>' +
        '</div>';
    }

    // ══════════════════════════════════════════════════════════════
    // BEAM CALC SHEET
    // ══════════════════════════════════════════════════════════════
    /**
     * Render a full A4 calc sheet for a steel beam.
     *
     * @param {Object} ctx
     * @param {Object} ctx.beamEl       — the beam element { x1, y1, x2, y2, typeRef, level, ... }
     * @param {Object} ctx.check        — result from runEnhancedBeamCheck()
     * @param {Object} ctx.analysis     — (optional) FloorBeamAnalysis.analyseSSBeam() output for diagrams
     * @param {Object} ctx.section      — getSectionProperties() result
     * @param {Object} [ctx.meta]       — { project, engineer, jobNo, sheetNo }
     * @returns {string} HTML
     */
    function renderBeamCalcSheet(ctx) {
        var beamEl  = ctx.beamEl;
        var check   = ctx.check;
        var analysis = ctx.analysis || null;
        var section = ctx.section;
        var meta    = ctx.meta || {};

        if (!beamEl || !check) {
            return '<div class="cs-sheet"><h1>No beam selected</h1></div>';
        }

        // ── Location info ──
        var dx = beamEl.x2 - beamEl.x1;
        var dy = beamEl.y2 - beamEl.y1;
        var span_mm = Math.sqrt(dx * dx + dy * dy);
        var span_m = span_mm / 1000;
        var startM = '(' + (beamEl.x1 / 1000).toFixed(2) + ', ' + (beamEl.y1 / 1000).toFixed(2) + ')';
        var endM   = '(' + (beamEl.x2 / 1000).toFixed(2) + ', ' + (beamEl.y2 / 1000).toFixed(2) + ')';

        // Beam angle → approximate direction
        var angle_deg = Math.abs(Math.atan2(dy, dx) * 180 / Math.PI);
        var dirText = (angle_deg < 45 || angle_deg > 135) ? 'X-direction' : 'Y-direction';

        // Section info
        var sizeStr = check.size || '—';
        var grade   = check.grade || '300';
        var fy      = check.fy || 300;
        var sw_kNm  = check.selfWeight || 0;
        var tribWidth_mm = check.tribWidth || 0;

        // ── Loading info ──
        var isLoaded = tribWidth_mm > 100; // more than 100mm trib = loaded beam
        var loadedText = isLoaded
            ? 'Loaded beam (perpendicular to joist span)'
            : 'Self-weight only (tie beam or no tributary area)';

        // Peak loads
        var peak_w_G_text = '—', peak_w_Q_text = '—';
        if (isLoaded && analysis) {
            peak_w_G_text = fmt2(analysis.peak_w_G_kNm) + ' kN/m (incl. ' + fmt2(sw_kNm) + ' SW)';
            peak_w_Q_text = fmt2(analysis.peak_w_Q_kNm) + ' kN/m';
        } else if (isLoaded) {
            peak_w_G_text = fmt2(sw_kNm) + ' kN/m (self-weight only — no zone data)';
        }

        // ── Diagrams ──
        var diagramsHtml = '';
        if (analysis && analysis.diagrams) {
            var d = analysis.diagrams;
            var comboLabel = check.governingCombo || 'governing';
            var M_curve = comboLabel === '1.35G' ? d.M_135G : d.M_12G15Q;
            var V_curve = comboLabel === '1.35G' ? d.V_135G : d.V_12G15Q;
            diagramsHtml =
                '<div class="cs-diagrams">' +
                svgDiagram({ x: d.x_grid_m, y: M_curve, title: 'Moment diagram — ' + comboLabel, units: 'kNm', color: '#c0392b' }) +
                svgDiagram({ x: d.x_grid_m, y: V_curve, title: 'Shear diagram — ' + comboLabel, units: 'kN', color: '#2980b9' }) +
                '</div>';
        } else {
            // Analytical parabola for self-weight-only
            var N = 21;
            var w = sw_kNm;
            var L = span_m;
            if (w > 0 && L > 0) {
                var xArr = [], Marr = [], Varr = [];
                for (var i = 0; i < N; i++) {
                    var xi = (i / (N - 1)) * L;
                    xArr.push(xi);
                    Marr.push((w * xi * (L - xi)) / 2);
                    Varr.push(w * (L / 2 - xi));
                }
                diagramsHtml =
                    '<div class="cs-diagrams">' +
                    svgDiagram({ x: xArr, y: Marr, title: 'Moment diagram — self-weight', units: 'kNm', color: '#c0392b' }) +
                    svgDiagram({ x: xArr, y: Varr, title: 'Shear diagram — self-weight', units: 'kN', color: '#2980b9' }) +
                    '</div>';
            }
        }

        // ── Restraint description ──
        var restraintText = check.assumptions
            ? check.assumptions.find(function (a) { return a.indexOf('LTB') >= 0 || a.indexOf('restraint') >= 0; }) || 'Full lateral restraint'
            : 'Full lateral restraint';

        // ── Level loads ──
        var levelId = beamEl.level || 'L1';
        var designSettings = (typeof getLevelDesignSettings === 'function')
            ? getLevelDesignSettings(levelId)
            : { G: 1.2, Q: 1.5 };

        // ── Design check rows ──
        var checks = [];

        // Flexure (member — LTB)
        if (check.phiMbx !== undefined) {
            checks.push({
                label: 'Flexure (member)',
                a: 'M* = ' + fmt1(check.Mstar) + ' kNm',
                b: 'φMbx = ' + fmt0(check.phiMbx) + ' kNm',
                util: check.bendingUtil
            });
        }
        // Flexure (section)
        checks.push({
            label: 'Flexure (section)',
            a: 'M* = ' + fmt1(check.Mstar) + ' kNm',
            b: 'φMsx = ' + fmt0(check.phiMsx) + ' kNm',
            util: check.phiMsx > 0 ? check.Mstar / check.phiMsx : 999
        });
        // Shear
        if (check.phiVv !== undefined) {
            checks.push({
                label: 'Shear',
                a: 'V* = ' + fmt1(check.Vstar) + ' kN',
                b: 'φVv = ' + fmt0(check.phiVv) + ' kN',
                util: check.shearUtil || 0
            });
        }
        // Deflection — total
        checks.push({
            label: 'Deflection — total (L/250)',
            a: 'δ = ' + fmt1(check.delta) + ' mm',
            b: 'limit = ' + fmt1(check.deltaLimit) + ' mm',
            util: check.deflectionUtil || 0
        });
        // Deflection — live
        if (check.deflectionLiveUtil !== undefined) {
            var liveLimitMm = span_mm / 300; // L/300 for live
            checks.push({
                label: 'Deflection — live (L/300)',
                a: 'δ_live = ' + fmt1((check.deflectionLiveUtil || 0) * liveLimitMm) + ' mm',
                b: 'limit = ' + fmt1(liveLimitMm) + ' mm',
                util: check.deflectionLiveUtil
            });
        }

        var checksHtml = '';
        for (var ci = 0; ci < checks.length; ci++) {
            var c = checks[ci];
            var pass = c.util <= 1.0;
            checksHtml += '<div class="cs-check-row ' + (pass ? 'cs-pass' : 'cs-fail') + '">' +
                '<span>' + c.label + '</span>' +
                '<span>' + c.a + '  ·  ' + c.b + '</span>' +
                '<span>' + pct(c.util) + '</span>' +
                '<span class="cs-pill">' + (pass ? 'PASS' : 'FAIL') + '</span>' +
            '</div>';
        }

        var overallPass = check.ok;
        var overallUtil = pct(check.maxUtil || check.bendingUtil);

        // ── Engine constants for footnote ──
        var E_MPa = 200000, G_MPa_val = 80000, phi_val = 0.9;
        if (typeof FloorCalcEngine !== 'undefined' && FloorCalcEngine.constants) {
            E_MPa = FloorCalcEngine.constants.E_MPa || E_MPa;
            G_MPa_val = FloorCalcEngine.constants.G_MPa || G_MPa_val;
            phi_val = FloorCalcEngine.constants.PHI || phi_val;
        }

        var sectionName = sizeStr;

        // ── Assembly ──
        return '<div class="cs-sheet">' +
            headerBlock(meta.project, meta.engineer, meta.jobNo, meta.sheetNo || sizeStr) +

            '<h1>Beam — ' + esc(sectionName) + ' Grade ' + esc(grade) + '</h1>' +
            '<p style="color:#666; font-size:9.5pt;">' +
                esc(dirText) + '  ·  ' + esc(loadedText) + '  ·  length ' + fmt2(span_m) + ' m' +
            '</p>' +

            '<h2>Geometry &amp; Location</h2>' +
            kvRows([
                { k: 'Section',          v: esc(sectionName) + '  (d=' + (section ? section.d : '—') + ', bf=' + (section ? section.bf : '—') + ', tf=' + (section ? section.tf : '—') + ', tw=' + (section ? section.tw : '—') + ' mm)' },
                { k: 'Grade / fy',       v: 'Grade ' + esc(grade) + ', fy = ' + fy + ' MPa', clause: 'AS/NZS 3679.1' },
                { k: 'Span (clear)',     v: fmt3(span_m) + ' m' },
                { k: 'Start → End',      v: startM + ' → ' + endM + ' (m)' },
                { k: 'Self-weight',      v: (section ? section.mass || section.m : '—') + ' kg/m = ' + fmt3(sw_kNm) + ' kN/m' },
                { k: 'Tributary width',  v: fmt2(tribWidth_mm / 1000) + ' m' },
                { k: 'Restraint',        v: restraintText, clause: 'Cl 5.4 / 5.6' },
            ]) +

            '<h2>Loading &amp; Combinations</h2>' +
            kvRows([
                { k: 'Background G / Q', v: fmt2(designSettings.G) + ' / ' + fmt2(designSettings.Q) + ' kPa', clause: 'AS 1170.1' },
                { k: 'Peak w_G on beam', v: peak_w_G_text },
                { k: 'Peak w_Q on beam', v: peak_w_Q_text },
                { k: 'Combinations',     v: '1.35G  and  1.2G + 1.5Q', clause: 'AS 1170.0 Cl 4.2.2' },
                { k: 'Governing combo',  v: esc(check.governingCombo || check.governing || '—') },
                { k: 'M*',              v: fmt1(check.Mstar) + ' kNm' },
                { k: 'V*',              v: fmt1(check.Vstar || 0) + ' kN' },
            ]) +

            diagramsHtml +

            '<h2>Section &amp; Member Capacity  (AS 4100)</h2>' +
            kvRows([
                { k: 'Section classification', v: esc(check.sectionClass || '—'), clause: 'Cl 5.2 / Table 5.2' },
                { k: 'Zex',                    v: (section && section._raw_Zx ? fmt0(section._raw_Zx) : '—') + ' × 10³ mm³' },
                { k: 'φMsx',                   v: fmt1(check.phiMsx) + ' kNm', clause: 'Cl 5.2' },
                { k: 'φMbx',                   v: fmt1(check.phiMbx || check.phiMsx) + ' kNm', clause: 'Cl 5.6.1.1' },
                { k: 'φVv',                    v: fmt1(check.phiVv || 0) + ' kN', clause: 'Cl 5.11.4' },
            ]) +

            '<h2>Serviceability  (AS 1170.0 App C)</h2>' +
            kvRows([
                { k: 'δ_total',                v: fmt2(check.delta) + ' mm' },
                { k: 'limit — total (L/250)',  v: fmt1(check.deltaLimit) + ' mm' },
            ]) +

            '<h2>Design Checks</h2>' +
            checksHtml +

            '<div class="cs-check-row ' + (overallPass ? 'cs-pass' : 'cs-fail') + '" style="margin-top:10pt; font-weight:700; font-size:10.5pt;">' +
                '<span>OVERALL</span>' +
                '<span>Governing: ' + esc(check.governing || '—') + '</span>' +
                '<span>' + overallUtil + '</span>' +
                '<span class="cs-pill">' + (overallPass ? 'PASS' : 'FAIL') + '</span>' +
            '</div>' +

            '<div class="cs-footnote">' +
                'AS 4100 check — single-span simply-supported beam, E = ' + E_MPa + ' MPa, G = ' + G_MPa_val + ' MPa, φ = ' + phi_val + '. ' +
                (analysis ? 'Analysis uses piecewise numerical integration (N = 200) over the beam\'s tributary strip with floor load zones applied. ' : 'Analysis uses uniform UDL from tributary floor area. ') +
                'Section data per AISC Design Capacity Tables Vol 1 (Red Book). Engineer to verify independently before construction.' +
            '</div>' +

            '<div class="cs-sheet-footer">' +
                '<span>StructuralSketch · ' + new Date().toLocaleDateString('en-AU') + '</span>' +
                '<span>' + esc(meta.sheetNo || sizeStr) + '</span>' +
            '</div>' +
        '</div>';
    }

    // ══════════════════════════════════════════════════════════════
    // MODAL CONTROLLER
    // ══════════════════════════════════════════════════════════════
    // Opens and manages the calc sheet overlay.
    // The modal HTML + CSS are injected by index.html (Slice 4).
    // This code just wires the show/hide/print behaviour.

    /**
     * Show the calc sheet modal for a given beam element.
     * Runs the enhanced check, builds the A4 HTML, and injects it.
     *
     * @param {Object} beamEl — the beam element
     */
    function showCalcSheetModal(beamEl) {
        if (!beamEl) return;

        var modal = document.getElementById('cs-modal');
        var wrap  = document.getElementById('cs-sheet-wrap');
        if (!modal || !wrap) {
            console.warn('[calc-sheet] Modal not found in DOM');
            return;
        }

        // Run the design check
        var check;
        if (typeof runEnhancedBeamCheck === 'function') {
            check = runEnhancedBeamCheck(beamEl);
        } else if (typeof runBeamDesignCheck === 'function') {
            check = runBeamDesignCheck(beamEl);
        } else {
            wrap.innerHTML = '<div class="cs-sheet"><h1>Design engine not loaded</h1></div>';
            modal.removeAttribute('hidden');
            return;
        }

        if (!check || check.errors && check.errors.length > 0) {
            wrap.innerHTML = '<div class="cs-sheet"><h1>Design check failed</h1><p>' +
                (check && check.errors ? check.errors.join('<br>') : 'Unknown error') + '</p></div>';
            modal.removeAttribute('hidden');
            return;
        }

        // Get section properties
        var section = null;
        if (check.size && typeof getSectionProperties === 'function') {
            section = getSectionProperties(check.size);
        }

        // Get analysis diagrams if available (piecewise path)
        var analysis = null;
        if (check._engine === 'FloorCalcEngine' && typeof FloorBeamAnalysis !== 'undefined') {
            // Re-run analysis to get full diagrams
            analysis = _buildAnalysisForCalcSheet(beamEl, check, section);
        }

        // Project metadata
        var meta = {};
        if (typeof project !== 'undefined' && project) {
            var info = project.projectInfo || {};
            meta.project = info.projectName || info.name || '';
            meta.engineer = info.engineer || '';
            meta.jobNo = info.jobNumber || info.jobNo || '';
        }
        meta.sheetNo = check.size || 'S-01';

        // Render
        wrap.innerHTML = renderBeamCalcSheet({
            beamEl: beamEl,
            check: check,
            analysis: analysis,
            section: section,
            meta: meta,
        });

        // Show modal
        modal.removeAttribute('hidden');
    }

    /**
     * Re-run the piecewise analysis to get diagram arrays.
     * (The enhanced check only stores peak values, not full M/V arrays.)
     */
    function _buildAnalysisForCalcSheet(beamEl, check, section) {
        try {
            var dx = beamEl.x2 - beamEl.x1;
            var dy = beamEl.y2 - beamEl.y1;
            var span_mm = Math.sqrt(dx * dx + dy * dy);
            var span_m = span_mm / 1000;
            if (span_m <= 0) return null;

            var Ix_mm4 = section ? section.Ix : 0;
            if (!Ix_mm4 || Ix_mm4 <= 0) return null;
            var EI_Nmm2 = 200000 * Ix_mm4;

            // Resolve floor zones
            var zones = [];
            if (typeof FloorLoadResolver !== 'undefined' && typeof project !== 'undefined') {
                var floorLoadSchedule = project.scheduleTypes.floorLoad || {};
                zones = FloorLoadResolver.resolveZonesFromElements(project.elements, floorLoadSchedule);
            }

            // Level design settings for defaults
            var levelId = beamEl.level || 'L1';
            var ds = (typeof getLevelDesignSettings === 'function')
                ? getLevelDesignSettings(levelId)
                : { G: 1.2, Q: 1.5 };

            var tribWidth_mm = check.tribWidth || 0;

            if (zones.length > 0 && typeof FloorLoadResolver.buildPiecewiseLoadFns === 'function') {
                // Piecewise path
                var loadFns = FloorLoadResolver.buildPiecewiseLoadFns(
                    beamEl, tribWidth_mm, zones, ds.G, ds.Q
                );
                var sw_kNm = (section ? (section.mass || section.m || 0) : 0) * 9.81 / 1000;
                var w_G_with_sw = function (x_m) {
                    return loadFns.w_G_of_x_kNm(x_m) + sw_kNm;
                };

                return FloorBeamAnalysis.analyseSSBeam({
                    span_m: span_m,
                    w_G_of_x_kNm: w_G_with_sw,
                    w_Q_of_x_kNm: loadFns.w_Q_of_x_kNm,
                    EI_Nmm2: EI_Nmm2,
                });
            } else {
                // Uniform UDL path — reconstruct from check data
                var tribW_m = tribWidth_mm / 1000;
                var wG = ds.G * tribW_m + ((section ? (section.mass || section.m || 0) : 0) * 9.81 / 1000);
                var wQ = ds.Q * tribW_m;

                return FloorBeamAnalysis.analyseSSBeam({
                    span_m: span_m,
                    w_G_of_x_kNm: function () { return wG; },
                    w_Q_of_x_kNm: function () { return wQ; },
                    EI_Nmm2: EI_Nmm2,
                });
            }
        } catch (e) {
            console.warn('[calc-sheet] Error building analysis for diagrams:', e);
            return null;
        }
    }

    // ══════════════════════════════════════════════════════════════
    // BAY CALC SHEET — hySPAN joist lookup (Slice 5)
    // ══════════════════════════════════════════════════════════════

    /**
     * Render a full A4 calc sheet for a joist bay.
     *
     * @param {Object} ctx
     * @param {Object} ctx.bay         — joistBay element
     * @param {Object} ctx.globals     — { spanDirection, G_kPa, Q_kPa, joistSpacing_mm, joistSpanType }
     * @param {Object} [ctx.meta]      — { project, engineer, jobNo, sheetNo }
     * @returns {string} HTML
     */
    function renderBayCalcSheet(ctx) {
        var bay = ctx.bay;
        var globals = ctx.globals || {};
        var meta = ctx.meta || {};

        if (!bay || !bay.joistResult) {
            return '<div class="cs-sheet"><h1>No bay selected</h1></div>';
        }

        var jr = bay.joistResult;
        var spacing = bay.spacing_mm || 450;
        var spanType = bay.spanType || 'single';

        var bayRect = '(' + (bay.x0 / 1000).toFixed(2) + ', ' + (bay.y0 / 1000).toFixed(2) + ')  →  (' +
                      (bay.x1 / 1000).toFixed(2) + ', ' + (bay.y1 / 1000).toFixed(2) + ')';
        var bayDims = ((bay.x1 - bay.x0) / 1000).toFixed(2) + ' m × ' + ((bay.y1 - bay.y0) / 1000).toFixed(2) + ' m';
        var joistDirText = (bay.spanDir === 'X') ? 'X (horizontal)' : 'Y (vertical)';
        var joistSpan_m = (bay.joistSpan_mm || 0) / 1000;

        // Load info
        var loadG = (bay.joistResult && bay.joistResult.loadCheck) ? globals.G_kPa : 1.2;
        var loadQ = (bay.joistResult && bay.joistResult.loadCheck) ? globals.Q_kPa : 1.5;

        // Build Dindas Table 3 rows
        var tableRows = '';
        if (typeof HyspanJoists !== 'undefined' && HyspanJoists.sections) {
            var secs = HyspanJoists.sections;
            for (var si = 0; si < secs.length; si++) {
                var sec = secs[si];
                var isPicked = jr.ok && jr.section && jr.section.name === sec.name && jr.section.product === sec.product;
                var maxSpan = sec.max[spanType][spacing] ? sec.max[spanType][spacing].span : 0;
                var passMark = maxSpan >= joistSpan_m ? '✓' : '·';
                tableRows += '<tr class="' + (isPicked ? 'cs-highlight' : '') + '">' +
                    '<td>' + sec.name + '</td>' +
                    '<td>' + sec.product + '</td>' +
                    '<td>' + sec.max.single[450].span.toFixed(1) + '</td>' +
                    '<td>' + sec.max.single[600].span.toFixed(1) + '</td>' +
                    '<td>' + sec.max.continuous[450].span.toFixed(1) + '</td>' +
                    '<td>' + sec.max.continuous[600].span.toFixed(1) + '</td>' +
                    '<td>' + passMark + '</td>' +
                '</tr>';
            }
        }

        // Result block
        var resultBlockHtml;
        if (jr.ok) {
            resultBlockHtml = '<div class="cs-check-row cs-pass">' +
                '<span>Joist selection</span>' +
                '<span>' + jr.section.name + ' ' + jr.section.product + ' @ ' + spacing + ' mm c/c ' + spanType + '</span>' +
                '<span>' + pct(jr.utilisation) + ' of ' + jr.maxSpan_m.toFixed(1) + ' m max</span>' +
                '<span class="cs-pill">OK</span>' +
            '</div>';
        } else {
            resultBlockHtml = '<div class="cs-check-row cs-fail">' +
                '<span>Joist selection</span>' +
                '<span>' + esc(jr.reason || 'No size available') + '</span>' +
                '<span>—</span>' +
                '<span class="cs-pill">FAIL</span>' +
            '</div>';
        }

        // Scope note
        var scopeNote;
        if (jr.loadCheck && jr.loadCheck.ok) {
            scopeNote = 'Within AS 1684 residential envelope (Q ≤ ' + (typeof HyspanJoists !== 'undefined' ? HyspanJoists.LIMIT_Q_KPA : 1.6) +
                        ' kPa, G_super ≤ ' + (typeof HyspanJoists !== 'undefined' ? HyspanJoists.LIMIT_G_SUP_KPA : 0.6) + ' kPa) — Dindas Table 3 applies.';
        } else {
            scopeNote = 'Outside the AS 1684 residential envelope — Table 3 does not apply. Engineered design required (AS 1720.1 LVL or alternative floor system).';
        }

        // Result details
        var resultDetailsHtml = '';
        if (jr.ok) {
            resultDetailsHtml = kvRows([
                { k: 'Selected section',    v: jr.section.name + ' ' + jr.section.product },
                { k: 'Tabulated max span',  v: jr.maxSpan_m.toFixed(1) + ' m', clause: spanType + ', ' + spacing + ' c/c' },
                { k: 'Required / max',      v: joistSpan_m.toFixed(3) + ' / ' + jr.maxSpan_m.toFixed(1) + ' m  =  ' + pct(jr.utilisation) },
                { k: 'Joist self-weight',   v: fmt2(jr.section.mass_kg_m) + ' kg/m  (≈ ' + (typeof HyspanJoists !== 'undefined' ? HyspanJoists.DENSITY_KG_M3 : 600) + ' kg/m³)' },
            ]);
        }

        return '<div class="cs-sheet">' +
            headerBlock(meta.project, meta.engineer, meta.jobNo, meta.sheetNo || ('FJ-' + (bay.typeRef || ''))) +

            '<h1>Bay — hySPAN Floor Joists</h1>' +
            '<p style="color:#666; font-size:9.5pt;">Per Dindas Span Tables (Oct 2012), Table 3 — Floor Joists supporting floor loads only.</p>' +

            '<h2>Bay Geometry</h2>' +
            kvRows([
                { k: 'Bay location (plan)',  v: bayRect },
                { k: 'Bay dimensions',       v: bayDims },
                { k: 'Joist direction',      v: joistDirText },
                { k: 'Joist span (clear)',   v: fmt3(joistSpan_m) + ' m' },
                { k: 'Joist spacing',        v: spacing + ' mm c/c' },
                { k: 'Span condition',       v: spanType === 'single' ? 'Single span' : 'Continuous (≥ 2 spans)' },
            ]) +

            '<h2>Loading  (worst case inside bay)</h2>' +
            kvRows([
                { k: 'Background G / Q',     v: fmt2(globals.G_kPa || 1.2) + ' / ' + fmt2(globals.Q_kPa || 1.5) + ' kPa' },
                { k: 'Scope check',          v: scopeNote, clause: 'AS 1684 / AS 1170.1' },
            ]) +

            '<h2>Dindas Table 3 Lookup</h2>' +
            '<p style="font-size:9pt; color:#555;">Walk the table shallowest-first. Pick the first row whose tabulated maximum span equals or exceeds the required ' + fmt3(joistSpan_m) + ' m.</p>' +
            '<table class="cs-dindas-table">' +
                '<thead>' +
                    '<tr>' +
                        '<th rowspan="2">Section<br>D × B (mm)</th>' +
                        '<th rowspan="2">Product</th>' +
                        '<th colspan="2">Single-span (m)</th>' +
                        '<th colspan="2">Continuous (m)</th>' +
                        '<th rowspan="2">≥ ' + fmt2(joistSpan_m) + ' m</th>' +
                    '</tr>' +
                    '<tr>' +
                        '<th>@ 450</th>' +
                        '<th>@ 600</th>' +
                        '<th>@ 450</th>' +
                        '<th>@ 600</th>' +
                    '</tr>' +
                '</thead>' +
                '<tbody>' + tableRows + '</tbody>' +
            '</table>' +

            '<h2>Result</h2>' +
            resultBlockHtml +
            resultDetailsHtml +

            '<div class="cs-footnote">' +
                'Dindas hySPAN Span Tables (Oct 2012), Table 3 — Floor Joists supporting floor loads only. ' +
                'Tabulated spans are maximum CLEAR distances between supports per AS 1684 residential framing, including strength and deflection limits. ' +
                'Minimum bearing = 30 mm end, 45 mm internal (continuous). ' +
                'Joists with D/B &gt; 4 require lateral restraint at supports (AS 1684.2 Cl 4.2.2.3). ' +
                'Scope is strictly AS 1684 residential loads (Q ≤ 1.5 kPa, G_super ≈ 0.4 kPa) — outside this envelope, an engineered LVL design is required.' +
            '</div>' +

            '<div class="cs-sheet-footer">' +
                '<span>StructuralSketch · ' + new Date().toLocaleDateString('en-AU') + '</span>' +
                '<span>' + esc(meta.sheetNo || ('FJ-' + (bay.typeRef || ''))) + '</span>' +
            '</div>' +
        '</div>';
    }

    /**
     * Show bay calc sheet modal for a joistBay element.
     */
    function showBayCalcSheetModal(bayEl) {
        if (!bayEl || bayEl.type !== 'joistBay') return;

        var modal = document.getElementById('cs-modal');
        var wrap  = document.getElementById('cs-sheet-wrap');
        if (!modal || !wrap) return;

        // Re-size the joist if result is stale
        if (!bayEl.joistResult && typeof HyspanJoists !== 'undefined') {
            var ds = (typeof getLevelDesignSettings === 'function')
                ? getLevelDesignSettings(bayEl.level || 'L1')
                : { G: 1.2, Q: 1.5 };
            bayEl.joistResult = HyspanJoists.sizeJoist({
                span_m: (bayEl.joistSpan_mm || 0) / 1000,
                spacing_mm: bayEl.spacing_mm || 450,
                spanType: bayEl.spanType || 'single',
                G_kPa: ds.G || 1.2,
                Q_kPa: ds.Q || 1.5,
            });
        }

        var ds = (typeof getLevelDesignSettings === 'function')
            ? getLevelDesignSettings(bayEl.level || 'L1')
            : { G: 1.2, Q: 1.5, spanDirection: 0 };

        var meta = {};
        if (typeof project !== 'undefined' && project) {
            var info = project.projectInfo || {};
            meta.project = info.projectName || info.name || '';
            meta.engineer = info.engineer || '';
            meta.jobNo = info.jobNumber || info.jobNo || '';
        }

        wrap.innerHTML = renderBayCalcSheet({
            bay: bayEl,
            globals: {
                spanDirection: bayEl.spanDir || 'X',
                G_kPa: ds.G || 1.2,
                Q_kPa: ds.Q || 1.5,
                joistSpacing_mm: bayEl.spacing_mm || 450,
                joistSpanType: bayEl.spanType || 'single',
            },
            meta: meta,
        });

        modal.removeAttribute('hidden');
    }

    /**
     * Hide the calc sheet modal.
     */
    function hideCalcSheetModal() {
        var modal = document.getElementById('cs-modal');
        if (modal) modal.setAttribute('hidden', '');
    }

    /**
     * Print the calc sheet (A4 PDF via browser print dialog).
     */
    function printCalcSheet() {
        window.print();
    }

    // ── Export ────────────────────────────────────────────────────
    window.FloorCalcSheet = {
        version: '1.1-slice5',
        renderBeamCalcSheet: renderBeamCalcSheet,
        renderBayCalcSheet:  renderBayCalcSheet,
        showCalcSheetModal:  showCalcSheetModal,
        showBayCalcSheetModal: showBayCalcSheetModal,
        hideCalcSheetModal:  hideCalcSheetModal,
        printCalcSheet:      printCalcSheet,
    };

    if (typeof console !== 'undefined') {
        console.log('[calc-sheet] 1.1-slice5 loaded');
    }
})();
