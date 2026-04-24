// ═══════════════════════════════════════════════════════════════════
// 33-hyspan-joists.js — hySPAN joist lookup engine (Dindas Table 3)
//
// Ported from Floor Designer prototype hyspan-joists.js.
// Pure JS, no DOM, no framework dependencies.
//
// Namespace: window.HyspanJoists
//
// Source: Carter Holt Harvey / Dindas hySPAN Span Tables (Oct 2012)
//         Table 3 — Floor Joists supporting floor loads only
//         Per AS 1684 residential framing scope.
//
// Changes from prototype:
//   - Namespace: HyspanJoists (not FloorDesigner.hyspanJoists)
//   - Scope gate returns warning instead of hard-blocking
//   - Per-bay spacing support (450 or 600 per bay)
//
// Reference: AS 1684.2-2010, AS/NZS 1170.1:2002, Dindas Oct 2012.
// ═══════════════════════════════════════════════════════════════════

(function () {
    'use strict';

    // ── Constants ────────────────────────────────────────────────
    var DENSITY_KG_M3     = 600;   // approximate for mass display
    var LIMIT_Q_KPA       = 1.6;   // AS 1684 residential max live
    var LIMIT_G_SUP_KPA   = 0.6;   // AS 1684 residential max superimposed dead
    var VALID_SPACINGS    = [450, 600];
    var VALID_SPAN_TYPES  = ['single', 'continuous'];

    // ── Dindas Table 3 ──────────────────────────────────────────
    // Each row: [name, isPlus, ss450, ss600, cont450, cont600]
    // name: "DxB" in mm
    // isPlus: true = hySPAN+, false = hySPAN
    // ssNNN / contNNN: { span: max clear span (m), oh: max overhang (m) }
    // Sorted shallowest-first.
    var TABLE_3 = [
        ['150x45', false, { span: 2.6, oh: 0.7 }, { span: 2.2, oh: 0.6 }, { span: 3.2, oh: 0.8 }, { span: 2.8, oh: 0.7 }],
        ['170x45', false, { span: 3.0, oh: 0.8 }, { span: 2.6, oh: 0.7 }, { span: 3.7, oh: 0.9 }, { span: 3.2, oh: 0.8 }],
        ['200x45', false, { span: 3.6, oh: 0.9 }, { span: 3.2, oh: 0.8 }, { span: 4.3, oh: 1.1 }, { span: 3.8, oh: 1.0 }],
        ['240x45', true,  { span: 5.2, oh: 1.5 }, { span: 4.9, oh: 1.3 }, { span: 6.1, oh: 1.5 }, { span: 5.6, oh: 1.3 }],
        ['240x45', false, { span: 4.3, oh: 1.1 }, { span: 3.8, oh: 1.0 }, { span: 5.2, oh: 1.3 }, { span: 4.6, oh: 1.2 }],
        ['240x63', true,  { span: 5.6, oh: 1.5 }, { span: 5.2, oh: 1.5 }, { span: 6.4, oh: 1.5 }, { span: 5.9, oh: 1.5 }],
        ['240x63', false, { span: 4.7, oh: 1.2 }, { span: 4.2, oh: 1.1 }, { span: 5.6, oh: 1.4 }, { span: 5.0, oh: 1.3 }],
        ['290x45', true,  { span: 5.8, oh: 1.5 }, { span: 5.4, oh: 1.5 }, { span: 6.8, oh: 1.5 }, { span: 6.2, oh: 1.5 }],
        ['290x45', false, { span: 5.0, oh: 1.3 }, { span: 4.5, oh: 1.1 }, { span: 5.9, oh: 1.5 }, { span: 5.4, oh: 1.4 }],
        ['290x63', true,  { span: 6.2, oh: 1.5 }, { span: 5.8, oh: 1.5 }, { span: 7.2, oh: 1.5 }, { span: 6.6, oh: 1.5 }],
        ['290x63', false, { span: 5.5, oh: 1.4 }, { span: 5.0, oh: 1.3 }, { span: 6.4, oh: 1.5 }, { span: 5.9, oh: 1.5 }],
        ['300x45', true,  { span: 5.9, oh: 1.5 }, { span: 5.5, oh: 1.5 }, { span: 7.0, oh: 1.5 }, { span: 6.4, oh: 1.5 }],
        ['300x45', false, { span: 5.2, oh: 1.3 }, { span: 4.7, oh: 1.2 }, { span: 6.1, oh: 1.5 }, { span: 5.6, oh: 1.4 }],
        ['300x63', true,  { span: 6.4, oh: 1.5 }, { span: 5.9, oh: 1.5 }, { span: 7.4, oh: 1.5 }, { span: 6.8, oh: 1.5 }],
        ['300x63', false, { span: 5.7, oh: 1.4 }, { span: 5.2, oh: 1.3 }, { span: 6.6, oh: 1.5 }, { span: 6.1, oh: 1.5 }],
    ];

    // ── Build processed sections array ──────────────────────────
    var sections = [];
    for (var i = 0; i < TABLE_3.length; i++) {
        var row = TABLE_3[i];
        var name = row[0];
        var isPlus = row[1];
        var parts = name.split('x');
        var D_mm = parseInt(parts[0], 10);
        var B_mm = parseInt(parts[1], 10);
        var A_mm2 = D_mm * B_mm;
        var mass_kg_m = (A_mm2 / 1e6) * DENSITY_KG_M3;

        sections.push({
            name: name,
            product: isPlus ? 'hySPAN+' : 'hySPAN',
            D_mm: D_mm,
            B_mm: B_mm,
            A_mm2: A_mm2,
            mass_kg_m: mass_kg_m,
            max: {
                single:     { 450: row[2], 600: row[3] },
                continuous: { 450: row[4], 600: row[5] }
            }
        });
    }

    // ── Scope gate ──────────────────────────────────────────────
    /**
     * Check whether loads fall within AS 1684 residential envelope.
     * Returns a warning object — does NOT block placement.
     *
     * @param {number} G_kPa — superimposed dead load (excluding joist SW)
     * @param {number} Q_kPa — imposed (live) load
     * @returns {{ ok: boolean, reason: string|null, limits: { Q: number, G: number } }}
     */
    function checkLoadValidity(G_kPa, Q_kPa) {
        var result = { ok: true, reason: null, limits: { Q: LIMIT_Q_KPA, G: LIMIT_G_SUP_KPA } };
        if (Q_kPa > LIMIT_Q_KPA) {
            result.ok = false;
            result.reason = 'Q = ' + Q_kPa.toFixed(1) + ' kPa exceeds AS 1684 residential limit of ' + LIMIT_Q_KPA + ' kPa';
        } else if (G_kPa > LIMIT_G_SUP_KPA) {
            result.ok = false;
            result.reason = 'G_super = ' + G_kPa.toFixed(1) + ' kPa exceeds AS 1684 limit of ' + LIMIT_G_SUP_KPA + ' kPa';
        }
        return result;
    }

    // ── Joist sizer ─────────────────────────────────────────────
    /**
     * Size a hySPAN joist for a given bay.
     *
     * @param {Object} opts
     * @param {number} opts.span_m       — required clear span (m)
     * @param {number} opts.spacing_mm   — joist c/c spacing (450 or 600)
     * @param {string} opts.spanType     — 'single' or 'continuous'
     * @param {number} opts.G_kPa        — superimposed dead load
     * @param {number} opts.Q_kPa        — imposed live load
     * @param {boolean} [opts.allowPlus=true] — include hySPAN+ sizes
     * @returns {{ ok, section?, maxSpan_m, utilisation, spanType, spacing_mm, loadCheck, reason }}
     */
    function sizeJoist(opts) {
        var span_m = opts.span_m;
        var spacing_mm = opts.spacing_mm;
        var spanType = opts.spanType || 'single';
        var G_kPa = opts.G_kPa;
        var Q_kPa = opts.Q_kPa;
        var allowPlus = opts.allowPlus !== false;

        // Validate spacing
        if (VALID_SPACINGS.indexOf(spacing_mm) === -1) {
            return {
                ok: false,
                reason: 'Invalid spacing ' + spacing_mm + ' mm — must be 450 or 600',
                loadCheck: checkLoadValidity(G_kPa, Q_kPa),
                spanType: spanType,
                spacing_mm: spacing_mm
            };
        }

        // Validate span type
        if (VALID_SPAN_TYPES.indexOf(spanType) === -1) {
            return {
                ok: false,
                reason: 'Invalid span type "' + spanType + '" — must be single or continuous',
                loadCheck: checkLoadValidity(G_kPa, Q_kPa),
                spanType: spanType,
                spacing_mm: spacing_mm
            };
        }

        // Scope check (warning-only — proceed regardless)
        var loadCheck = checkLoadValidity(G_kPa, Q_kPa);

        // Walk sections shallowest-first
        for (var i = 0; i < sections.length; i++) {
            var sec = sections[i];
            if (!allowPlus && sec.product === 'hySPAN+') continue;

            var entry = sec.max[spanType][spacing_mm];
            if (!entry) continue;

            if (entry.span >= span_m) {
                return {
                    ok: true,
                    section: sec,
                    maxSpan_m: entry.span,
                    utilisation: span_m / entry.span,
                    spanType: spanType,
                    spacing_mm: spacing_mm,
                    loadCheck: loadCheck,
                    reason: null
                };
            }
        }

        // No section found
        var largest = sections[sections.length - 1];
        var largestSpan = largest.max[spanType][spacing_mm].span;
        return {
            ok: false,
            reason: 'Span ' + span_m.toFixed(2) + ' m exceeds max table span of ' + largestSpan.toFixed(1) + ' m (' + largest.name + ' @ ' + spacing_mm + ' c/c ' + spanType + ')',
            loadCheck: loadCheck,
            spanType: spanType,
            spacing_mm: spacing_mm
        };
    }

    // ── Lookup by name ──────────────────────────────────────────
    function byName(name) {
        for (var i = 0; i < sections.length; i++) {
            if (sections[i].name === name) return sections[i];
        }
        return null;
    }

    // ── Export ────────────────────────────────────────────────────
    window.HyspanJoists = {
        version: '1.0-slice5',
        sections: sections,
        sizeJoist: sizeJoist,
        checkLoadValidity: checkLoadValidity,
        byName: byName,
        LIMIT_Q_KPA: LIMIT_Q_KPA,
        LIMIT_G_SUP_KPA: LIMIT_G_SUP_KPA,
        DENSITY_KG_M3: DENSITY_KG_M3,
    };

    if (typeof console !== 'undefined') {
        console.log('[hyspan-joists] 1.0-slice5 loaded — ' + sections.length + ' sections');
    }
})();
