// ═══════════════════════════════════════════════════════════════════
// 30-floor-calc-engine.js — AS 4100 steel beam check, restraint-aware
//
// Ported from the Floor Designer prototype's floor-calc-engine.js.
// Pure JS, no DOM, no framework dependencies.
//
// Namespace: window.FloorCalcEngine
//
// Changes from prototype:
//   - Section input uses getSectionProperties() from 22-beam-design.js
//     (which now returns Red Book values from 18a-steel-sections.js)
//   - getFy falls back to existing getYieldStress() from 22-beam-design.js
//   - autoSize iterates STEEL_CATALOGUE.UB via getSectionProperties
//   - Keeps both the uniform-UDL path AND the piecewise-override path
//
// Reference: AS 4100-1998, AS/NZS 1170.0:2002, AS/NZS 3679.1.
// ═══════════════════════════════════════════════════════════════════

(function () {
    'use strict';

    // ── Material constants ────────────────────────────────────────
    var E_MPa = 200000;   // Young's modulus, steel
    var G_MPa = 80000;    // Shear modulus, steel
    var PHI   = 0.9;      // Capacity reduction factor, AS 4100 Table 3.4
    var g     = 9.81;

    // ── Helpers ───────────────────────────────────────────────────
    function parseDeflectionLimit(input, span_mm) {
        if (typeof input === 'number') return span_mm / input;
        if (typeof input === 'string') {
            var m = input.replace(/\s/g, '').match(/^L\/(\d+(?:\.\d+)?)$/i);
            if (m) return span_mm / parseFloat(m[1]);
            var asNum = parseFloat(input);
            if (!isNaN(asNum)) return span_mm / asNum;
        }
        return span_mm / 250;
    }

    function fmt(v, dp) {
        if (dp === undefined) dp = 1;
        if (typeof v !== 'number' || !isFinite(v)) return String(v);
        return Math.abs(v) >= 1000 ? v.toFixed(0) : v.toFixed(dp);
    }
    function fmt2(v) { return fmt(v, 2); }
    function fmt3(v) { return fmt(v, 3); }

    // ── Section property resolver ────────────────────────────────
    // Accepts either:
    //   (a) A prototype-style section object with stored-unit properties
    //       (Ix in 10⁶ mm⁴, Zx in 10³ mm³, etc.)
    //   (b) A getSectionProperties() result (base-unit Ix in mm⁴ etc. +
    //       _raw_* fields with the stored-unit originals)
    // Returns a normalised object with all values in base SI units.
    function normSection(s) {
        if (!s) return null;
        // If _raw_Ix exists, this came from getSectionProperties —
        // use the _raw_ stored-unit values (matching prototype's format)
        // for the calc-sheet text formatting, but compute with base units.
        var Ix_mm4, Iy_mm4, Zx_mm3, Sx_mm3, J_mm4, Iw_mm6, Ag, mass_kgm, d1;

        if (s._raw_Ix != null) {
            // From getSectionProperties — stored units available
            Ix_mm4 = s._raw_Ix * 1e6;
            Iy_mm4 = (s._raw_Iy || 0) * 1e6;
            Zx_mm3 = s._raw_Zx * 1e3;
            Sx_mm3 = s._raw_Sx * 1e3;
            J_mm4  = (s._raw_J  || 0) * 1e3;
            Iw_mm6 = (s._raw_Iw || 0) * 1e9;
            Ag     = s.Ag || 0;
            mass_kgm = s.mass || s.m || 0;
            d1     = s.d1 || (s.d - 2 * s.tf);
        } else if (s.Ix != null && s.Ix < 50000) {
            // Prototype-style stored units (Ix in 10⁶ mm⁴ ≈ single/double digits)
            Ix_mm4 = s.Ix * 1e6;
            Iy_mm4 = (s.Iy || 0) * 1e6;
            Zx_mm3 = (s.Zx || 0) * 1e3;
            Sx_mm3 = (s.Sx || 0) * 1e3;
            J_mm4  = (s.J  || 0) * 1e3;
            Iw_mm6 = (s.Iw || 0) * 1e9;
            Ag     = s.Ag || 0;
            mass_kgm = s.m || 0;
            d1     = s.d1 || (s.d - 2 * s.tf);
        } else {
            // Already in base units (e.g. from calcIBeamProps)
            Ix_mm4 = s.Ix || 0;
            Iy_mm4 = s.Iy || 0;
            Zx_mm3 = s.Zx || 0;
            Sx_mm3 = s.Sx || 0;
            J_mm4  = s.J_mm4 || 0;
            Iw_mm6 = s.Iw_mm6 || 0;
            Ag     = s.Ag || 0;
            mass_kgm = s.mass || s.m || 0;
            d1     = s.d1 || (s.d - 2 * s.tf);
        }

        return {
            name: s.name || (s.type ? (s.d + s.type + (s.m || '')) : '?'),
            d: s.d, bf: s.bf, tf: s.tf, tw: s.tw, r1: s.r1, d1: d1,
            Ag: Ag, m: mass_kgm,
            Ix_mm4: Ix_mm4, Iy_mm4: Iy_mm4,
            Zx_mm3: Zx_mm3, Sx_mm3: Sx_mm3,
            J_mm4: J_mm4, Iw_mm6: Iw_mm6,
        };
    }

    // ── Effective length from restraint type (AS 4100 Cl 5.4/5.6) ─
    function effectiveLengths(restraint, joistSpacing_mm, span_mm) {
        if (restraint === 'full') {
            return {
                Le_pos_mm: 0, Le_neg_mm: 0,
                kl_pos: 1.0, kl_neg: 1.0,
                rationale: 'Full restraint assumed — αs = 1.0, no LTB reduction.'
            };
        }
        if (restraint === 'joists_framed') {
            return {
                Le_pos_mm: joistSpacing_mm,
                Le_neg_mm: joistSpacing_mm,
                kl_pos: 1.0, kl_neg: 1.0,
                rationale:
                    'Joists framed into beam web → F-restraint at each joist (Cl 5.4.2.1). ' +
                    'Le = joist spacing for ±ve bending. Load near shear centre, kl = 1.0.'
            };
        }
        // joists_over
        return {
            Le_pos_mm: joistSpacing_mm,
            Le_neg_mm: span_mm,
            kl_pos: 1.4, kl_neg: 1.4,
            rationale:
                'Joists bearing on top flange → L-restraint for +ve bending only (Cl 5.4.2.2). ' +
                'Bottom flange unrestrained for −ve bending, Le = span. ' +
                'Gravity above shear centre, kl = 1.4 (Cl 5.6.3(b)).'
        };
    }

    // ── getFy — delegate to existing getYieldStress if available ──
    function getFy(grade, tf) {
        if (typeof getYieldStress === 'function') {
            return getYieldStress(grade, 'UB', tf);
        }
        // Inline fallback (AS/NZS 3679.1 Grade 300)
        if (grade === '300' || !grade) {
            if (tf <= 11) return 320;
            if (tf <= 17) return 300;
            return 280;
        }
        if (grade === '350') {
            if (tf <= 11) return 360;
            if (tf <= 17) return 340;
            return 330;
        }
        return 300;
    }

    // ══════════════════════════════════════════════════════════════
    // checkBeam — full AS 4100 design check
    // ══════════════════════════════════════════════════════════════
    function checkBeam(input) {
        var rawSection = input.section;
        if (!rawSection) throw new Error('FloorCalcEngine.checkBeam: input.section is required');

        var s = normSection(rawSection);
        var grade   = input.grade || '300';
        var span_m  = Number(input.span_m);
        var span_mm = span_m * 1000;
        var udl_G   = Number(input.udl_G_kNm) || 0;
        var udl_Q   = Number(input.udl_Q_kNm) || 0;
        var restraint = input.restraint || 'joists_framed';
        var joistSpacing_mm = Number(input.joistSpacing_mm) || 450;
        var psi_s = input.psi_s_live != null ? Number(input.psi_s_live) : 0.7;
        var psi_l = input.psi_l_live != null ? Number(input.psi_l_live) : 0.4;
        var am_override = Number(input.am_override) || 0;

        var steps = [];

        // ── Derived section quantities ────────────────────────────
        var fy = getFy(grade, s.tf);
        var d1 = s.d1;
        var bf_out = (s.bf - s.tw) / 2;
        var Aw = d1 * s.tw;
        var sw = (s.m * g) / 1000; // self-weight kN/m

        // ═══ 1. LOAD EFFECTS ──────────────────────────────────────
        var M_star, V_star, w_G_total, governingCombo;
        var ov = input.analysis_override;

        if (ov && typeof ov === 'object') {
            M_star = Number(ov.M_star_kNm);
            V_star = Number(ov.V_star_kN);
            governingCombo = ov.governing_combo || '—';
            w_G_total = Number(ov.peak_w_G_kNm) || 0;
            steps.push({
                title: 'Load Effects — piecewise',
                clause: 'AS/NZS 1170.0 Cl 4.2.2',
                text:
                    'Load analysis:   ' + (ov.method || 'piecewise numerical') + '\n' +
                    'Self-weight:     sw = ' + fmt2(sw) + ' kN/m (' + s.name + ', included in G)\n' +
                    'Peak dead UDL:   G  = ' + fmt2(w_G_total) + ' kN/m  (incl SW)\n' +
                    'Peak live UDL:   Q  = ' + fmt2(ov.peak_w_Q_kNm || 0) + ' kN/m\n\n' +
                    'M*  = ' + fmt(M_star) + ' kNm  [' + governingCombo + ']\n' +
                    'V*  = ' + fmt(V_star) + ' kN\n\n' +
                    '(Direct integration over the beam\'s varying load — see floor zones.)'
            });
        } else {
            w_G_total = udl_G + sw;
            var w_135G = 1.35 * w_G_total;
            var w_12G_15Q = 1.2 * w_G_total + 1.5 * udl_Q;
            var w_uls = Math.max(w_135G, w_12G_15Q);
            governingCombo = w_135G > w_12G_15Q ? '1.35G' : '1.2G + 1.5Q';
            M_star = (w_uls * span_m * span_m) / 8;
            V_star = (w_uls * span_m) / 2;
            steps.push({
                title: 'Load Effects',
                clause: 'AS/NZS 1170.0 Cl 4.2.2',
                text:
                    'Self-weight:    sw = ' + fmt2(sw) + ' kN/m (' + s.name + ')\n' +
                    'Dead UDL:       G  = ' + fmt2(udl_G) + ' + ' + fmt2(sw) + ' = ' + fmt2(w_G_total) + ' kN/m (incl SW)\n' +
                    'Live UDL:       Q  = ' + fmt2(udl_Q) + ' kN/m\n\n' +
                    'Combinations:\n' +
                    '  1.35G        = ' + fmt2(w_135G) + ' kN/m\n' +
                    '  1.2G + 1.5Q  = ' + fmt2(w_12G_15Q) + ' kN/m\n' +
                    '  Governing w* = ' + fmt2(w_uls) + ' kN/m  [' + governingCombo + ']\n\n' +
                    'M* = w*L²/8 = ' + fmt2(w_uls) + ' × ' + fmt2(span_m) + '² / 8 = ' + fmt(M_star) + ' kNm\n' +
                    'V* = w*L/2  = ' + fmt2(w_uls) + ' × ' + fmt2(span_m) + ' / 2  = ' + fmt(V_star) + ' kN'
            });
        }

        // ═══ 2. SECTION CLASSIFICATION (AS 4100 Table 5.2) ────────
        var fy_ratio = Math.sqrt(fy / 250);
        var lam_ef = (bf_out / s.tf) * fy_ratio;
        var lam_epf = 9, lam_eyf = 16;
        var flangeClass = lam_ef <= lam_epf ? 'Compact' : lam_ef <= lam_eyf ? 'Non-compact' : 'Slender';

        var lam_ew = (d1 / s.tw) * fy_ratio;
        var lam_epw = 82, lam_eyw = 115;
        var webClass = lam_ew <= lam_epw ? 'Compact' : lam_ew <= lam_eyw ? 'Non-compact' : 'Slender';

        var sectionClass;
        if (flangeClass === 'Slender' || webClass === 'Slender') sectionClass = 'Slender';
        else if (flangeClass === 'Non-compact' || webClass === 'Non-compact') sectionClass = 'Non-compact';
        else sectionClass = 'Compact';

        steps.push({
            title: 'Section Classification',
            clause: 'AS 4100 Table 5.2',
            text:
                'fy = ' + fy + ' MPa  (Grade ' + grade + ', tf = ' + s.tf + ' mm)\n' +
                'Flange outstand: λe = ' + fmt2(lam_ef) + '  (λep=' + lam_epf + ', λey=' + lam_eyf + ') → ' + flangeClass + '\n' +
                'Web (bending):   λe = ' + fmt2(lam_ew) + '  (λep=' + lam_epw + ', λey=' + lam_eyw + ') → ' + webClass + '\n' +
                'Section class: ' + sectionClass.toUpperCase()
        });

        // ═══ 3. SECTION MOMENT CAPACITY φMsx (Cl 5.2) ────────────
        var Zc = Math.min(s.Sx_mm3, 1.5 * s.Zx_mm3);
        var Ze;
        if (sectionClass === 'Compact') {
            Ze = Zc;
        } else if (sectionClass === 'Non-compact') {
            var Ze_f = s.Zx_mm3 + ((lam_eyf - lam_ef) / (lam_eyf - lam_epf)) * (Zc - s.Zx_mm3);
            var Ze_w = s.Zx_mm3 + ((lam_eyw - lam_ew) / (lam_eyw - lam_epw)) * (Zc - s.Zx_mm3);
            Ze = Math.min(Ze_f, Ze_w);
        } else {
            var ratioF = lam_eyf / lam_ef, ratioW = lam_eyw / lam_ew;
            Ze = s.Zx_mm3 * Math.pow(Math.min(ratioF, ratioW), 2);
        }
        var Msx = (Ze * fy) / 1e6;
        var phiMsx = PHI * Msx;

        steps.push({
            title: 'Section Moment Capacity φMsx',
            clause: 'AS 4100 Cl 5.2',
            text:
                'Zc = min(Sx, 1.5·Zx) = ' + fmt(Zc / 1e3) + ' × 10³ mm³\n' +
                'Ze = ' + fmt(Ze / 1e3) + ' × 10³ mm³  (' + sectionClass + ')\n' +
                'Msx = Ze × fy = ' + fmt(Msx) + ' kNm\n' +
                'φMsx = ' + PHI + ' × ' + fmt(Msx) + ' = ' + fmt(phiMsx) + ' kNm'
        });

        // ═══ 4. MEMBER MOMENT CAPACITY φMbx (Cl 5.6) ─────────────
        var Le = effectiveLengths(restraint, joistSpacing_mm, span_mm);
        var Le_pos_eff = Le.kl_pos * Le.Le_pos_mm;
        var am = am_override > 0 ? am_override : 1.13;

        var phiMbx;
        if (Le_pos_eff <= 0) {
            phiMbx = phiMsx;
            steps.push({
                title: 'Member Moment Capacity φMbx',
                clause: 'AS 4100 Cl 5.6.1.1',
                text:
                    'Restraint: ' + restraint + '\n' +
                    Le.rationale + '\n\n' +
                    'Le = 0 → αs = 1.0 → φMbx = φMsx = ' + fmt(phiMbx) + ' kNm'
            });
        } else {
            var term1 = (Math.PI * Math.PI * E_MPa * s.Iy_mm4) / (Le_pos_eff * Le_pos_eff);
            var term2 = G_MPa * s.J_mm4 + (Math.PI * Math.PI * E_MPa * s.Iw_mm6) / (Le_pos_eff * Le_pos_eff);
            var Moa = Math.sqrt(term1 * term2) / 1e6;
            var ratio = Msx / Moa;
            var alpha_s = 0.6 * (Math.sqrt(ratio * ratio + 3) - ratio);
            var Mbx = am * alpha_s * Msx;
            phiMbx = Math.min(PHI * Mbx, phiMsx);

            steps.push({
                title: 'Member Moment Capacity φMbx',
                clause: 'AS 4100 Cl 5.6.1.1',
                text:
                    'Restraint: ' + restraint + '\n' +
                    Le.rationale + '\n\n' +
                    'Le(+) = kl × segment = ' + fmt2(Le.kl_pos) + ' × ' + fmt(Le.Le_pos_mm) + ' = ' + fmt(Le_pos_eff) + ' mm\n' +
                    'αm = ' + fmt2(am) + '  (Table 5.6.1, SS beam UDL)\n\n' +
                    'Moa = √[(π²EIy/Le²)(GJ + π²EIw/Le²)]\n' +
                    '    = ' + fmt(Moa) + ' kNm\n' +
                    'αs  = 0.6·{√[(Msx/Moa)² + 3] − (Msx/Moa)} = ' + fmt3(alpha_s) + '\n' +
                    'Mbx = αm · αs · Msx = ' + fmt(Mbx) + ' kNm\n' +
                    'φMbx = min(φ·Mbx, φMsx) = ' + fmt(phiMbx) + ' kNm'
            });
        }

        // ═══ 5. SHEAR CAPACITY φVv (Cl 5.11) ─────────────────────
        var dp_tw = d1 / s.tw;
        var shearLimit = 82 / fy_ratio;
        var webStocky = dp_tw <= shearLimit;
        var Vw;
        if (webStocky) {
            Vw = (0.6 * fy * Aw) / 1000;
        } else {
            var alpha_v = Math.pow(82 / (dp_tw * fy_ratio), 2);
            Vw = (alpha_v * 0.6 * fy * Aw) / 1000;
        }
        var phiVv = PHI * Vw;

        steps.push({
            title: 'Shear Capacity φVv',
            clause: 'AS 4100 Cl 5.11',
            text:
                'd1/tw = ' + fmt2(dp_tw) + '  vs limit 82/√(fy/250) = ' + fmt2(shearLimit) + '\n' +
                'Web is ' + (webStocky ? 'stocky (yield governs)' : 'slender (shear buckling)') + '\n' +
                'Vw = ' + fmt(Vw) + ' kN\n' +
                'φVv = ' + fmt(phiVv) + ' kN'
        });

        // ═══ 6. DEFLECTION CHECK (AS/NZS 1170.0 App C) ───────────
        var delta_G, delta_Q;
        if (ov && typeof ov === 'object') {
            delta_G = Number(ov.delta_G_mm) || 0;
            delta_Q = Number(ov.delta_Q_mm) || 0;
        } else {
            var defl_mm = function (w_kNm) {
                return (5 * w_kNm * Math.pow(span_mm, 4)) / (384 * E_MPa * s.Ix_mm4);
            };
            delta_G = defl_mm(w_G_total);
            delta_Q = defl_mm(udl_Q);
        }
        var delta_total = delta_G + psi_l * delta_Q;
        var delta_live  = psi_s * delta_Q;

        var limit_total = parseDeflectionLimit(input.deflectionLimitTotal || 'L/250', span_mm);
        var limit_live  = parseDeflectionLimit(input.deflectionLimitLive  || 'L/360', span_mm);

        var util_defl_total = delta_total / limit_total;
        var util_defl_live  = delta_live / limit_live;

        steps.push({
            title: 'Serviceability — Deflection',
            clause: 'AS/NZS 1170.0 Cl 4.3 & App C',
            text:
                'δ(G incl SW) = ' + fmt2(delta_G) + ' mm' + (ov ? '  (piecewise)' : '  [5wL⁴/(384EI)]') + '\n' +
                'δ(Q)         = ' + fmt2(delta_Q) + ' mm\n' +
                'Total long-term: δ_G + ψl·δ_Q = ' + fmt2(delta_G) + ' + ' + fmt2(psi_l) + '·' + fmt2(delta_Q) + ' = ' + fmt2(delta_total) + ' mm\n' +
                '   Limit = L/' + fmt(span_mm / limit_total) + ' = ' + fmt2(limit_total) + ' mm\n' +
                '   Utilisation = ' + fmt3(util_defl_total) + '\n\n' +
                'Short-term live: ψs·δ_Q = ' + fmt2(psi_s) + '·' + fmt2(delta_Q) + ' = ' + fmt2(delta_live) + ' mm\n' +
                '   Limit = L/' + fmt(span_mm / limit_live) + ' = ' + fmt2(limit_live) + ' mm\n' +
                '   Utilisation = ' + fmt3(util_defl_live)
        });

        // ═══ 7. SUMMARY ──────────────────────────────────────────
        var util_moment = M_star / phiMbx;
        var util_shear  = V_star / phiVv;
        var utilisations = {
            'Flexure (member)':   util_moment,
            'Shear':              util_shear,
            'Deflection (total)': util_defl_total,
            'Deflection (live)':  util_defl_live,
        };
        var governing_key = 'Flexure (member)', max_util = util_moment;
        for (var k in utilisations) {
            if (utilisations.hasOwnProperty(k) && utilisations[k] > max_util) {
                max_util = utilisations[k]; governing_key = k;
            }
        }
        var pass = max_util <= 1.0;

        steps.push({
            title: 'Design Summary',
            clause: 'AS 4100 Cl 3.4',
            text:
                'Flexure:           M*/φMbx = ' + fmt3(util_moment) + '\n' +
                'Shear:             V*/φVv  = ' + fmt3(util_shear) + '\n' +
                'Deflection total:  ' + fmt3(util_defl_total) + '\n' +
                'Deflection live:   ' + fmt3(util_defl_live) + '\n\n' +
                'Governing: ' + governing_key + '\n' +
                'Max utilisation: ' + (max_util * 100).toFixed(1) + '%  → ' + (pass ? 'PASS' : 'FAIL')
        });

        return {
            inputs: {
                section: s, grade: grade, fy: fy,
                span_m: span_m,
                udl_G_kNm: udl_G, udl_Q_kNm: udl_Q,
                restraint: restraint, joistSpacing_mm: joistSpacing_mm,
                psi_s: psi_s, psi_l: psi_l,
            },
            M_star_kNm: M_star,
            V_star_kN: V_star,
            governing_combo: governingCombo,
            phiMsx_kNm: phiMsx,
            phiMbx_kNm: phiMbx,
            phiVv_kN: phiVv,
            Le_pos_mm: Le_pos_eff,
            kl_pos: Le.kl_pos,
            section_class: sectionClass,
            deflection_G_mm: delta_G,
            deflection_Q_mm: delta_Q,
            deflection_total_mm: delta_total,
            deflection_live_mm: delta_live,
            limit_total_mm: limit_total,
            limit_live_mm: limit_live,
            util_moment: util_moment,
            util_shear: util_shear,
            util_defl_total: util_defl_total,
            util_defl_live: util_defl_live,
            max_util: max_util,
            governing: governing_key,
            pass: pass,
            steps: steps,
        };
    }

    // ── Auto-size ────────────────────────────────────────────────
    // Iterate UB catalogue in ascending mass order, return the first
    // section whose max_util ≤ target. Returns null if nothing passes.
    function autoSize(baseInput, target) {
        if (target === undefined) target = 0.95;
        // Build sorted UB list from STEEL_CATALOGUE
        var sortedNames = [];
        if (typeof STEEL_CATALOGUE !== 'undefined' && STEEL_CATALOGUE.UB) {
            for (var name in STEEL_CATALOGUE.UB) {
                if (STEEL_CATALOGUE.UB.hasOwnProperty(name)) {
                    sortedNames.push({ name: name, m: STEEL_CATALOGUE.UB[name].m || 999 });
                }
            }
            sortedNames.sort(function (a, b) { return a.m - b.m; });
        }
        for (var i = 0; i < sortedNames.length; i++) {
            var props = getSectionProperties(sortedNames[i].name);
            if (!props) continue;
            try {
                var res = checkBeam(Object.assign({}, baseInput, { section: props }));
                if (res.pass && res.max_util <= target) {
                    return { section: props, result: res };
                }
            } catch (e) {
                // Skip sections that error (e.g. missing torsional props)
                continue;
            }
        }
        return null;
    }

    // ── Export ────────────────────────────────────────────────────
    window.FloorCalcEngine = {
        checkBeam: checkBeam,
        autoSize: autoSize,
        effectiveLengths: effectiveLengths,
        normSection: normSection,
        constants: { E_MPa: E_MPa, G_MPa: G_MPa, PHI: PHI },
    };

    if (typeof console !== 'undefined') {
        console.log('[floor-calc-engine] AS 4100 beam check loaded');
    }
})();
