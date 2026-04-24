// ═══════════════════════════════════════════════════════════════════
// 31-floor-beam-analysis.js — Simply-supported beam, piecewise UDL
//
// Ported from prototype floor-beam-analysis.js.
// Pure JS, no DOM. Namespace: window.FloorBeamAnalysis
//
// When a floor has multiple load zones with different G/Q, a beam
// sees a VARYING UDL along its length. The closed-form wL²/8 no
// longer applies. This module does numerical integration of a SS
// beam under an arbitrary piecewise UDL and returns M*, V*, and
// mid-span deflection.
//
// Method: discretise beam into N equal segments, evaluate w(x) at
// each midpoint, compute reactions / shear / moment / deflection
// by running sums and virtual work (unit-load at midspan).
//
// Reproduces the closed-form 5wL⁴/(384EI) exactly for uniform UDL
// (verified in prototype test harness, 38 PASS / 0 FAIL).
//
// Units contract:
//   span_m         metres
//   w_G_of_x_kNm   function(x_m) → kN/m (INCLUDES beam self-weight)
//   w_Q_of_x_kNm   function(x_m) → kN/m
//   EI_Nmm2         N·mm² (= 200e3 MPa × Ix mm⁴)
//
// Outputs: kN, kNm, mm.
// ═══════════════════════════════════════════════════════════════════

(function () {
    'use strict';

    var DEFAULT_SEGMENTS = 200;

    /**
     * Analyse a simply-supported beam under a piecewise UDL.
     *
     * @param {object} opts
     * @param {number} opts.span_m
     * @param {function} opts.w_G_of_x_kNm — dead UDL (kN/m, INCL self-weight)
     * @param {function} opts.w_Q_of_x_kNm — live UDL (kN/m)
     * @param {number} opts.EI_Nmm2
     * @param {number} [opts.nSegments=200]
     * @param {number} [opts.psi_l_live=0.4]
     * @param {number} [opts.psi_s_live=0.7]
     * @returns {object}
     */
    function analyseSSBeam(opts) {
        var L = Number(opts.span_m);
        if (!(L > 0)) throw new Error('analyseSSBeam: span_m must be > 0');
        var N = opts.nSegments | 0 || DEFAULT_SEGMENTS;
        var EI = Number(opts.EI_Nmm2);
        if (!(EI > 0)) throw new Error('analyseSSBeam: EI_Nmm2 must be > 0');
        var psi_l = opts.psi_l_live != null ? Number(opts.psi_l_live) : 0.4;
        var psi_s = opts.psi_s_live != null ? Number(opts.psi_s_live) : 0.7;

        var dx = L / N;
        var x_mid = new Array(N);
        var w_G = new Array(N);
        var w_Q = new Array(N);
        var peak_w_G = 0, peak_w_Q = 0;

        var i;
        for (i = 0; i < N; i++) {
            x_mid[i] = (i + 0.5) * dx;
            w_G[i] = Number(opts.w_G_of_x_kNm(x_mid[i])) || 0;
            w_Q[i] = Number(opts.w_Q_of_x_kNm(x_mid[i])) || 0;
            if (w_G[i] > peak_w_G) peak_w_G = w_G[i];
            if (w_Q[i] > peak_w_Q) peak_w_Q = w_Q[i];
        }

        // Solve one load case → reactions, V(x), M(x), M_max, V_max
        function solveCase(w_i) {
            var R_A = 0;
            for (var ii = 0; ii < N; ii++) {
                R_A += w_i[ii] * (L - x_mid[ii]) / L * dx;
            }
            var V = new Array(N + 1);
            var M = new Array(N + 1);
            var x_grid = new Array(N + 1);

            V[0] = R_A;
            M[0] = 0;
            x_grid[0] = 0;

            var cum_w_dx = 0, cum_w_x_dx = 0;
            for (var k = 1; k <= N; k++) {
                cum_w_dx   += w_i[k - 1] * dx;
                cum_w_x_dx += w_i[k - 1] * x_mid[k - 1] * dx;
                var x_k = k * dx;
                x_grid[k] = x_k;
                V[k] = R_A - cum_w_dx;
                M[k] = R_A * x_k - x_k * cum_w_dx + cum_w_x_dx;
            }
            var M_max = 0, V_max = 0;
            for (var kk = 0; kk <= N; kk++) {
                if (M[kk] > M_max) M_max = M[kk];
                var va = Math.abs(V[kk]);
                if (va > V_max) V_max = va;
            }
            return { R_A: R_A, V: V, M: M, x_grid: x_grid, M_max: M_max, V_max: V_max };
        }

        // Factored load arrays
        var w_135G = new Array(N);
        var w_12G_15Q = new Array(N);
        for (i = 0; i < N; i++) {
            w_135G[i]    = 1.35 * w_G[i];
            w_12G_15Q[i] = 1.2  * w_G[i] + 1.5 * w_Q[i];
        }

        var caseG    = solveCase(w_G);
        var caseQ    = solveCase(w_Q);
        var case135  = solveCase(w_135G);
        var case125  = solveCase(w_12G_15Q);

        // Pick governing combination
        var M_star_kNm, V_star_kN, governing_combo;
        if (case125.M_max >= case135.M_max) {
            M_star_kNm = case125.M_max;
            governing_combo = '1.2G + 1.5Q';
        } else {
            M_star_kNm = case135.M_max;
            governing_combo = '1.35G';
        }
        V_star_kN = Math.max(case135.V_max, case125.V_max);

        // Mid-span deflection by virtual work
        function midspanDeflection(Mcase) {
            var Md = Mcase.M;
            var x = Mcase.x_grid;
            var L_half = L / 2;
            var integral_kNm3 = 0;
            for (var k = 0; k < N; k++) {
                var m0 = x[k]     <= L_half ? x[k]     / 2 : (L - x[k])     / 2;
                var m1 = x[k + 1] <= L_half ? x[k + 1] / 2 : (L - x[k + 1]) / 2;
                integral_kNm3 += 0.5 * (m0 * Md[k] + m1 * Md[k + 1]) * dx;
            }
            var EI_kNm2 = EI * 1e-9;
            return (integral_kNm3 / EI_kNm2) * 1000; // mm
        }

        var delta_G_mm = midspanDeflection(caseG);
        var delta_Q_mm = midspanDeflection(caseQ);

        return {
            method: 'piecewise_numerical',
            nSegments: N,
            span_m: L,
            governing_combo: governing_combo,
            M_star_kNm: M_star_kNm,
            V_star_kN: V_star_kN,
            delta_G_mm: delta_G_mm,
            delta_Q_mm: delta_Q_mm,
            delta_total_mm: delta_G_mm + psi_l * delta_Q_mm,
            delta_live_mm: psi_s * delta_Q_mm,
            peak_w_G_kNm: peak_w_G,
            peak_w_Q_kNm: peak_w_Q,
            diagrams: {
                x_grid_m: case125.x_grid,
                M_G: caseG.M, V_G: caseG.V,
                M_Q: caseQ.M, V_Q: caseQ.V,
                M_135G: case135.M, V_135G: case135.V,
                M_12G15Q: case125.M, V_12G15Q: case125.V,
            },
        };
    }

    window.FloorBeamAnalysis = {
        analyseSSBeam: analyseSSBeam,
        DEFAULT_SEGMENTS: DEFAULT_SEGMENTS,
    };

    if (typeof console !== 'undefined') {
        console.log('[floor-beam-analysis] piecewise solver loaded');
    }
})();
