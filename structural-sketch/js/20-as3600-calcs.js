// ============================================================
// 20-as3600-calcs.js — AS 3600-2018 Beam Design Engine
// ============================================================
// Pure calculation module — no UI, no DOM, no canvas.
// Provides: bending capacity, ductility check (Cl 8.1.6),
// shear capacity (unreinforced), and cracking moment.
// Designed for raft slab beams per AS 2870 Cl 4.4, but
// reusable for any RC beam design to AS 3600.
// All clause references are to AS 3600-2018 unless noted.
// ============================================================

const AS3600 = (() => {
    'use strict';

    // -----------------------------------------------------------------
    // MATERIAL PROPERTIES
    // -----------------------------------------------------------------

    /**
     * Concrete properties per AS 3600 Table 3.1.2.
     * fc in MPa, Ec in MPa, f'ct.f (flexural tensile) in MPa.
     */
    const CONCRETE_PROPS = {
        20:  { fc: 20,  Ec: 24000, fctf: 2.68, fct: 1.55, alpha2: 0.85, gamma: 0.85 },
        25:  { fc: 25,  Ec: 26700, fctf: 3.00, fct: 1.73, alpha2: 0.85, gamma: 0.826 },
        32:  { fc: 32,  Ec: 30100, fctf: 3.39, fct: 1.96, alpha2: 0.85, gamma: 0.796 },
        40:  { fc: 40,  Ec: 32800, fctf: 3.79, fct: 2.19, alpha2: 0.85, gamma: 0.766 },
        50:  { fc: 50,  Ec: 34800, fctf: 4.24, fct: 2.45, alpha2: 0.85, gamma: 0.736 },
        65:  { fc: 65,  Ec: 37400, fctf: 4.83, fct: 2.79, alpha2: 0.85, gamma: 0.696 },
    };

    /**
     * Get concrete properties. Accepts "N25" or 25 format.
     */
    function getConcreteProps(grade) {
        const fc = typeof grade === 'string' ? parseInt(grade.replace('N', '')) : grade;
        return CONCRETE_PROPS[fc] || CONCRETE_PROPS[25];
    }

    /** Reinforcement properties (AS/NZS 4671) */
    const REBAR = {
        fsy: 500,       // MPa — D500N grade
        Es: 200000,     // MPa
        // Standard bar areas (mm²)
        barArea: {
            'N10': 78.5,
            'N12': 113,
            'N16': 201,
            'N20': 314,
            'N24': 452,
            'N28': 616,
            'N32': 804,
            'N36': 1018,
        },
    };

    /** Common trench mesh areas (mm² total, per AS/NZS 4671) */
    const TRENCH_MESH = {
        '3-L8TM':   { Ast: 3 * 50.3,   bars: 3, dia: 8,  desc: '3 × L8' },
        '3-L11TM':  { Ast: 3 * 95.0,   bars: 3, dia: 11, desc: '3 × L11' },
        '3-L12TM':  { Ast: 3 * 113.1,  bars: 3, dia: 12, desc: '3 × L12' },
        '4-L11TM':  { Ast: 4 * 95.0,   bars: 4, dia: 11, desc: '4 × L11' },
        '4-L12TM':  { Ast: 4 * 113.1,  bars: 4, dia: 12, desc: '4 × L12' },
        '5-L12TM':  { Ast: 5 * 113.1,  bars: 5, dia: 12, desc: '5 × L12' },
        // Individual bar options (for deeper beams)
        '2-N12':    { Ast: 2 * 113,     bars: 2, dia: 12, desc: '2 × N12' },
        '3-N12':    { Ast: 3 * 113,     bars: 3, dia: 12, desc: '3 × N12' },
        '2-N16':    { Ast: 2 * 201,     bars: 2, dia: 16, desc: '2 × N16' },
        '3-N16':    { Ast: 3 * 201,     bars: 3, dia: 16, desc: '3 × N16' },
        '4-N16':    { Ast: 4 * 201,     bars: 4, dia: 16, desc: '4 × N16' },
        '2-N20':    { Ast: 2 * 314,     bars: 2, dia: 20, desc: '2 × N20' },
        '3-N20':    { Ast: 3 * 314,     bars: 3, dia: 20, desc: '3 × N20' },
        '4-N20':    { Ast: 4 * 314,     bars: 4, dia: 20, desc: '4 × N20' },
        '2-N24':    { Ast: 2 * 452,     bars: 2, dia: 24, desc: '2 × N24' },
        '3-N24':    { Ast: 3 * 452,     bars: 3, dia: 24, desc: '3 × N24' },
        '2-N28':    { Ast: 2 * 616,     bars: 2, dia: 28, desc: '2 × N28' },
        '3-N28':    { Ast: 3 * 616,     bars: 3, dia: 28, desc: '3 × N28' },
    };

    // -----------------------------------------------------------------
    // BENDING CAPACITY (Cl 8.1)
    // -----------------------------------------------------------------

    /**
     * Ultimate bending capacity of a singly-reinforced rectangular beam.
     * Uses rectangular stress block (Cl 8.1.3).
     *
     * @param {object} params
     * @param {number} params.b - width of compression zone (mm)
     * @param {number} params.d - effective depth to tension steel (mm)
     * @param {number} params.Ast - tension steel area (mm²)
     * @param {number} params.fc - characteristic concrete strength (MPa)
     * @param {number} params.fsy - steel yield strength (MPa), default 500
     * @returns {{ Mu, phi, phiMu, dn, ku, a, log }}
     */
    function bendingCapacity(params) {
        const { b, d, Ast, fc } = params;
        const fsy = params.fsy || 500;
        const conc = getConcreteProps(fc);
        const alpha2 = conc.alpha2;
        const gamma = conc.gamma;

        const log = [];

        // Depth of neutral axis from equilibrium: Cc = Ts
        // alpha2 × f'c × gamma × dn × b = Ast × fsy
        const dn = (Ast * fsy) / (alpha2 * fc * gamma * b);
        const ku = dn / d;
        const a = gamma * dn;  // stress block depth

        // Check steel yields (ku should be < 0.36 for ductile)
        const steelYields = ku < 0.6;

        // Mu = Ast × fsy × (d - a/2)
        const Mu = Ast * fsy * (d - a / 2); // N.mm
        const Mu_kNm = Mu / 1e6;

        // Capacity reduction factor (Table 2.2.2)
        // For bending: phi = 0.85 for ku ≤ 0.12, reduces linearly to 0.65 at ku = 0.36
        // For raft beams typically ku is low → phi ≈ 0.85
        let phi;
        if (ku <= 0.12) {
            phi = 0.85;
        } else if (ku <= 0.36) {
            phi = 0.85 - (ku - 0.12) * (0.85 - 0.65) / (0.36 - 0.12);
        } else {
            phi = 0.65;
        }

        const phiMu = phi * Mu;
        const phiMu_kNm = phiMu / 1e6;

        log.push('--- Bending Capacity (Cl 8.1) ---');
        log.push(`b = ${b} mm, d = ${d.toFixed(0)} mm, Ast = ${Ast.toFixed(0)} mm²`);
        log.push(`f'c = ${fc} MPa, fsy = ${fsy} MPa`);
        log.push(`α₂ = ${alpha2}, γ = ${gamma.toFixed(3)}`);
        log.push(`dn = Ast×fsy/(α₂×f'c×γ×b) = ${dn.toFixed(1)} mm`);
        log.push(`ku = dn/d = ${ku.toFixed(3)}`);
        log.push(`a = γ×dn = ${a.toFixed(1)} mm`);
        log.push(`Mu = Ast×fsy×(d - a/2) = ${Mu_kNm.toFixed(1)} kN.m`);
        log.push(`φ = ${phi.toFixed(2)} (Table 2.2.2, ku = ${ku.toFixed(3)})`);
        log.push(`φMu = ${phiMu_kNm.toFixed(1)} kN.m`);

        return {
            Mu: Mu,             // N.mm
            Mu_kNm: Mu_kNm,
            phi: phi,
            phiMu: phiMu,      // N.mm
            phiMu_kNm: phiMu_kNm,
            dn: dn,
            ku: ku,
            a: a,
            steelYields: steelYields,
            log: log,
        };
    }

    // -----------------------------------------------------------------
    // DUCTILITY CHECK — Cl 4.4.2 of AS 2870 / Cl 8.1.6 of AS 3600
    // -----------------------------------------------------------------

    /**
     * Cracking moment of a section.
     * Mcr = Z × f'ct.f  where f'ct.f = 0.6√f'c (Cl 3.1.1.3)
     *
     * @param {number} Z - section modulus to tension face (mm³)
     * @param {number} fc - concrete strength (MPa)
     * @returns {{ Mcr, Mcr_kNm, fctf, log }}
     */
    function crackingMoment(Z, fc) {
        const fctf = 0.6 * Math.sqrt(fc);
        const Mcr = Z * fctf;  // N.mm
        const Mcr_kNm = Mcr / 1e6;

        return {
            Mcr: Mcr,
            Mcr_kNm: Mcr_kNm,
            fctf: fctf,
            log: [
                `f'ct.f = 0.6√${fc} = ${fctf.toFixed(2)} MPa (Cl 3.1.1.3)`,
                `Mcr = Z × f'ct.f = ${(Z/1e3).toFixed(0)}×10³ × ${fctf.toFixed(2)} = ${Mcr_kNm.toFixed(1)} kN.m`,
            ],
        };
    }

    /**
     * AS 2870 Cl 4.4.2 ductility check: Mu ≥ 1.2 × Mcr
     * This is a HARD requirement — non-negotiable.
     *
     * @param {number} Mu - ultimate bending capacity (N.mm)
     * @param {number} Mcr - cracking moment (N.mm)
     * @returns {{ pass, ratio, Mu_kNm, Mcr_x_1_2_kNm, log }}
     */
    function ductilityCheck(Mu, Mcr) {
        const required = 1.2 * Mcr;
        const pass = Mu >= required - 1;  // 1 N.mm tolerance
        const ratio = Mu / Mcr;

        return {
            pass: pass,
            ratio: ratio,
            Mu_kNm: Mu / 1e6,
            Mcr_kNm: Mcr / 1e6,
            required_kNm: required / 1e6,
            log: [
                '--- Ductility Check (AS 2870 Cl 4.4.2) ---',
                `Mu = ${(Mu / 1e6).toFixed(1)} kN.m`,
                `1.2 × Mcr = 1.2 × ${(Mcr / 1e6).toFixed(1)} = ${(required / 1e6).toFixed(1)} kN.m`,
                `Mu / Mcr = ${ratio.toFixed(2)}`,
                `${(Mu / 1e6).toFixed(1)} ${pass ? '≥' : '<'} ${(required / 1e6).toFixed(1)} → ${pass ? 'PASS ✓' : 'FAIL ✗'}`,
                pass ? '' : '⚠ DUCTILITY REQUIREMENT NOT MET — increase reinforcement or reduce section',
            ].filter(Boolean),
        };
    }

    // -----------------------------------------------------------------
    // SHEAR CAPACITY (Cl 8.2) — Unreinforced (Vuc only)
    // -----------------------------------------------------------------

    /**
     * Shear capacity of concrete without shear reinforcement (Cl 8.2.4.1).
     * Per AS 2870 Cl 4.4.3: raft beams rely on Vuc alone — no fitments.
     * If φVuc < V*, increase beam section.
     *
     * Simplified formula (AS 3600 Cl 8.2.4.1):
     *   Vuc = β₁ × β₂ × β₃ × bv × do × fcv
     * where:
     *   β₁ = 1.1 × (1.6 - do/1000) ≥ 1.1   [size effect]
     *   β₂ = 1.0 (no axial force in raft beams)
     *   β₃ = 1.0 (standard support conditions)
     *   fcv = f'c^(1/3)  ≤ 4.0 MPa
     *   bv = effective web width (mm)
     *   do = effective depth (mm)
     *
     * @param {object} params
     * @param {number} params.bv - effective shear width (mm), typically beam width
     * @param {number} params.do_ - effective depth (mm)
     * @param {number} params.Ast - tension steel area (mm²)
     * @param {number} params.fc - concrete strength (MPa)
     * @returns {{ Vuc, phiVuc, Vuc_kN, phiVuc_kN, log }}
     */
    function shearCapacityVuc(params) {
        const { bv, do_, Ast, fc } = params;
        const log = [];

        // β₁ = 1.1(1.6 - do/1000) ≥ 1.1
        const beta1 = Math.max(1.1, 1.1 * (1.6 - do_ / 1000));
        const beta2 = 1.0;
        const beta3 = 1.0;

        // fcv = f'c^(1/3) ≤ 4.0
        const fcv = Math.min(Math.pow(fc, 1 / 3), 4.0);

        // Ast/bv.do term (reinforcement ratio effect)
        const rho = Ast / (bv * do_);

        // Vuc = β₁ × β₂ × β₃ × bv × do × fcv × (Ast/(bv.do))^(1/3)
        // But the full Cl 8.2.4.1 formula is more nuanced. Using the simplified approach:
        // Vuc = [β₁ × β₂ × β₃ × bv × do] × [fcv] × [max(Ast/(bv×do), 0.01)]^(1/3)
        // Per exact AS 3600 Cl 8.2.4.1:
        const rhoTerm = Math.pow(Math.max(rho, 0.001), 1 / 3);
        const Vuc = beta1 * beta2 * beta3 * bv * do_ * fcv * rhoTerm; // N

        // Simplified form commonly used: Vuc = 0.17 × √f'c × bv × do (ACI-like lower bound)
        // But we'll use the AS 3600 form above.

        const phi = 0.7; // Table 2.2.2 for shear
        const phiVuc = phi * Vuc;

        log.push('--- Shear Capacity (Cl 8.2.4.1) — Vuc only ---');
        log.push(`bv = ${bv} mm, do = ${do_.toFixed(0)} mm`);
        log.push(`Ast = ${Ast.toFixed(0)} mm², ρ = Ast/(bv×do) = ${(rho * 100).toFixed(3)}%`);
        log.push(`β₁ = 1.1(1.6 - ${do_}/1000) = ${beta1.toFixed(3)}`);
        log.push(`fcv = ${fc}^(1/3) = ${fcv.toFixed(2)} MPa`);
        log.push(`Vuc = ${beta1.toFixed(3)} × ${bv} × ${do_.toFixed(0)} × ${fcv.toFixed(2)} × ${rhoTerm.toFixed(3)}`);
        log.push(`    = ${(Vuc / 1000).toFixed(1)} kN`);
        log.push(`φ = ${phi} (Table 2.2.2)`);
        log.push(`φVuc = ${(phiVuc / 1000).toFixed(1)} kN`);
        log.push('');
        log.push('NOTE: Per AS 2870 Cl 4.4.3, raft beams use Vuc only — no shear reinforcement.');
        log.push('If φVuc < V*, increase beam width or depth.');

        return {
            Vuc: Vuc,               // N
            Vuc_kN: Vuc / 1000,
            phiVuc: phiVuc,          // N
            phiVuc_kN: phiVuc / 1000,
            phi: phi,
            beta1: beta1,
            fcv: fcv,
            rho: rho,
            log: log,
        };
    }

    // -----------------------------------------------------------------
    // FULL BEAM DESIGN CHECK
    // -----------------------------------------------------------------

    /**
     * Complete beam design check for a raft beam per AS 2870 Cl 4.4.
     * Checks bending capacity, ductility, and shear for BOTH faces
     * (top and bottom reinforcement for both mound shapes).
     *
     * @param {object} params
     * @param {number} params.beamWidth - mm
     * @param {number} params.beamDepth - total depth from top of slab (mm)
     * @param {number} params.slabThickness - mm
     * @param {number} params.fc - concrete strength (MPa) or grade string "N25"
     * @param {string} params.topReo - key into TRENCH_MESH, e.g. '3-L12TM'
     * @param {string} params.botReo - key into TRENCH_MESH
     * @param {number} params.topCover - cover to top steel (mm), default 30
     * @param {number} params.botCover - cover to bottom steel (mm), default 40
     * @param {number} params.Mstar_hog - design hogging moment (kN.m), tension at top
     * @param {number} params.Mstar_sag - design sagging moment (kN.m), tension at bottom
     * @param {number} params.Vstar - design shear force (kN)
     * @param {number} params.flangeWidth - effective flange width (mm), for Mcr calculation
     * @returns {object} comprehensive results
     */
    function raftBeamDesign(params) {
        const fc = typeof params.fc === 'string' ? parseInt(params.fc.replace('N', '')) : params.fc;
        const conc = getConcreteProps(fc);
        const bw = params.beamWidth;
        const D = params.beamDepth;
        const tf = params.slabThickness;
        const topCover = params.topCover || 30;
        const botCover = params.botCover || 40;

        const topMesh = TRENCH_MESH[params.topReo];
        const botMesh = TRENCH_MESH[params.botReo];
        if (!topMesh || !botMesh) {
            return { error: `Unknown reinforcement: top=${params.topReo}, bot=${params.botReo}` };
        }

        const topDia = topMesh.dia;
        const botDia = botMesh.dia;

        // Effective depths
        const d_top = D - topCover - topDia / 2;  // for hogging (tension at top, compression at bottom)
        const d_bot = D - botCover - botDia / 2;  // for sagging (tension at bottom, compression at top)

        const log = [];
        log.push('=== RAFT BEAM DESIGN — AS 3600 / AS 2870 Cl 4.4 ===');
        log.push(`Section: ${bw}W × ${D}D, slab ${tf}mm, f'c = ${fc} MPa`);
        log.push(`Top reo: ${params.topReo} (Ast = ${topMesh.Ast.toFixed(0)} mm²), cover = ${topCover} mm`);
        log.push(`Bot reo: ${params.botReo} (Ast = ${botMesh.Ast.toFixed(0)} mm²), cover = ${botCover} mm`);
        log.push(`d_top = ${d_top.toFixed(0)} mm, d_bot = ${d_bot.toFixed(0)} mm`);
        log.push('');

        const results = {};

        // === SAGGING (tension at bottom — centre heave edge, edge heave centre) ===
        if (params.Mstar_sag !== undefined && params.Mstar_sag > 0) {
            const bendSag = bendingCapacity({
                b: params.flangeWidth || bw,  // T-beam: use flange in compression
                d: d_bot,
                Ast: botMesh.Ast,
                fc: fc,
            });
            log.push('--- SAGGING CHECK (tension at bottom) ---');
            log.push(...bendSag.log);
            log.push(`M* = ${params.Mstar_sag.toFixed(1)} kN.m`);
            const sagPass = bendSag.phiMu_kNm >= params.Mstar_sag - 0.1;
            log.push(`φMu = ${bendSag.phiMu_kNm.toFixed(1)} kN.m ${sagPass ? '≥' : '<'} ${params.Mstar_sag.toFixed(1)} kN.m → ${sagPass ? 'PASS ✓' : 'FAIL ✗'}`);
            log.push('');
            results.sagging = { ...bendSag, Mstar: params.Mstar_sag, pass: sagPass };
        }

        // === HOGGING (tension at top — edge heave edge, centre heave centre) ===
        if (params.Mstar_hog !== undefined && params.Mstar_hog > 0) {
            const bendHog = bendingCapacity({
                b: bw,  // Rectangular: flange in tension, web in compression
                d: d_top,
                Ast: topMesh.Ast,
                fc: fc,
            });
            log.push('--- HOGGING CHECK (tension at top) ---');
            log.push(...bendHog.log);
            log.push(`M* = ${params.Mstar_hog.toFixed(1)} kN.m`);
            const hogPass = bendHog.phiMu_kNm >= params.Mstar_hog - 0.1;
            log.push(`φMu = ${bendHog.phiMu_kNm.toFixed(1)} kN.m ${hogPass ? '≥' : '<'} ${params.Mstar_hog.toFixed(1)} kN.m → ${hogPass ? 'PASS ✓' : 'FAIL ✗'}`);
            log.push('');
            results.hogging = { ...bendHog, Mstar: params.Mstar_hog, pass: hogPass };
        }

        // === DUCTILITY CHECK (AS 2870 Cl 4.4.2) ===
        // Mcr based on gross section (uncracked) — use the LARGER Z (tension face)
        const flangeW = params.flangeWidth || bw;
        const section = AS2870.tSectionProperties({
            beamWidth: bw,
            beamDepth: D,
            slabThickness: tf,
            flangeWidth: flangeW,
        });

        // For sagging: tension at bottom → use Zb
        // For hogging: tension at top → use Zt
        const Mcr_sag = crackingMoment(section.Zb, fc);
        const Mcr_hog = crackingMoment(section.Zt, fc);

        log.push('--- DUCTILITY — Mu ≥ 1.2Mcr (AS 2870 Cl 4.4.2) ---');
        log.push(`Gross section: I = ${(section.I / 1e6).toFixed(1)}×10⁶ mm⁴, ȳ = ${section.yBar.toFixed(1)} mm`);
        log.push(`Zt = ${(section.Zt / 1e3).toFixed(0)}×10³ mm³, Zb = ${(section.Zb / 1e3).toFixed(0)}×10³ mm³`);

        if (results.sagging) {
            const ductSag = ductilityCheck(results.sagging.Mu, Mcr_sag.Mcr);
            log.push('Sagging: ' + Mcr_sag.log.join(', '));
            log.push(...ductSag.log);
            results.ductility_sag = ductSag;
        }
        if (results.hogging) {
            const ductHog = ductilityCheck(results.hogging.Mu, Mcr_hog.Mcr);
            log.push('Hogging: ' + Mcr_hog.log.join(', '));
            log.push(...ductHog.log);
            results.ductility_hog = ductHog;
        }
        log.push('');

        // === SHEAR CHECK (AS 2870 Cl 4.4.3) — Vuc only ===
        if (params.Vstar !== undefined && params.Vstar > 0) {
            // Use the larger Ast (governs for Vuc — more conservative with less steel)
            const Ast_shear = Math.min(topMesh.Ast, botMesh.Ast);
            const do_shear = Math.min(d_top, d_bot);

            const shear = shearCapacityVuc({
                bv: bw,
                do_: do_shear,
                Ast: Ast_shear,
                fc: fc,
            });
            log.push(`V* = ${params.Vstar.toFixed(1)} kN`);
            const shearPass = shear.phiVuc_kN >= params.Vstar - 0.1;
            log.push(`φVuc = ${shear.phiVuc_kN.toFixed(1)} kN ${shearPass ? '≥' : '<'} ${params.Vstar.toFixed(1)} kN → ${shearPass ? 'PASS ✓' : 'FAIL ✗'}`);
            if (!shearPass) {
                log.push('⚠ Increase beam width or depth — do NOT add shear reinforcement (AS 2870 Cl 4.4.3)');
            }
            log.push('');
            results.shear = { ...shear, Vstar: params.Vstar, pass: shearPass };
        }

        // === OVERALL ===
        const allChecks = [];
        if (results.sagging) allChecks.push(results.sagging.pass);
        if (results.hogging) allChecks.push(results.hogging.pass);
        if (results.ductility_sag) allChecks.push(results.ductility_sag.pass);
        if (results.ductility_hog) allChecks.push(results.ductility_hog.pass);
        if (results.shear) allChecks.push(results.shear.pass);

        results.overallPass = allChecks.length > 0 && allChecks.every(Boolean);
        results.section = section;
        results.log = log;

        return results;
    }

    // -----------------------------------------------------------------
    // DESIGN MOMENT ESTIMATION (simplified, for preliminary design)
    // -----------------------------------------------------------------

    /**
     * Estimate design moments from the simplified method.
     * For a continuous raft beam on a mounding soil, the bending moments
     * are a function of the differential movement (Δ) and raft stiffness.
     *
     * Approximate moments (per Appendix G principles):
     *   M ≈ (3 × EI × Δ) / L²
     * for both hog and sag, distributed along the raft.
     *
     * @param {number} EI - beam stiffness (N.mm²) for one beam
     * @param {number} delta_mm - differential movement (mm)
     * @param {number} L_mm - raft dimension (mm)
     * @returns {{ Mstar_kNm: number }}
     */
    function estimateDesignMoment(EI, delta_mm, L_mm) {
        // M* ≈ 3 × EI × Δ / L²  (simplified from beam-on-mound theory)
        const Mstar = (3 * EI * delta_mm) / (L_mm * L_mm);  // N.mm
        return {
            Mstar: Mstar,
            Mstar_kNm: Mstar / 1e6,
        };
    }

    /**
     * Estimate design shear from the simplified method.
     * V* ≈ 2 × M* / L  (approximate for uniformly loaded continuous beam)
     *
     * @param {number} Mstar - design moment (N.mm)
     * @param {number} L_mm - raft dimension (mm)
     * @returns {{ Vstar: number, Vstar_kN: number }}
     */
    function estimateDesignShear(Mstar, L_mm) {
        const Vstar = (2 * Mstar) / L_mm;  // N
        return {
            Vstar: Vstar,
            Vstar_kN: Vstar / 1000,
        };
    }

    // -----------------------------------------------------------------
    // REINFORCEMENT SELECTOR (auto-suggest)
    // -----------------------------------------------------------------

    /**
     * Suggest minimum reinforcement for a raft beam to satisfy ductility.
     * Works backwards from Mu ≥ 1.2 × Mcr to find required Ast.
     *
     * @param {object} params
     * @param {number} params.beamWidth - mm
     * @param {number} params.beamDepth - mm
     * @param {number} params.slabThickness - mm
     * @param {number} params.fc - MPa
     * @param {number} params.cover - cover to steel (mm)
     * @param {number} params.flangeWidth - effective flange width (mm)
     * @param {'top'|'bottom'} params.face - which face
     * @returns {{ requiredAst, suggestedReo, log }}
     */
    function suggestReinforcement(params) {
        const fc = typeof params.fc === 'string' ? parseInt(params.fc.replace('N', '')) : params.fc;
        const D = params.beamDepth;
        const tf = params.slabThickness;
        const bw = params.beamWidth;
        const cover = params.cover || (params.face === 'top' ? 30 : 40);
        const flangeW = params.flangeWidth || bw;

        // Section properties for Mcr
        const section = AS2870.tSectionProperties({
            beamWidth: bw, beamDepth: D, slabThickness: tf, flangeWidth: flangeW,
        });

        const Z = params.face === 'top' ? section.Zt : section.Zb;
        const fctf = 0.6 * Math.sqrt(fc);
        const Mcr = Z * fctf;
        const Mu_required = 1.2 * Mcr;

        // Back-solve for Ast from Mu = Ast × fsy × (d - a/2)
        // Approximate: assume a is small, so d - a/2 ≈ d - (Ast×fsy)/(2×alpha2×fc×b)
        // Iterative solution:
        const fsy = 500;
        const b = params.face === 'bottom' ? flangeW : bw; // compression zone width
        const conc = getConcreteProps(fc);
        let Ast = Mu_required / (fsy * (D - cover - 8)); // initial guess, assume 16mm bar, d ≈ D-cover-8

        for (let iter = 0; iter < 10; iter++) {
            const dn = (Ast * fsy) / (conc.alpha2 * fc * conc.gamma * b);
            const a = conc.gamma * dn;
            const d = D - cover - 8; // approximate
            const Mu_calc = Ast * fsy * (d - a / 2);
            if (Mu_calc >= Mu_required) break;
            Ast *= Mu_required / Mu_calc * 1.05; // 5% overestimate for convergence
        }

        // Find smallest matching reo from TRENCH_MESH
        const candidates = Object.entries(TRENCH_MESH)
            .filter(([_, tm]) => tm.Ast >= Ast)
            .sort((a, b) => a[1].Ast - b[1].Ast);

        const suggested = candidates.length > 0 ? candidates[0][0] : null;

        return {
            requiredAst: Math.ceil(Ast),
            Mcr_kNm: Mcr / 1e6,
            Mu_required_kNm: Mu_required / 1e6,
            suggestedReo: suggested,
            suggestedAst: suggested ? TRENCH_MESH[suggested].Ast : null,
            log: [
                `${params.face === 'top' ? 'Hogging' : 'Sagging'} face:`,
                `Z = ${(Z / 1e3).toFixed(0)}×10³ mm³, Mcr = ${(Mcr / 1e6).toFixed(1)} kN.m`,
                `Required Mu ≥ 1.2Mcr = ${(Mu_required / 1e6).toFixed(1)} kN.m`,
                `Required Ast ≈ ${Math.ceil(Ast)} mm²`,
                suggested ? `Suggested: ${suggested} (${TRENCH_MESH[suggested].Ast.toFixed(0)} mm²)` : 'No suitable reo found',
            ],
        };
    }

    // -----------------------------------------------------------------
    // PUBLIC API
    // -----------------------------------------------------------------

    return {
        // Material data
        CONCRETE_PROPS,
        getConcreteProps,
        REBAR,
        TRENCH_MESH,

        // Bending
        bendingCapacity,

        // Cracking & ductility
        crackingMoment,
        ductilityCheck,

        // Shear
        shearCapacityVuc,

        // Full beam design
        raftBeamDesign,

        // Moment/shear estimation
        estimateDesignMoment,
        estimateDesignShear,

        // Auto-suggest reinforcement
        suggestReinforcement,
    };
})();
