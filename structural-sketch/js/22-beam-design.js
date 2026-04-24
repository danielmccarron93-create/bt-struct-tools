/**
 * 22-beam-design.js
 * Steel beam design checks per AS 4100-2020
 *
 * V1 scope:
 *   - Simply supported UB/UC in bending (full lateral restraint assumed)
 *   - φMsx capacity check per Cl 5.1 / Cl 5.2
 *   - Deflection check: δ ≤ L/250 for floors, L/125 for cantilevers
 *   - Auto tributary width from adjacent parallel beams
 *   - Per-level floor loads (G, Q in kPa)
 *   - Per-level one-way span direction
 *
 * Assumptions flagged:
 *   - Full lateral restraint (αm = 1.0, no lateral buckling reduction)
 *   - Simply supported end conditions (M* = wL²/8, δ = 5wL⁴/384EI)
 *   - Self-weight included automatically from section mass
 *   - Load combination: 1.2G + 1.5Q per AS/NZS 1170.0 Cl 4.2.2
 *
 * Loads AFTER: 18a-steel-sections.js, 12-properties.js, 10-levels.js
 */

// ============================================================================
// SECTION PROPERTY CALCULATOR
// ============================================================================
// Computes Zx (plastic), Sx (elastic), Ix, Ag, and mass from catalogue dims.
// Includes fillet contribution for UB/UC sections.

/**
 * Compute section properties for UB/UC I-sections.
 * All inputs in mm, outputs in mm³ (Zx, Sx), mm⁴ (Ix), mm² (Ag), kg/m (mass).
 *
 * Formulae include web-flange fillet region (approximated as quarter-circle fillets
 * at each of the 4 web-flange junctions). This adds ~2-5% to Ix for typical UBs.
 *
 * @param {object} p - { d, bf, tf, tw, r1 } all in mm
 * @returns {object} { Ix, Sx, Zx, Ag, mass }
 */
function calcIBeamProps(p) {
    const { d, bf, tf, tw, r1 } = p;
    const hw = d - 2 * tf; // clear web height between flanges

    // ── Gross area ──
    // Flanges + web + 4 root fillet regions.
    // Each fillet fills the concave web-flange corner: area = r1²(1 - π/4)
    const Af = bf * tf;                          // one flange
    const Aw = tw * hw;                          // web
    const filletArea = (1 - Math.PI / 4) * r1 * r1; // one root fillet region (positive)
    // 4 fillets: 2 per flange × top and bottom
    const Ag = 2 * Af + Aw + 4 * filletArea;

    // ── Second moment of area Ix (about major axis, centroid at d/2) ──
    // Flanges (parallel axis theorem)
    const Ix_flanges = 2 * (bf * Math.pow(tf, 3) / 12 + bf * tf * Math.pow((d - tf) / 2, 2));
    // Web
    const Ix_web = tw * Math.pow(hw, 3) / 12;
    // Fillets — centroid of fillet region (r²(1-π/4) shape) from the web-flange
    // corner is at 0.2234×r1 from each edge. Distance from NA:
    // yFillet = hw/2 - r1 + 0.2234×r1 = hw/2 - 0.7766×r1
    // But the fillet has material spread across its depth, so also need self-Ix.
    // Self-Ix of fillet region ≈ filletArea × (0.1955×r1)² (RMS radius estimate)
    const yFillet = hw / 2 - 0.7766 * r1;
    const Ix_fillet_own = filletArea * Math.pow(0.1955 * r1, 2);
    const Ix_fillets = 4 * (filletArea * yFillet * yFillet + Ix_fillet_own);

    const Ix = Ix_flanges + Ix_web + Ix_fillets;

    // ── Elastic section modulus Sx ──
    const Sx = Ix / (d / 2);

    // ── Plastic section modulus Zx ──
    // For doubly symmetric I-section: Zx = bf*tf*(d-tf) + tw*(d-2tf)²/4 + fillet contribution
    const Zx_flanges = bf * tf * (d - tf);
    const Zx_web = tw * Math.pow(hw, 2) / 4;
    // Each fillet contributes: filletArea × yFillet to plastic modulus
    const Zx_fillets = 4 * filletArea * yFillet;
    const Zx = Zx_flanges + Zx_web + Zx_fillets;

    // ── Mass per metre (steel density = 7850 kg/m³) ──
    const mass = Ag * 7850 / 1e6; // Ag in mm² → m² factor

    return { Ix, Sx, Zx, Ag, mass };
}

/**
 * Compute section properties for SHS/RHS hollow sections.
 * @param {object} p - { d, bf, t, rExt }
 * @returns {object} { Ix, Sx, Zx, Ag, mass }
 */
function calcHollowRectProps(p) {
    const { d, t } = p;
    const bf = p.bf || d; // SHS: bf = d
    const rExt = p.rExt || t * 2.5;
    const rInt = Math.max(rExt - t, 0);

    // Simplified: outer rect minus inner rect (ignoring corner radii for V1)
    const bInner = bf - 2 * t;
    const dInner = d - 2 * t;
    const Ag = bf * d - bInner * dInner;
    const Ix = (bf * Math.pow(d, 3) - bInner * Math.pow(dInner, 3)) / 12;
    const Sx = Ix / (d / 2);
    // Plastic modulus for hollow rect
    const Zx = (bf * Math.pow(d, 2) - bInner * Math.pow(dInner, 2)) / 4;
    const mass = Ag * 7850 / 1e6;

    return { Ix, Sx, Zx, Ag, mass };
}

/**
 * Compute section properties for CHS circular hollow sections.
 * @param {object} p - { od, t }
 * @returns {object} { Ix, Sx, Zx, Ag, mass }
 */
function calcCHSProps(p) {
    const { od, t } = p;
    const di = od - 2 * t;
    const Ag = Math.PI / 4 * (od * od - di * di);
    const Ix = Math.PI / 64 * (Math.pow(od, 4) - Math.pow(di, 4));
    const Sx = Ix / (od / 2);
    const Zx = (Math.pow(od, 3) - Math.pow(di, 3)) / 6;
    const mass = Ag * 7850 / 1e6;

    return { Ix, Sx, Zx, Ag, mass };
}

/**
 * Compute section properties for PFC (parallel flange channel).
 * Approximated as half an I-section (asymmetric about minor axis, symmetric about major).
 * For bending about major axis the Ix/Zx calcs are identical to UB.
 * @param {object} p - { d, bf, tf, tw, r1 }
 * @returns {object} { Ix, Sx, Zx, Ag, mass }
 */
function calcPFCProps(p) {
    // PFC has same major-axis bending properties formula as UB/UC
    // (both flanges contribute equally to Ix about centroid at d/2)
    return calcIBeamProps(p);
}


// ── Property cache ──────────────────────────────────────────
// Computed once per section, stored for reuse.
var _sectionPropsCache = {};

/**
 * Get section properties for a given size string.
 * Returns null if section not found in catalogue.
 * @param {string} sizeStr - e.g., "360UB56.7"
 * @returns {object|null} { type, d, bf, tf, tw, r1, Ix, Sx, Zx, Ag, mass, ... }
 */
function getSectionProperties(sizeStr) {
    if (!sizeStr) return null;
    if (_sectionPropsCache[sizeStr]) return _sectionPropsCache[sizeStr];

    const profile = lookupSteelSection(sizeStr);
    if (!profile) return null;

    // If the catalogue entry already has Red Book properties (Ix, Zx, Sx, Ag, m)
    // use them directly — they're more accurate than computing from geometry.
    // The UB catalogue in 18a-steel-sections.js now carries the full Red Book
    // property set (Ix, Iy, Zx, Sx, J, Iw, Ag, m, d1) sourced from AISC
    // Design Capacity Tables Vol 1, Table 3.1-3(A).
    // Units in catalogue: Ix/Iy in 10⁶ mm⁴, Zx/Sx in 10³ mm³, J in 10³ mm⁴,
    //                     Iw in 10⁹ mm⁶, Ag in mm², m in kg/m.
    if (profile.Ix && profile.Zx && profile.Sx && profile.Ag) {
        // Catalogue has full property set — normalise to base units (mm⁴ etc.)
        // and present the same interface as computed props.
        const result = {
            ...profile,
            // Base-unit conversions for consumers that expect mm⁴/mm³ directly
            Ix: profile.Ix * 1e6,               // 10⁶ mm⁴ → mm⁴
            Sx: profile.Sx * 1e3,               // 10³ mm³ → mm³
            Zx: profile.Zx * 1e3,               // 10³ mm³ → mm³
            mass: profile.m,                     // kg/m
            // Keep the original stored values accessible under _raw prefix
            // for the new calc engine that expects the stored (×10⁶ etc.) form
            _raw_Ix: profile.Ix,
            _raw_Iy: profile.Iy,
            _raw_Zx: profile.Zx,
            _raw_Sx: profile.Sx,
            _raw_J:  profile.J,
            _raw_Iw: profile.Iw,
        };
        // Iy, J, Iw may also be present — convert to base units
        if (profile.Iy != null) result.Iy = profile.Iy * 1e6;
        if (profile.J  != null) result.J_mm4 = profile.J * 1e3;
        if (profile.Iw != null) result.Iw_mm6 = profile.Iw * 1e9;
        _sectionPropsCache[sizeStr] = result;
        return result;
    }

    // Fallback — compute from geometry (UC, PFC, hollows, estimated sections)
    let computed;
    switch (profile.type) {
        case 'UB':
        case 'UC':
            computed = calcIBeamProps(profile);
            break;
        case 'PFC':
            computed = calcPFCProps(profile);
            break;
        case 'SHS':
        case 'RHS':
            computed = calcHollowRectProps(profile);
            break;
        case 'CHS':
            computed = calcCHSProps(profile);
            break;
        default:
            return null; // EA/UA — not typically used as beams, skip for V1
    }

    const result = { ...profile, ...computed };
    _sectionPropsCache[sizeStr] = result;
    return result;
}

/**
 * Get yield stress fy (MPa) for a given grade string per AS 4100 Table 2.1.
 * For hot-rolled sections (UB/UC/PFC): fy depends on flange thickness.
 * For hollow sections (SHS/RHS/CHS): fy depends on wall thickness.
 *
 * AS/NZS 3679.1 Grade 300: fy = 320 MPa (tf ≤ 11mm), 300 MPa (11 < tf ≤ 17mm), 280 MPa (tf > 17mm)
 * AS/NZS 3679.1 Grade 350: fy = 360 MPa (tf ≤ 11mm), 340 MPa (11 < tf ≤ 17mm), 330 MPa (tf > 17mm)
 * AS/NZS 3679.2 Grade 350 (hollow): fy = 350 MPa all thicknesses
 * Grade 450 (hollow): fy = 450 MPa
 *
 * @param {string} grade - '300', '350', '450'
 * @param {string} sectionType - 'UB', 'UC', 'PFC', 'SHS', 'RHS', 'CHS'
 * @param {number} tf - flange/wall thickness in mm
 * @returns {number} fy in MPa
 */
function getYieldStress(grade, sectionType, tf) {
    const isHollow = (sectionType === 'SHS' || sectionType === 'RHS' || sectionType === 'CHS');

    if (isHollow) {
        // AS/NZS 1163 hollow sections
        if (grade === '450') return 450;
        if (grade === '350') return 350;
        return 350; // C350 is standard for hollows
    }

    // Hot-rolled open sections (AS/NZS 3679.1)
    if (grade === '350') {
        if (tf <= 11) return 360;
        if (tf <= 17) return 340;
        return 330;
    }
    if (grade === '450') {
        // Not standard for hot-rolled UB/UC — use 450 as specified
        return 450;
    }
    // Grade 300 (default)
    if (tf <= 11) return 320;
    if (tf <= 17) return 300;
    return 280;
}


// ============================================================================
// LEVEL DESIGN SETTINGS
// ============================================================================
// Extends level objects with floor loading and span direction.

/**
 * Initialise design settings on level system.
 * Called once on module load. Safe to call multiple times (idempotent).
 */
function initDesignSettings() {
    if (!levelSystem || !levelSystem.levels) return;

    // Add designSettings to project if not present
    if (!project.designSettings) {
        project.designSettings = {};
    }

    // Set defaults per level
    for (const lv of levelSystem.levels) {
        if (!project.designSettings[lv.id]) {
            const isRoof = (lv.id === 'RF' || lv.name.toLowerCase().includes('roof'));
            project.designSettings[lv.id] = {
                G: isRoof ? 0.5 : 1.2,     // kPa — permanent action (self-weight of floor system)
                Q: isRoof ? 0.25 : 1.5,    // kPa — imposed action (AS 1170.1 Table 3.1 Category A)
                spanDirection: 0,            // degrees from X-axis (0 = joists span in X, beams in Y carry load)
                                             // More precisely: floor spans in this direction, beams PERPENDICULAR carry load
            };
        }
    }
}

// Initialise on load
initDesignSettings();

/**
 * Get design settings for a given level.
 * @param {string} levelId
 * @returns {object} { G, Q, spanDirection }
 */
function getLevelDesignSettings(levelId) {
    if (!project.designSettings || !project.designSettings[levelId]) {
        initDesignSettings();
    }
    return project.designSettings[levelId] || { G: 1.2, Q: 1.5, spanDirection: 0 };
}


// ============================================================================
// TRIBUTARY WIDTH CALCULATION
// ============================================================================

/**
 * Calculate the tributary width for a beam element.
 *
 * Algorithm:
 * 1. Find all beams on the same level
 * 2. Determine this beam's direction vector
 * 3. Filter to beams running roughly parallel (within 15°)
 * 4. For each side (left/right perpendicular), find the nearest parallel beam
 * 5. Tributary width = half-distance-left + half-distance-right
 * 6. If no beam found on one side, check for slab edge. If no slab, use the
 *    half-distance from the other side (mirror assumption) or flag for override.
 *
 * @param {object} beamEl - beam element { x1, y1, x2, y2, level, ... }
 * @returns {object} { tribWidth (mm), tribLeft (mm), tribRight (mm), method, warnings[] }
 */
function calculateTributaryWidth(beamEl) {
    const warnings = [];
    const levelId = beamEl.level || getActiveLevel().id;

    // ── Beam direction ──
    const dx = beamEl.x2 - beamEl.x1;
    const dy = beamEl.y2 - beamEl.y1;
    const beamLen = Math.sqrt(dx * dx + dy * dy);
    if (beamLen < 1) return { tribWidth: 0, tribLeft: 0, tribRight: 0, method: 'error', warnings: ['Zero-length beam'] };

    // Unit direction vector (along beam)
    const ux = dx / beamLen;
    const uy = dy / beamLen;
    // Perpendicular vector (to the "right" of beam direction)
    const nx = -uy;
    const ny = ux;

    // Beam midpoint (used as reference for distance measurement)
    const mx = (beamEl.x1 + beamEl.x2) / 2;
    const my = (beamEl.y1 + beamEl.y2) / 2;

    // ── Collect all beams on this level (excluding this one) ──
    const beams = project.elements.filter(el =>
        el.type === 'line' &&
        el.layer === 'S-BEAM' &&
        el.id !== beamEl.id &&
        (el.level === levelId || (!el.level && levelId === getActiveLevel().id))
    );

    // ── Filter to parallel beams (within 15° tolerance) ──
    const PARALLEL_TOL = Math.cos(15 * Math.PI / 180); // ~0.966
    const parallelBeams = [];

    for (const b of beams) {
        const bdx = b.x2 - b.x1;
        const bdy = b.y2 - b.y1;
        const bLen = Math.sqrt(bdx * bdx + bdy * bdy);
        if (bLen < 1) continue;

        // Normalise direction
        const bux = bdx / bLen;
        const buy = bdy / bLen;

        // Check parallelism (absolute dot product — allows anti-parallel)
        const dot = Math.abs(ux * bux + uy * buy);
        if (dot >= PARALLEL_TOL) {
            // Signed perpendicular distance from this beam's midpoint to reference beam's line
            const bMx = (b.x1 + b.x2) / 2;
            const bMy = (b.y1 + b.y2) / 2;
            const perpDist = (bMx - mx) * nx + (bMy - my) * ny;
            parallelBeams.push({ el: b, perpDist });
        }
    }

    // ── Find nearest beam on each side ──
    let nearestLeft = null;  // negative perpDist (to the "left")
    let nearestRight = null; // positive perpDist (to the "right")

    for (const pb of parallelBeams) {
        if (pb.perpDist < -50) { // at least 50mm away to count as separate beam
            if (!nearestLeft || pb.perpDist > nearestLeft.perpDist) {
                nearestLeft = pb;
            }
        } else if (pb.perpDist > 50) {
            if (!nearestRight || pb.perpDist < nearestRight.perpDist) {
                nearestRight = pb;
            }
        }
    }

    // ── Calculate tributary distances ──
    let tribLeft, tribRight;
    let method = 'auto';

    if (nearestLeft) {
        tribLeft = Math.abs(nearestLeft.perpDist) / 2;
    } else {
        // No beam to the left — check for slab edge or use right mirror
        tribLeft = null;
        warnings.push('No parallel beam found on left side — using mirror of right side');
    }

    if (nearestRight) {
        tribRight = nearestRight.perpDist / 2;
    } else {
        tribRight = null;
        warnings.push('No parallel beam found on right side — using mirror of left side');
    }

    // Resolve nulls
    if (tribLeft === null && tribRight !== null) {
        tribLeft = tribRight;
        method = 'auto-mirror';
    } else if (tribRight === null && tribLeft !== null) {
        tribRight = tribLeft;
        method = 'auto-mirror';
    } else if (tribLeft === null && tribRight === null) {
        // No parallel beams found at all — flag for manual input
        warnings.length = 0;
        warnings.push('No parallel beams found — please set tributary width manually');
        return { tribWidth: 0, tribLeft: 0, tribRight: 0, method: 'manual-required', warnings };
    }

    const tribWidth = tribLeft + tribRight;

    return { tribWidth, tribLeft, tribRight, method, warnings };
}


// ============================================================================
// AS 4100 BEAM DESIGN CHECK
// ============================================================================

/**
 * Run a full design check on a beam element.
 *
 * @param {object} beamEl - beam element
 * @param {object} [overrides] - optional overrides { tribWidth, G, Q, spanDirection }
 * @returns {object} Design check results
 */
function runBeamDesignCheck(beamEl, overrides) {
    const result = {
        ok: false,
        errors: [],
        warnings: [],
        assumptions: [
            'Simply supported end conditions',
            'Full lateral restraint (no lateral-torsional buckling)',
            'UDL from tributary floor area'
        ],
        // Inputs
        span: 0,
        size: '',
        grade: '',
        fy: 0,
        tribWidth: 0,
        G: 0,
        Q: 0,
        selfWeight: 0,
        // ULS
        wULS: 0,
        Mstar: 0,
        phiMsx: 0,
        bendingUtil: 0,
        bendingOk: false,
        // SLS
        wSLS: 0,
        delta: 0,
        deltaLimit: 0,
        deflectionUtil: 0,
        deflectionOk: false,
        // Section classification
        compact: true,
        sectionClass: 'compact'
    };

    // ── Get beam span ──
    const dx = beamEl.x2 - beamEl.x1;
    const dy = beamEl.y2 - beamEl.y1;
    const span = Math.sqrt(dx * dx + dy * dy); // mm
    result.span = span;
    if (span < 100) {
        result.errors.push('Beam span too short');
        return result;
    }

    // ── Get section from schedule ──
    const typeRef = beamEl.typeRef || beamEl.tag;
    if (!typeRef) {
        result.errors.push('No beam type assigned');
        return result;
    }
    const schedData = project.scheduleTypes.beam[typeRef] || {};
    const sizeStr = schedData.size;
    if (!sizeStr) {
        result.errors.push('No section size assigned — set size in beam schedule');
        return result;
    }
    result.size = sizeStr;
    result.grade = schedData.grade || '300';

    // ── Get section properties ──
    const props = getSectionProperties(sizeStr);
    if (!props) {
        result.errors.push(`Section "${sizeStr}" not found in catalogue`);
        return result;
    }

    // ── Yield stress ──
    const tf = props.tf || props.t || 10;
    const fy = getYieldStress(result.grade, props.type, tf);
    result.fy = fy;

    // ── Section compactness check (AS 4100 Table 5.2) ──
    // For UB/UC flanges: λe = (bf/2)/tf * √(fy/250)
    // Compact limit for flange outstand: λep = 9
    // For web: λe = (d-2tf)/tw * √(fy/250)
    // Compact limit for web in bending: λep = 82
    if (props.type === 'UB' || props.type === 'UC' || props.type === 'PFC') {
        const fyRatio = Math.sqrt(fy / 250);
        const flangeSlender = ((props.bf / 2) / props.tf) * fyRatio;
        const webSlender = ((props.d - 2 * props.tf) / props.tw) * fyRatio;

        if (flangeSlender > 16 || webSlender > 115) {
            result.sectionClass = 'slender';
            result.compact = false;
            result.warnings.push('Section is slender — Zx reduced per AS 4100 Cl 5.2.4');
        } else if (flangeSlender > 9 || webSlender > 82) {
            result.sectionClass = 'non-compact';
            result.compact = false;
            result.warnings.push('Section is non-compact — Zx interpolated per AS 4100 Cl 5.2.3');
        } else {
            result.sectionClass = 'compact';
            result.compact = true;
        }
    }

    // ── Effective section modulus ──
    let Zex;
    if (result.compact) {
        Zex = Math.min(props.Zx, 1.5 * props.Sx); // AS 4100 Cl 5.2.2 cap
    } else if (result.sectionClass === 'non-compact') {
        // Linear interpolation between Sx and Zx
        // Simplified: use Sx as conservative approximation for V1
        Zex = props.Sx;
        result.assumptions.push('Non-compact section: using elastic modulus Sx (conservative)');
    } else {
        // Slender: use reduced effective section modulus
        // Simplified: 0.85 * Sx as conservative approximation for V1
        Zex = 0.85 * props.Sx;
        result.assumptions.push('Slender section: using 0.85×Sx (conservative estimate)');
    }

    // ── Loading ──
    const levelId = beamEl.level || getActiveLevel().id;
    const designSettings = getLevelDesignSettings(levelId);
    const G = (overrides && overrides.G !== undefined) ? overrides.G : designSettings.G; // kPa
    const Q = (overrides && overrides.Q !== undefined) ? overrides.Q : designSettings.Q; // kPa
    result.G = G;
    result.Q = Q;

    // ── Tributary width ──
    let tribWidth;
    if (overrides && overrides.tribWidth > 0) {
        tribWidth = overrides.tribWidth;
    } else {
        const tribResult = calculateTributaryWidth(beamEl);
        tribWidth = tribResult.tribWidth;
        result.warnings.push(...tribResult.warnings);
        if (tribResult.method === 'manual-required') {
            result.errors.push('Cannot auto-detect tributary width — no parallel beams found');
            return result;
        }
    }
    result.tribWidth = tribWidth;

    // Self-weight of beam (kN/m)
    const selfWeight = props.mass * 9.81 / 1000; // kg/m → kN/m
    result.selfWeight = selfWeight;

    // ── ULS load (kN/m) ──
    // Floor area load → line load: w = pressure × tribWidth(m)
    const tribWidthM = tribWidth / 1000;
    const wG = G * tribWidthM;          // kN/m from permanent floor load
    const wQ = Q * tribWidthM;          // kN/m from imposed floor load
    const wSW = selfWeight;             // kN/m beam self-weight (permanent)

    const wULS = 1.2 * (wG + wSW) + 1.5 * wQ; // AS/NZS 1170.0 Cl 4.2.2(a)
    result.wULS = wULS;

    // ── ULS bending moment ──
    const spanM = span / 1000;
    const Mstar = wULS * spanM * spanM / 8; // kN·m for simply supported UDL
    result.Mstar = Mstar;

    // ── Bending capacity ──
    const phi = 0.9; // AS 4100 Table 3.4
    const phiMsx = phi * Zex * fy / 1e6; // Zex in mm³, fy in MPa → kN·m
    result.phiMsx = phiMsx;
    result.bendingUtil = Mstar / phiMsx;
    result.bendingOk = (Mstar <= phiMsx);

    // ── SLS deflection ──
    const E = 200000; // MPa (Young's modulus for steel)
    const wSLS = (wG + wSW) + 0.7 * wQ; // Short-term combination (AS 1170.0 Appendix C)
    // Note: 0.7 is ψs for Category A (residential) per AS 1170.0 Table C2
    result.wSLS = wSLS;

    const wSLS_Nm = wSLS * 1000 / 1000; // kN/m → N/mm (wSLS in kN/m × 1000/1000 = N/mm)
    // Actually: kN/m = N/mm directly. 1 kN/m = 1 N/mm. So wSLS is already in N/mm.
    const delta = 5 * wSLS * Math.pow(span, 4) / (384 * E * props.Ix); // mm
    result.delta = delta;

    const deltaLimit = span / 250; // AS 1170.0 / NCC guidance for floors
    result.deltaLimit = deltaLimit;
    result.deflectionUtil = delta / deltaLimit;
    result.deflectionOk = (delta <= deltaLimit);

    // ── Overall result ──
    result.ok = result.bendingOk && result.deflectionOk && result.errors.length === 0;

    return result;
}


// ============================================================================
// PROPERTIES PANEL INTEGRATION
// ============================================================================

/**
 * Generate HTML for the design check section in the properties panel.
 * Called from the beam section of updatePropsPanel().
 *
 * @param {object} beamEl - the selected beam element
 * @returns {string} HTML string to append to properties panel
 */
function buildDesignCheckHTML(beamEl) {
    let html = '';

    // ── Design Check button / header ──
    html += `<div style="margin-top:8px; border-top:1px solid var(--border-color,#ddd); padding-top:8px;">`;
    html += `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">`;
    html += `<span style="font-weight:700; font-size:11px; color:var(--text-primary,#333); letter-spacing:0.5px;">DESIGN CHECK</span>`;
    html += `<button onclick="toggleDesignCheckPanel('${beamEl.id}')" style="font-size:9px; padding:2px 8px; border:1px solid var(--accent,#2563EB); color:var(--accent,#2563EB); background:transparent; border-radius:3px; cursor:pointer;">AS 4100</button>`;
    html += `</div>`;

    // ── Check if we have enough data to run ──
    const typeRef = beamEl.typeRef || beamEl.tag;
    let schedData = typeRef ? (project.scheduleTypes.beam[typeRef] || {}) : {};
    if (!schedData.size && project.scheduleTypes.floorBeam) {
        schedData = project.scheduleTypes.floorBeam[typeRef] || schedData;
    }

    if (!typeRef || !schedData.size) {
        html += `<div style="font-size:10px; color:#999; padding:4px 0;">Assign a section size to enable design checks</div>`;
        html += `</div>`;
        return html;
    }

    // ── Run the check — prefer enhanced engine if available ──
    const check = (typeof runEnhancedBeamCheck === 'function')
        ? runEnhancedBeamCheck(beamEl)
        : runBeamDesignCheck(beamEl);

    if (check.errors && check.errors.length > 0) {
        html += `<div style="font-size:10px; color:#DC2626; padding:4px 0;">`;
        for (const err of check.errors) {
            html += `⚠ ${err}<br>`;
        }
        html += `</div></div>`;
        return html;
    }

    // ── Results display ──
    const bendColor = check.bendingUtil <= 0.9 ? '#16A34A' : check.bendingUtil <= 1.0 ? '#D97706' : '#DC2626';
    const deflColor = check.deflectionUtil <= 0.9 ? '#16A34A' : check.deflectionUtil <= 1.0 ? '#D97706' : '#DC2626';
    const overallColor = check.ok ? '#16A34A' : '#DC2626';
    const overallLabel = check.ok ? 'PASS' : 'FAIL';

    // Overall status + utilisation + auto-size
    const maxUtilPct = (check.maxUtil * 100).toFixed(0);
    html += `<div style="display:flex; justify-content:center; align-items:center; gap:6px; margin:4px 0 8px 0;">`;
    html += `<span style="font-weight:700; font-size:12px; color:${overallColor}; background:${overallColor}15; padding:2px 12px; border-radius:3px; border:1px solid ${overallColor}40;">${overallLabel} · ${maxUtilPct}%</span>`;
    html += `<span style="font-size:9px; color:#666;">${check.governing || ''}</span>`;
    html += `</div>`;
    // Auto-size button
    html += `<div style="text-align:center; margin-bottom:6px;">`;
    html += `<button onclick="applyAutoSize('${beamEl.id}')" style="font-size:9px; padding:2px 10px; border:1px solid #8B5CF6; color:#8B5CF6; background:transparent; border-radius:3px; cursor:pointer;" title="Find lightest passing UB section">Auto-size</button>`;
    html += `</div>`;

    // Loading inputs (editable)
    const levelId = beamEl.level || getActiveLevel().id;
    const ds = getLevelDesignSettings(levelId);
    html += `<div style="font-size:10px; color:#666; margin-bottom:4px;">`;
    html += `<div style="display:flex; justify-content:space-between;"><span>G (dead)</span><span class="prop-click-edit" onclick="editDesignLoad('${levelId}','G')" style="cursor:pointer;">${ds.G.toFixed(1)} kPa</span></div>`;
    html += `<div style="display:flex; justify-content:space-between;"><span>Q (live)</span><span class="prop-click-edit" onclick="editDesignLoad('${levelId}','Q')" style="cursor:pointer;">${ds.Q.toFixed(1)} kPa</span></div>`;
    html += `<div style="display:flex; justify-content:space-between;"><span>Self-wt</span><span>${(check.selfWeight || 0).toFixed(2)} kN/m</span></div>`;
    html += `<div style="display:flex; justify-content:space-between;"><span>Trib. width</span><span>${((check.tribWidth || 0) / 1000).toFixed(2)} m${check.isTieBeam ? ' <em style="color:#999;font-size:9px;">(tie beam)</em>' : ''}</span></div>`;
    html += `<div style="display:flex; justify-content:space-between;"><span>Span</span><span>${((check.span || 0) / 1000).toFixed(2)} m</span></div>`;
    if (check.governingCombo) {
        html += `<div style="display:flex; justify-content:space-between;"><span>Governing</span><span>${check.governingCombo}</span></div>`;
    }
    html += `</div>`;

    // Bending check (member capacity if LTB available, otherwise section)
    const bendLabel = check.phiMbx !== undefined ? 'Bending (LTB)' : 'Bending';
    const bendCapacity = check.phiMbx !== undefined ? check.phiMbx : check.phiMsx;
    html += `<div style="font-size:10px; margin-top:6px; padding:4px; background:${bendColor}08; border-left:3px solid ${bendColor}; border-radius:0 3px 3px 0;">`;
    html += `<div style="font-weight:600; color:${bendColor};">${bendLabel} — ${(check.bendingUtil * 100).toFixed(0)}%</div>`;
    html += `<div style="color:#666;">M* = ${(check.Mstar || 0).toFixed(1)} kN·m</div>`;
    html += `<div style="color:#666;">φM = ${(bendCapacity || 0).toFixed(1)} kN·m</div>`;
    html += `<div style="color:#666; font-size:9px;">fy = ${check.fy} MPa (Grade ${check.grade}${check.sectionClass ? ', ' + check.sectionClass : ''})</div>`;
    html += `</div>`;

    // Shear check (if available from enhanced engine)
    if (check.phiVv !== undefined && check.Vstar !== undefined) {
        const shearColor = (check.shearUtil || 0) <= 0.9 ? '#16A34A' : (check.shearUtil || 0) <= 1.0 ? '#D97706' : '#DC2626';
        html += `<div style="font-size:10px; margin-top:4px; padding:4px; background:${shearColor}08; border-left:3px solid ${shearColor}; border-radius:0 3px 3px 0;">`;
        html += `<div style="font-weight:600; color:${shearColor};">Shear — ${((check.shearUtil || 0) * 100).toFixed(0)}%</div>`;
        html += `<div style="color:#666;">V* = ${(check.Vstar || 0).toFixed(1)} kN  ·  φVv = ${(check.phiVv || 0).toFixed(1)} kN</div>`;
        html += `</div>`;
    }

    // Deflection check
    html += `<div style="font-size:10px; margin-top:4px; padding:4px; background:${deflColor}08; border-left:3px solid ${deflColor}; border-radius:0 3px 3px 0;">`;
    html += `<div style="font-weight:600; color:${deflColor};">Deflection — ${((check.deflectionUtil || 0) * 100).toFixed(0)}%</div>`;
    html += `<div style="color:#666;">δ = ${(check.delta || 0).toFixed(1)} mm (limit ${(check.deltaLimit || 0).toFixed(1)} mm = L/250)</div>`;
    if (check.wSLS !== undefined && check.wSLS > 0) {
        html += `<div style="color:#666; font-size:9px;">SLS: w = ${check.wSLS.toFixed(2)} kN/m (G + ψs·Q, ψs = 0.7)</div>`;
    }
    html += `</div>`;

    // Warnings
    if (check.warnings && check.warnings.length > 0) {
        html += `<div style="font-size:9px; color:#D97706; margin-top:4px;">`;
        for (const w of check.warnings) {
            html += `⚠ ${w}<br>`;
        }
        html += `</div>`;
    }

    // Calc sheet button (if FloorCalcSheet is available)
    if (typeof FloorCalcSheet !== 'undefined' && FloorCalcSheet.showCalcSheetModal) {
        html += `<div style="margin-top:6px; text-align:center;">`;
        html += `<button onclick="FloorCalcSheet.showCalcSheetModal(selectedElement)" style="font-size:10px; padding:3px 14px; border:1px solid var(--accent,#2B7CD0); color:var(--accent,#2B7CD0); background:transparent; border-radius:3px; cursor:pointer;">View Calc Sheet</button>`;
        html += `</div>`;
    }

    // Assumptions
    if (check.assumptions && check.assumptions.length > 0) {
        html += `<div style="font-size:9px; color:#999; margin-top:6px; cursor:pointer;" onclick="this.nextElementSibling.style.display = this.nextElementSibling.style.display === 'none' ? 'block' : 'none'">`;
        html += `▸ Assumptions (click to expand)`;
        html += `</div>`;
        html += `<div style="display:none; font-size:9px; color:#999; padding-left:8px;">`;
        for (const a of check.assumptions) {
            html += `• ${a}<br>`;
        }
        html += `<br>AS 4100-2020 | AS/NZS 1170.0 | AS/NZS 1170.1`;
        html += `</div>`;
    }

    html += `</div>`;
    return html;
}

// ── Toggle / refresh design check ──
function toggleDesignCheckPanel(elId) {
    // Simply re-render the props panel to refresh results
    _lastPropsElement = null; // force refresh
    updatePropsPanel();
}

// ── Edit design load ──
function editDesignLoad(levelId, field) {
    const ds = getLevelDesignSettings(levelId);
    const label = field === 'G' ? 'Dead load G (kPa)' : 'Live load Q (kPa)';
    const current = ds[field];
    const input = prompt(label, current.toFixed(1));
    if (input === null) return;
    const val = parseFloat(input);
    if (isNaN(val) || val < 0) return;
    ds[field] = val;
    // Refresh panel
    _lastPropsElement = null;
    updatePropsPanel();
}


// ============================================================================
// ENHANCED BEAM CHECK — delegates to FloorCalcEngine when available
// ============================================================================
// Slice 3: If the new engine modules are loaded AND the beam has floor zones
// overlapping it, run the piecewise analysis + full AS 4100 check (with LTB
// and shear). Otherwise fall back to the original runBeamDesignCheck above.
//
// This is the function that the properties panel and future calc sheet will
// call. It returns an object compatible with the original result shape PLUS
// the new engine's richer output (steps[], util_*, etc.).

/**
 * Run the best-available design check on a beam element.
 *
 * @param {object} beamEl — beam element { x1, y1, x2, y2, typeRef, level, ... }
 * @param {object} [overrides] — optional { tribWidth, G, Q, restraint, joistSpacing_mm }
 * @returns {object} — design check result
 */
function runEnhancedBeamCheck(beamEl, overrides) {
    // Gate: new engine must be loaded
    if (typeof FloorCalcEngine === 'undefined' || typeof FloorCalcEngine.checkBeam !== 'function') {
        return runBeamDesignCheck(beamEl, overrides);
    }

    // ── Basics ──
    const dx = beamEl.x2 - beamEl.x1;
    const dy = beamEl.y2 - beamEl.y1;
    const span_mm = Math.sqrt(dx * dx + dy * dy);
    if (span_mm < 100) return runBeamDesignCheck(beamEl, overrides);

    const typeRef = beamEl.typeRef || beamEl.tag;
    if (!typeRef) return runBeamDesignCheck(beamEl, overrides);

    // Look up schedule — try beam first, then floorBeam
    let schedData = (project.scheduleTypes.beam && project.scheduleTypes.beam[typeRef]) || {};
    if (!schedData.size && project.scheduleTypes.floorBeam) {
        schedData = project.scheduleTypes.floorBeam[typeRef] || schedData;
    }
    const sizeStr = schedData.size;
    if (!sizeStr) return runBeamDesignCheck(beamEl, overrides);

    const props = getSectionProperties(sizeStr);
    if (!props) return runBeamDesignCheck(beamEl, overrides);

    // ── Tributary width ──
    let tribWidth_mm;
    let tribWidth_geometric_mm = 0; // geometric width (for display, even on tie beams)
    let isTieBeam = false;
    if (overrides && overrides.tribWidth > 0) {
        tribWidth_mm = overrides.tribWidth;
        tribWidth_geometric_mm = tribWidth_mm;
    } else {
        const tribResult = calculateTributaryWidth(beamEl);
        if (tribResult.method === 'manual-required') {
            return runBeamDesignCheck(beamEl, overrides);
        }
        tribWidth_geometric_mm = tribResult.tribWidth;
        tribWidth_mm = tribResult.tribWidth;
    }

    const levelId = beamEl.level || (typeof getActiveLevel === 'function' ? getActiveLevel().id : 'L1');

    // ── Span direction check: is this beam a tie beam? ──
    // A beam running parallel to joist span carries no floor load (self-weight only).
    // A beam perpendicular to joist span carries full tributary floor load.
    {
        const designSettings = getLevelDesignSettings(levelId);
        const spanDir_deg = designSettings.spanDirection || 0; // 0° = X, 90° = Y

        // Beam direction angle in degrees (0° = horizontal E-W, 90° = vertical N-S)
        const bDx = (beamEl.x2 || 0) - (beamEl.x1 || 0);
        const bDy = (beamEl.y2 || 0) - (beamEl.y1 || 0);
        const beamAngle_deg = Math.abs(Math.atan2(bDy, bDx) * 180 / Math.PI);
        // Normalise to 0–180 range (beam direction is unsigned)
        const beamDir = beamAngle_deg > 90 ? 180 - beamAngle_deg : beamAngle_deg;
        // spanDir normalised similarly
        const spanDirNorm = (spanDir_deg % 180);

        // Angular difference between beam direction and joist span direction
        let angleDiff = Math.abs(beamDir - spanDirNorm);
        if (angleDiff > 90) angleDiff = 180 - angleDiff;

        // If beam is within 15° of span direction → tie beam → no floor load
        // If beam is within 15° of perpendicular to span → loaded beam → full trib
        if (angleDiff < 15) {
            // Beam runs parallel to joist span — it's a tie beam
            isTieBeam = true;
            tribWidth_mm = 0;
        }
        // else: beam is perpendicular to span or angled — keep full trib width
    }

    // ── Check for floor zones ──
    // If floor zones exist on this level, use the piecewise path.
    // Otherwise fall back to the uniform-UDL legacy check.
    let hasFloorZones = false;
    let zones = [];
    if (typeof FloorLoadResolver !== 'undefined' && typeof FloorLoadResolver.resolveZonesFromElements === 'function') {
        const floorLoadSchedule = project.scheduleTypes.floorLoad || {};
        zones = FloorLoadResolver.resolveZonesFromElements(project.elements, floorLoadSchedule);
        hasFloorZones = zones.length > 0;
    }

    const span_m = span_mm / 1000;
    const grade = schedData.grade || '300';
    const restraint = (overrides && overrides.restraint) || 'full'; // default full until joist wiring (Slice 5)
    const joistSpacing_mm = (overrides && overrides.joistSpacing_mm) || 450;

    let analysis_override = null;

    if (hasFloorZones
        && typeof FloorBeamAnalysis !== 'undefined'
        && typeof FloorBeamAnalysis.analyseSSBeam === 'function'
        && typeof FloorLoadResolver.buildPiecewiseLoadFns === 'function') {
        // ── Piecewise path ──
        const designSettings = getLevelDesignSettings(levelId);
        const defaultG = (overrides && overrides.G !== undefined) ? overrides.G : designSettings.G;
        const defaultQ = (overrides && overrides.Q !== undefined) ? overrides.Q : designSettings.Q;

        const loadFns = FloorLoadResolver.buildPiecewiseLoadFns(
            beamEl, tribWidth_mm, zones, defaultG, defaultQ
        );

        // Self-weight must be added to w_G
        const sw_kNm = (props.mass || 0) * 9.81 / 1000;
        const w_G_with_sw = function (x_m) {
            return loadFns.w_G_of_x_kNm(x_m) + sw_kNm;
        };

        const Ix_mm4 = props.Ix || 0;
        const EI_Nmm2 = 200000 * Ix_mm4;

        analysis_override = FloorBeamAnalysis.analyseSSBeam({
            span_m: span_m,
            w_G_of_x_kNm: w_G_with_sw,
            w_Q_of_x_kNm: loadFns.w_Q_of_x_kNm,
            EI_Nmm2: EI_Nmm2,
        });
    }

    // ── Call the full AS 4100 engine ──
    const engineInput = {
        section: props,
        grade: grade,
        span_m: span_m,
        restraint: restraint,
        joistSpacing_mm: joistSpacing_mm,
    };
    if (analysis_override) {
        engineInput.analysis_override = analysis_override;
    } else {
        // Uniform UDL path — get loads from level settings
        const designSettings = getLevelDesignSettings(levelId);
        const G = (overrides && overrides.G !== undefined) ? overrides.G : designSettings.G;
        const Q = (overrides && overrides.Q !== undefined) ? overrides.Q : designSettings.Q;
        const tribW_m = tribWidth_mm / 1000;
        engineInput.udl_G_kNm = G * tribW_m;
        engineInput.udl_Q_kNm = Q * tribW_m;
    }

    try {
        const engineResult = FloorCalcEngine.checkBeam(engineInput);
        // Map back to the shape expected by buildDesignCheckHTML
        return {
            ok: engineResult.pass,
            errors: [],
            warnings: [],
            assumptions: [
                'Simply supported end conditions',
                engineResult.Le_pos_mm > 0
                    ? 'LTB check per AS 4100 Cl 5.6 (Le = ' + Math.round(engineResult.Le_pos_mm) + ' mm)'
                    : 'Full lateral restraint (no LTB reduction)',
                analysis_override ? 'Piecewise load from floor zones' : 'UDL from tributary floor area',
            ],
            span: span_mm,
            size: sizeStr,
            grade: grade,
            fy: engineResult.inputs.fy,
            tribWidth: tribWidth_mm,
            tribWidthGeometric: tribWidth_geometric_mm,
            isTieBeam: isTieBeam,
            G: engineResult.inputs.udl_G_kNm || 0,
            Q: engineResult.inputs.udl_Q_kNm || 0,
            selfWeight: (props.mass || 0) * 9.81 / 1000,
            // ULS — map to legacy field names
            wULS: 0, // not directly available from engine, M*/V* used instead
            Mstar: engineResult.M_star_kNm,
            phiMsx: engineResult.phiMsx_kNm,
            phiMbx: engineResult.phiMbx_kNm,
            bendingUtil: engineResult.util_moment,
            bendingOk: engineResult.util_moment <= 1.0,
            // Shear
            Vstar: engineResult.V_star_kN,
            phiVv: engineResult.phiVv_kN,
            shearUtil: engineResult.util_shear,
            shearOk: engineResult.util_shear <= 1.0,
            // SLS
            wSLS: 0,
            delta: engineResult.deflection_total_mm,
            deltaLimit: engineResult.limit_total_mm,
            deflectionUtil: engineResult.util_defl_total,
            deflectionOk: engineResult.util_defl_total <= 1.0,
            deflectionLiveUtil: engineResult.util_defl_live,
            // New engine extras
            sectionClass: engineResult.section_class,
            compact: engineResult.section_class === 'Compact',
            maxUtil: engineResult.max_util,
            governing: engineResult.governing,
            governingCombo: engineResult.governing_combo,
            steps: engineResult.steps,
            // Flag for UI to know which engine ran
            _engine: 'FloorCalcEngine',
        };
    } catch (e) {
        // If the new engine throws, fall back to legacy
        console.warn('[beam-design] FloorCalcEngine error, falling back to legacy:', e);
        return runBeamDesignCheck(beamEl, overrides);
    }
}

// ============================================================================
// AUTO-SIZE BEAM — find lightest passing UB section
// ============================================================================

/**
 * Find the lightest UB section that passes all AS 4100 checks for a beam.
 * Iterates from lightest to heaviest, returns first passing section.
 *
 * @param {object} beamEl — beam element from project.elements
 * @returns {object|null} { size, result } or null if nothing passes
 */
function autoSizeBeam(beamEl) {
    if (!beamEl || beamEl.type !== 'line' || beamEl.layer !== 'S-BEAM') return null;

    const typeRef = beamEl.typeRef || beamEl.tag;
    let schedData = typeRef ? (project.scheduleTypes.beam[typeRef] || {}) : {};
    if (!schedData.size && project.scheduleTypes.floorBeam) {
        schedData = project.scheduleTypes.floorBeam[typeRef] || schedData;
    }

    const origSize = schedData.size;
    const sections = typeof STEEL_SECTIONS !== 'undefined' ? STEEL_SECTIONS.UB : [];

    for (let i = 0; i < sections.length; i++) {
        schedData.size = sections[i];
        try {
            const result = runEnhancedBeamCheck(beamEl);
            if (result && result.ok) {
                schedData.size = origSize; // restore
                return { size: sections[i], result: result };
            }
        } catch (e) { /* skip section */ }
    }

    // Nothing passed — restore original
    schedData.size = origSize;
    return null;
}

/**
 * Auto-size a beam and apply the result to its schedule entry.
 * Called from the properties panel UI.
 * @param {string} beamId — element ID
 */
function applyAutoSize(beamId) {
    const el = project.elements.find(function (e) { return e.id == beamId; });
    if (!el) return;

    const result = autoSizeBeam(el);
    if (!result) {
        console.warn('[beam-design] No passing UB section found for beam ' + beamId);
        return;
    }

    const typeRef = el.typeRef || el.tag;
    let schedData = project.scheduleTypes.beam[typeRef];
    if (!schedData && project.scheduleTypes.floorBeam) {
        schedData = project.scheduleTypes.floorBeam[typeRef];
    }
    if (!schedData) return;

    const oldSize = schedData.size;
    schedData.size = result.size;

    // Invalidate utilisation cache
    _beamUtilCache = {};
    _beamUtilHash = '';

    // Re-render
    if (typeof engine !== 'undefined' && engine.requestRender) engine.requestRender();
    if (typeof updatePropertiesPanel === 'function') updatePropertiesPanel();

    console.log('[beam-design] Auto-sized: ' + oldSize + ' → ' + result.size +
        ' (util ' + (result.result.maxUtil * 100).toFixed(0) + '%)');
}
