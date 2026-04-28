"""
AS 3600:2018 Concrete Column Design Engine
============================================
Calculations follow AS 3600:2018 Section 10 and are cross-referenced
with Reinforced Concrete Basics 3e, Chapter 5 (Gilbert/Mickleborough/
Ranzi/Foster, Pearson 2021, ISBN 9780655703662).

Sign conventions:
- Compressive forces are positive (N, Cc, Cs all +).
- Moments are taken about the plastic centroid, unless stated.
- For the section capacity routine, the "compressive face" is the
  face where the extreme compressive fibre is (strain = 0.003 at ULS).
- For M1/M2, M2 is the larger end moment (absolute value).
  M1/M2 is NEGATIVE for single curvature, POSITIVE for double.

Units: SI — mm, N, MPa (= N/mm^2), kN, kNm.
"""

from __future__ import annotations
import math
from dataclasses import dataclass, field
from typing import List, Tuple, Optional


# =============================================================================
# MATERIAL COEFFICIENTS (AS 3600 Cl 10.6.2)
# =============================================================================

def alpha_1(fc: float) -> float:
    """Squash-load concrete stress coefficient (AS 3600 Cl 10.6.2.2)."""
    a = 1.0 - 0.003 * fc
    return max(0.72, min(0.85, a))


def alpha_2(fc: float) -> float:
    """Equivalent rectangular stress-block intensity (AS 3600 Cl 10.6.2.5 Eq (1))."""
    a = 0.85 - 0.0015 * fc
    return max(0.67, a)


def gamma(fc: float) -> float:
    """Equivalent rectangular stress-block depth ratio (AS 3600 Cl 10.6.2.5 Eq (2))."""
    g = 0.97 - 0.0025 * fc
    return max(0.67, g)


# =============================================================================
# BAR DATABASE
# =============================================================================

BAR_AREAS = {  # mm^2
    12: 113.1,
    16: 201.1,
    20: 314.2,
    24: 452.4,
    28: 615.8,
    32: 804.2,
    36: 1017.9,
    40: 1256.6,
}


# =============================================================================
# SECTION DEFINITION
# =============================================================================

@dataclass
class RebarLayer:
    """One horizontal layer of reinforcement in the cross-section.

    d_from_comp_face : distance from the extreme compressive fibre, mm
    area             : total steel area in the layer, mm^2
    """
    d_from_comp_face: float
    area: float


@dataclass
class Void:
    """Internal void (PT duct, conduit) in the cross-section.

    d_from_comp_face : depth of void centre from extreme compressive fibre, mm
    diameter         : void diameter, mm
    """
    d_from_comp_face: float
    diameter: float

    def area(self) -> float:
        return math.pi * self.diameter ** 2 / 4.0


@dataclass
class Section:
    """Generic cross-section (rectangular, square, or circular).

    For rectangular/square:
        shape = "rect", width b, depth D (bending about x-axis — D is
        the dimension parallel to the applied moment).
    For circular:
        shape = "circ", diameter = D (b is ignored).

    Reinforcement is specified as a list of horizontal layers with the
    distance from the extreme compressive fibre. This works for any
    layout — symmetric, asymmetric, multiple layers.

    fc   : characteristic concrete strength, MPa
    fsy  : yield stress of longitudinal steel, MPa
    Es   : elastic modulus of steel, MPa (default 200,000)
    """
    shape: str
    D: float
    b: float
    fc: float
    fsy: float
    layers: List[RebarLayer]
    Es: float = 200_000.0
    voids: List[Void] = field(default_factory=list)

    def void_area(self) -> float:
        return sum(v.area() for v in self.voids)

    def gross_area(self) -> float:
        if self.shape == "circ":
            base = math.pi * (self.D ** 2) / 4.0
        else:
            base = self.b * self.D
        return base - self.void_area()

    def total_steel_area(self) -> float:
        return sum(L.area for L in self.layers)

    # ----- Plastic centroid (AS 3600, from squash condition) --------------
    def plastic_centroid(self) -> float:
        """Depth to the plastic centroid measured from the extreme
        compressive fibre (RCB §5.4.2, Eq 5.8). For a symmetric section
        without voids this equals D/2; voids and asymmetric reinforcement
        shift it. Validated against RCB Example 5.1 (600×800, 8N32, 150 mm
        void at d=500 → 397 mm)."""
        a1 = alpha_1(self.fc)
        sigma_s = min(self.fsy, 0.0025 * self.Es)
        if self.shape == "circ":
            Ag_solid = math.pi * self.D ** 2 / 4.0
        else:
            Ag_solid = self.b * self.D
        # Concrete first moment about compressive face (gross), then subtract
        # voids and steel-displaced areas.
        Mc = a1 * self.fc * Ag_solid * (self.D / 2)
        for v in self.voids:
            Mc -= a1 * self.fc * v.area() * v.d_from_comp_face
        for L in self.layers:
            Mc -= a1 * self.fc * L.area * L.d_from_comp_face
        Ms = sum(L.area * sigma_s * L.d_from_comp_face for L in self.layers)
        return (Mc + Ms) / self.nuo()

    # ----- Squash load (Cl 10.6.2.2, RCB Eq 5.6) --------------------------
    def nuo(self) -> float:
        """Squash load, N (compression positive).

        Cl 10.6.2.2 limits the maximum compressive strain in the longitudinal
        reinforcement at squash to 0.0025, so the contributing steel stress is
        capped at εs·Es = 0.0025·Es (=500 MPa for Es=200 GPa). For Class N
        500 MPa bars the cap doesn't bind; for 600 MPa bars it does.
        """
        a1 = alpha_1(self.fc)
        As = self.total_steel_area()
        Ag = self.gross_area()
        sigma_s_squash = min(self.fsy, 0.0025 * self.Es)
        return a1 * self.fc * (Ag - As) + As * sigma_s_squash

    # ----- Pure tension (RCB Eq 5.17) -------------------------------------
    def nuo_t(self) -> float:
        """Pure tension capacity (steel only, concrete cracked)."""
        return self.total_steel_area() * self.fsy

    # ----- Width at a given depth from compressive face -------------------
    def width_at_depth(self, y_from_comp_face: float) -> float:
        """Width perpendicular to the bending axis at depth y from the
        extreme compressive fibre. For circular sections, derived from
        the chord at that height."""
        if self.shape == "circ":
            R = self.D / 2
            # Compressive face is the top of the circle; y = 0 is the top.
            # Distance from circle centre is (y - R).
            offset = y_from_comp_face - R
            if abs(offset) >= R:
                return 0.0
            return 2.0 * math.sqrt(R * R - offset * offset)
        return self.b

    # ----- Compressed concrete force Cc + its centroid --------------------
    def compressed_concrete(self, dn: float) -> Tuple[float, float]:
        """Return (Cc, dC) — force in compressed concrete in N and depth
        to its centroid from the extreme compressive fibre in mm.

        Uses the equivalent rectangular stress block (Cl 10.6.2.5) of
        intensity alpha_2*fc acting over a depth of gamma*dn. Voids whose
        centre lies within the stress block are subtracted (point approx).
        """
        a2 = alpha_2(self.fc)
        g = gamma(self.fc)
        if self.shape == "circ":
            a2 = 0.95 * a2
        stress_block_depth = g * dn
        if stress_block_depth <= 0:
            return 0.0, 0.0
        if self.shape == "rect":
            Cc = a2 * self.fc * self.b * stress_block_depth
            moment = Cc * (stress_block_depth / 2.0)
        else:
            n = 40
            dy = stress_block_depth / n
            Cc = 0.0
            moment = 0.0
            for i in range(n):
                y_mid = (i + 0.5) * dy
                w = self.width_at_depth(y_mid)
                dF = a2 * self.fc * w * dy
                Cc += dF
                moment += dF * y_mid
        # ── PRO ── Subtract voids inside the stress block (point approximation)
        for v in self.voids:
            if v.d_from_comp_face <= stress_block_depth:
                dF = a2 * self.fc * v.area()
                Cc -= dF
                moment -= dF * v.d_from_comp_face
        if Cc <= 0:
            return 0.0, 0.0
        return Cc, moment / Cc

    # ----- Interaction point at a given neutral axis depth -----------------
    def interaction_point(self, dn: float) -> Tuple[float, float]:
        """For a given neutral axis depth dn (from extreme compressive
        fibre), return (Nu, Mu) in N and Nmm. Compression positive, Mu
        taken about the plastic centroid.

        Handles the general case where some steel layers may be in
        compression and others in tension, based on dn relative to each
        layer's depth.
        """
        Cc, dC = self.compressed_concrete(dn)
        dpc = self.plastic_centroid()

        # Steel force contributions
        Fs_total = 0.0   # net steel force (compression positive)
        Ms_total = 0.0   # steel moment about plastic centroid
        for L in self.layers:
            eps = 0.003 * (dn - L.d_from_comp_face) / dn
            eps_yield = self.fsy / self.Es
            if eps >= 0:
                stress = min(eps * self.Es, self.fsy)  # compression
            else:
                stress = max(eps * self.Es, -self.fsy)  # tension
            F = stress * L.area
            Fs_total += F
            # Moment arm: (dpc - L.d) — positive when layer is above pc
            Ms_total += F * (dpc - L.d_from_comp_face)

        Nu = Cc + Fs_total
        Mu = Cc * (dpc - dC) + Ms_total
        return Nu, Mu

    # ----- Balanced point (ku = kub = 0.545 for 500 MPa steel, RCB) -------
    def balanced_point(self) -> Tuple[float, float, float]:
        """Return (Nub, Mub, kub) at balanced failure — tension steel
        yielding simultaneously with concrete crushing at 0.003.

        kub is computed per RCB Eq 3.71 which for fsy = 500 MPa gives
        0.545. For other fsy values we compute it from strain geometry.
        """
        eps_sy = self.fsy / self.Es
        kub = 0.003 / (0.003 + eps_sy)
        # dnb is measured from compressive face to the OUTERMOST tensile bar
        # (deepest layer). In RCB this uses "do" = depth to tensile steel.
        do = max(L.d_from_comp_face for L in self.layers)
        dnb = kub * do
        Nu, Mu = self.interaction_point(dnb)
        return Nu, Mu, kub

    # ----- Decompression point (ku = 1.0 at outermost tensile bar) --------
    def decompression_point(self) -> Tuple[float, float]:
        """Point B in RCB §5.3.3 — neutral axis at the outermost layer."""
        do = max(L.d_from_comp_face for L in self.layers)
        return self.interaction_point(do)

    # ----- Pure bending (iterate dn until Nu = 0) -------------------------
    def pure_bending(self, tol: float = 1e-3) -> Tuple[float, float, float]:
        """Return (Nu ≈ 0, Muo, dn). Iterates dn using bisection."""
        # Search bounds: dn small (mostly tensile) to do (balanced-ish)
        dn_lo = 1.0
        dn_hi = max(L.d_from_comp_face for L in self.layers)
        # Ensure Nu(lo) < 0 and Nu(hi) > 0
        N_lo, _ = self.interaction_point(dn_lo)
        N_hi, _ = self.interaction_point(dn_hi)
        if N_lo > 0:
            # Section might be heavily reinforced; return decompression
            # point as closest we can get — user will see it in plot
            N, M = self.interaction_point(dn_lo)
            return N, M, dn_lo
        if N_hi < 0:
            # Section cannot reach positive N even at decompression —
            # extremely unusual; return tensile side
            N, M = self.interaction_point(dn_hi)
            return N, M, dn_hi
        # Bisection
        for _ in range(200):
            dn_mid = 0.5 * (dn_lo + dn_hi)
            N, M = self.interaction_point(dn_mid)
            if abs(N) < tol:
                return N, M, dn_mid
            if N < 0:
                dn_lo = dn_mid
            else:
                dn_hi = dn_mid
        N, M = self.interaction_point(dn_mid)
        return N, M, dn_mid

    # ----- Full interaction diagram ---------------------------------------
    def interaction_diagram(self, n_points: int = 60) -> List[Tuple[float, float]]:
        """Generate a list of (Nu, Mu) points from squash to pure
        tension. The A–B segment (neutral axis outside section) uses
        linear interpolation per Cl 10.6.2.4 between squash (A) and
        decompression (B). Below D, also linear interpolation to E.
        """
        # Squash (A)
        A = (self.nuo(), 0.0)
        # Decompression (B)
        B_N, B_M = self.decompression_point()
        # Pure bending (D)
        D_N, D_M, _ = self.pure_bending()
        # Pure tension (E)
        E = (-self.nuo_t(), 0.0)

        pts: List[Tuple[float, float]] = []
        # Segment A → B: linear (Cl 10.6.2.4)
        for i in range(6):
            t = i / 5.0
            pts.append((A[0] * (1 - t) + B_N * t, A[1] * (1 - t) + B_M * t))
        # Segment B → D: iterate dn from do down
        do = max(L.d_from_comp_face for L in self.layers)
        _, _, dn_pb = self.pure_bending()
        k_steps = n_points - 10
        for i in range(1, k_steps):
            t = i / k_steps
            dn = do * (1 - t) + dn_pb * t
            pts.append(self.interaction_point(dn))
        # Segment D → E: linear
        for i in range(1, 5):
            t = i / 4.0
            pts.append((D_N * (1 - t) + E[0] * t, D_M * (1 - t) + E[1] * t))
        return pts


# =============================================================================
# CAPACITY REDUCTION FACTOR φ (AS 3600 Table 2.2.2)
# =============================================================================

def phi_bending_only(kuo: float, class_N: bool = True) -> float:
    """φ for pure-bending point (Item (b)(i) of Table 2.2.2).

    φ' = 1.24 − 13·kuo/12, bounded 0.65 ≤ φ ≤ 0.85 for Class N.
    """
    if not class_N:
        return 0.65
    p = 1.24 - 13.0 * kuo / 12.0
    return max(0.65, min(0.85, p))


def phi_compression(Nu: float, Nub: float, phi_o: float,
                    phi_prime: float) -> float:
    """φ for bending + compression (Item (d) of Table 2.2.2).

    - If Nu ≥ Nub: φ = φo
    - If Nu < Nub: φ = φo + (φ' − φo)(1 − Nu/Nub)

    φo is 0.65 for stocky short columns with Q/G ≥ 0.25, else 0.60.
    φ' is the pure-bending value from phi_bending_only.
    """
    if Nu >= Nub:
        return phi_o
    ratio = 1.0 - (Nu / Nub)
    return phi_o + (phi_prime - phi_o) * ratio


def phi_tension(Nu: float, Nuot: float, phi_prime: float) -> float:
    """φ for bending + axial tension (Item (c) of Table 2.2.2).

    φ = φ' + (0.85 − φ')(Nu/Nuot)  [Nu is tension, treat magnitude]
    """
    t = abs(Nu) / Nuot
    t = max(0.0, min(1.0, t))
    return phi_prime + (0.85 - phi_prime) * t


# =============================================================================
# PRO ── Cl 10.7.5 splice provisions (top & bottom of column)
# =============================================================================

def check_splice(splice_type: str, lap_ratio: float = 1.0,
                 face_in_tension_ever: bool = False) -> dict:
    """Cl 10.7.5 splice check — same logic as the JS engine."""
    if splice_type == 'none':
        return dict(ok=True, mode='none', capacityRatio=float('inf'),
                    note='no splice')
    if splice_type == 'mech':
        return dict(ok=True, mode='mechanical', capacityRatio=1.0,
                    note='rated >= fsy')
    if splice_type == 'endb':
        if face_in_tension_ever:
            return dict(ok=False, mode='end-bearing', capacityRatio=0,
                        note='face in tension under at least one combo — not permitted (Cl 10.7.5.4)')
        return dict(ok=True, mode='end-bearing', capacityRatio=0,
                    note='permitted: face in compression for all combos')
    # lap
    required = 0.25
    ok = lap_ratio >= required
    return dict(ok=ok, mode='lap', capacityRatio=lap_ratio,
                note=f"develops {lap_ratio*100:.0f}% of fsy ({'≥' if ok else '<'} 25% required)")


# =============================================================================
# PRO ── Cl 10.7.3.1 / Fig 5.31 — special confinement region length
# =============================================================================

def special_region_length(M_high: float, M_other: float, L: float, D: float,
                          M_thresh: float) -> float:
    """Distance from the high-moment end where |M(x)| drops to M_thresh.

    M_high, M_other are SIGNED (BMD convention); L is column length, D the
    column dimension perp. to the bending axis being checked. Lower-bounded by
    1.2·D per RCB Fig 5.31. Validated against RCB Example 5.7 (Mt=+225, Mb=-190
    → 586 mm).
    """
    Mhi_abs = abs(M_high)
    if Mhi_abs <= M_thresh:
        return max(1.2 * D, 0)
    slope = (M_other - M_high) / L
    if abs(slope) < 1e-9:
        return max(L, 1.2 * D)
    target = (1 if M_high >= 0 else -1) * M_thresh
    x = (target - M_high) / slope
    if x < 0:
        return max(1.2 * D, 0)
    return max(x, 1.2 * D)


# =============================================================================
# PRO ── Cl 10.7.4.1 — lateral restraint pattern
# =============================================================================

def restraint_pattern_required(s: float, N_star: float, fc: float, Ag: float) -> str:
    """Returns 'every-bar' or 'alternate-bar' per Cl 10.7.4.1."""
    if s > 150:
        return 'every-bar'
    if N_star > 0.3 * fc * Ag:
        return 'every-bar'
    return 'alternate-bar'


# =============================================================================
# PRO ── Cl 10.6.3 — biaxial concession (skip αn check)
# =============================================================================

def biaxial_concession(b: float, D: float, Mx_ratio: float, My_ratio: float) -> dict:
    """Returns {applies, reason}."""
    if b <= 0 or D <= 0:
        return dict(applies=False, reason='invalid section')
    aspect = max(b, D) / min(b, D)
    if aspect > 3:
        return dict(applies=False, reason=f'aspect ratio {aspect:.2f} > 3')
    minRatio = min(Mx_ratio, My_ratio)
    if minRatio <= 0.06:
        return dict(applies=True, reason=f'aspect {aspect:.2f}; min M*/φMu = {minRatio:.3f} ≤ 0.06')
    return dict(applies=False, reason=f'aspect {aspect:.2f} OK but both M*/φMu > 0.06')


# =============================================================================
# PRO ── Cl 10.8 — floor-joint transmission of axial force
# =============================================================================

def joint_transmission_check(fc_col: float, fc_slab: float, h: float, c: float,
                             restraint: str = '4-sides') -> dict:
    """Returns dict with applies, fce, ratio, note."""
    if fc_slab >= 0.75 * fc_col:
        return dict(applies=False, fce=fc_col, ratio=1.0,
                    note=f'slab fc′={fc_slab} ≥ 0.75·col fc′={0.75*fc_col:.0f} — no check required')
    dh = h / (2 * c) if c > 0 else 0
    if restraint == '4-sides':
        fce = min(fc_col, 0.75 * fc_col + dh * (fc_col - fc_slab))
    elif restraint == '2-opposite':
        fce = min(fc_col, 0.85 * fc_col + dh * (fc_col - fc_slab))
    else:
        fce = 0.65 * fc_col
    ratio = fce / fc_col
    return dict(applies=True, fce=fce, ratio=ratio,
                note=f'fce = {fce:.1f} MPa ({ratio*100:.0f}% of col fc′)')


# =============================================================================
# SLENDERNESS (AS 3600 Cl 10.3.1, 10.5)
# =============================================================================

def radius_of_gyration(shape: str, D: float) -> float:
    """Cl 10.5.2: r = 0.3D (rectangular) or 0.25D (circular)."""
    return 0.25 * D if shape == "circ" else 0.3 * D


def is_short_column(
    Le_over_r: float,
    braced: bool,
    N_star: float,
    Nuo: float,
    M1_over_M2: float,  # NEGATIVE single curvature, POSITIVE double
    fc: float,
) -> Tuple[bool, float]:
    """Cl 10.3.1 — returns (is_short, limit)."""
    if braced:
        ratio = N_star / Nuo
        if ratio >= 0.15:
            alpha_c = math.sqrt(2.25 - 2.5 * N_star / (0.65 * Nuo))
        else:
            alpha_c = math.sqrt(1.0 / (3.5 * N_star / (0.65 * Nuo))) if N_star > 0 else 10.0
        limit2 = alpha_c * (38.0 - fc / 15.0) * (1.0 + M1_over_M2)
        limit = max(25.0, limit2)
    else:
        limit = 22.0
    return (Le_over_r <= limit, limit)


# =============================================================================
# EFFECTIVE LENGTH (AS 3600 Cl 10.5.3) — simplified lookup
# =============================================================================

def effective_length_factor_braced(gamma_1: float, gamma_2: float) -> float:
    """Approximate closed-form for Fig 10.5.3(B) (braced).

    Uses the Wood (1974)/Galambos (1968) alignment-chart fit — for
    braced columns:
      k = [1 + 0.145(γ1 + γ2) − 0.265·γ1·γ2] /
          [2 − 0.364(γ1 + γ2) − 0.247·γ1·γ2]
    Bounded 0.5 ≤ k ≤ 1.0 (braced column theoretical range).
    """
    a = gamma_1
    b = gamma_2
    num = 1.0 + 0.145 * (a + b) - 0.265 * a * b
    den = 2.0 - 0.364 * (a + b) - 0.247 * a * b
    k = num / den if den != 0 else 1.0
    return max(0.5, min(1.0, k))


def effective_length_factor_unbraced(gamma_1: float, gamma_2: float) -> float:
    """Approximate closed-form for Fig 10.5.3(C) (unbraced/sway).

    Wood formula:
      k = sqrt((1.6·γ1·γ2 + 4·(γ1 + γ2) + 7.5) /
              (γ1 + γ2 + 7.5))
    Always ≥ 1.0. For very large γ (near free), k approaches ∞;
    we cap at 5.0 for practical display.
    """
    a = min(gamma_1, 50.0)
    b = min(gamma_2, 50.0)
    num = 1.6 * a * b + 4.0 * (a + b) + 7.5
    den = a + b + 7.5
    k = math.sqrt(num / den) if den > 0 else 5.0
    return max(1.0, min(5.0, k))


# =============================================================================
# MOMENT MAGNIFIER (AS 3600 Cl 10.4)
# =============================================================================

def buckling_load_Nc(
    section: Section,
    Le: float,
    beta_d: float,
) -> float:
    """Cl 10.4.4: Nc = (π²/Le²) · 182·do·φMub / (1 + βd)
    with φ = 0.65 hard-wired, Mub at ku = 0.545.
    """
    do = max(L.d_from_comp_face for L in section.layers)
    _, Mub, _ = section.balanced_point()
    phi_Mub = 0.65 * Mub  # Nmm
    # do is in mm; Le in mm; phi_Mub in Nmm; Nc in N
    Nc = (math.pi ** 2 / (Le ** 2)) * (182.0 * do * phi_Mub / (1.0 + beta_d))
    return Nc


def moment_magnifier_braced(
    km: float,
    N_star: float,
    Nc: float,
) -> float:
    """Cl 10.4.2: δb = km / (1 − N*/Nc) ≥ 1.0."""
    if Nc <= 0:
        return 1.0
    den = 1.0 - N_star / Nc
    if den <= 0:
        return float("inf")
    return max(1.0, km / den)


def km_factor(M1: float, M2: float, transverse_loading: bool = False) -> float:
    """Cl 10.4.2: km = 0.6 − 0.4·(M1/M2) ≥ 0.4.
    If significant transverse loading, km = 1.0.
    """
    if transverse_loading:
        return 1.0
    if M2 == 0:
        return 1.0
    # M1/M2 signed: negative single-curvature, positive double
    ratio = M1 / M2
    km = 0.6 - 0.4 * ratio
    return max(0.4, km)


def moment_magnifier_unbraced_storey(
    sum_N_star: float,
    sum_Nc: float,
) -> float:
    """Cl 10.4.3(b): δs = 1 / (1 − ΣN*/ΣNc)."""
    if sum_Nc <= 0:
        return 1.0
    den = 1.0 - sum_N_star / sum_Nc
    if den <= 0:
        return float("inf")
    return max(1.0, 1.0 / den)


# =============================================================================
# BIAXIAL BENDING (AS 3600 Cl 10.6.4)
# =============================================================================

def biaxial_utilisation(
    Mx_star: float, My_star: float,
    phi_Mux: float, phi_Muy: float,
    N_star: float, Nuo: float,
) -> Tuple[float, float]:
    """Return (utilisation, alpha_n).

    (M*x / φMux)^αn + (M*y / φMuy)^αn ≤ 1.0
    αn = 0.7 + 1.7·N*/Nuo, bounded [1, 2]
    """
    alpha_n = 0.7 + 1.7 * N_star / Nuo
    alpha_n = max(1.0, min(2.0, alpha_n))
    if phi_Mux <= 0 or phi_Muy <= 0:
        return float("inf"), alpha_n
    util = (abs(Mx_star) / phi_Mux) ** alpha_n + (abs(My_star) / phi_Muy) ** alpha_n
    return util, alpha_n


# =============================================================================
# HSC CONFINEMENT (AS 3600 Cl 10.7.3) — only required for fc > 50 MPa
# =============================================================================

@dataclass
class HSCConfinementInput:
    bc: float              # core width between outermost fitments (mm)
    dc: float              # core depth between outermost fitments (mm)
    n: int                 # number of laterally restrained longitudinal bars
    w: float               # average clear spacing between restrained bars (mm)
    Ab_fit: float          # one fitment leg area (mm^2)
    fsy_f: float           # fitment yield stress (MPa)
    fc: float              # concrete strength (MPa)
    m: int                 # fitment legs crossing the confinement plane
    shape: str = "rect"    # "rect" or "circ"
    ds: float = 0.0        # core diameter / side dim for ke (mm)
    theta_deg: float = 90.0  # fitment leg angle to confinement plane (°)


def hsc_fitment_spacing_simplified(inp: HSCConfinementInput) -> float:
    """Cl 10.7.3.3 — returns maximum fitment spacing s (mm) such that
    fr.eff ≥ 0.01·fc.

    fr.eff = ke · fr   where  fr = (m·Ab·fsy·sinθ)/(ds·s)
    For rect:  ke = (1 − n·w²/(6·Ac)) · (1 − s/(2·bc))·(1 − s/(2·dc))
               ≈ ke(s)  — solved iteratively for s.
    """
    # fr term coefficient (independent of s): fr·s = K1
    theta = math.radians(inp.theta_deg)
    K1 = inp.m * inp.Ab_fit * inp.fsy_f * math.sin(theta) / inp.ds
    # Required fr.eff
    fr_eff_req = 0.01 * inp.fc
    if inp.shape == "circ":
        # ke = (1 − s/(2·ds))²
        # fr.eff = (K1/s) · (1 − s/(2·ds))²   ≥ fr_eff_req
        # Bisection on s
        s_lo, s_hi = 1.0, 2.0 * inp.ds - 0.1
        def lhs(s):
            return (K1 / s) * (1 - s / (2 * inp.ds)) ** 2 - fr_eff_req
        # lhs is large at s_lo, small at s_hi — find root
        if lhs(s_lo) <= 0:
            return 0.0
        if lhs(s_hi) > 0:
            return s_hi
        for _ in range(100):
            sm = 0.5 * (s_lo + s_hi)
            v = lhs(sm)
            if abs(v) < 1e-6:
                return sm
            if v > 0:
                s_lo = sm
            else:
                s_hi = sm
        return 0.5 * (s_lo + s_hi)
    # Rectangular
    Ac = inp.bc * inp.dc
    k_base = 1 - inp.n * inp.w ** 2 / (6 * Ac)
    def lhs(s):
        ke_s = k_base * (1 - s / (2 * inp.bc)) * (1 - s / (2 * inp.dc))
        fr = K1 / s
        return ke_s * fr - fr_eff_req
    s_lo, s_hi = 1.0, min(inp.bc, inp.dc) - 0.1
    if lhs(s_lo) <= 0:
        return 0.0
    if lhs(s_hi) > 0:
        return s_hi
    for _ in range(100):
        sm = 0.5 * (s_lo + s_hi)
        v = lhs(sm)
        if abs(v) < 1e-6:
            return sm
        if v > 0:
            s_lo = sm
        else:
            s_hi = sm
    return 0.5 * (s_lo + s_hi)


def hsc_fitment_spacing_deemed(inp: HSCConfinementInput) -> float:
    """Cl 10.7.3.4 — deemed-to-comply spacing.

    Rect:  s ≤ 15·n·Ab·fsy.f / (fc·bc)
    Circ:  s ≤ 100·Ab·fsy.f / (fc·ds)
    """
    if inp.shape == "circ":
        return 100.0 * inp.Ab_fit * inp.fsy_f / (inp.fc * inp.ds)
    return 15.0 * inp.n * inp.Ab_fit * inp.fsy_f / (inp.fc * inp.bc)


# =============================================================================
# LOAD COMBINATIONS (AS/NZS 1170.0)
# =============================================================================

def load_combinations_uls(
    G: float, Q: float, Wu: float = 0.0, Eu: float = 0.0,
    psi_c: float = 0.4, psi_l: float = 0.4,
) -> List[Tuple[str, float]]:
    """Return standard ULS combos as (name, factor) pairs applied to
    (N, M) action effects.

    Assumes each of G, Q, Wu, Eu is the action effect (axial or moment)
    due to that load alone; the combinations factor them directly.

    AS/NZS 1170.0 Cl 4.2:
      1.35G
      1.2G + 1.5Q
      1.2G + 1.5·ψl·Q          (long-term)
      1.2G + Wu + ψc·Q         (wind primary)
      0.9G + Wu                (overturning)
      G + Eu + ψc·Q            (earthquake)
    """
    return [
        ("1.35G",               1.35 * G),
        ("1.2G + 1.5Q",         1.2 * G + 1.5 * Q),
        ("1.2G + 1.5·ψl·Q",     1.2 * G + 1.5 * psi_l * Q),
        ("1.2G + Wu + ψc·Q",    1.2 * G + Wu + psi_c * Q),
        ("0.9G + Wu",           0.9 * G + Wu),
        ("G + Eu + ψc·Q",       G + Eu + psi_c * Q),
    ]


# =============================================================================
# HIGH-LEVEL DESIGN CHECK
# =============================================================================

@dataclass
class DesignResult:
    # Capacity envelope
    Nuo: float
    balanced_Nub: float
    balanced_Mub: float
    kub: float
    # Reductions
    phi_o: float
    phi_prime: float
    # Interaction point at N*
    phi_Mu_at_Nstar: float
    # Check
    utilisation: float
    passes: bool
    # Reference values
    governing_combo: str = ""
    notes: List[str] = field(default_factory=list)


def capacity_at_N_star(
    section: Section,
    N_star: float,
    stocky: bool = True,
    Q_over_G: float = 1.0,
) -> Tuple[float, float, float]:
    """Find φMu at N*.

    Returns (Mu, phi, phi_Mu).
    Scans the interaction diagram (segment B–D, where the rectangular
    stress block applies) for the dn that yields Nu = N*.
    For N* above Nub we're in the compression-failure region;
    for N* below Nub we're in the tension-failure region.
    """
    phi_prime = phi_bending_only(0.2)  # approximate — refined when we know kuo
    phi_o = 0.65 if (stocky and Q_over_G >= 0.25) else 0.60
    Nub, Mub, kub = section.balanced_point()

    # Bisect dn to find Nu ≈ N*
    do = max(L.d_from_comp_face for L in section.layers)
    Nuo = section.nuo()
    _, _, dn_pb = section.pure_bending()

    if N_star >= Nuo:
        return 0.0, phi_o, 0.0
    if N_star >= Nub:
        # Segment B–C: dn between do and dn_balanced
        dn_balanced = kub * do
        dn_lo, dn_hi = dn_balanced, do
    else:
        # Segment C–D
        dn_lo, dn_hi = dn_pb, kub * do
    # Ensure monotonicity
    N_lo, _ = section.interaction_point(dn_lo)
    N_hi, _ = section.interaction_point(dn_hi)
    if (N_star - N_lo) * (N_star - N_hi) > 0:
        # Outside this segment; clamp to nearest
        if abs(N_star - N_lo) < abs(N_star - N_hi):
            dn = dn_lo
        else:
            dn = dn_hi
    else:
        for _ in range(100):
            dn_mid = 0.5 * (dn_lo + dn_hi)
            N_mid, _ = section.interaction_point(dn_mid)
            if abs(N_mid - N_star) < 1.0:
                break
            # Find which side
            if (N_mid - N_star) * (N_lo - N_star) < 0:
                dn_hi = dn_mid
                N_hi = N_mid
            else:
                dn_lo = dn_mid
                N_lo = N_mid
        dn = 0.5 * (dn_lo + dn_hi)
    Nu, Mu = section.interaction_point(dn)

    # Refine phi_prime using actual kuo from pure bending
    kuo = dn_pb / do
    phi_prime = phi_bending_only(kuo)

    phi = phi_compression(Nu, Nub, phi_o, phi_prime)
    return Mu, phi, phi * Mu


# =============================================================================
# END
# =============================================================================
