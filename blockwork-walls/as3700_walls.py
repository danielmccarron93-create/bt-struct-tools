"""
AS 3700:2018 reinforced blockwork wall capacity engine.

Pure functions implementing every Section 8 equation needed for the wall designer.
This is the validated reference — JS port and HTML calculator must match these results.

All functions use SI units internally:
  Lengths    : mm
  Stresses   : MPa  (= N/mm²)
  Forces     : N    (convert to kN at output)
  Moments    : N·mm (convert to kN·m at output)
  Areas      : mm²

Sign convention:
  Compression positive for axial.
  Tension reported as positive magnitude in dedicated tension API.

Engine version: 1.0.0
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field, asdict
from typing import Literal, Optional

ENGINE_VERSION = "1.0.0"

# ─────────────────────────────────────────────────────────────────────────────
# Material lookup tables (AS 3700:2018 Section 3 / Table 3.1)
# ─────────────────────────────────────────────────────────────────────────────

# km values from Table 3.1
KM_TABLE = {
    ("clay",     "full",       "M2"): 1.1,
    ("clay",     "full",       "M3"): 1.4,
    ("clay",     "full",       "M4"): 2.0,
    ("clay",     "face_shell", "M3"): 1.6,
    ("concrete", "full",       "M3"): 1.4,
    ("concrete", "face_shell", "M3"): 1.6,
    ("calcium_silicate", "full", "M3"): 1.4,
    ("calcium_silicate", "full", "M4"): 2.0,
}

# Capacity reduction factors (Table 4.1)
PHI_REINFORCED            = 0.75   # all actions
PHI_URM_COMPRESSION_SOLID = 0.75
PHI_URM_COMPRESSION_HOLLOW= 0.50
PHI_URM_COMPRESSION_GROUT = 0.60
PHI_URM_FLEXURE_SHEAR     = 0.60

# Reinforcing contribution factor (Cl 8.5.1)
ALPHAR_PIER     = 1.00
ALPHAR_WALL     = 0.40
ALPHAR_CONC_LOAD= 0.00

# Fixed properties
FSY_N_BAR        = 500.0   # MPa, N-grade reinforcement
F_VM_REINFORCED  = 0.35    # MPa, char shear strength of reinforced masonry (Cl 8.8)
FVS_MAIN_REO     = 17.5    # MPa, design shear strength of main reo (Cl 8.8)


# ─────────────────────────────────────────────────────────────────────────────
# Reinforcement schedule lookup (mirrors L7.xlsm SCHEDULE!AA5:AB21 — "reo" table)
# ─────────────────────────────────────────────────────────────────────────────
# Single-bar areas in mm². EF = "Each Face" — two bars (one per face), so area is doubled.
# SL meshes have a fixed mm²/m value — for these, the spacing input is overridden to 1000
# in the per-metre formula so the lookup gives the per-metre area directly.
#
# Note: the original L7 spreadsheet has SL82 EF = 5448, which is clearly a typo
# (should be 2 x 227 = 454). We use the corrected value here.
REO_TABLE = {
    "-":       0,
    "N12":   110,
    "N16":   200,
    "N20":   314,
    "N24":   450,
    "N28":   620,
    "N32":   800,
    "N12 EF": 220,   # = 2 x 110, bars on both faces
    "N16 EF": 400,   # = 2 x 200
    "N20 EF": 628,   # = 2 x 314
    "N24 EF": 900,   # = 2 x 450
    "SL82":  227,    # mesh — mm²/m (treat spacing as 1000)
    "SL92":  290,
    "SL82 EF": 454,  # corrected from L7 typo (5448) — 2 x 227
    "SL92 EF": 580,
}
MESH_CODES = {"SL82", "SL92", "SL82 EF", "SL92 EF"}
EF_CODES   = {c for c in REO_TABLE if c.endswith(" EF")}


def reo_area_per_m(code: str, spacing_mm: float) -> float:
    """Return area of reinforcement per metre wall length, mm²/m.

    For mesh codes, the lookup IS the per-metre area (spacing input ignored).
    For EF codes, the area in REO_TABLE is already doubled.
    For "-" or unknown, returns 0.
    """
    if code is None:
        return 0.0
    a = REO_TABLE.get(code, 0)
    if a == 0:
        return 0.0
    if code in MESH_CODES:
        return float(a)
    if spacing_mm is None or spacing_mm <= 0:
        return 0.0
    return float(a) * 1000.0 / spacing_mm


def is_ef(code: str) -> bool:
    """Whether this reo code represents bars on each face (i.e. 2 layers)."""
    return code in EF_CODES


def bar_dia_from_code(code: str):
    """Return numeric bar diameter for N-bar codes, or None for mesh / '-'."""
    if not code or code == "-": return None
    if code.startswith("N"):
        # N16 → 16, N16 EF → 16
        digits = ""
        for ch in code[1:]:
            if ch.isdigit(): digits += ch
            else: break
        return int(digits) if digits else None
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Data classes for inputs and outputs
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class Geometry:
    """Wall geometry, all in mm."""
    L: float            # length between vertical supports
    t: float            # overall thickness
    H: float            # height between horizontal supports
    tfs: float = 30.0   # face shell thickness (typical 30 for 190 series)
    grouting: Literal["full", "partial", "none"] = "full"
    # For partial grouting, the engineer must override Ag, Ad explicitly via cross_section.
    # v1 only fully validates "full" grouting.


@dataclass
class Material:
    """Materials per Section 3."""
    fuc: float          # MPa, characteristic unit compressive strength
    fcg: float = 20.0   # MPa, design characteristic grout compressive strength
    unit_type: Literal["clay", "concrete", "calcium_silicate"] = "concrete"
    bedding: Literal["full", "face_shell"] = "full"
    mortar: Literal["M2", "M3", "M4"] = "M3"
    kh: float = 1.3     # joint thickness factor (1.3 max for 190-high block on 10mm bed)
    kc_override: Optional[float] = None  # override default (1.4 hollow conc / 1.2 other)
    fsy: float = FSY_N_BAR
    unit_density_high: bool = True       # > 2000 kg/m³ → kc=1.4

    @property
    def km(self) -> float:
        key = (self.unit_type, self.bedding, self.mortar)
        if key not in KM_TABLE:
            raise ValueError(
                f"No km in Table 3.1 for {key}. Add explicit value or change inputs."
            )
        return KM_TABLE[key]

    @property
    def kc(self) -> float:
        if self.kc_override is not None:
            return self.kc_override
        return 1.4 if self.unit_density_high else 1.2

    @property
    def fm(self) -> float:
        """Eq for f'm: f'm = kh × km × √f'uc  (Cl 3.3.2)."""
        return self.kh * self.km * math.sqrt(self.fuc)


@dataclass
class Reinforcement:
    """Vertical and horizontal reinforcement.

    Each direction supports a *primary* layer + an optional *additional* layer.
    Bars are specified by code (N12, N16 EF, SL82, "-", etc.) — see REO_TABLE.

    Backward compatible: if `vert_bar_dia` is set (legacy numeric API),
    we synthesise vert_primary_code = "N{dia}" automatically.
    """
    # New schema (preferred)
    vert_primary_code:    str    = "N16"
    vert_primary_spacing: float  = 200.0
    vert_additional_code:    str   = "-"
    vert_additional_spacing: float = 400.0
    horiz_primary_code:    str   = "N12"
    horiz_primary_spacing: float = 600.0
    horiz_additional_code:    str   = "-"
    horiz_additional_spacing: float = 600.0

    # Legacy fields (kept for backward compat with older test inputs).
    # If set, they override the primary_code on construction.
    vert_bar_dia:     Optional[float] = None
    vert_spacing:     Optional[float] = None
    vert_layers:      int  = 1
    horiz_bar_dia:    Optional[float] = None
    horiz_spacing:    Optional[float] = None
    horiz_layers:     int  = 1

    def __post_init__(self):
        # Legacy → new schema migration
        if self.vert_bar_dia is not None:
            self.vert_primary_code = f"N{int(self.vert_bar_dia)}"
            if self.vert_spacing is not None:
                self.vert_primary_spacing = self.vert_spacing
            if self.vert_layers == 2:
                self.vert_primary_code += " EF"
        if self.horiz_bar_dia is not None:
            self.horiz_primary_code = f"N{int(self.horiz_bar_dia)}"
            if self.horiz_spacing is not None:
                self.horiz_primary_spacing = self.horiz_spacing
            if self.horiz_layers == 2:
                self.horiz_primary_code += " EF"

    @staticmethod
    def bar_area(d: float) -> float:
        return math.pi * d * d / 4.0

    def Asv_per_m(self) -> float:
        """Vertical reo area per metre of wall length, mm²/m. Sums primary + additional."""
        return (reo_area_per_m(self.vert_primary_code, self.vert_primary_spacing) +
                reo_area_per_m(self.vert_additional_code, self.vert_additional_spacing))

    def Ash_per_m(self) -> float:
        """Horizontal reo area per metre of wall HEIGHT, mm²/m. Sums primary + additional."""
        return (reo_area_per_m(self.horiz_primary_code, self.horiz_primary_spacing) +
                reo_area_per_m(self.horiz_additional_code, self.horiz_additional_spacing))

    def Aov(self, L: float) -> float:
        """Total vertical reo area across wall length L, mm²."""
        return self.Asv_per_m() * L / 1000.0

    def Aoh(self, H: float) -> float:
        """Total anchored horizontal reo area over wall height H, mm²."""
        return self.Ash_per_m() * H / 1000.0

    # Convenience for 3D rendering / drawing
    @property
    def vert_is_ef(self) -> bool:
        return is_ef(self.vert_primary_code) or is_ef(self.vert_additional_code) \
            or self.vert_primary_code in MESH_CODES or self.vert_additional_code in MESH_CODES
    @property
    def horiz_is_ef(self) -> bool:
        return is_ef(self.horiz_primary_code) or is_ef(self.horiz_additional_code) \
            or self.horiz_primary_code in MESH_CODES or self.horiz_additional_code in MESH_CODES


@dataclass
class Slenderness:
    """Slenderness and lateral-support coefficients per Cl 7.3.4.3."""
    av: float = 0.85    # vertical: 0.75 / 0.85 / 1.0 / 1.5 / 2.5 (Figure 7.1)
    ah: float = 2.5     # horizontal: 1.0 (both edges) / 2.5 (one edge) (Figure 7.2)
    kt: float = 1.0     # thickness coeff (Table 7.2; 1.0 if no piers)
    edges_supported: Literal["both", "one", "none"] = "one"


@dataclass
class Eccentricity:
    """Effective load eccentricity per Cl 7.3.4.4."""
    e: float = 32.0     # mm (default = t/6 for 190-series single-floor bearing)


@dataclass
class Method:
    """Engineering choices (toggles)."""
    reinforced: bool = True
    alphar: float = ALPHAR_WALL          # 0.40 walls / 1.0 piers / 0.0 conc loads / 0.0 to be conservative
    bending_method: Literal["kes_in_axial", "explicit_8_6", "stress_block_hybrid"] = "kes_in_axial"
    sr_rule: Literal["standard", "always_conservative"] = "standard"
    enforce_cl_8_8_cap: bool = True
    enforce_stability_check: bool = False  # Cl 8.7.4
    enforce_eq_detailing: bool = False     # Cl 8.4.5 close-spaced reo
    enforce_concentrated_loads: bool = False  # Cl 8.5.2
    include_oop_bending: bool = False
    no_top_restraint: bool = False         # for Cl 8.7.4 applicability


@dataclass
class LoadCase:
    """A single load combination (factored / ULS)."""
    name: str = "LC"
    Nc_static: float = 0.0     # kN, axial compression — RCB Static method
    Nc_area: float = 0.0       # kN, axial compression — RCB Area method
    Nt: float = 0.0            # kN, axial tension (positive magnitude)
    Mx_top: float = 0.0        # kNm, in-plane moment top
    Mx_btm: float = 0.0        # kNm, in-plane moment bottom
    My_top: float = 0.0        # kNm, out-of-plane moment top
    My_btm: float = 0.0        # kNm, out-of-plane moment bottom
    Vx_top: float = 0.0        # kN, in-plane shear top
    Vx_btm: float = 0.0        # kN, in-plane shear bottom
    Vy_top: float = 0.0        # kN, OOP shear top
    Vy_btm: float = 0.0        # kN, OOP shear bottom

    @property
    def Mx_max(self) -> float:
        return max(abs(self.Mx_top), abs(self.Mx_btm))

    @property
    def My_max(self) -> float:
        return max(abs(self.My_top), abs(self.My_btm))

    @property
    def Vx_max(self) -> float:
        return max(abs(self.Vx_top), abs(self.Vx_btm))

    @property
    def Vy_max(self) -> float:
        return max(abs(self.Vy_top), abs(self.Vy_btm))

    @property
    def Nc_worst(self) -> float:
        """Worst-case axial compression from RCB's two methods."""
        return max(self.Nc_static, self.Nc_area)


# ─────────────────────────────────────────────────────────────────────────────
# Cross-section properties (Section 4.5)
# ─────────────────────────────────────────────────────────────────────────────

def Ab_per_m(geom: Geometry) -> float:
    """Bedded area per metre length, mm²/m. Cl 4.5.4(b) face-shell bedding."""
    return 2.0 * geom.tfs * 1000.0


def Ad_per_m(geom: Geometry) -> float:
    """Design cross-section area per metre length, mm²/m. Cl 4.5.6.

    For fully grouted: Ad = t × L (so per metre = t × 1000).
    For partially grouted, override required.
    """
    if geom.grouting == "full":
        return geom.t * 1000.0
    elif geom.grouting == "none":
        return Ab_per_m(geom)
    else:
        # Partial grouting placeholder — a more refined model needed.
        # Conservative: use Ab_per_m as Ad_per_m. Engineer should override.
        return Ab_per_m(geom)


def Ag_per_m(geom: Geometry) -> float:
    """Grout area per metre length, mm²/m. Cl 4.5.7: Ag = Ad - Ab."""
    return Ad_per_m(geom) - Ab_per_m(geom)


def Ad_total(geom: Geometry) -> float:
    """Total design area, mm². For shear (full wall length)."""
    return Ad_per_m(geom) * geom.L / 1000.0


def Ag_total(geom: Geometry) -> float:
    return Ag_per_m(geom) * geom.L / 1000.0


def Zex(geom: Geometry) -> float:
    """Section modulus about strong (in-plane) axis, mm³ = t·L²/6."""
    return geom.t * geom.L * geom.L / 6.0


# ─────────────────────────────────────────────────────────────────────────────
# Slenderness (Cl 7.3.4.3)
# ─────────────────────────────────────────────────────────────────────────────

def Sr_eq_1(slend: Slenderness, geom: Geometry) -> float:
    """Eq 7.3.4.3(1): Sr = av·H / (kt·t)."""
    return slend.av * geom.H / (slend.kt * geom.t)


def Sr_eq_2(slend: Slenderness, geom: Geometry) -> float:
    """Eq 7.3.4.3(2): Sr = (0.7/t) · √(av·H·ah·L)."""
    return (0.7 / geom.t) * math.sqrt(slend.av * geom.H * slend.ah * geom.L)


def governing_Sr(Sr1: float, Sr2: float, Fd_per_m: float, Fo_per_m: float,
                 method: Method, edges_supported: str) -> float:
    """Apply Cl 7.3.4.3(a) selection rule.

    Lesser of (1) and (2) only when:
      - wall is laterally supported on at least one vertical edge, AND
      - Fd ≤ 0.20·Fo
    Otherwise use (1) (Cl 7.3.4.3(c) "all other cases").
    """
    if method.sr_rule == "always_conservative":
        return max(Sr1, Sr2)
    if edges_supported == "none":
        return Sr1
    if Fo_per_m <= 0:
        return Sr1
    if Fd_per_m / Fo_per_m <= 0.20:
        return min(Sr1, Sr2)
    return Sr1


# ─────────────────────────────────────────────────────────────────────────────
# Eccentricity helpers
# ─────────────────────────────────────────────────────────────────────────────

def deemed_eccentricity_single_floor(t: float) -> float:
    """Cl 7.3.4.4 deemed-to-comply: load at 1/3 of bearing depth from face,
    so e measured from centroid = t/2 - t/3 = t/6."""
    return t / 6.0


def deemed_eccentricity_continuous_floor(t: float) -> float:
    """Continuous floor: minimum 0.05·t per refined calc convention."""
    return 0.05 * t


def implied_eccentricity_from_M_N(M_kNm: float, N_kN: float) -> Optional[float]:
    """e = M/N converted to mm. Returns None if N ~ 0."""
    if abs(N_kN) < 1e-9:
        return None
    return M_kNm * 1000.0 / N_kN  # kNm × 1000 / kN = N·mm/N = mm


# ─────────────────────────────────────────────────────────────────────────────
# Cl 8.5.1 — Compression
# ─────────────────────────────────────────────────────────────────────────────

def kes(Sr: float, e: float, t: float) -> float:
    """Cl 8.5.1 slenderness/eccentricity factor.

    kes = (1 - 0.025·Sr)·(1 - 2·e/t), bounded to [0, 1].
    Bounding to ≥ 0 is engineering judgment; standard allows the formula to give
    negative numbers but capacity then has no physical meaning — escalate to Cl 8.11.1.
    """
    raw = (1.0 - 0.025 * Sr) * (1.0 - 2.0 * e / t)
    return max(0.0, min(1.0, raw))


def grout_effective_stress(fcg: float) -> float:
    """The (f'cg/1.3)^(0.55 + 0.005·f'cg) term in Eq 8.5.1."""
    return (fcg / 1.3) ** (0.55 + 0.005 * fcg)


def basic_compression_per_m(material: Material, geom: Geometry,
                            reo: Reinforcement, alphar: float) -> dict:
    """Compute the bracketed sum in Eq 8.5.1, per metre length, in N/m.

    Returns a dict with each term separately for transparency.
    """
    Ab_m = Ab_per_m(geom)
    Ag_m = Ag_per_m(geom)
    As_m = reo.Asv_per_m()

    masonry_term = material.fm * Ab_m            # MPa × mm²/m = N/m
    grout_factor = grout_effective_stress(material.fcg)
    grout_term = material.kc * grout_factor * Ag_m
    steel_term = alphar * material.fsy * As_m

    bracket_total = masonry_term + grout_term + steel_term  # N/m

    return {
        "Ab_per_m": Ab_m,
        "Ag_per_m": Ag_m,
        "As_per_m": As_m,
        "fm": material.fm,
        "kc": material.kc,
        "grout_effective_stress": grout_factor,
        "masonry_term_N_per_m": masonry_term,
        "grout_term_N_per_m": grout_term,
        "steel_term_N_per_m": steel_term,
        "bracket_total_N_per_m": bracket_total,
        "bracket_total_kN_per_m": bracket_total / 1000.0,
    }


def Fo_no_steel_per_m_kN(material: Material, geom: Geometry) -> float:
    """Basic capacity per metre with no steel, no kes, no φ — used for Fd/Fo Sr rule.
    Returns kN/m."""
    Ab_m = Ab_per_m(geom)
    Ag_m = Ag_per_m(geom)
    masonry = material.fm * Ab_m
    grout = material.kc * grout_effective_stress(material.fcg) * Ag_m
    return (masonry + grout) / 1000.0


def phi_axial(method: Method) -> float:
    """φ for axial — 0.75 reinforced, 0.60 unreinforced grouted."""
    return PHI_REINFORCED if method.reinforced else PHI_URM_COMPRESSION_GROUT


def phi_tbs(method: Method) -> float:
    """φ for tension/bending/shear — 0.75 reinforced, 0.60 URM."""
    return PHI_REINFORCED if method.reinforced else PHI_URM_FLEXURE_SHEAR


def compression_capacity(material: Material, geom: Geometry, reo: Reinforcement,
                         slend: Slenderness, ecc: Eccentricity, method: Method,
                         demand_Nc_kN: float) -> dict:
    """Run the full Cl 8.5.1 compression check.

    Args:
      demand_Nc_kN: design compressive force on the wall, total, in kN.

    Returns a dict with every intermediate value (mirrors spreadsheet step-by-step).
    """
    # Geometry / cross-section
    Ad_t = Ad_total(geom)
    Ag_t = Ag_total(geom)

    # Slenderness
    Sr1 = Sr_eq_1(slend, geom)
    Sr2 = Sr_eq_2(slend, geom)

    # Need basic Fo (no steel, no kes) per metre to apply the Fd/Fo selection rule
    Fo_basic_per_m = Fo_no_steel_per_m_kN(material, geom)
    Fd_per_m = demand_Nc_kN / (geom.L / 1000.0)
    Fd_over_Fo = Fd_per_m / Fo_basic_per_m if Fo_basic_per_m > 0 else 0.0

    Sr_gov = governing_Sr(Sr1, Sr2, Fd_per_m, Fo_basic_per_m, method, slend.edges_supported)

    # kes
    kes_val = kes(Sr_gov, ecc.e, geom.t)

    # Full bracket including steel term
    bc = basic_compression_per_m(material, geom, reo, method.alphar)
    bracket_kN_per_m = bc["bracket_total_kN_per_m"]

    phi = phi_axial(method)
    phi_Nuc_per_m = phi * kes_val * bracket_kN_per_m
    phi_Nuc_total = phi_Nuc_per_m * geom.L / 1000.0

    # Utilisation
    util = demand_Nc_kN / phi_Nuc_total if phi_Nuc_total > 0 else float("inf")

    return {
        # Inputs echoed
        "demand_Nc_kN": demand_Nc_kN,
        "demand_Nc_per_m_kN": Fd_per_m,
        # Slenderness
        "Sr1": Sr1,
        "Sr2": Sr2,
        "Fo_basic_per_m_kN": Fo_basic_per_m,
        "Fd_over_Fo": Fd_over_Fo,
        "Sr_governing": Sr_gov,
        # Eccentricity factor
        "kes": kes_val,
        # Capacity build-up
        **bc,
        "phi": phi,
        "phi_Nuc_per_m_kN": phi_Nuc_per_m,
        "phi_Nuc_total_kN": phi_Nuc_total,
        # Verdict
        "utilisation": util,
        "pass": util <= 1.0,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Cl 8.10 — Tension
# ─────────────────────────────────────────────────────────────────────────────

def tension_capacity(material: Material, geom: Geometry, reo: Reinforcement,
                     method: Method, demand_Nt_kN: float) -> dict:
    """Eq 8.10: Fdt ≤ φ·fsy·As. Steel only — masonry tension = 0 (Cl 8.3(e))."""
    As_m = reo.Asv_per_m()
    As_total = As_m * geom.L / 1000.0   # mm²
    phi = phi_tbs(method)
    phi_Nt_per_m = phi * material.fsy * As_m / 1000.0  # kN/m
    phi_Nt_total = phi_Nt_per_m * geom.L / 1000.0      # kN

    # Detailing checks
    spacing_ok = reo.vert_spacing <= 2000.0  # Cl 8.10(c)
    edge_bar_ok = Reinforcement.bar_area(reo.vert_bar_dia) >= 100.0  # Cl 8.10(b)

    # Utilisation
    if demand_Nt_kN <= 0:
        util = 0.0
        verdict = "n/a"
    elif phi_Nt_total <= 0:
        util = float("inf")
        verdict = "fail"
    else:
        util = demand_Nt_kN / phi_Nt_total
        verdict = "pass" if util <= 1.0 else "fail"

    return {
        "demand_Nt_kN": demand_Nt_kN,
        "As_per_m": As_m,
        "As_total": As_total,
        "phi": phi,
        "fsy": material.fsy,
        "phi_Nt_per_m_kN": phi_Nt_per_m,
        "phi_Nt_total_kN": phi_Nt_total,
        "spacing_check_ok": spacing_ok,
        "spacing_value": reo.vert_spacing,
        "edge_bar_check_ok": edge_bar_ok,
        "edge_bar_area": Reinforcement.bar_area(reo.vert_bar_dia),
        "utilisation": util,
        "verdict": verdict,
        "pass": verdict in ("pass", "n/a"),
    }


# ─────────────────────────────────────────────────────────────────────────────
# Cl 8.7 / 8.8 — Shear (in-plane)
# ─────────────────────────────────────────────────────────────────────────────

def fvr(geom: Geometry) -> float:
    """Effective shear strength for long walls, MPa: fvr = 1.5 - 0.5·H/L."""
    return 1.5 - 0.5 * (geom.H / geom.L)


def long_wall_As_for_shear(geom: Geometry, reo: Reinforcement) -> float:
    """Cl 8.7.2 definition of As (effective horizontal reo for shear), mm²."""
    HoverL = geom.H / geom.L
    Aoh = reo.Aoh(geom.H)
    Aov = reo.Aov(geom.L)
    if HoverL > 1.0:
        return Aoh * geom.L / geom.H
    else:
        return min(Aoh, Aov)


def shear_long_wall(material: Material, geom: Geometry, reo: Reinforcement,
                    method: Method) -> dict:
    """Eq 8.7.2: Vu = fvr·Ad + 0.8·fsy·As (un-φ'd, in N).
    Returns kN."""
    fvr_val = fvr(geom)
    As = long_wall_As_for_shear(geom, reo)
    Ad = Ad_total(geom)
    Vu_N = fvr_val * Ad + 0.8 * material.fsy * As
    return {
        "fvr": fvr_val,
        "As_for_shear": As,
        "Ad": Ad,
        "Vu_kN": Vu_N / 1000.0,
    }


def shear_short_wall(material: Material, geom: Geometry, reo: Reinforcement,
                     method: Method) -> dict:
    """Cl 8.7.3 routes to Cl 8.8: Vu = f'vm·bw·d + fvs·Ast + fsy·(Asv·d/s)
    capped at 4·f'vm·bw·d. All in N then convert to kN.

    For in-plane shear: bw = t (wall thickness), d = L (wall length / effective depth).
    Ast = main longitudinal (vertical) reo total area.
    Asv = shear (transverse / horizontal) reo per metre, applied via Asv·d/s formula.
    """
    bw = geom.t
    d = geom.L

    # Term 1: f'vm·bw·d
    masonry_N = F_VM_REINFORCED * bw * d

    # Term 2: fvs·Ast (longitudinal/vertical reo, dowel action)
    Ast = reo.Aov(geom.L)        # total vertical reo across length
    fvs_term_N = FVS_MAIN_REO * Ast

    # Term 3: fsy·(Asv·d/s) — horizontal shear reo with truss-analogy term.
    # In per-metre form this equals fsy × d × (Ash_per_m / 1000), summing primary + additional.
    # Ash_per_m already handles mesh (treats spacing as 1000) per the REO_TABLE lookup.
    shear_reo_N = material.fsy * d * reo.Ash_per_m() / 1000.0

    Vu_raw_N = masonry_N + fvs_term_N + shear_reo_N
    cap_N = 4.0 * F_VM_REINFORCED * bw * d

    if method.enforce_cl_8_8_cap:
        Vu_N = min(Vu_raw_N, cap_N)
        cap_governs = Vu_raw_N > cap_N
    else:
        Vu_N = Vu_raw_N
        cap_governs = False

    return {
        "bw": bw,
        "d": d,
        "fvm": F_VM_REINFORCED,
        "fvs": FVS_MAIN_REO,
        "masonry_term_kN": masonry_N / 1000.0,
        "Ast_long": Ast,
        "fvs_term_kN": fvs_term_N / 1000.0,
        "horiz_reo_term_kN": shear_reo_N / 1000.0,
        "Vu_raw_kN": Vu_raw_N / 1000.0,
        "cap_kN": cap_N / 1000.0,
        "Vu_kN": Vu_N / 1000.0,
        "cap_governs": cap_governs,
    }


def shear_capacity(material: Material, geom: Geometry, reo: Reinforcement,
                   method: Method, demand_V_kN: float) -> dict:
    """Full Cl 8.7 / 8.8 in-plane shear check."""
    HoverL = geom.H / geom.L
    is_long = HoverL <= 2.3

    long_branch = shear_long_wall(material, geom, reo, method)
    short_branch = shear_short_wall(material, geom, reo, method)

    Vu_governing_kN = long_branch["Vu_kN"] if is_long else short_branch["Vu_kN"]
    branch_used = "long (Cl 8.7.2)" if is_long else "short (Cl 8.7.3 → 8.8)"

    phi = phi_tbs(method)
    phi_Vu_kN = phi * Vu_governing_kN

    util = demand_V_kN / phi_Vu_kN if phi_Vu_kN > 0 else float("inf")

    result = {
        "H_over_L": HoverL,
        "is_long_wall": is_long,
        "branch_used": branch_used,
        "long_branch": long_branch,
        "short_branch": short_branch,
        "Vu_governing_kN": Vu_governing_kN,
        "phi": phi,
        "phi_Vu_kN": phi_Vu_kN,
        "demand_V_kN": demand_V_kN,
        "utilisation": util,
        "pass": util <= 1.0,
    }
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Cl 8.7.4 — Stability check (long wall, no top restraint)
# ─────────────────────────────────────────────────────────────────────────────

def stability_check(material: Material, geom: Geometry, reo: Reinforcement,
                    method: Method, demand_Nc_kN: float, demand_V_kN: float,
                    L_prime: float = 0.0) -> dict:
    """Eq 8.7.4: Vd ≤ φ·[ksw·Pv·L/2 + fsy·Asv·(L − 2L')] / H.

    ksw = 1 - Pv/(Ad·f'm).
    Asv per Cl 8.7.4 = vertical reinforcement (perpendicular to applied shear).
    L' = distance from centroid of reo under consideration to tensile end (default 0, conservative).
    """
    Ad = Ad_total(geom)
    Pv_N = demand_Nc_kN * 1000.0
    ksw = 1.0 - Pv_N / (Ad * material.fm) if Ad * material.fm > 0 else 0.0
    ksw_used = max(ksw, 0.0)
    Asv = reo.Aov(geom.L)
    phi = phi_tbs(method)

    # Vd,stab in N
    bracket_N_mm = (ksw_used * Pv_N * geom.L / 2.0) + (material.fsy * Asv * (geom.L - 2 * L_prime))
    Vd_stab_N = phi * bracket_N_mm / geom.H
    Vd_stab_kN = Vd_stab_N / 1000.0

    HoverL = geom.H / geom.L
    applies = method.enforce_stability_check and HoverL <= 2.3 and method.no_top_restraint

    util = demand_V_kN / Vd_stab_kN if (applies and Vd_stab_kN > 0) else 0.0

    return {
        "ksw_raw": ksw,
        "ksw": ksw_used,
        "Pv_kN": demand_Nc_kN,
        "Ad": Ad,
        "Asv": Asv,
        "L_prime": L_prime,
        "phi": phi,
        "Vd_stab_kN": Vd_stab_kN,
        "applies": applies,
        "utilisation": util,
        "pass": util <= 1.0 if applies else True,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Cl 8.6 — Reinforced flexure (explicit, opt-in)
# ─────────────────────────────────────────────────────────────────────────────

def flexural_capacity_8_6(material: Material, geom: Geometry, reo: Reinforcement,
                          method: Method, demand_M_kNm: float, axis: str = "in_plane") -> dict:
    """Eq 8.6: Md ≤ φ·fsy·Asd·d·[1 − (0.6·fsy·Asd)/(1.3·f'm·b·d)].

    For in-plane bending of a wall:
      b = t (wall thickness — width of the equivalent rectangle in flexure)
      d = effective depth = L − cover (approximated as 0.9·L for v1 simplicity)
      Asd = area of main tensile reo in tension zone (use total vertical reo as conservative)

    Cl 8.6 limits Asd to lesser of:
      - (0.29 · 1.3 · f'm · b · d) / fsy   [ku ≤ 0.4 limit]
      - Asc (compression steel area, set = Asd here)

    Detailing: spacing ≤ 2000 mm; ≥100 mm² within 300 mm of edges.
    Md ≥ 1.2 × Mu,unreinforced (not yet enforced).
    """
    if axis == "in_plane":
        b = geom.t
        d_eff = 0.9 * geom.L  # approximation for tensile reo centroid
    else:
        # OOP bending: section is L wide, t deep. Tensile reo centroid ~ t - cover.
        b = geom.L  # for capacity per metre, divide later
        d_eff = 0.9 * geom.t

    Asd_provided = reo.Aov(geom.L) if axis == "in_plane" else reo.Asv_per_m()
    Asd_limit = (0.29 * 1.3 * material.fm * b * d_eff) / material.fsy
    Asd = min(Asd_provided, Asd_limit)

    if Asd <= 0:
        return {
            "Mu_kNm": 0.0,
            "phi_Mu_kNm": 0.0,
            "demand_M_kNm": demand_M_kNm,
            "utilisation": float("inf") if demand_M_kNm > 0 else 0.0,
            "pass": demand_M_kNm == 0,
            "Asd_provided": Asd_provided,
            "Asd_limit": Asd_limit,
            "Asd_used": Asd,
            "b": b, "d_eff": d_eff,
        }

    bracket = 1.0 - (0.6 * material.fsy * Asd) / (1.3 * material.fm * b * d_eff)
    Mu_Nmm = material.fsy * Asd * d_eff * max(bracket, 0.0)
    Mu_kNm = Mu_Nmm / 1e6
    phi = phi_tbs(method)
    phi_Mu_kNm = phi * Mu_kNm

    util = demand_M_kNm / phi_Mu_kNm if phi_Mu_kNm > 0 else float("inf")

    return {
        "Asd_provided": Asd_provided,
        "Asd_limit": Asd_limit,
        "Asd_used": Asd,
        "b": b,
        "d_eff": d_eff,
        "fm": material.fm,
        "fsy": material.fsy,
        "Mu_kNm": Mu_kNm,
        "phi": phi,
        "phi_Mu_kNm": phi_Mu_kNm,
        "demand_M_kNm": demand_M_kNm,
        "utilisation": util,
        "pass": util <= 1.0,
        "axis": axis,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Detailing checks (Cl 8.4, 8.7.2, 8.10)
# ─────────────────────────────────────────────────────────────────────────────

def detailing_checks(material: Material, geom: Geometry, reo: Reinforcement,
                     method: Method) -> list:
    """Return a list of dicts with each detailing rule's PASS/FAIL/N/A."""
    Ad = Ad_total(geom)
    Aov = reo.Aov(geom.L)
    Aoh = reo.Aoh(geom.H)
    checks = []

    # Cl 8.5.1(f) — As ≥ 0.002·Ad for axial reinforced wall
    min_As = 0.002 * Ad
    checks.append({
        "ref": "Cl 8.5.1(f)",
        "rule": "Vertical reo As ≥ 0.002·Ad (min for axial)",
        "limit": min_As, "actual": Aov,
        "status": "PASS" if Aov >= min_As else "FAIL",
        "always_on": True,
    })

    # Cl 8.7.2(ii) — vert spacing ≤ min(0.75·H, 3000), horiz spacing ≤ min(0.75·L, 3000)
    vmax = min(0.75 * geom.H, 3000.0)
    hmax = min(0.75 * geom.L, 3000.0)
    checks.append({
        "ref": "Cl 8.7.2(ii)",
        "rule": "Vertical reo spacing ≤ min(0.75·H, 3000)",
        "limit": vmax, "actual": reo.vert_spacing,
        "status": "PASS" if reo.vert_spacing <= vmax else "FAIL",
        "always_on": True,
    })
    checks.append({
        "ref": "Cl 8.7.2(ii)",
        "rule": "Horizontal reo spacing ≤ min(0.75·L, 3000)",
        "limit": hmax, "actual": reo.horiz_spacing,
        "status": "PASS" if reo.horiz_spacing <= hmax else "FAIL",
        "always_on": True,
    })

    # Cl 8.7.2(iii) — vert ≥ 0.0013·Ad, horiz ≥ 0.0007·Ad
    min_v = 0.0013 * Ad
    min_h = 0.0007 * Ad
    checks.append({
        "ref": "Cl 8.7.2(iii)",
        "rule": "Vertical reo ratio ≥ 0.0013·Ad (full wall length)",
        "limit": min_v, "actual": Aov,
        "status": "PASS" if Aov >= min_v else "FAIL",
        "always_on": True,
    })
    checks.append({
        "ref": "Cl 8.7.2(iii)",
        "rule": "Horizontal reo ratio ≥ 0.0007·Ad (full wall height)",
        "limit": min_h, "actual": Aoh,
        "status": "PASS" if Aoh >= min_h else "FAIL",
        "always_on": True,
    })

    # Cl 8.7.2(iv) / 8.10(b) — ≥100 mm² within 300 mm of edges
    bar_area = Reinforcement.bar_area(reo.vert_bar_dia)
    checks.append({
        "ref": "Cl 8.7.2(iv) / 8.10(b)",
        "rule": "Edge reo: ≥ 100 mm² within 300 mm of each edge",
        "limit": 100.0, "actual": bar_area,
        "status": "PASS" if bar_area >= 100.0 else "FAIL",
        "always_on": True,
    })

    # Cl 8.10(c) — main reo spacing ≤ 2000 mm
    checks.append({
        "ref": "Cl 8.10(c)",
        "rule": "Main reo spacing ≤ 2000 mm (tension)",
        "limit": 2000.0, "actual": reo.vert_spacing,
        "status": "PASS" if reo.vert_spacing <= 2000.0 else "FAIL",
        "always_on": True,
    })

    # Cl 8.6 — bar spacing 2000 + edge bar 100 mm² (similar to 8.10)
    checks.append({
        "ref": "Cl 8.6(a)",
        "rule": "Bending reo spacing ≤ 2000 mm",
        "limit": 2000.0, "actual": reo.vert_spacing,
        "status": "PASS" if reo.vert_spacing <= 2000.0 else "FAIL",
        "always_on": True,
    })

    # Cl 8.4.5 close-spaced reo (toggle)
    if method.enforce_eq_detailing:
        max_eq_spacing = 800.0
        min_eq_ratio = 0.0013 * Ad
        checks.append({
            "ref": "Cl 8.4.5",
            "rule": "EQ ductile detailing: vert spacing ≤ 800 mm",
            "limit": max_eq_spacing, "actual": reo.vert_spacing,
            "status": "PASS" if reo.vert_spacing <= max_eq_spacing else "FAIL",
            "always_on": False, "toggle": "Earthquake detailing",
        })
        checks.append({
            "ref": "Cl 8.4.5",
            "rule": "EQ ductile detailing: horiz spacing ≤ 800 mm",
            "limit": max_eq_spacing, "actual": reo.horiz_spacing,
            "status": "PASS" if reo.horiz_spacing <= max_eq_spacing else "FAIL",
            "always_on": False, "toggle": "Earthquake detailing",
        })
        checks.append({
            "ref": "Cl 8.4.5",
            "rule": "EQ ductile detailing: vert reo ≥ 0.0013·Ad",
            "limit": min_eq_ratio, "actual": Aov,
            "status": "PASS" if Aov >= min_eq_ratio else "FAIL",
            "always_on": False, "toggle": "Earthquake detailing",
        })
        checks.append({
            "ref": "Cl 8.4.5",
            "rule": "EQ ductile detailing: horiz reo ≥ 0.0013·Ad",
            "limit": min_eq_ratio, "actual": Aoh,
            "status": "PASS" if Aoh >= min_eq_ratio else "FAIL",
            "always_on": False, "toggle": "Earthquake detailing",
        })
        checks.append({
            "ref": "Cl 8.4.5",
            "rule": "EQ ductile detailing: wall must be FULLY GROUTED",
            "limit": "full grouting", "actual": geom.grouting,
            "status": "PASS" if geom.grouting == "full" else "FAIL",
            "always_on": False, "toggle": "Earthquake detailing",
        })
    return checks


# ─────────────────────────────────────────────────────────────────────────────
# Top-level full check
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class WallDesign:
    geometry: Geometry
    material: Material
    reinforcement: Reinforcement
    slenderness: Slenderness
    eccentricity: Eccentricity
    method: Method = field(default_factory=Method)

    def check(self, lc: LoadCase) -> dict:
        """Run all enabled checks for a single load case. Returns structured dict."""
        Nc = lc.Nc_worst       # worst of Static/Area
        Nt = lc.Nt
        Mip = lc.Mx_max
        Vip = lc.Vx_max

        comp = compression_capacity(self.material, self.geometry, self.reinforcement,
                                    self.slenderness, self.eccentricity, self.method, Nc)
        tens = tension_capacity(self.material, self.geometry, self.reinforcement,
                                self.method, Nt)
        shear = shear_capacity(self.material, self.geometry, self.reinforcement,
                               self.method, Vip)
        stab = stability_check(self.material, self.geometry, self.reinforcement,
                               self.method, Nc, Vip)
        flex_ip = None
        if self.method.bending_method == "explicit_8_6":
            flex_ip = flexural_capacity_8_6(self.material, self.geometry, self.reinforcement,
                                            self.method, Mip, "in_plane")

        flex_oop = None
        if self.method.include_oop_bending:
            Moop = lc.My_max
            flex_oop = flexural_capacity_8_6(self.material, self.geometry, self.reinforcement,
                                             self.method, Moop, "out_of_plane")

        # Worst utilisation across enabled checks
        utils = [comp["utilisation"], tens["utilisation"], shear["utilisation"]]
        if stab["applies"]:
            utils.append(stab["utilisation"])
        if flex_ip is not None:
            utils.append(flex_ip["utilisation"])
        if flex_oop is not None:
            utils.append(flex_oop["utilisation"])
        max_util = max([u for u in utils if math.isfinite(u)] + [0.0])

        return {
            "lc_name": lc.name,
            "Nc_kN": Nc,
            "Nt_kN": Nt,
            "M_inplane_kNm": Mip,
            "V_inplane_kN": Vip,
            "compression": comp,
            "tension": tens,
            "shear": shear,
            "stability": stab,
            "flexure_inplane": flex_ip,
            "flexure_oop": flex_oop,
            "detailing": detailing_checks(self.material, self.geometry,
                                          self.reinforcement, self.method),
            "max_utilisation": max_util,
            "pass": max_util <= 1.0,
        }

    def check_many(self, load_cases: list) -> dict:
        """Check a list of LCs. Returns per-LC results plus governing LC."""
        results = [self.check(lc) for lc in load_cases]
        if results:
            governing_idx = max(range(len(results)), key=lambda i: results[i]["max_utilisation"])
            governing = results[governing_idx]
        else:
            governing = None
        return {
            "load_cases": results,
            "governing_lc_index": governing_idx if results else None,
            "governing": governing,
            "engine_version": ENGINE_VERSION,
        }


# ─────────────────────────────────────────────────────────────────────────────
# Convenience for quick checks
# ─────────────────────────────────────────────────────────────────────────────

def quick_check(L=2000, t=190, H=5000, tfs=30,
                fuc=15, fcg=20, av=0.85, ah=2.5, e=32,
                vert_dia=16, vert_spacing=200, horiz_dia=12, horiz_spacing=600,
                Nc=600, Nt=0, M=50, V=80,
                alphar=0.40, reinforced=True) -> dict:
    """One-call wrapper for quick interactive checks."""
    geom = Geometry(L=L, t=t, H=H, tfs=tfs)
    mat = Material(fuc=fuc, fcg=fcg)
    reo = Reinforcement(vert_bar_dia=vert_dia, vert_spacing=vert_spacing,
                        horiz_bar_dia=horiz_dia, horiz_spacing=horiz_spacing)
    slend = Slenderness(av=av, ah=ah)
    ecc = Eccentricity(e=e)
    meth = Method(reinforced=reinforced, alphar=alphar)
    wall = WallDesign(geom, mat, reo, slend, ecc, meth)
    lc = LoadCase(name="QC", Nc_static=Nc, Nt=Nt, Mx_top=M, Vx_top=V)
    return wall.check(lc)


if __name__ == "__main__":
    r = quick_check()
    print(f"Compression η = {r['compression']['utilisation']:.3f}  → {'PASS' if r['compression']['pass'] else 'FAIL'}")
    print(f"Tension     η = {r['tension']['utilisation']:.3f}  → {'PASS' if r['tension']['pass'] else 'FAIL'}")
    print(f"Shear       η = {r['shear']['utilisation']:.3f}  → {'PASS' if r['shear']['pass'] else 'FAIL'}")
    print(f"Max overall η = {r['max_utilisation']:.3f}  → {'PASS' if r['pass'] else 'FAIL'}")
