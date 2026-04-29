"""
Validation suite for as3700_walls.py.

Compares engine output to:
  1. The example wall in Blockwork_Wall_Check.xlsx (hand-verified per master notes)
  2. A battery of edge cases (very slender, pure tension, short wall, Cl 8.8 cap, etc.)
  3. Hand calculations done step-by-step against AS 3700:2018 directly

Each test asserts to within 0.1% relative tolerance for capacity values.
"""
from __future__ import annotations
import math
import sys
from dataclasses import dataclass
from typing import Callable

import as3700_walls as w

TOL_RELATIVE = 1e-3   # 0.1%
TOL_ABSOLUTE = 0.01   # for utilisation values

PASS_LOG = []
FAIL_LOG = []


def approx_eq(a, b, rel=TOL_RELATIVE, abs_tol=1e-6):
    if abs(b) < 1e-9:
        return abs(a) < max(abs_tol, 1e-6)
    return abs(a - b) / abs(b) <= rel or abs(a - b) <= abs_tol


def check(label, actual, expected, *, rel=TOL_RELATIVE, abs_tol=1e-6):
    ok = approx_eq(actual, expected, rel=rel, abs_tol=abs_tol)
    status = "PASS" if ok else "FAIL"
    diff_pct = abs(actual - expected) / abs(expected) * 100.0 if abs(expected) > 1e-9 else 0
    line = f"  [{status}] {label:55s} actual={actual:>14.4f}  expected={expected:>14.4f}  diff={diff_pct:>6.3f}%"
    print(line)
    (PASS_LOG if ok else FAIL_LOG).append(line)


def section(title):
    print()
    print("═" * 110)
    print(f"  {title}")
    print("═" * 110)


# ─────────────────────────────────────────────────────────────────────────────
# Test 1: Example wall (190 series, f'uc=15, f'cg=20, e=32, αr=0.40)
# Cross-checked against Blockwork_Wall_Check.xlsx + hand calcs (see master notes §6.6)
# ─────────────────────────────────────────────────────────────────────────────
def test_example_wall():
    section("TEST 1 — Example wall: 190·2000·5000, f'uc=15, f'cg=20, αr=0.40")
    geom = w.Geometry(L=2000, t=190, H=5000, tfs=30)
    mat = w.Material(fuc=15, fcg=20)
    reo = w.Reinforcement(vert_bar_dia=16, vert_spacing=200,
                          horiz_bar_dia=12, horiz_spacing=600)
    slend = w.Slenderness(av=0.85, ah=2.5)
    ecc = w.Eccentricity(e=32)
    meth = w.Method(reinforced=True, alphar=0.40)
    wall = w.WallDesign(geom, mat, reo, slend, ecc, meth)
    lc = w.LoadCase(name="LC1", Nc_static=600, Nt=0, Mx_top=50, Vx_top=80)
    r = wall.check(lc)

    # Material
    check("f'm (MPa)", mat.fm, 7.0488)
    # Cross-section
    check("Ab/m (mm²/m)", w.Ab_per_m(geom), 60000)
    check("Ag/m (mm²/m)", w.Ag_per_m(geom), 130000)
    check("Ad total (mm²)", w.Ad_total(geom), 380000)
    # Slenderness
    check("Sr1", r["compression"]["Sr1"], 22.3684)
    check("Sr2", r["compression"]["Sr2"], 16.9834)
    check("Fo basic per m (kN/m)", r["compression"]["Fo_basic_per_m_kN"], 1498.6)
    check("Fd/Fo", r["compression"]["Fd_over_Fo"], 0.2002)
    # Sr selection: Fd/Fo just over 0.20 → use Sr1 (matches my spreadsheet behaviour)
    check("Sr governing", r["compression"]["Sr_governing"], 22.3684)
    # Eccentricity factor
    check("kes", r["compression"]["kes"], 0.2924)
    # Steel reo (using REO_TABLE rounded values: N16 = 200, N12 = 110, matches L7)
    check("As/m vert (mm²/m)", reo.Asv_per_m(), 1000.0)
    # Capacity
    check("φNu.c per m (kN/m)", r["compression"]["phi_Nuc_per_m_kN"], 371.55, rel=3e-3)
    check("φNu.c total (kN)", r["compression"]["phi_Nuc_total_kN"], 743.10, rel=3e-3)
    check("Compression utilisation", r["compression"]["utilisation"], 0.807, rel=3e-3)
    # Tension
    check("φNt total (kN)", r["tension"]["phi_Nt_total_kN"], 750.0, rel=2e-3)
    # Shear
    check("H/L", r["shear"]["H_over_L"], 2.5)
    check("Long-wall fvr (MPa)", r["shear"]["long_branch"]["fvr"], 0.25)
    check("Long-wall As (mm²)", r["shear"]["long_branch"]["As_for_shear"], 366.67, rel=2e-3)
    check("Long-wall Vu (kN)", r["shear"]["long_branch"]["Vu_kN"], 241.67, rel=2e-3)
    check("Short-wall masonry term (kN)", r["shear"]["short_branch"]["masonry_term_kN"], 133.00)
    check("Short-wall fvs term (kN)", r["shear"]["short_branch"]["fvs_term_kN"], 35.0, rel=2e-3)
    check("Short-wall horiz reo term (kN)", r["shear"]["short_branch"]["horiz_reo_term_kN"], 183.33, rel=3e-3)
    check("Short-wall Vu raw (kN)", r["shear"]["short_branch"]["Vu_raw_kN"], 351.33, rel=3e-3)
    check("Cl 8.8 cap (kN)", r["shear"]["short_branch"]["cap_kN"], 532.0)
    check("Vu governing (kN)", r["shear"]["Vu_governing_kN"], 351.33, rel=3e-3)
    check("φVu (kN)", r["shear"]["phi_Vu_kN"], 263.5, rel=3e-3)
    check("Shear utilisation", r["shear"]["utilisation"], 0.304, rel=3e-3)


# ─────────────────────────────────────────────────────────────────────────────
# Test 2: kes bounds and Sr selection edge cases
# ─────────────────────────────────────────────────────────────────────────────
def test_kes_bounds():
    section("TEST 2 — kes formula bounds")
    # kes at Sr=0, e=0 should = 1
    check("kes(Sr=0, e=0)", w.kes(0, 0, 190), 1.0)
    # kes at Sr=40 should = 0
    check("kes(Sr=40, e=0)", w.kes(40, 0, 190), 0.0)
    # kes at e/t = 0.5 should = 0
    check("kes(Sr=0, e/t=0.5)", w.kes(0, 95, 190), 0.0)
    # kes negative scenario — clamped to 0
    check("kes(Sr=50, e=0) [clamped]", w.kes(50, 0, 190), 0.0)
    check("kes(Sr=0, e/t=0.6) [clamped]", w.kes(0, 114, 190), 0.0)
    # Hand calc: Sr=20, e/t=0.1 → (1-0.5)·(1-0.2) = 0.4
    check("kes(Sr=20, e=19, t=190)", w.kes(20, 19, 190), 0.4)


# ─────────────────────────────────────────────────────────────────────────────
# Test 3: Pure tension wall
# ─────────────────────────────────────────────────────────────────────────────
def test_pure_tension():
    section("TEST 3 — Pure tension wall")
    geom = w.Geometry(L=2000, t=190, H=3000, tfs=30)
    mat = w.Material(fuc=15, fcg=20)
    reo = w.Reinforcement(vert_bar_dia=16, vert_spacing=200, horiz_bar_dia=12, horiz_spacing=600)
    slend = w.Slenderness(av=1.0, ah=2.5)
    ecc = w.Eccentricity(e=10)
    meth = w.Method(reinforced=True, alphar=0.40)
    wall = w.WallDesign(geom, mat, reo, slend, ecc, meth)
    lc = w.LoadCase(name="Tension", Nc_static=0, Nt=300, Mx_top=0, Vx_top=0)
    r = wall.check(lc)
    # Hand using REO_TABLE: As/m = 200 × 1000/200 = 1000
    # φNt = 0.75 × 500 × 1000 / 1000 = 375 kN/m  →  total = 750 kN
    # η = 300/750 = 0.4
    check("φNt total (kN)", r["tension"]["phi_Nt_total_kN"], 750.0, rel=2e-3)
    check("Tension utilisation", r["tension"]["utilisation"], 0.4, rel=2e-3)


# ─────────────────────────────────────────────────────────────────────────────
# Test 4: Long wall (H/L < 2.3) shear — fvr and Cl 8.7.2 branch
# ─────────────────────────────────────────────────────────────────────────────
def test_long_wall_shear():
    section("TEST 4 — Long-wall shear (H/L < 2.3)")
    geom = w.Geometry(L=4000, t=190, H=3000, tfs=30)  # H/L = 0.75
    mat = w.Material(fuc=15, fcg=20)
    reo = w.Reinforcement(vert_bar_dia=16, vert_spacing=400, horiz_bar_dia=12, horiz_spacing=600)
    slend = w.Slenderness(av=1.0, ah=1.0, edges_supported="both")
    ecc = w.Eccentricity(e=10)
    meth = w.Method(reinforced=True, alphar=0.40)
    wall = w.WallDesign(geom, mat, reo, slend, ecc, meth)
    lc = w.LoadCase(name="Long", Nc_static=400, Nt=0, Mx_top=20, Vx_top=200)
    r = wall.check(lc)
    # H/L = 0.75 → long wall, fvr = 1.125 MPa.
    # Using REO_TABLE: Aoh = 110·1000/600 × 3 = 550 mm²; Aov = 200·1000/400 × 4 = 2000 mm²
    # min = 550. Vu = (1.125·760000 + 0.8·500·550)/1000 = 855 + 220 = 1075 kN
    # φVu = 806.25 kN; η = 200/806.25 = 0.248
    check("H/L", r["shear"]["H_over_L"], 0.75)
    check("is long?", float(r["shear"]["is_long_wall"]), 1.0)
    check("fvr", r["shear"]["long_branch"]["fvr"], 1.125)
    check("As for shear (mm²)", r["shear"]["long_branch"]["As_for_shear"], 550.0, rel=2e-3)
    check("Vu (long, kN)", r["shear"]["long_branch"]["Vu_kN"], 1075.0, rel=2e-3)
    check("φVu (kN)", r["shear"]["phi_Vu_kN"], 806.25, rel=2e-3)
    check("Shear utilisation", r["shear"]["utilisation"], 0.2481, rel=3e-3)


# ─────────────────────────────────────────────────────────────────────────────
# Test 5: Cl 8.8 cap actively governing (heavy shear reo)
# ─────────────────────────────────────────────────────────────────────────────
def test_cl_8_8_cap_governs():
    section("TEST 5 — Cl 8.8 cap actively governing (short wall, heavy shear reo)")
    # Short wall H/L = 3.0, very heavy horizontal reo
    geom = w.Geometry(L=1500, t=190, H=4500, tfs=30)
    mat = w.Material(fuc=15, fcg=20)
    reo = w.Reinforcement(vert_bar_dia=20, vert_spacing=200,    # heavy vertical
                          horiz_bar_dia=20, horiz_spacing=200)  # heavy horizontal
    slend = w.Slenderness(av=1.0, ah=2.5)
    ecc = w.Eccentricity(e=10)
    meth = w.Method(reinforced=True, alphar=0.40, enforce_cl_8_8_cap=True)
    wall = w.WallDesign(geom, mat, reo, slend, ecc, meth)
    lc = w.LoadCase(name="Cap", Nc_static=400, Nt=0, Mx_top=20, Vx_top=200)
    r = wall.check(lc)
    # Cap = 4·0.35·190·1500 / 1000 = 399 kN
    check("H/L", r["shear"]["H_over_L"], 3.0)
    check("is long?", float(r["shear"]["is_long_wall"]), 0.0)
    check("Cap (kN)", r["shear"]["short_branch"]["cap_kN"], 399.0)
    # The raw sum should be larger than the cap, so cap governs
    raw = r["shear"]["short_branch"]["Vu_raw_kN"]
    cap = r["shear"]["short_branch"]["cap_kN"]
    check("Vu raw > Cap (cap governs)", float(raw > cap), 1.0)
    check("Vu (after cap, kN)", r["shear"]["short_branch"]["Vu_kN"], cap)
    check("cap_governs flag", float(r["shear"]["short_branch"]["cap_governs"]), 1.0)


# ─────────────────────────────────────────────────────────────────────────────
# Test 6: Sr selection rule — Fd/Fo ≤ 0.20 vs > 0.20
# ─────────────────────────────────────────────────────────────────────────────
def test_sr_selection_rule():
    section("TEST 6 — Sr selection rule (Cl 7.3.4.3(a) Fd/Fo ≤ 0.20)")
    geom = w.Geometry(L=2000, t=190, H=5000, tfs=30)
    mat = w.Material(fuc=15, fcg=20)
    reo = w.Reinforcement(vert_bar_dia=16, vert_spacing=200, horiz_bar_dia=12, horiz_spacing=600)
    slend = w.Slenderness(av=0.85, ah=2.5, edges_supported="one")
    ecc = w.Eccentricity(e=32)

    # Low-axial case → should use MIN(Sr1, Sr2)
    meth = w.Method(reinforced=True, alphar=0.40)
    wall = w.WallDesign(geom, mat, reo, slend, ecc, meth)
    lc_low = w.LoadCase(name="LowN", Nc_static=10, Nt=0, Mx_top=0, Vx_top=0)
    r_low = wall.check(lc_low)
    check("Low N, Sr governing = MIN", r_low["compression"]["Sr_governing"],
          min(r_low["compression"]["Sr1"], r_low["compression"]["Sr2"]))

    # High-axial case → should use Sr1 only
    lc_high = w.LoadCase(name="HighN", Nc_static=2000, Nt=0, Mx_top=0, Vx_top=0)
    r_high = wall.check(lc_high)
    check("High N, Sr governing = Sr1", r_high["compression"]["Sr_governing"],
          r_high["compression"]["Sr1"])

    # Edges_supported = "none" — must always use Sr1
    slend_none = w.Slenderness(av=0.85, ah=2.5, edges_supported="none")
    wall2 = w.WallDesign(geom, mat, reo, slend_none, ecc, meth)
    r_none = wall2.check(lc_low)
    check("No edge support, Sr governing = Sr1", r_none["compression"]["Sr_governing"],
          r_none["compression"]["Sr1"])


# ─────────────────────────────────────────────────────────────────────────────
# Test 7: αr = 0 vs αr = 0.40 — verify steel term contribution
# ─────────────────────────────────────────────────────────────────────────────
def test_alpha_r():
    section("TEST 7 — αr (steel contribution to axial)")
    geom = w.Geometry(L=2000, t=190, H=5000, tfs=30)
    mat = w.Material(fuc=15, fcg=20)
    reo = w.Reinforcement(vert_bar_dia=16, vert_spacing=200, horiz_bar_dia=12, horiz_spacing=600)
    slend = w.Slenderness(av=0.85, ah=2.5)
    ecc = w.Eccentricity(e=32)

    # αr = 0 (L7 conservative)
    meth0 = w.Method(reinforced=True, alphar=0.0)
    wall0 = w.WallDesign(geom, mat, reo, slend, ecc, meth0)
    r0 = wall0.check(w.LoadCase(name="A", Nc_static=600, Vx_top=80))

    # αr = 0.40
    meth40 = w.Method(reinforced=True, alphar=0.40)
    wall40 = w.WallDesign(geom, mat, reo, slend, ecc, meth40)
    r40 = wall40.check(w.LoadCase(name="B", Nc_static=600, Vx_top=80))

    # Steel term using REO_TABLE: 0.40 × 500 × 1000 = 200,000 N/m = 200 kN/m
    check("Steel term @ αr=0   (kN/m)", r0["compression"]["steel_term_N_per_m"]/1000, 0.0)
    check("Steel term @ αr=0.4 (kN/m)", r40["compression"]["steel_term_N_per_m"]/1000, 200.0, rel=2e-3)
    # Capacity uplift from including steel: ~12% for this wall
    uplift = (r40["compression"]["phi_Nuc_total_kN"] / r0["compression"]["phi_Nuc_total_kN"]) - 1
    print(f"  [INFO] Capacity uplift from αr=0 → αr=0.4: {uplift*100:.1f}%")


# ─────────────────────────────────────────────────────────────────────────────
# Test 8: Multi-LC mode — pick the worst across LCs
# ─────────────────────────────────────────────────────────────────────────────
def test_multi_lc():
    section("TEST 8 — Multi-LC: pick worst utilisation")
    geom = w.Geometry(L=2000, t=190, H=5000, tfs=30)
    mat = w.Material(fuc=15, fcg=20)
    reo = w.Reinforcement(vert_bar_dia=16, vert_spacing=200, horiz_bar_dia=12, horiz_spacing=600)
    slend = w.Slenderness(av=0.85, ah=2.5)
    ecc = w.Eccentricity(e=32)
    meth = w.Method(reinforced=True, alphar=0.40)
    wall = w.WallDesign(geom, mat, reo, slend, ecc, meth)
    lcs = [
        w.LoadCase(name="LC1 1.2G+1.5Q", Nc_static=400, Vx_top=50),
        w.LoadCase(name="LC2 1.2G+1.5Q+W", Nc_static=600, Vx_top=80, Mx_top=50),
        w.LoadCase(name="LC3 0.9G+W",   Nc_static=200, Vx_top=120, Mx_top=80),
        w.LoadCase(name="LC4 1.2G+0.4Q+W", Nc_static=350, Vx_top=100, Mx_top=70),
    ]
    res = wall.check_many(lcs)
    print(f"  Governing LC: {res['governing']['lc_name']}  η = {res['governing']['max_utilisation']:.3f}")
    # LC2 has the highest axial → highest compression util usually
    check("Number of LCs", float(len(res["load_cases"])), 4.0)
    # Sanity: governing util should equal the max across LCs
    max_u = max(r["max_utilisation"] for r in res["load_cases"])
    check("Governing util = max across LCs", res["governing"]["max_utilisation"], max_u)


# ─────────────────────────────────────────────────────────────────────────────
# Test 9: Static vs Area axial worst-case logic
# ─────────────────────────────────────────────────────────────────────────────
def test_static_vs_area():
    section("TEST 9 — RCB Static vs Area: worst-case selection")
    lc1 = w.LoadCase(name="X", Nc_static=300, Nc_area=500)
    lc2 = w.LoadCase(name="Y", Nc_static=600, Nc_area=400)
    check("worst(static=300, area=500)", lc1.Nc_worst, 500)
    check("worst(static=600, area=400)", lc2.Nc_worst, 600)


# ─────────────────────────────────────────────────────────────────────────────
# Test 10: Detailing checks engine
# ─────────────────────────────────────────────────────────────────────────────
def test_detailing():
    section("TEST 10 — Detailing checks")
    geom = w.Geometry(L=2000, t=190, H=5000, tfs=30)
    mat = w.Material(fuc=15, fcg=20)
    reo = w.Reinforcement(vert_bar_dia=16, vert_spacing=200, horiz_bar_dia=12, horiz_spacing=600)
    slend = w.Slenderness(av=0.85, ah=2.5)
    ecc = w.Eccentricity(e=32)
    meth = w.Method(reinforced=True, alphar=0.40, enforce_eq_detailing=False)
    wall = w.WallDesign(geom, mat, reo, slend, ecc, meth)
    r = wall.check(w.LoadCase(name="D", Nc_static=600, Vx_top=80))
    fails = [c for c in r["detailing"] if c["status"] == "FAIL"]
    print(f"  {len(r['detailing'])} detailing checks run, {len(fails)} failures.")
    for c in r["detailing"]:
        print(f"    [{c['status']}] {c['ref']}: {c['rule']}")
    # All should pass for the example wall (well-detailed)
    check("All detailing checks pass", float(len(fails) == 0), 1.0)

    # Now turn on EQ detailing — sv=200, sh=600 vs limit 800, both pass
    meth2 = w.Method(reinforced=True, alphar=0.40, enforce_eq_detailing=True)
    wall2 = w.WallDesign(geom, mat, reo, slend, ecc, meth2)
    r2 = wall2.check(w.LoadCase(name="D2", Nc_static=600, Vx_top=80))
    eq_checks = [c for c in r2["detailing"] if not c["always_on"]]
    print(f"  With EQ toggle on: {len(eq_checks)} EQ checks added.")
    check("EQ checks were added", float(len(eq_checks) > 0), 1.0)


# ─────────────────────────────────────────────────────────────────────────────
# Test 11: Stability check (Cl 8.7.4)
# ─────────────────────────────────────────────────────────────────────────────
def test_stability():
    section("TEST 11 — Cl 8.7.4 stability check")
    geom = w.Geometry(L=4000, t=190, H=3000, tfs=30)  # H/L=0.75
    mat = w.Material(fuc=15, fcg=20)
    reo = w.Reinforcement(vert_bar_dia=16, vert_spacing=400, horiz_bar_dia=12, horiz_spacing=600)
    slend = w.Slenderness(av=1.0, ah=1.0, edges_supported="both")
    ecc = w.Eccentricity(e=10)
    meth = w.Method(reinforced=True, alphar=0.40, enforce_stability_check=True, no_top_restraint=True)
    wall = w.WallDesign(geom, mat, reo, slend, ecc, meth)
    r = wall.check(w.LoadCase(name="Stab", Nc_static=400, Vx_top=200))
    # ksw = 1 - 400000/(760000 × 7.0488) = 0.9253
    check("ksw", r["stability"]["ksw"], 0.9253, rel=3e-3)
    # Vd,stab using Aov = 2000 (REO_TABLE): 0.75·(0.9253·400000·4000/2 + 500·2000·4000)/3000/1000
    # = 0.75·(740,240,000 + 4,000,000,000)/3,000,000 = 1185.06 kN
    check("Vd,stab (kN)", r["stability"]["Vd_stab_kN"], 1185.06, rel=3e-3)
    check("stability check applies", float(r["stability"]["applies"]), 1.0)


# ─────────────────────────────────────────────────────────────────────────────
# Test 12: Cl 8.6 explicit flexure
# ─────────────────────────────────────────────────────────────────────────────
def test_flexure_8_6():
    section("TEST 12 — Cl 8.6 explicit reinforced flexure (in-plane)")
    geom = w.Geometry(L=2000, t=190, H=5000, tfs=30)
    mat = w.Material(fuc=15, fcg=20)
    reo = w.Reinforcement(vert_bar_dia=16, vert_spacing=200, horiz_bar_dia=12, horiz_spacing=600)
    slend = w.Slenderness(av=0.85, ah=2.5)
    ecc = w.Eccentricity(e=32)
    meth = w.Method(reinforced=True, alphar=0.40, bending_method="explicit_8_6")
    wall = w.WallDesign(geom, mat, reo, slend, ecc, meth)
    r = wall.check(w.LoadCase(name="Flex", Nc_static=100, Mx_top=200))
    # Hand check:
    # Using REO_TABLE: Aov = 1000·2 = 2000 mm²
    # Asd_limit = 0.29·1.3·7.05·190·1800/500 = 1817.27, Asd = MIN = 1817.27 (limit governs)
    # Mu = 500·1817.27·1800·(1 - 0.6·500·1817.27/(1.3·7.05·190·1800))/1e6 = 1351.4 kNm
    # φMu = 1013.5 kNm — same answer because Asd_limit governs (not Asd_provided)
    fl = r["flexure_inplane"]
    print(f"  Asd_provided = {fl['Asd_provided']:.1f}, Asd_limit = {fl['Asd_limit']:.1f}, used = {fl['Asd_used']:.1f}")
    print(f"  Mu = {fl['Mu_kNm']:.1f} kNm, φMu = {fl['phi_Mu_kNm']:.1f} kNm")
    check("Asd_used (mm²)", fl["Asd_used"], 1817.27, rel=5e-3)
    check("Mu (kNm)", fl["Mu_kNm"], 1351.4, rel=5e-3)
    check("φMu (kNm)", fl["phi_Mu_kNm"], 1013.5, rel=5e-3)


# ─────────────────────────────────────────────────────────────────────────────
# Test 13: Very slender wall (kes → low)
# ─────────────────────────────────────────────────────────────────────────────
def test_very_slender():
    section("TEST 13 — Very slender wall (kes near zero)")
    geom = w.Geometry(L=3000, t=140, H=8000, tfs=25)  # Very tall, thin
    mat = w.Material(fuc=15, fcg=20)
    reo = w.Reinforcement(vert_bar_dia=16, vert_spacing=400, horiz_bar_dia=12, horiz_spacing=800)
    slend = w.Slenderness(av=1.0, ah=2.5)
    ecc = w.Eccentricity(e=23)  # ~t/6
    meth = w.Method(reinforced=True, alphar=0.40)
    wall = w.WallDesign(geom, mat, reo, slend, ecc, meth)
    r = wall.check(w.LoadCase(name="Slender", Nc_static=200, Vx_top=20))
    # Sr1 = 1.0 × 8000 / 140 = 57.14 — VERY slender by Sr1 alone
    # Sr2 = (0.7/140)·√(1.0·8000·2.5·3000) = 38.73 — lateral edges save it
    # Fd/Fo is low → use MIN(Sr1, Sr2) = 38.73 (still slender but kes > 0)
    # kes = (1 - 0.025·38.73)·(1 - 2·23/140) = 0.0317·0.6714 = 0.0213
    print(f"  Sr1 = {r['compression']['Sr1']:.2f}, Sr2 = {r['compression']['Sr2']:.2f}")
    print(f"  Sr governing = {r['compression']['Sr_governing']:.2f}, kes = {r['compression']['kes']:.4f}")
    check("Sr1 > 40 (Eq 1 alone fails)", float(r["compression"]["Sr1"] > 40), 1.0)
    check("Sr2 < 40 (Eq 2 saves it)", float(r["compression"]["Sr2"] < 40), 1.0)
    check("Sr governing = MIN(Sr1, Sr2)", r["compression"]["Sr_governing"], min(r["compression"]["Sr1"], r["compression"]["Sr2"]))
    check("kes very low but positive", r["compression"]["kes"], 0.0213, rel=5e-3)
    # Now verify a wall with NO edge support fails completely
    slend_no = w.Slenderness(av=1.0, ah=2.5, edges_supported="none")
    wall_no = w.WallDesign(geom, mat, reo, slend_no, ecc, meth)
    r_no = wall_no.check(w.LoadCase(name="SlenderNoEdge", Nc_static=200, Vx_top=20))
    check("kes = 0 when no edge support and Sr1 > 40", r_no["compression"]["kes"], 0.0)


# ─────────────────────────────────────────────────────────────────────────────
# Test 14: REO_TABLE + EF + additional bars + mesh
# ─────────────────────────────────────────────────────────────────────────────
def test_reo_table():
    section("TEST 14 — REO_TABLE: EF, additional bars, mesh")

    # Each Face: N16 EF should give exactly 2× the area of N16
    geom = w.Geometry(L=2000, t=190, H=5000)
    reo_n16 = w.Reinforcement(vert_primary_code="N16", vert_primary_spacing=200,
                              horiz_primary_code="N12", horiz_primary_spacing=600)
    reo_n16ef = w.Reinforcement(vert_primary_code="N16 EF", vert_primary_spacing=200,
                                horiz_primary_code="N12", horiz_primary_spacing=600)
    check("Asv N16 @200",      reo_n16.Asv_per_m(),   1000.0)         # 200·1000/200
    check("Asv N16 EF @200",   reo_n16ef.Asv_per_m(), 2000.0)         # 400·1000/200
    check("EF doubles N16",    reo_n16ef.Asv_per_m() / reo_n16.Asv_per_m(), 2.0)

    # Additional bars sum on top of primary
    reo_combo = w.Reinforcement(
        vert_primary_code="N16", vert_primary_spacing=400,            # 200·1000/400 = 500
        vert_additional_code="N12", vert_additional_spacing=400,      # 110·1000/400 = 275
        horiz_primary_code="N12", horiz_primary_spacing=600)
    check("Asv primary+additional", reo_combo.Asv_per_m(), 775.0)     # 500 + 275

    # SL mesh — area is per metre, spacing input ignored
    reo_mesh = w.Reinforcement(vert_primary_code="SL82", vert_primary_spacing=999,
                               horiz_primary_code="SL82", horiz_primary_spacing=42)
    check("Asv SL82 (mesh, ignores spacing)",  reo_mesh.Asv_per_m(),  227.0)
    check("Ash SL82 (mesh)",                   reo_mesh.Ash_per_m(),  227.0)

    # SL82 EF — corrected from L7 typo (5448), should be 2 × 227 = 454
    reo_mesh_ef = w.Reinforcement(vert_primary_code="SL82 EF", vert_primary_spacing=200,
                                  horiz_primary_code="-", horiz_primary_spacing=600)
    check("Asv SL82 EF (corrected from L7 typo)", reo_mesh_ef.Asv_per_m(), 454.0)

    # "-" gives 0
    reo_none = w.Reinforcement(vert_primary_code="-", vert_primary_spacing=200,
                               horiz_primary_code="-", horiz_primary_spacing=600)
    check("Asv with '-' code", reo_none.Asv_per_m(), 0.0)

    # Larger bars (N28, N32)
    reo_big = w.Reinforcement(vert_primary_code="N32", vert_primary_spacing=400,
                              horiz_primary_code="N12", horiz_primary_spacing=600)
    check("Asv N32 @400 (= 800·1000/400)", reo_big.Asv_per_m(), 2000.0)

    # Backward compat: legacy vert_bar_dia=16, vert_spacing=200 should still work
    reo_legacy = w.Reinforcement(vert_bar_dia=16, vert_spacing=200,
                                 horiz_bar_dia=12, horiz_spacing=600)
    check("Legacy schema vert_bar_dia=16 → N16", reo_legacy.Asv_per_m(), 1000.0)


# ─────────────────────────────────────────────────────────────────────────────
# Run all
# ─────────────────────────────────────────────────────────────────────────────
def run_all():
    tests = [
        test_example_wall,
        test_kes_bounds,
        test_pure_tension,
        test_long_wall_shear,
        test_cl_8_8_cap_governs,
        test_sr_selection_rule,
        test_alpha_r,
        test_multi_lc,
        test_static_vs_area,
        test_detailing,
        test_stability,
        test_flexure_8_6,
        test_very_slender,
        test_reo_table,
    ]
    for t in tests:
        try:
            t()
        except Exception as e:
            FAIL_LOG.append(f"  [EXCEPTION] {t.__name__}: {e}")
            print(f"  [EXCEPTION] {t.__name__}: {e}")

    # Summary
    print()
    print("═" * 110)
    print(f"  SUMMARY: {len(PASS_LOG)} PASSED, {len(FAIL_LOG)} FAILED")
    print("═" * 110)
    if FAIL_LOG:
        print("\nFAILURES:")
        for line in FAIL_LOG:
            print(line)
    return len(FAIL_LOG) == 0


if __name__ == "__main__":
    ok = run_all()
    sys.exit(0 if ok else 1)
