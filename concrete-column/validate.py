"""
Validation of as3600_columns.py against the worked examples in
Reinforced Concrete Basics (3e), Chapter 5.

Examples validated:
  5.2 — Section capacity line for 400 x 600 column with 1200 mm^2/face
  5.5 — Full column design (slender unbraced) — moment magnifier
  5.6 — HSC core confinement (high axial)
  5.7 — HSC core confinement (moderate axial, high moment)
"""
import math
from engine_reference import (
    Section, RebarLayer, Void,
    alpha_1, alpha_2, gamma,
    phi_bending_only, phi_compression, phi_tension,
    radius_of_gyration, is_short_column,
    effective_length_factor_braced, effective_length_factor_unbraced,
    buckling_load_Nc, moment_magnifier_braced, km_factor,
    moment_magnifier_unbraced_storey,
    biaxial_utilisation,
    HSCConfinementInput, hsc_fitment_spacing_simplified,
    hsc_fitment_spacing_deemed,
    capacity_at_N_star,
    # ── PRO ──
    check_splice, special_region_length, restraint_pattern_required,
    biaxial_concession, joint_transmission_check,
)


def check(name, computed, expected, tol_rel=0.01, unit="", expected_str=None):
    err = abs(computed - expected) / abs(expected) if expected else abs(computed)
    status = "PASS" if err <= tol_rel else "FAIL"
    exp_disp = expected_str if expected_str else f"{expected:.4g}"
    print(f"  {status}  {name:38s}  computed = {computed:>10.4g} {unit:5s}"
          f" expected = {exp_disp:>10s} {unit:5s} err = {err*100:>5.2f}%")
    return err <= tol_rel


def section_5_2():
    """Example 5.2 — 400x600 with 1200 mm^2 in each face, fc'=40, fsy=500."""
    print("\n=== EXAMPLE 5.2 — Section capacity line ===")
    print("400x600 RC column, 1200 mm^2 in each face, fc'=40 MPa, fsy=500 MPa")
    print("d = 526 mm (cover-to-bar-centre 74 mm)")
    sec = Section(
        shape="rect", D=600.0, b=400.0, fc=40.0, fsy=500.0,
        layers=[
            RebarLayer(d_from_comp_face=74.0,  area=1200.0),
            RebarLayer(d_from_comp_face=526.0, area=1200.0),
        ],
    )
    # Material coefficients
    print(f"  α1 = {alpha_1(40):.3f}  (book: 0.85)")
    print(f"  α2 = {alpha_2(40):.3f}  (book: 0.79)")
    print(f"  γ  = {gamma(40):.3f}  (book: 0.87)")

    # Plastic centroid — symmetric, should be 300 mm
    dpc = sec.plastic_centroid()
    print(f"\n  Plastic centroid dpc = {dpc:.1f} mm  (book: 300 mm)")

    results = []
    # Point A — squash. AS 3600 Cl 10.6.2.2: uses (Ag - As) and caps steel
    # stress at εs·Es = 0.0025·Es = 500 MPa for Es=200 GPa. RCB uses Ag (no
    # subtraction), so the book's 9360 kN is ~0.9% higher than AS 3600.
    Nuo = sec.nuo() / 1000.0
    results.append(check("A: Nuo (squash, AS 3600)", Nuo, 9278, unit="kN"))

    # Point B — decompression (ku = 1.0)
    Nb_N, Mb_Nmm = sec.decompression_point()
    results.append(check("B: Nu @ ku=1.0", Nb_N / 1000.0, 6385, unit="kN"))
    results.append(check("B: Mu @ ku=1.0", Mb_Nmm / 1e6, 547, unit="kNm"))

    # Point C — balanced
    Nub_N, Mub_Nmm, kub = sec.balanced_point()
    results.append(check("C: Nub (balanced)", Nub_N / 1000.0, 3090, unit="kN"))
    results.append(check("C: Mub (balanced)", Mub_Nmm / 1e6, 809, unit="kNm"))
    results.append(check("C: kub", kub, 0.545))

    # Point D — pure bending
    Nd, Md, dn_pb = sec.pure_bending()
    print(f"  D: dn (pure bending) = {dn_pb:.1f} mm  (book: 64 mm)")
    results.append(check("D: Muo (pure bending)", Md / 1e6, 302, tol_rel=0.02, unit="kNm"))

    # Point E — pure tension
    Nuot = sec.nuo_t() / 1000.0
    results.append(check("E: Nuo,t (pure tension)", Nuot, 1200, unit="kN"))

    return all(results)


def section_5_5():
    """Example 5.5 — slender unbraced column.
    400x600, fc'=40, p ≈ 0.01 chosen, βd = 0.7."""
    print("\n=== EXAMPLE 5.5 — Slender unbraced column, moment magnifier ===")
    print("400x600 column C2-3 in unbraced frame, βd=0.7, p≈0.01")
    sec = Section(
        shape="rect", D=600.0, b=400.0, fc=40.0, fsy=500.0,
        layers=[
            # p = 0.01, equal each face: As_total = 0.01 × 400 × 600 = 2400 mm²
            RebarLayer(d_from_comp_face=74.0,  area=1200.0),
            RebarLayer(d_from_comp_face=526.0, area=1200.0),
        ],
    )
    results = []

    # Effective length: book gives 6340 mm interior (already computed)
    # Le/r check:
    r = radius_of_gyration("rect", 600)
    Le = 6340.0
    print(f"  r = {r:.0f} mm     Le = {Le:.0f}     Le/r = {Le/r:.1f}  (book: 35.2)")

    # φMub for the section, used in Nc
    Nub_N, Mub_Nmm, kub = sec.balanced_point()
    phi_Mub_AS = 0.65 * Mub_Nmm  # per AS 3600 Cl 10.4.4
    phi_Mub_RCB = 0.60 * Mub_Nmm  # what RCB Example 5.5 used
    print(f"  φMub per AS 3600 Cl 10.4.4 (φ=0.65) = {phi_Mub_AS/1e6:.0f} kNm")
    print(f"  φMub per RCB Example 5.5 (φ=0.6)   = {phi_Mub_RCB/1e6:.0f} kNm  (book: 485)")
    print(f"  → RCB applied the slender k_φ=12/13 reduction in the Nc formula,")
    print(f"    but AS 3600 Cl 10.4.4 hard-wires φ=0.65. ENGINE FOLLOWS AS 3600.")
    results.append(check("φMub (AS 3600 Cl 10.4.4)", phi_Mub_AS / 1e6, 526, tol_rel=0.01, unit="kNm"))

    # Buckling load Nc — interior, βd = 0.7
    # NOTE: per AS 3600, Nc should be 7272 kN (book 6706 = book × 0.6/0.65).
    Nc_int = buckling_load_Nc(sec, Le=6340.0, beta_d=0.7) / 1000.0
    Nc_int_book = 7272 * (485 / 525.9)  # adjusted to book's φMub
    results.append(check("Nc interior (AS 3600)", Nc_int, 7272, tol_rel=0.01, unit="kN"))

    # Buckling load Nc — exterior, Le = 8640 mm
    Nc_ext = buckling_load_Nc(sec, Le=8640.0, beta_d=0.7) / 1000.0
    results.append(check("Nc exterior (AS 3600)", Nc_ext, 3916, tol_rel=0.01, unit="kN"))

    # Verify δb by feeding the BOOK'S Nc value — this checks the magnifier
    # formula in isolation from the φMub disagreement.
    km = km_factor(M1=-270.0, M2=360.0)
    results.append(check("km", km, 0.90, tol_rel=0.01))

    db_book_inputs = moment_magnifier_braced(km=km, N_star=2000.0, Nc=6706.0)
    results.append(check("δb with book's Nc=6706", db_book_inputs, 1.28, tol_rel=0.02))

    # Storey magnifier δs — also verified with book's Nc values
    sumN = 2 * 1100 + 4 * 2000
    sumNc = 2 * 3611 + 4 * 6706
    ds = moment_magnifier_unbraced_storey(sum_N_star=sumN, sum_Nc=sumNc)
    results.append(check("δs with book's Nc values", ds, 1.43, tol_rel=0.01))

    return all(results)


def section_5_4():
    """Example 5.4 — biaxial check (Cl 10.6.4)."""
    print("\n=== EXAMPLE 5.4 — Biaxial bending check ===")
    print("400x600 with 1N28 in each corner, N*=4000 kN, M*x=280, M*y=140")
    sec = Section(
        shape="rect", D=600.0, b=400.0, fc=40.0, fsy=500.0,
        layers=[
            RebarLayer(d_from_comp_face=74.0,  area=2 * 615.8),
            RebarLayer(d_from_comp_face=526.0, area=2 * 615.8),
        ],
    )
    # AS 3600 Cl 10.6.4: αn = 0.7 + 1.7·N*/Nuo  (NO 0.65 factor)
    # RCB Eq 5.21 (textbook): αn = 0.7 + 1.7·N*/(0.65·Nuo)
    # The book's formula gives αn = 1.82, AS 3600 gives αn = 1.43.
    # Engine follows AS 3600 — more conservative (smaller αn).
    util_AS, alpha_n_AS = biaxial_utilisation(
        Mx_star=280.0, My_star=140.0,
        phi_Mux=390.0, phi_Muy=230.0,
        N_star=4000.0, Nuo=9360.0,
    )
    alpha_n_book = 0.7 + 1.7 * 4000 / (0.65 * 9360)
    util_book = (280/390)**alpha_n_book + (140/230)**alpha_n_book
    print(f"  αn per AS 3600 Cl 10.6.4 = {alpha_n_AS:.3f}  (engine)")
    print(f"  αn per RCB Eq 5.21       = {alpha_n_book:.3f}  (book: 1.82)")
    print(f"  Utilisation per AS 3600  = {util_AS:.3f}  (more conservative)")
    print(f"  Utilisation per RCB      = {util_book:.3f}  (book: 0.95)")
    print(f"  → ENGINE FOLLOWS AS 3600 — RCB's αn formula has an extra 0.65 factor.")
    results = []
    # Verify the engine reproduces AS 3600 formula exactly
    expected_alpha_AS = 0.7 + 1.7 * 4000 / 9360
    results.append(check("αn (AS 3600 Cl 10.6.4)", alpha_n_AS, expected_alpha_AS, tol_rel=0.001))
    expected_util_AS = (280/390)**expected_alpha_AS + (140/230)**expected_alpha_AS
    results.append(check("util (AS 3600 Cl 10.6.4)", util_AS, expected_util_AS, tol_rel=0.001))
    # And verify the book's numbers reproduce when fed book formula
    results.append(check("RCB formula reproduces book", util_book, 0.95, tol_rel=0.02))
    return all(results)


def section_5_6():
    """Example 5.6 — HSC confinement, 700x700, 12N40, fc'=80."""
    print("\n=== EXAMPLE 5.6 — HSC core confinement (high axial) ===")
    print("700x700 sq, fc'=80, 12N40 bars, N12 fitments, cover 30 mm")
    # bc = dc = 700 - 2×30 - 12 = 628 mm
    inp = HSCConfinementInput(
        bc=628.0, dc=628.0,
        n=12,
        # 12 bars on perimeter: 4 per side. Clear spacing w:
        # 700 - 2×30 - 2×12 - 4×40 = 456 mm over 3 spaces = 152 mm
        w=152.0,
        Ab_fit=110.0,
        fsy_f=500.0,
        fc=80.0,
        m=4,
        shape="rect",
        ds=628.0,
    )
    s_simp = hsc_fitment_spacing_simplified(inp)
    s_deemed = hsc_fitment_spacing_deemed(inp)
    results = []
    results.append(check("Simplified s_max", s_simp, 249, tol_rel=0.05, unit="mm"))
    results.append(check("Deemed-to-comply s_max", s_deemed, 197, tol_rel=0.02, unit="mm"))
    return all(results)


def section_5_7():
    """Example 5.7 — HSC confinement, 350x350, 8N20, fc'=80."""
    print("\n=== EXAMPLE 5.7 — HSC core confinement (moderate axial, high moment) ===")
    print("350x350 sq, fc'=80, 8N20 bars, N10 fitments, cover 30 mm")
    # bc = dc = 350 - 2×30 - 10 = 280 mm
    # n = 4 (only corner bars laterally restrained)
    # w (clear spacing between restrained corner bars) = 350 - 2×30 - 2×10 - 2×20 = 230 mm
    inp = HSCConfinementInput(
        bc=280.0, dc=280.0,
        n=4,
        w=230.0,
        Ab_fit=80.0,   # N10 = 78.5 mm² ≈ 80 mm² used in book
        fsy_f=500.0,
        fc=80.0,
        m=2,
        shape="rect",
        ds=280.0,
    )
    s_simp = hsc_fitment_spacing_simplified(inp)
    s_deemed = hsc_fitment_spacing_deemed(inp)
    results = []
    results.append(check("Simplified s_max", s_simp, 121, tol_rel=0.05, unit="mm"))
    results.append(check("Deemed-to-comply s_max", s_deemed, 107, tol_rel=0.05, unit="mm"))
    return all(results)


def e2e_b1_fsy_cap():
    """End-to-end: 600 MPa longitudinal steel must NOT raise Nuo above the
    500 MPa cap (Cl 10.6.2.2). Engine should give same Nuo regardless of fsy
    once fsy ≥ 500."""
    print("\n=== E2E B1 — Squash-load 500 MPa cap (fsy=600) ===")
    sec500 = Section(
        shape="rect", D=600.0, b=400.0, fc=40.0, fsy=500.0,
        layers=[
            RebarLayer(d_from_comp_face=74.0,  area=1200.0),
            RebarLayer(d_from_comp_face=526.0, area=1200.0),
        ],
    )
    sec600 = Section(
        shape="rect", D=600.0, b=400.0, fc=40.0, fsy=600.0,
        layers=[
            RebarLayer(d_from_comp_face=74.0,  area=1200.0),
            RebarLayer(d_from_comp_face=526.0, area=1200.0),
        ],
    )
    results = []
    n500 = sec500.nuo() / 1000.0
    n600 = sec600.nuo() / 1000.0
    print(f"  Nuo with fsy=500 MPa: {n500:.0f} kN")
    print(f"  Nuo with fsy=600 MPa: {n600:.0f} kN  (must equal fsy=500 case after cap)")
    results.append(check("Nuo cap binds at 500 MPa", n600, n500, tol_rel=0.001, unit="kN"))
    return all(results)


def e2e_b3_axis_slenderness():
    """End-to-end: rectangular column 400×600 — Le/r about each axis differs."""
    print("\n=== E2E B3 — Per-axis radius of gyration (rect 400×600) ===")
    rx = radius_of_gyration("rect", 600.0)
    ry = radius_of_gyration("rect", 400.0)
    Le = 2550.0
    print(f"  rx = 0.3·600 = {rx:.0f} mm,  Le/rx = {Le/rx:.2f}")
    print(f"  ry = 0.3·400 = {ry:.0f} mm,  Le/ry = {Le/ry:.2f}")
    results = []
    results.append(check("rx",  rx, 180.0, tol_rel=0.001, unit="mm"))
    results.append(check("ry",  ry, 120.0, tol_rel=0.001, unit="mm"))
    results.append(check("Le/rx",  Le/rx, 14.17, tol_rel=0.005))
    results.append(check("Le/ry",  Le/ry, 21.25, tol_rel=0.005))
    print(f"  → governing slenderness uses min(rx, ry) = ry = {ry:.0f} mm")
    return all(results)


def e2e_b2_per_combo_ds():
    """End-to-end: storey magnifier δs is per-combo, not averaged across combos."""
    print("\n=== E2E B2 — Per-combo storey magnifier δs ===")
    Nc = 10_000.0  # kN, illustrative
    combos = [
        ("1.35G",          1080.0),
        ("1.2G+1.5Q",      1485.0),
        ("1.2G+1.5*psi*Q", 1170.0),
        ("1.2G+Wu+psi*Q",  1100.0),
        ("0.9G+Wu",         720.0),
        ("G+Eu+psi*Q",      940.0),
    ]
    print(f"  Nc = {Nc:.0f} kN")
    results = []
    for name, NStar in combos:
        ds_correct = 1.0 / (1.0 - NStar / Nc)
        ds_engine = moment_magnifier_unbraced_storey(
            sum_N_star=NStar, sum_Nc=Nc
        )
        ok = abs(ds_correct - ds_engine) / ds_correct < 0.001
        print(f"  {name:18s}: N*={NStar:>5.0f}  δs={ds_engine:.3f}  (expected {ds_correct:.3f})  {'PASS' if ok else 'FAIL'}")
        results.append(ok)
    return all(results)


# ─────────────────────────────────────────────────────────────────────
# PRO validation tests (engine extensions for the comprehensive build)
# ─────────────────────────────────────────────────────────────────────

def pro_voids_example_5_1():
    """RCB Example 5.1 — 600×800 column with 8N32 + 150 mm dia void at d=500.
    Expected dpc ≈ 397 mm. Engine Nuo differs from book by ~1% because the
    book uses Ag (not Ag-As) and σs=fsy without the 500 MPa cap."""
    print("\n=== PRO ── Voids — RCB Example 5.1 ===")
    sec = Section(
        shape="rect", D=800.0, b=600.0, fc=40.0, fsy=500.0,
        layers=[
            RebarLayer(d_from_comp_face= 66.0, area=3*800.0),
            RebarLayer(d_from_comp_face=400.0, area=2*800.0),
            RebarLayer(d_from_comp_face=734.0, area=3*800.0),
        ],
        voids=[Void(d_from_comp_face=500.0, diameter=150.0)],
    )
    results = []
    results.append(check("dpc with void", sec.plastic_centroid(), 397, tol_rel=0.005, unit="mm"))
    # Sanity — without void, dpc = D/2 = 400
    secNoVoid = Section(
        shape="rect", D=800.0, b=600.0, fc=40.0, fsy=500.0,
        layers=sec.layers,
    )
    results.append(check("dpc without void = D/2", secNoVoid.plastic_centroid(), 400, tol_rel=0.001, unit="mm"))
    return all(results)


def pro_special_region_5_7():
    """RCB Example 5.7 — special-region length from top end ≈ 586 mm."""
    print("\n=== PRO ── Special-region length — RCB Example 5.7 ===")
    x = special_region_length(M_high=+225, M_other=-190, L=3200, D=350, M_thresh=149)
    results = []
    results.append(check("Top-end zone length", x, 586, tol_rel=0.01, unit="mm"))
    # Bottom end: M_high becomes -190, M_other becomes +225 → also needs distance
    # from bot end. By symmetry of the calc with signs flipped, distance from bot
    # = (190-149)/(225+190) * 3200 = 316 mm. Lower bound 1.2D = 420.
    x2 = special_region_length(M_high=-190, M_other=+225, L=3200, D=350, M_thresh=149)
    results.append(check("Bot-end zone length (1.2D bound)", x2, 420, tol_rel=0.005, unit="mm"))
    return all(results)


def pro_splice_provisions():
    """Cl 10.7.5 — splice mode logic."""
    print("\n=== PRO ── Splice provisions Cl 10.7.5 ===")
    results = []
    cases = [
        ('lap',   1.00, False, True,  'full lap'),
        ('lap',   0.50, False, True,  '50% lap'),
        ('lap',   0.30, False, True,  '30% lap (just satisfies 25%)'),
        ('lap',   0.15, False, False, '15% lap fails 25% min'),
        ('mech',  1.00, False, True,  'mechanical'),
        ('endb',  1.00, False, True,  'end-bearing, no tension ever'),
        ('endb',  1.00, True,  False, 'end-bearing, face goes tension — must FAIL'),
        ('none',  1.00, True,  True,  'no splice'),
    ]
    for typ, ratio, tens, expectOk, label in cases:
        res = check_splice(typ, ratio, tens)
        ok = (res['ok'] == expectOk)
        status = "PASS" if ok else "FAIL"
        print(f"  {status}  {label:50s} mode={res['mode']:12s} ok={res['ok']} expected={expectOk}")
        results.append(ok)
    return all(results)


def pro_restraint_pattern():
    """Cl 10.7.4.1 — restraint pattern triggers."""
    print("\n=== PRO ── Lateral restraint pattern Cl 10.7.4.1 ===")
    results = []
    tests = [
        # s, N* (N), fc, Ag (mm²), expected
        (120, 1_000_000,  40, 240_000, 'alternate-bar', 'low s, low N*'),
        (160, 1_000_000,  40, 240_000, 'every-bar',     's > 150'),
        (120, 4_000_000,  40, 240_000, 'every-bar',     'N* > 0.3 fc Ag'),
    ]
    for s, N, fc, Ag, expected, label in tests:
        out = restraint_pattern_required(s, N, fc, Ag)
        ok = (out == expected)
        status = "PASS" if ok else "FAIL"
        print(f"  {status}  {label:30s} → {out} (expected {expected})")
        results.append(ok)
    return all(results)


def pro_joint_transmission():
    """Cl 10.8 — fce calc."""
    print("\n=== PRO ── Floor-joint Cl 10.8 ===")
    results = []
    # Slab fc' >= 0.75 col fc' → no check
    j1 = joint_transmission_check(40, 32, 400, 400, '4-sides')
    results.append(check("slab fc=32 >= 0.75*40 → no check", 1 if not j1['applies'] else 0, 1))
    # Strict ratio fce = 0.65*col for 2-adjacent
    j2 = joint_transmission_check(80, 20, 400, 400, '2-adjacent')
    results.append(check("2-adjacent fce = 0.65 * 80", j2['fce'], 52, tol_rel=0.005))
    return all(results)


def pro_biaxial_concession():
    """Cl 10.6.3 — concession applies only when aspect ≤ 3 AND one ratio ≤ 0.06."""
    print("\n=== PRO ── Biaxial concession Cl 10.6.3 ===")
    results = []
    c1 = biaxial_concession(400, 600, 0.05, 0.5)
    results.append(check("aspect 1.5, min 0.05 → applies", 1 if c1['applies'] else 0, 1))
    c2 = biaxial_concession(400, 1500, 0.05, 0.5)
    results.append(check("aspect 3.75 → does NOT apply", 1 if not c2['applies'] else 0, 1))
    c3 = biaxial_concession(400, 600, 0.30, 0.50)
    results.append(check("both > 0.06 → does NOT apply", 1 if not c3['applies'] else 0, 1))
    return all(results)


def pro_phi_tension():
    """Cl Table 2.2.2 row (c) — interpolation φ' → 0.85 at pure tension."""
    print("\n=== PRO ── φ for axial tension + bending ===")
    results = []
    phiPrime = 0.85
    Nuot = 1200
    # At Nu = 0 → φ = φ'
    results.append(check("φ at N=0", phi_tension(0, Nuot, phiPrime), 0.85))
    # At Nu = -Nuot/2 → φ = 0.85 (since φ' = 0.85 already)
    results.append(check("φ at N=-Nuot/2 (φ'=0.85)", phi_tension(-Nuot/2, Nuot, phiPrime), 0.85))
    # With φ' = 0.65 → at -Nuot, φ = 0.85
    results.append(check("φ at N=-Nuot (φ'=0.65)", phi_tension(-Nuot, Nuot, 0.65), 0.85))
    # Halfway: φ = 0.65 + 0.5·0.20 = 0.75
    results.append(check("φ at N=-600 (half, φ'=0.65)", phi_tension(-Nuot/2, Nuot, 0.65), 0.75))
    return all(results)


if __name__ == "__main__":
    print("=" * 72)
    print("AS 3600:2018 COLUMN ENGINE — VALIDATION AGAINST RCB 3e EXAMPLES")
    print("=" * 72)
    out = []
    out.append(("5.2", section_5_2()))
    out.append(("5.4", section_5_4()))
    out.append(("5.5", section_5_5()))
    out.append(("5.6", section_5_6()))
    out.append(("5.7", section_5_7()))
    print("\n" + "=" * 72)
    print("END-TO-END TESTS (orchestration / fix verification)")
    print("=" * 72)
    out.append(("E2E B1 — fsy cap",            e2e_b1_fsy_cap()))
    out.append(("E2E B3 — axis r",             e2e_b3_axis_slenderness()))
    out.append(("E2E B2 — per-combo δs",       e2e_b2_per_combo_ds()))
    print("\n" + "=" * 72)
    print("PRO TESTS (comprehensive coverage of RCB Chapter 5)")
    print("=" * 72)
    out.append(("PRO Voids — Ex 5.1",          pro_voids_example_5_1()))
    out.append(("PRO Special region — Ex 5.7", pro_special_region_5_7()))
    out.append(("PRO Splices",                 pro_splice_provisions()))
    out.append(("PRO Restraint pattern",       pro_restraint_pattern()))
    out.append(("PRO Joint Cl 10.8",           pro_joint_transmission()))
    out.append(("PRO Biaxial concession",      pro_biaxial_concession()))
    out.append(("PRO φ tension",               pro_phi_tension()))
    print("\n" + "=" * 72)
    print("OVERALL VALIDATION SUMMARY")
    print("=" * 72)
    for name, ok in out:
        print(f"  {name}: {'PASS' if ok else 'FAIL'}")
    n_pass = sum(1 for _, ok in out if ok)
    print(f"\n  {n_pass}/{len(out)} test groups passing")
