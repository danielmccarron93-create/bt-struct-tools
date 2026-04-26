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
from as3600_columns import (
    Section, RebarLayer,
    alpha_1, alpha_2, gamma,
    phi_bending_only, phi_compression,
    radius_of_gyration, is_short_column,
    effective_length_factor_braced, effective_length_factor_unbraced,
    buckling_load_Nc, moment_magnifier_braced, km_factor,
    moment_magnifier_unbraced_storey,
    biaxial_utilisation,
    HSCConfinementInput, hsc_fitment_spacing_simplified,
    hsc_fitment_spacing_deemed,
    capacity_at_N_star,
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
    # Point A — squash
    Nuo = sec.nuo() / 1000.0
    results.append(check("A: Nuo (squash)", Nuo, 9360, unit="kN"))

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
    print("OVERALL VALIDATION SUMMARY")
    print("=" * 72)
    for name, ok in out:
        print(f"  Example {name}: {'PASS' if ok else 'FAIL'}")
    n_pass = sum(1 for _, ok in out if ok)
    print(f"\n  {n_pass}/{len(out)} examples passing")
