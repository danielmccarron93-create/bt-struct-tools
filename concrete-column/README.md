# Concrete Column Designer — AS 3600:2018

A single-file HTML application for designing reinforced concrete columns to AS 3600:2018.

## Contents

- **`concrete_column_designer.html`** — the application. Open in any modern browser (no install, no backend, works offline). Drag-and-drop or double-click to launch.
- **`as3600_columns.py`** — the Python reference implementation of the engine, used to develop and validate the calculation logic.
- **`validate_engine.py`** — Python validation script that reproduces the worked examples in *Reinforced Concrete Basics* 3e (Gilbert/Mickleborough/Ranzi/Foster, 2021).
- **`validate_js_port.js`** — Node.js validation that the JS engine in the HTML matches the Python engine.
- **`validation_report.txt`** — log of the most recent validation run (Python + JS).

## Validation summary

The engine has been validated against five worked examples in RCB 3e Chapter 5,
plus three end-to-end orchestration tests covering the fixes from the April 2026
technical review:

| Test | Topic | Status |
|---|---|---|
| 5.2 | Section capacity line — five key points (squash, decompression, balanced, pure bending, pure tension) | PASS |
| 5.4 | Biaxial bending check (Cl 10.6.4) | PASS |
| 5.5 | Slender unbraced column — moment magnifier (δb, δs) | PASS |
| 5.6 | HSC core confinement — high axial (Cl 10.7.3) | PASS |
| 5.7 | HSC core confinement — moderate axial, high moment | PASS |
| E2E B1 | Squash-load 500 MPa cap binds at fsy=600 (Cl 10.6.2.2) | PASS |
| E2E B3 | Per-axis radius of gyration for rectangular columns (Cl 10.5.2) | PASS |
| E2E B2 | Storey magnifier δs computed per combo, not averaged | PASS |

### Standard-vs-textbook reconciliations

1. **Cl 10.6.2.2 (squash steel stress)** — at squash, max steel strain = 0.0025 → max stress = 500 MPa even for 600 MPa bars. RCB Eq 5.6 uses `Ag` (not `Ag − As`) which the engine does not. Engine Nuo is ~0.9 % below RCB for typical ρ = 1 %.
2. **Cl 10.4.4 (buckling load Nc)** — RCB Example 5.5 used φ = 0.6 (slender k_φ reduction) where AS 3600 hard-wires φ = 0.65. The engine uses 0.65.
3. **Cl 10.6.4 (biaxial αn)** — RCB Eq 5.21 includes an extra 0.65 factor in the αn formula that AS 3600 does not include. The engine follows AS 3600 (more conservative).

All three are documented in the print-PDF assumptions section.

### Fixes applied April 2026 (post technical review)

- **B1** — squash steel stress capped at 0.0025·Es per Cl 10.6.2.2 (fixes ~2.6 % over-prediction with 600 MPa bars).
- **B2** — δs computed per combination, not averaged across all six combos.
- **B3** — separate r_x, r_y, Nc_x, Nc_y for non-square rectangular columns; weak-axis governance now captured.
- **B4** — wind moment top and bottom entered separately (was a single value applied to both ends, which forced same-sign = single-curvature behaviour incorrectly).
- **B5** — φo selection uses the proper Cl 10.3.1 short-column criterion across all combos (was a fixed `Le/r ≤ 25` proxy).
- **B7** — durability cover is now a 2-D lookup of exposure × fc′ per AS 3600:2018 Table 4.10.3.2.
- **C7** — δs > 1.5 (Cl 10.4.3 stiffening trigger) flagged in the slenderness tile.
- **C8** — Cl 10.7.3.1 special-confinement-region trigger checked before applying HSC spacing limits.
- **C10** — automatic check that fitment-bar diameter meets Table 10.7.4.3 minimum.

## Scope

**v1 covers:**
- Rectangular, square, and circular columns
- Material grades fc′ = 25–100 MPa, including HSC core confinement (Cl 10.7.3)
- Both simplified (Cl 10.7.3.3) and deemed-to-comply (Cl 10.7.3.4) confinement methods
- Braced and unbraced columns
- Manual or γ-derived effective length factor k
- Auto-generated AS/NZS 1170.0 ULS load combinations (six standard combos)
- Live N–M interaction diagram, biaxial contour, strain profile
- Three-view geometry (3D rotatable, 2D section, 2D elevation)
- Detailing checks (longitudinal %, fitment spacing, lateral restraint, cover, fire)
- Print-to-PDF calculation report with full AS 3600 clause references and audit hash

**v1 does not cover:**
- Prestressed columns (AS 3600 Section 13)
- Composite steel-concrete columns (different standard)
- Frame analysis — design loads (G, Q, W, E moments and axial) are user-entered
- Section 14 earthquake-detailing capacity design (use a dedicated tool)

## Known limitations

- Moment magnifier δs is computed for the column treated as a single-column storey. For multi-column storey magnification, use a hand calculation with ΣN\*/ΣNc.
- Effective length factor `k` and unsupported length `Lu` are entered once and applied to both axes. For frames where x-axis and y-axis k or Lu differ, evaluate each axis manually.
- Cover lookup for fire (Cl 5.5) assumes a "compression member with ≥ 3 sides exposed and load level μfi ≤ 0.7" — for other geometries refer to AS 3600 Table 5.5.2 directly.
- Cc uses the gross stress-block area (does not subtract concrete displaced by compression bars). Common simplification per the rectangular stress block method; over-predicts capacity by ~0.5–1 % at ρ ≈ 1 % (proportional to ρ).

## Usage tips

- Click "Example" in the top bar to load RCB Example 5.5 inputs as a starting point.
- "Save" exports a JSON of all inputs; "Load" re-imports them. Use this for project archiving.
- Click any check tile to toggle a detail panel.
- The interaction diagram shows both Mu (dashed grey) and φMu (solid black) curves with all load combinations plotted; the governing combo is highlighted orange.
- The PDF includes a hash of all inputs in the header for audit/version-control purposes.

---

*Engine v1.0 — built and validated April 2026. AS 3600:2018 incorporating Amendment No. 1.*
