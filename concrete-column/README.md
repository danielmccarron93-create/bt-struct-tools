# Concrete Column Designer — AS 3600:2018

A single-file HTML application for designing reinforced concrete columns to AS 3600:2018.

## Contents

- **`concrete_column_designer.html`** — the application. Open in any modern browser (no install, no backend, works offline). Drag-and-drop or double-click to launch.
- **`as3600_columns.py`** — the Python reference implementation of the engine, used to develop and validate the calculation logic.
- **`validate_engine.py`** — Python validation script that reproduces the worked examples in *Reinforced Concrete Basics* 3e (Gilbert/Mickleborough/Ranzi/Foster, 2021).
- **`validate_js_port.js`** — Node.js validation that the JS engine in the HTML matches the Python engine.
- **`validation_report.txt`** — log of the most recent validation run (Python + JS).

## Validation summary

The engine has been validated against five worked examples in RCB 3e Chapter 5:

| Example | Topic | Status |
|---|---|---|
| 5.2 | Section capacity line — five key points (squash, decompression, balanced, pure bending, pure tension) | PASS |
| 5.4 | Biaxial bending check (Cl 10.6.4) | PASS |
| 5.5 | Slender unbraced column — moment magnifier (δb, δs) | PASS |
| 5.6 | HSC core confinement — high axial (Cl 10.7.3) | PASS |
| 5.7 | HSC core confinement — moderate axial, high moment | PASS |

Two minor numerical inconsistencies were identified between the textbook and AS 3600:2018; in both cases the engine follows the standard:

1. **Cl 10.4.4 (buckling load Nc)** — RCB Example 5.5 used φ = 0.6 (slender k_φ reduction) where AS 3600 hard-wires φ = 0.65. The engine uses 0.65, giving Nc values approximately 8% higher than the book.
2. **Cl 10.6.4 (biaxial αn)** — RCB Eq 5.21 includes an extra 0.65 factor in the αn formula that AS 3600 does not include. The engine follows AS 3600, which is more conservative (smaller αn, larger utilisation).

Both are documented in the print-PDF assumptions section.

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
- Buckling load Nc is computed using the radius of gyration in the x-axis direction; for rectangular columns where the y-axis governs slenderness, this can be conservative on the x-axis but should be verified for sensitive cases.

## Usage tips

- Click "Example" in the top bar to load RCB Example 5.5 inputs as a starting point.
- "Save" exports a JSON of all inputs; "Load" re-imports them. Use this for project archiving.
- Click any check tile to toggle a detail panel.
- The interaction diagram shows both Mu (dashed grey) and φMu (solid black) curves with all load combinations plotted; the governing combo is highlighted orange.
- The PDF includes a hash of all inputs in the header for audit/version-control purposes.

---

*Engine v1.0 — built and validated April 2026. AS 3600:2018 incorporating Amendment No. 1.*
