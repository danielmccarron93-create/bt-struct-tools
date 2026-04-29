# Blockwork Wall Designer — AS 3700:2018

A single-file HTML application for designing reinforced (core-filled) concrete blockwork walls to AS 3700:2018.

## Contents

- **`blockwork_wall_designer.html`** — the application. Open in any modern browser (no install, no backend, works offline).
- **`as3700_walls.py`** — Python reference implementation of the engine, used to develop and validate the calculation logic.
- **`validate_engine.py`** — Python validation script: 13 tests, 70 assertions covering the example wall + edge cases.
- **`validate_js_port.js`** — Node validation that the JS engine in the HTML matches the Python engine.
- **`validation_report.txt`** — log of the most recent validation run (Python).

## Validation summary

The engine has been validated against:

| Test | Topic | Status |
|---|---|---|
| 1 | Example wall (190·2000·5000, f'uc=15, αr=0.40) — matches `Blockwork_Wall_Check.xlsx` | PASS (24 assertions) |
| 2 | kes formula bounds [0..1] | PASS |
| 3 | Pure tension wall (Cl 8.10) | PASS |
| 4 | Long-wall shear, H/L < 2.3 (Cl 8.7.2 with **Ad** not Ag) | PASS |
| 5 | Cl 8.8 cap actively governing | PASS |
| 6 | Sr selection rule (Cl 7.3.4.3a Fd/Fo ≤ 0.20) | PASS |
| 7 | αr = 0 vs αr = 0.40 — steel contribution to axial | PASS |
| 8 | Multi-LC mode, governing pick | PASS |
| 9 | Static vs Area axial worst-case logic | PASS |
| 10 | All detailing checks (Cl 8.5.1f, 8.7.2, 8.10, 8.4.5) | PASS |
| 11 | Cl 8.7.4 stability check (units fixed vs L7) | PASS |
| 12 | Cl 8.6 explicit reinforced flexure | PASS |
| 13 | Very slender wall — kes near zero | PASS |

**Python:** 70/70 assertions PASS · **JS port:** 37/37 assertions PASS (matches Python to 0.1%).

## Scope

**v1 covers (live):**
- Compression — Eq 8.5.1 with full strain-compatibility-equivalent kes treatment
- Tension — Eq 8.10
- Shear — Cl 8.7.2 (long, H/L ≤ 2.3) and Cl 8.7.3 → Cl 8.8 (short)
- **Cl 8.8 cap (4·φ·f'vm·bw·d) explicitly enforced** — code requirement that L7 omits
- Cl 8.7.4 stability check (toggle for unrestrained walls)
- Cl 8.6 explicit reinforced flexure (toggle)
- Cl 8.4.5 close-spaced earthquake reo (toggle)
- All Cl 8.7.2 / 8.10 detailing checks always enforced

**Two load modes:**
- **Manual** — single-case quick check (default)
- **Spreadsheet upload** — drop an XLSX matching L7's Input sheet column order; engine processes every LC, picks Static/Area worst per LC, picks worst LC overall

**Visualisations:**
- Wall plan section (face shells, grouted cores, vertical bars, eccentricity arrow)
- Wall elevation (vertical + horizontal reo, support symbols, dimensions)
- Stress block diagram (P/A ± M/Z, compression length lc highlighted)
- Capacity bar chart (compression / tension / shear demand vs capacity)
- N–M envelope (kes-based, Cl 8.5.1)

**Print PDF:** A4 portrait, BT-branded, 7 pages, audit hash, every step traceable to a spreadsheet row.

## Engineering notes

The validated engine (`as3700_walls.py`) is the single source of truth. The HTML's JS engine matches it to within 0.1% across all test cases.

A few intentional engineering choices documented in the HTML method banner:

- **αr default = 0.40** (per Cl 8.5.1 for walls). L7 hard-wires αr = 0 — switch the chip to "0 (L7-conservative)" to match L7 outputs.
- **kes bounded to [0, 1]** — if the formula gives a negative value, kes is clamped to 0 (capacity = 0 → escalate to Cl 8.11.1 strain compatibility).
- **Sr selection rule per Cl 7.3.4.3(a)** — uses MIN(Sr1, Sr2) only when wall has lateral edge support AND Fd ≤ 0.20·Fo, else Sr1 only. L7 uses MAX (slightly conservative).
- **Bending check defaults to "kes-in-axial"** (eccentricity captured in Cl 8.5.1 only). Toggle "explicit Cl 8.6" to additionally run reinforced flexure capacity per Eq 8.6 and the Cl 8.11.1 combined-loading requirement.

## Usage tips

- Click "Example" in the top bar to load the L7 example wall as a starting point.
- "Save" exports a JSON of all inputs; "Load" re-imports them. Use this for archiving.
- Click any check tile (Compression / Tension / Shear) to expand a step-by-step detail panel.
- The N–M envelope shows all LCs as dots with the governing case highlighted orange.
- The PDF includes an audit hash of all inputs in the header — useful for version control.
- Dark mode toggle (top-right) for evening work.

## Development

To re-validate after editing the engine:

```bash
# Python
python3 validate_engine.py
# JS port
node validate_js_port.js
```

Both must report 0 failures.

---

*Engine v1.0 — built and validated April 2026. AS 3700:2018.*
