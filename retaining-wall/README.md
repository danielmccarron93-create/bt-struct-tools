# Retaining Wall Designer — AS 4678 / AS 3600

A single-file HTML application for designing reinforced concrete cantilever retaining walls. Theory follows Foster, Kilpatrick & Warner, *Reinforced Concrete Basics* 3e §6.3 (Pearson Australia, 2021), with stability framework per AS 4678 + AS/NZS 1170.0 and component design per AS 3600.

## Contents

- **`index.html`** — the application. Open in any modern browser (no install, no backend, works offline).
- **`validate.js`** — Node validation harness reproducing Worked Example 6.3 from RCB 3e (pp 425–428) numerically. Run with `node validate.js`.

## Validation summary

Engine validated against the chapter's full Worked Example 6.3 (4 m cantilever wall on gravelly clay, surcharge 5 kPa, gravel backfill).

| Quantity                | Engine | Textbook | Status |
|---|---:|---:|---|
| k_a (Rankine–Bell, β=0, φ=30°) | 0.333 | 0.33  | PASS |
| W₁ stem self-weight (kN/m)     | 30.0  | 30.0  | PASS |
| W₂ base self-weight            | 22.5  | 22.5  | PASS |
| W₃ soil over heel              | 142.8 | 142.8 | PASS |
| Surcharge over heel S          | 10.5  | 10.5  | PASS |
| R = ΣV                         | 205.8 | 205.8 | PASS |
| P_A1                           | 52.4  | 51.9  | PASS |
| P_A2                           | 7.17  | 7.1   | PASS |
| Eccentricity e (m)             | 0.21  | 0.21  | PASS |
| q_max (kPa)                    | 97.9  | 97    | PASS |
| q_min (kPa)                    | 39.3  | 40    | PASS |
| Overturning util               | 0.40  | 0.40  | PASS |
| Sliding util                   | 0.95  | 0.96  | PASS¹ |
| Stem M* (kNm/m)                | 110.7 | 110   | PASS |
| Toe M* (kNm/m)                 | 23.4  | 24.7  | PASS² |

¹ 1% diff: book rounds δ = 0.75·φ to 22° (engine: 22.5° exactly).
² Engine subtracts base self-weight from toe upward bearing; book omits this — engine value is more accurate but slightly less conservative. Difference is within engineering tolerance.

## Scope

**Covers:**
- Cantilever RC retaining walls with optional shear key
- Rankine–Bell active pressure, level or sloping fill (β > 0)
- At-rest pressure (k₀ = 1.0) for basement / propped walls
- AS 4678 Class A/B/C with auto-imposed minimum surcharge (1.5 / 2.5 / 5 kPa)
- Stability per Eq 6.20 simplified: 0.9·G_stb ≥ γ_e·F_e + ψ_c·Q_d,dst
- Bearing pressure: trapezoidal (e ≤ L/6) or triangular partial contact (e > L/6)
- Reinforcement design via the empirical p ≈ 2.5·M*/d² shortcut + ρ_min and ρ_shrinkage per AS 3600
- Stem flexure & shear, heel flexure, toe flexure
- Live update of section, force diagram, reinforcement, bearing-pressure plot, active-pressure prism
- PDF report with full free-body table, stability calc, component design, and audit hash
- Save/load JSON of all inputs

**Does not cover:**
- Global slope-stability (Fig 6.11a) — use a specialist tool (Aysen 2005 / Slide / Plaxis)
- Bearing-capacity analysis from soil parameters (q_a is user-input)
- Counterfort or buttress walls (Fig 6.12f, g)
- Reinforced-earth walls (Fig 6.12i)
- Section 14 earthquake-detailing capacity design
- Construction-stage compaction-induced lateral pressure

## Usage tips

- Click **Example 6.3** in the top bar to load the textbook problem.
- Click **Auto-size** to apply H/14 thickness rule + L = 0.6H starting trial dimensions.
- Click any check tile to toggle a detail panel.
- Switch between **Section / Forces / Reinforcement** tabs to see the wall, free-body diagram, or bar layout.
- Switch between **Bearing / Active / Stem M** chart tabs to inspect the pressure or moment distributions.
- The PDF report has a hash of all inputs in the header for audit/version control.

## Theory references

- AS 4678–2002 — Earth-retaining structures (master document)
- AS/NZS 1170.0–2002 — Structural design actions: general principles, stability checks
- AS/NZS 1170.1–2002 — Permanent, imposed and other actions
- AS 3600 — Concrete structures
- RCB 3e §6.3 (Foster, Kilpatrick & Warner, 2021) — pp 412–428

---

*Engine v1.0 — built April 2026. Validated against RCB 3e Worked Example 6.3.*
