# Concrete Column Design to AS 3600:2018 — Learnings

A reference document capturing everything learned while building the Concrete Column Designer tool. Covers the AS 3600:2018 design methodology, cross-references with *Reinforced Concrete Basics* 3e (Gilbert/Mickleborough/Ranzi/Foster, 2021), worked validation cases, two textbook-vs-standard inconsistencies, and engineering judgements made during the build.

> Sources: AS 3600:2018 (Section 10 plus supporting clauses), *Reinforced Concrete Basics* 3rd edition Chapter 5 (RCB 3e). Both PDFs reside in the project library.

---

## 1. The design problem in one paragraph

A column resists combined axial force `N*` and bending moments `M*x`, `M*y`. Bending exists even in nominally axial columns due to frame moments, eccentricity from imperfections, and unbalanced gravity. AS 3600 Cl 10.1.2 enforces a **minimum eccentricity** of `0.05 D` on each axis — `M*min = N* × 0.05 D` — which you cannot design below regardless of analysis output. Design is almost entirely strength limit state; serviceability rarely governs columns. The controlling inequalities are `φNu ≥ N*` and `φMu ≥ M*` acting *simultaneously* on the N–M interaction surface.

---

## 2. Two early classifications that drive everything else

### Braced vs unbraced (Cl 10.1.3.1, RCB §5.2.1)

A **braced** column relies on shear walls, cores, or bracing to resist storey-wide lateral actions. An **unbraced** column resists those actions through column moments — sway between top and bottom is free.

This drives whether the storey magnifier `δs` (Cl 10.4.3) is applied on top of the per-column `δb`.

### Short vs slender (Cl 10.3.1, RCB §5.5.2)

A column is **short** if:

- Braced: `Le/r ≤ max(25, αc·(38 − fc′/15)·(1 + M*1/M*2))`
- Unbraced: `Le/r ≤ 22`

with:

- `αc = √(2.25 − 2.5·N*/(0.65·Nuo))` for `N*/Nuo ≥ 0.15`
- `αc = √(1 / (3.5·N*/(0.65·Nuo)))` for `N*/Nuo < 0.15`

Radius of gyration `r = 0.3 D` (rectangular) or `0.25 D` (circular) — Cl 10.5.2.

> **Sign trap:** `M1/M2` is **negative for single curvature**, **positive for double curvature**. This is the opposite of what many engineers assume; getting it wrong inflates `km` and over-magnifies the moment.

`Le/r ≤ 120` is a hard maximum (Cl 10.5.1) — beyond that, only rigorous nonlinear analysis is permitted.

---

## 3. Effective length factor `k`

Cl 10.5.3 — `Le = k·Lu`. Three routes to `k`:

1. **Simple end conditions** — read `k` from Fig 10.5.3(A) (e.g. `k = 1.0` for pinned-pinned, `0.7` for fixed-fixed, etc.).
2. **Framed structure** — compute end restraint coefficients `γ1`, `γ2` and look up Fig 10.5.3(B) braced or 10.5.3(C) unbraced.
3. **Critical buckling load analysis** — derive `Le` from a frame-level eigenvalue analysis.

For framed structures (Cl 10.5.4):

```
γ = Σ(I/L)c / Σ(β · I/L)b
```

with `β` from Table 10.5.4:

| Far-end fixity of beam/slab | Braced | Unbraced |
|---|---|---|
| Pinned | 1.5 | 0.5 |
| Rigid to column | 1.0 | 1.0 |
| Fixed | 2.0 | 0.67 |

For analysis stiffness, RCB suggests `0.8 Ig` for columns and `0.4 Ig` for beams/slabs in this calculation (matching the higher-end values of Table 6.2.4). Don't confuse this with the `Ieff` used in the analysis model — those serve different purposes.

For a footing (Cl 10.5.6): `γ = 10` (negligible rotational restraint) or `γ = 1.0` (designed fixed).

The closed-form fits I used in the engine (Wood/Galambos) are:

```
k_braced   = (1 + 0.145(γ1+γ2) − 0.265 γ1 γ2)
           / (2 − 0.364(γ1+γ2) − 0.247 γ1 γ2)        bounded 0.5 ≤ k ≤ 1.0

k_unbraced = √((1.6 γ1 γ2 + 4(γ1+γ2) + 7.5) / (γ1+γ2+7.5))     ≥ 1.0
```

These match the AS 3600 alignment charts to within ~3% across the practical range.

---

## 4. The N–M interaction diagram — the heart of the check

Cl 10.6 and RCB §5.4 agree on the assumptions:

- Plane sections remain plane.
- Concrete has zero tensile strength.
- `εcu = 0.003` at the extreme compressive fibre when `ku ≤ 1`.
- Equivalent rectangular stress block (Cl 10.6.2.5):
    - `α2 = 0.85 − 0.0015 fc′`, bounded `0.67 ≤ α2 ≤ 0.85`
    - `γ = 0.97 − 0.0025 fc′`, bounded `0.67 ≤ γ ≤ 0.85`
    - For circular or width-tapering sections, reduce `α2` by 5% (Cl 10.6.2.5 Note 3).
- For squash load `Nuo` (Cl 10.6.2.2): uniform concrete stress `α1·fc′` with:
    - `α1 = 1.0 − 0.003 fc′`, bounded `0.72 ≤ α1 ≤ 0.85`
    - Maximum steel strain at squash = 0.0025.

> **Watch out:** The AS 3600 0.9 fc′ in-situ modifier (Cl 3.1.4) is **already baked into** `α1` and `α2`. Don't apply it twice.

### Five key points on the diagram (RCB §5.3.3)

| Point | Description | How to compute |
|---|---|---|
| **A** — Squash | `e = 0`, uniform compressive strain at failure | `Nuo = α1·fc′·(Ag − Ast) + Ast·fsy` |
| **B** — Decompression | `ku = do/d = 1.0`, neutral axis at outermost tensile bar | Compute `Cc`, `Cs`, `T` using rectangular block at `dn = do` |
| **C** — Balanced | Tension steel yields *simultaneously* with concrete crushing at 0.003 | `kub = 0.003 / (0.003 + εsy)`. For `fsy = 500 MPa`, `kub = 0.545` |
| **D** — Pure bending | `N = 0`, `Muo` | Iterate `dn` until `Cc + Cs − T = 0` |
| **E** — Pure tension | Concrete cracked, only steel resists | `Nuo,t = (Asc + Ast)·fsy` |

### Computing the diagram

- **A → B segment** (neutral axis outside the section): the rectangular stress block doesn't strictly apply. Cl 10.6.2.4 permits **linear interpolation** between A and B. This is a conservative approximation; rigorous integration of the curvilinear concrete stress is more accurate but rarely worth the complexity.
- **B → D segment**: iterate `dn` from `do` down to the pure-bending value. At each `dn`:
    - `Cc = α2·fc′·b·γ·dn` (or integrated for circular)
    - For each steel layer at depth `d_i`: `εi = 0.003·(dn − d_i)/dn`, then stress = bounded by `±fsy`
    - `Nu = Cc + ΣFs`
    - `Mu = Cc·(dpc − γdn/2) + Σ Fi·(dpc − di)`
- **D → E segment**: linear interpolation again per Cl 10.6.2.4 implicitly.

### Plastic centroid (RCB §5.4.2)

For asymmetric reinforcement or sections with voids, the plastic centroid is **not** the geometric centroid. It's the point about which the squash-load resultant acts:

```
dpc = [α1·fc′·Ac·(centroid of concrete) + Σ Asi·fsy·di] / Nuo
```

For symmetric sections (equal As top and bottom, no voids), this just lands on the geometric centroid. For everything else, compute it explicitly — otherwise `Mu` is wrong.

---

## 5. Capacity reduction factor `φ` (Table 2.2.2)

This is the most nuanced part of the AS 3600 column workflow — `φ` is **not** a single number; it varies along the interaction curve.

### For bending with axial compression

- If `Nu ≥ Nub`: `φ = φo`
    - `φo = 0.65` for stocky columns with `Q/G ≥ 0.25`
    - `φo = 0.60` otherwise *(equivalent to 0.65 × kφ where kφ = 12/13 for slender or low-Q/G cases)*
- If `Nu < Nub`: linear interpolation
    - `φ = φo + (φ′ − φo)·(1 − Nu/Nub)`
- `φ′` is the pure-bending value: `φ′ = 1.24 − 13·kuo/12`, bounded `0.65 ≤ φ′ ≤ 0.85`

### For axial tension with bending

`φ = φ′ + (0.85 − φ′)·(Nu/Nuo,t)` — interpolates from `φ′` at pure bending up to `0.85` at pure tension.

### Practical implications

- The `φMu` curve **pinches inward** near the balanced point relative to the un-factored `Mu` curve (because `φ` is at its minimum there).
- The `φMu` curve **bulges outward** near pure bending (because `φ′` can reach 0.85 for under-reinforced sections).
- The φ-curve is piecewise — discontinuities of slope at the zone boundaries are normal and not a calc error.

---

## 6. Slenderness magnifier — Cl 10.4

For slender columns, the design moment becomes `M*max = δ·M*2`, where `δ` is a moment magnifier accounting for the P-Δ effect along the column length.

### Buckling load `Nc` (Cl 10.4.4)

```
Nc = (π² / Le²) · 182·do·φMub / (1 + βd)
```

where:

- `do` = depth to outermost tensile reinforcement layer
- `φMub` is the balanced-point moment **with φ = 0.65 fixed** (Cl 10.4.4 hard-wires this — don't apply slender k_φ here, see Inconsistency #1 below)
- `βd = G/(G + Q)` — sustained-load creep modifier; reduces stiffness for sustained loading

This formulation comes from MacGregor et al. 1975 (RCB §5.5.3), inferring `EI` from `Mub/κub` rather than integrating the cracked-section stiffness.

### Per-column magnifier `δb` (Cl 10.4.2)

```
δb = km / (1 − N*/Nc)         ≥ 1.0
km = 0.6 − 0.4·(M1/M2)        ≥ 0.4
```

If significant transverse loading exists between column ends: `km = 1.0`.

> **Iteration trap:** `Nc` depends on reinforcement ratio (through `Mub`), so there's an implicit iteration: pick `ρ`, compute `Nc`, magnify the moment, check the section, adjust `ρ`. Most cases converge in 1–2 iterations starting from `ρ = 0.01`.

### Storey magnifier `δs` (Cl 10.4.3, unbraced only)

```
δs = 1 / (1 − ΣN*/ΣNc)
```

summing over all columns in the storey. AS 3600 caps `δs ≤ 1.5` — over that, the frame must be stiffened. The single-column tool I built treats one column as the storey, which is conservative for true multi-column situations.

### Final design moment

```
M*max = δ · M*2,        where δ = max(δb, δs)
                        and  M*max ≥ 0.05·D·N* (minimum eccentricity)
```

---

## 7. Biaxial bending — Cl 10.6.4

For rectangular sections with aspect ratio ≤ 3 and the resultant load within a diamond around the centroid (Fig 10.6.3), each axis can be checked independently (Cl 10.6.3 concession).

Otherwise:

```
(M*x / φMux)^αn  +  (M*y / φMuy)^αn  ≤  1.0

αn = 0.7 + 1.7·N*/Nuo,    bounded  1 ≤ αn ≤ 2
```

`φMux` and `φMuy` are evaluated **at the same `N*`**. Each can have a different `φ` because the balanced force differs between axes.

---

## 8. HSC core confinement — Cl 10.7.3 (only when fc′ > 50 MPa)

Above 50 MPa, concrete is brittle and cover spalling is consequential. Special confinement regions need `fr,eff ≥ 0.01·fc′`.

### Special confinement regions (Fig 10.7.3.1(A)) are where:

- `N* ≥ 0.75 Nuo`, **OR**
- `N* ≥ 0.3 fc′·Ag` AND `M* ≥ 0.6 Mu`

In those regions, fitment spacing ≤ `min(0.6 Dc, 300 mm, Cl 10.7.4 limit)`.

### Two routes to designing the fitments

**Route 1 — Simplified (Cl 10.7.3.3):**

```
fr     = (m·Ab,fit·fsy,f·sin θ) / (ds·s)
ke     = effectiveness factor (0–1, depends on fitment arrangement)
fr,eff = ke · fr  ≥  0.01·fc′
```

For rectangular sections:

```
ke = (1 − n·w² / (6·Ac)) · (1 − s/(2·bc)) · (1 − s/(2·dc))
```

For circular sections: `ke = (1 − s/(2·ds))²`.

Solve for the maximum `s` iteratively (closed-form quadratic in `s` for rect).

**Route 2 — Deemed-to-comply (Cl 10.7.3.4):**

```
Rect:  s ≤ 15·n·Ab,fit·fsy,f / (fc′·bc)
Circ:  s ≤ 100·Ab,fit·fsy,f / (fc′·ds)
```

Faster, more conservative — typically gives ~20% tighter spacing than the simplified method.

### Practical observation

For `fc′ = 80 MPa` columns I checked: simplified gave `s ≈ 250 mm` for the high-axial example, deemed-to-comply gave `s ≈ 197 mm`. Both are well within the 15·db buckling limit, so the deemed-to-comply approach is usually fine for routine design.

---

## 9. Detailing — Cl 10.7

### Longitudinal reinforcement (Cl 10.7.1)

- `0.01 Ag ≤ As ≤ 0.04 Ag` (1–4%)
- Exception: oversize columns (`Asc·fsy > 0.15 N*`) can drop below 1%
- Bundles: max 4 bars, must be tied together

### Fitment diameter (Table 10.7.4.3)

| Longitudinal bar | Min fitment |
|---|---|
| ≤ N20 | R6 |
| N24–N28 | N10 |
| N32–N36 | N12 |
| N40 | N16 |
| Bundled | N12 |

### Fitment spacing (Cl 10.7.4.3)

- Single bars: `s ≤ min(Dc, 15·db)`
- Bundled: `s ≤ min(0.5·Dc, 7.5·db)`

### Lateral restraint (Cl 10.7.4.1)

- Every corner bar
- Every bar if spacing > 150 mm or if `N* > 0.3 Ag fc′`
- Every alternate bar if spacing ≤ 150 mm
- 135° hooks for external fitments (90° hooks only allowed for internal cross-ties under specific conditions)

### Splices (Cl 10.7.5)

- Min tensile capacity 0.25·fsy·As per face
- End-bearing splices acceptable in pure compression with sleeve + adjacent fitments

### Floor-joint transmission (Cl 10.8)

If slab `fc′ < 0.75 × column fc′`, you cannot rely on the slab strength — compute `fce` for the joint and add through-joint reinforcement. Three formulas depending on whether the column is restrained on 4, 2-opposite, or 2-adjacent sides.

---

## 10. Two textbook-vs-standard inconsistencies (engine follows AS 3600 in both cases)

These are the kind of subtle traps that catch you when you're working from a textbook but expected to design to a standard.

### Inconsistency 1 — Cl 10.4.4 buckling load `Nc`

RCB Example 5.5 used `φ = 0.6` in the `Nc` formula, citing the slender-column `kφ = 12/13` reduction. But:

- AS 3600 Cl 10.4.4 **explicitly fixes** `φ = 0.65` in the `Nc` formula.
- RCB's own Eq 5.35 states `φ = 0.65`.

So the Example 5.5 numerical worked solution is internally inconsistent with both AS 3600 and the book's own equation. The engine uses 0.65, giving `Nc` values about 8% higher than the book's worked solution. **The engine matches AS 3600.**

### Inconsistency 2 — Cl 10.6.4 biaxial `αn`

RCB Eq 5.21 includes a 0.65 factor in `αn`:

```
αn (RCB)     = 0.7 + 1.7·N* / (0.65·Nuo)
αn (AS 3600) = 0.7 + 1.7·N* / Nuo            ← no 0.65
```

For Example 5.4 (N* = 4000 kN, Nuo = 9360 kN):

- RCB formula: `αn = 1.82` → util = 0.95 (PASS)
- AS 3600 formula: `αn = 1.43` → util = 1.12 (FAIL!)

These give materially different answers. AS 3600 is more conservative (smaller `αn`, larger utilisation). **The engine follows AS 3600.**

If you have any older calculations following RCB's `αn`, they may be on the unconservative side relative to the current standard.

---

## 11. Validation cases — what we benchmarked against

The engine reproduces all five RCB worked examples in Chapter 5 to within 1% (allowing for the two documented inconsistencies above):

| Example | Topic | Coverage |
|---|---|---|
| 5.2 | Section capacity line | All five points (squash, decompression, balanced, pure bending, pure tension) |
| 5.4 | Biaxial bending | `αn` formula + interaction equation |
| 5.5 | Slender unbraced column | `Nc`, `δb`, `δs`, `km` |
| 5.6 | HSC confinement (high axial) | Simplified + deemed-to-comply spacings |
| 5.7 | HSC confinement (moderate axial, high moment) | Simplified + deemed-to-comply spacings |

Validation runs nightly via `validate_engine.py` (Python) and `validate_js_port.js` (Node). Both must pass 21/21 for any engine change.

---

## 12. Engineering judgements made during the build

These are decisions the standard doesn't directly answer; flagging them so they can be revisited if practice changes.

1. **A → B segment of the interaction diagram**: linear interpolation per Cl 10.6.2.4, conservative but defensible. Rigorous curvilinear integration is more accurate but adds complexity for marginal gain.

2. **Default `φo`**: 0.65 if column is stocky (Le/r ≤ 25) AND Q/G ≥ 0.25; otherwise 0.60. The slender vs stocky boundary is the same threshold as the short/slender check.

3. **Single-column storey approximation for δs**: the tool treats the column as the entire storey when computing `δs`. For multi-column storeys, the user should compute `ΣN*/ΣNc` by hand or use the per-column `δb` if `δb > δs`.

4. **`r` based on `D` only**: the tool uses `r = 0.3D` for the column overall, which is correct for x-axis bending. For rectangular columns where y-axis governs slenderness, this is conservative on the x-axis but may need refinement. v1.1 will compute per-axis `r`, `Le`, and `Nc`.

5. **Minimum eccentricity 0.05D**: applied as a floor on `M*max` *after* magnification, not as a floor on `M*2`. This is the standard interpretation and matches RCB.

6. **Cover auto-derivation**: from exposure class (Cl 4.10.3.2 Table 4.10.3.2) plus FRL (Section 5). The tool's lookup is simplified — for unusual concretes (light, high-strength), refer to the standard directly.

7. **HSC confinement is computed both ways and the tighter spacing wins**: rather than picking one method. Gives the user both numbers transparently.

---

## 13. Out-of-scope items (deliberately not in v1)

- **Prestressed columns** (AS 3600 Section 13) — different geometry assumptions, time-dependent prestress losses.
- **Composite steel-concrete columns** — different standard (AS/NZS 2327).
- **Section 14 earthquake capacity design** — limited-ductile detailing requirements; needs its own dedicated tool with hierarchy of strength checks.
- **Fibre-reinforced concrete columns** (Section 16) — emerging area; the SFRC residual tensile strength contribution to flexure isn't yet routine.
- **Frame analysis** — loads (G, Q, Wu, Eu and corresponding moments) are user-entered, derived externally from RAPT/Microstran/SpaceGass.

---

## 14. Useful reading & references

- AS 3600:2018 (and Amendment 1) — Sections 2, 4, 5, 6, 10, 14
- AS 3600 Supplement 1 (2021) — commentary
- AS/NZS 1170.0 — load combinations
- *Reinforced Concrete Basics* 3e (Pearson, 2021) — Chapter 5 for columns, Chapter 1 for fundamentals
- Foster, S.J. — "Design and Detailing of High Strength Concrete Columns" (UNICIV Report R375, UNSW 1999) — basis of the HSC confinement provisions in Cl 10.7.3
- Foster, S.J. — *Concrete in Australia*, V35 No.3 (2009) — practical guidance on HSC detailing
- MacGregor, Oelhafen & Hage (1975) — "A Re-Examination of the EI Value for Slender Columns" (ACI SP-50) — basis of the `Nc` formula
- Galambos, T.V. (1968) — *Structural Members and Frames* — basis of effective length factor charts
- Concrete Institute of Australia (2010) — *Reinforcement Detailing Handbook* (CIA-26-2010)

---

## 15. Workflow for the next column you design

1. **Establish bracing** — does the column rely on shear walls / cores? → braced. Otherwise → unbraced.
2. **Get loads** from frame analysis: G, Q, Wu, Eu axial and end moments about both axes.
3. **Apply minimum eccentricity** — `0.05·D·N*` is the floor.
4. **Effective length** — manual `k` from Fig 10.5.3(A) for simple end conditions, or `γ`-based lookup for framed structures.
5. **Slenderness check** — short or slender? Le/r vs limit per Cl 10.3.1.
6. **Pick a trial section** and reinforcement (start at ρ = 0.01).
7. **If slender**, compute `Nc`, then `δb` (and `δs` if unbraced). Multiply M2 by δ.
8. **Check section capacity** — for each load combo, compute `φMux` and `φMuy` at that combo's `N*`. Ratio gives the per-axis utilisation.
9. **Biaxial check** — `(M*x/φMux)^αn + (M*y/φMuy)^αn ≤ 1`.
10. **Detailing checks** — ρ within 1–4%, fitment spacing ≤ min(Dc, 15db), cover meets durability and fire.
11. **HSC confinement** if `fc′ > 50 MPa` — simplified or deemed-to-comply.
12. **Document** — save to PDF with full clause references.

The tool automates 5–11; you supply 1–4 and review 12.

---

*Document compiled April 2026 during the build of the Concrete Column Designer. Updated as the tool evolves.*
