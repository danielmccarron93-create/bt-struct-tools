# Drafter ↔ StructDraw Mapping

**Purpose:** cross-reference between the two StructDraw knowledge bases.

- **`Structural_Drafter.md`** (in the `Structural Drafter` repo) — the house-style / AS 1100 / drafting knowledge base. Treated as the single source of truth for *how a Bligh Tanner steel detail should look*.
- **`dev/index.html`** (this repo) — the interactive StructDraw application. Treated as the single source of truth for *how the tool behaves*.

This file records which sections of `Structural_Drafter.md` have been implemented in `dev/index.html`, where they live in the code, and at what version they were last sync'd. It is the bridge the feedback loop runs across.

**Maintained by:** updated at the end of every session that touches either side. The rule is: if a change in StructDraw implements (or drifts from) a drafter rule, the relevant row gets bumped here before the session closes.

---

## How this file relates to the other docs

| File | Purpose | When to update |
|---|---|---|
| `Structural_Drafter.md` | House style, AS 1100 conventions, drafting patterns, pitfalls, lessons log | Whenever a *general drafting insight* is discovered — including during StructDraw work. §11 (pitfalls) and §12 (lessons log) are append-only. |
| `dev/STRUCTDRAW_PROJECT_BRIEF.md` | StructDraw architecture, data model, rendering pipeline, interaction model | Whenever the *application architecture* shifts (new object type, new subsystem, new transform, new interaction mode). |
| `dev/CHANGELOG.md` | Per-version release notes for StructDraw | Every approved version bump. One line per change. |
| `dev/DRAFTER_MAPPING.md` *(this file)* | Cross-reference table | Every session that lands drafter → StructDraw integration work. |

**Rule of thumb for where a lesson lands:**
- *"How do I draw a steel detail to AS 1100 house style?"* → `Structural_Drafter.md`
- *"How does the StructDraw engine work?"* → `STRUCTDRAW_PROJECT_BRIEF.md`
- *"What version of StructDraw implements which part of the drafter knowledge?"* → this file

---

## Mapping Table

| Drafter section | Topic | StructDraw implementation | Line(s) in `dev/index.html` (V13) | Status | Last sync |
|---|---|---|---|---|---|
| §3.2 | Member size notation (`360UB50` vs `360UB 50.7`) | `UB_DB` keys use formal notation | ~414–465 | ✅ matches | V13 |
| §3.3 | Bolt group labelling (`N/M20 8.8/S`) | `drawBoltCallout2D` + "Bolt Callout" library tool. Auto-counts selected bolts, picks dominant size. | ~5920 | ✅ matches (V18) | V18 |
| §3.4 | Bolt geometry defaults (M20) — edge/pitch/gauge | Not codified as constants; user draws bolt-by-bolt | — | ⏳ pending (V15 connection library) | — |
| §3.5 | Plate defaults (cap plate, base plate, cleat, splice gap) | Not codified as constants | — | ⏳ pending (V15) | — |
| **§3.6** | **Lineweights — 1.2 / 0.7 / 0.65 / 0.40 / 0.30 + MW overlay** | `const LW = { CUT:1.20, VIS_HEAVY:0.70, VIS:0.65, MW:0.50, DIM:0.40, HID:0.30, CL:0.30, HATCH:0.18 }` | ~632 | ✅ matches (V15.2) | V15.2 |
| §3.7 | Line styles (dash patterns for hidden / centreline) | `const DASH` table with SOLID, CL, CL_BOLT, SECTION, HIDDEN, THREAD, SNAP, UI_CHAIN, UI_ALT, UI_ROT. All call sites reference the table. | ~622 (table), ~45 sites | ✅ matches (V15.4) | V15.4 |
| §3.8 | Dimension conventions (arrowheads, text, witness, stagger) | `drawDim2D` now supports horizontal, aligned, vertical, angular, chain, baseline. Baseline type tiers stagger 10mm per stop. | ~5581 | ✅ matches (V18) | V18 |
| §3.9 | Section cut symbols (chain-dash, A-A labels) | `SECTION CUT LINES` section | 2071–2191 | ✅ working | V13 |
| **§3.10** | **Welds (AS 1101.3) — fillet triangle, tick, hatch on interface face** | Full vocabulary: `drawWeld2D` + `drawWeldGlyph` (fillet, square, single-V, double-V, partial-pen, bevel). Both-sides support via `otherType`/`otherSize`. Modifiers: `allAround`, `siteWeld`, `tail`, `length`. Auto-weld hatching on interface face only (`drawWeldHatch` respects `seg.hatchSide`). `#weldDialog` UI. | ~5442 (drawWeld2D), ~5570 (drawWeldGlyph), ~3723 (hatch), ~408 (dialog) | ✅ matches (V15.3) | V15.3 |
| §3.11 | Plan / elevation / section view conventions | `DetailBlock` with `viewKey` switch in every draw function | 574+ | ✅ matches | V13 |
| §3.12 | PFC open-face default (away from column) | No PFC support in StructDraw yet | — | ⏳ pending | — |
| §3.13 | Drawing scale / sheet / view order | A1 sheet at user-selected drawingScale (1:5 / 1:10 / 1:20) | 389–412 | ✅ matches (A1 vs drafter's A3) | V13 |
| §5 | Two-theme visual language (sketch / classic) | CSS variables `--bg`, `--ink`, etc. under `body.theme-dark` and `body.theme-classic`; plus canvas-layer wobble + grain toggles | 11–56, ~632 | ✅ matches (V17) | V17 |
| §5.5 | SVG filters — paper grain + sketch wobble | Canvas-equivalent: `_lineW` (deterministic Mulberry32-hash jitter wrapping rLine/rRect) + `_paperGrain` (cached 200px noise tile). Both off by default; toggles in library. Suppressed during PDF export for crisp output. | ~4208 (wobble), ~2844 (grain) | ✅ matches (V17) | V17 |
| §6.1 | SVG engine primitives (`sLine`, `sRect`, `sPath`, `dimH`) | Canvas equivalents: `rLine`, `rRect`, `rPath`, `rLineOcc` | 2465–2514 | ✅ equivalent | V13 |
| §6.3–6.7 | Member draw recipes (UB / SHS in elev / plan / section) | `drawUB`, `drawSHS`, `drawPlate`, `drawPolyPlate` | 2667+, 2821+, 2964+, 3025+ | ✅ matches (inherits new LW values via §3.6 cascade) | V15.2 |
| §6.8 | Clean-overlay technique for sketch theme | Wobble is applied as a per-line subdivision (not a double-draw clean overlay). Different technique but same *visual intent* — stays legible at A1 scale. | ~4208 | ✅ equivalent (V17 — different approach) | V17 |
| §6.10 | Stiffener plates (Type B) | Not implemented as a standard element | — | ⏳ pending | — |
| §6.11 | Cleat plate (elev, plan) | User draws as polygon plate | — | ⏳ standardisable | — |
| §6.14 | Title block | Fixed 30mm bottom strip rendered by `drawSheet` | ~1972+ | ✅ matches | V13 |
| **§7.2** | **M20 8.8 canonical bolt dimensions table** | `BOLT_DB` with `d, pitch, headAF, headH, nutAF, nutH, washOD, washT, minorD, threadL` (+ legacy `head`/`nut`/`p` aliases) | ~467–490 | ✅ matches | V14 |
| **§7.3** | **Chamfered hex profiles (`hexPathV/H`)** | `hexPointsAlongU` / `hexPointsAlongV` helpers → polygon path consumed by `_drawBoltSectionA_V14` / `_drawBoltPlanB_V14` | ~3207–3250, 3311+, 3376+ | ✅ matches (canvas polygon equivalent of drafter's SVG path) | V14 |
| §7.4 | clipPath-based hatching on hex heads | Not implemented — canvas `clip()` equivalent deferred until chamfered hex is visually confirmed | — | ⏳ deferred to V14.x (add as optional overlay once §7.3 is approved) | V14 |
| **§7.5** | **Sawtooth thread profile with pitch exaggeration** | `drawThreadAlongU` / `drawThreadAlongV` with `pitchReal = Math.max(0.9*drawingScale, realPitch*1.8)` exaggeration | ~3252–3306, used at 3341/3406 | ✅ matches (drafter §7.5 formula ported) | V14 |
| §7.6 | `drawBoltV` / `drawBoltH` assembly pattern with head-orientation flag | Two view-specific renderers (`_drawBoltSectionA_V14`, `_drawBoltPlanB_V14`) gated by `V14_NEW_BOLTS` feature flag at line ~505; head-outer face is always away from joint (outer=chamfered) | 3311+, 3376+ | ⚠️ no head-orientation flag yet — slip-joint support still pending | V14 |
| §7.7 | Slotted holes (22×40 for M20) | `drawSlot2D` renders a stadium shape + chain-dash centreline. Drawable via "Slotted Hole" library item. Defaults 22×40 matches M20 standard. | ~5789 | ✅ matches (V17) | V17 |
| **§7.8** | **Weld helpers — `weldTick`, `weldHatch`** | `drawWeldHatch` (chevron hatching along interface) + `drawWeldGlyph` (per-type glyph helper). Both driven from `computeWeldInterfaces` / `drawAutoWelds` pipeline. | ~3723, ~5570 | ✅ matches (V15.3) | V15.3 |
| §8.2 | Three.js materials (r128, MeshStandardMaterial + edges) | Three.js r128 engine in `3D ISOMETRIC VIEW ENGINE` section | ~3500+ | ✅ matches | V13 |
| §8.3 | Bounding-box beam positioning | StructDraw uses object `{x, y, z}` centroid — no extrude-direction pitfall | — | ✅ N/A (different model) | V13 |
| §8.7 | Bolts in 3D — Y- or Z-aligned cylinders only | Same limitation | — | ⚠️ known gap — no horizontal-X bolt | — |
| §9.1 | WSP standard parameters (ae / e / pitch / bolt count per UB depth) | `buildWSP(beam, params)` derives bolt count from beam depth + `WSP_PITCH`; plate from `WSP_AE` + `WSP_EDGE_BEAM` | ~7650 | ✅ matches (V16) | V16 |
| §9.2 | Column cap plate — `STD_BOLT_COL_GAP=40`, `STD_CAP_AE=35`, bolt-first derivation | `buildCapPlate(col, params)` with `CAP_BOLT_COL_GAP`, `CAP_AE`, `CAP_PLATE_THK` in `CONN_DEFAULTS` | ~7580, defaults ~629 | ✅ matches (V16) | V16 |
| §9.3 | UB moment splice — `SPLICE_DB` per-beam lookup | `buildSplice(beam, params)` — two end plates across `SPLICE_GAP`, bolt rows straddling the web | ~7700 | ✅ matches (V16 — not per-beam lookup, uses general derivation) | V16 |
| §9.4 | Column base plate — cast-in | `buildBaseplate(col, params)` with `BASE_OVERHANG`, `BASE_PLATE_THK`, `BASE_HD_BOLT_LEN` | ~7620 | ✅ matches (V16) | V16 |
| §9.5 | Mullion cap plate — cleat-and-tab slip joint (6011.6) | Not encoded | — | ⏳ V15 | — |
| §9.6 | Through-cleat (RHS to CHS) | No CHS support in StructDraw yet | — | ⏳ pending | — |
| §11.6 | **Pitfall:** weld hatch on interface face only | `drawWeldHatch` uses `seg.hatchSide` (signed perpendicular) computed in `computeWeldInterfaces`. Hatch is correctly side-aware. | ~3723 (hatch), ~3528 (compute) | ✅ mitigated (V15.3) | V15.3 |
| §11.7 | **Pitfall:** Three.js `rotateY(-PI/2)` + bounding-box translate | Different model — not applicable | — | ✅ N/A | V13 |
| §11.10 | **Pitfall:** cap plate sized from width instead of bolt position | Not yet a risk — no standard-detail builder | — | capture when V15 lands | — |
| §11.14 | **Pitfall:** mass rounding `360UB50` vs `360UB 50.7` | `UB_DB` uses formal | — | ✅ matches | V13 |
| §11.16 | **Pitfall:** thread pitch illegible at scale | Handled via `pitchReal = Math.max(0.9 * drawingScale, realPitch * 1.8)` in `drawThreadAlongU/V` — exaggerates pitch to ≥0.9 sheet-mm regardless of drawingScale | ~3262, 3286 | ✅ mitigated | V14 |

---

## Legend

- ✅ **matches** — StructDraw implements the drafter rule faithfully. Nothing to do.
- ⚠️ **drifts** — StructDraw does something related but differs from the drafter rule. Flagged for review.
- ❌ **stub** — StructDraw has a placeholder (or nothing at all) for this piece of drafter knowledge. Candidate for the next upgrade.
- ⏳ **pending** — explicitly deferred to a named future version.

---

## Feedback loop at session close

At the end of every session that touched either the drafter knowledge or the StructDraw code:

1. Update `dev/CHANGELOG.md` with a one-line entry for the version bump.
2. Update `dev/STRUCTDRAW_PROJECT_BRIEF.md` only if application architecture shifted.
3. Update *this file* — bump the Status and Last-sync columns for any drafter section that was implemented, corrected, or deferred.
4. Update `Structural_Drafter.md` §11 (pitfalls) or §12 (lessons log) with any general drafting insight discovered during StructDraw work. Remember: a drafter lesson learned *while building StructDraw* still counts as a drafter lesson and belongs in the master MD file.
5. Update `Structural_Drafter.md` §10 (Catalogue) only if StructDraw's Standard Connection library gained a new detail type in this session — note alongside the reference visualiser.

---

*Last updated: 2026-04-18 — V22 landed: catalogue expansion. 5 new structural section types (PFC / RHS / CHS / EA / UA) with 94 standard-range profiles across AS/NZS 3679.1 and AS 1163. Unified `sectionProfile()` and `isMemberType()` helpers back the geometry pipeline so existing UB/SHS codepaths stay untouched. `drawSectionMember()` handles all 5 new types per view with cut/projected states, AS 1100 hatching on section views. V22 also activated every palette placeholder: Aligned + Angular dims, Arc (3-point), Polygon (regular N-sided), Offset (parallel copy), Fillet + Chamfer (polygon-plate vertex operations), Grid Line (bubble + chain-dash), Note (leader + text block), Hatch (steel/cross/concrete patterns with polygon clip), MText (wrapped text block), Rev Schedule (auto-populated table aggregating all revisionTriangle entities across all views). Deferred to V22.1b: 3D Three.js builders for new section types, true profile-specific DXF entities (CHS arcs, PFC C-profile), CHS through-cleat + welded plate-girder parametric connections.

---

Previously — 2026-04-18 — V21 landed: complete UI rebuild to Revit/Bluebeam-grade polish. Three themes (Classic B&W default, Dark neutral, BT-Red brand) via CSS variables with legacy aliases for zero engine churn. ~50 hand-drawn SVG section/tool icons as `<symbol>` bank — technical end-view profiles for UB/UC/PFC/SHS/RHS/CHS/EA/UA so a UB icon literally looks like a UB. Three-mode intent model (Model/Draw/Annotate) with top-right segmented pill switcher. Chrome-style sheet tabs in top bar (replaces sidebar browser). 78×78 tile palette with grouped sub-sections, "soon" placeholders for V22/V23 features. Grouped-by-series searchable size picker (610 series → 610UB 125/113/101). Right-docked Inspector replacing floating props — auto-switches between sheet info / member properties / multi-selection / tool options. Keyboard chord layer (M/D/A then letter). Favourites strip persisted to localStorage with pin/unpin. Status bar moved Snap/Ortho/Grid toggles here per AutoCAD convention. Engine, data model, renderers, exports all untouched. No drafter §-row semantic changes. File ~12,500 lines.

---

Previously — 2026-04-18 — V19.5 and V20 landed in one session. V19.5: full multi-sheet project model (`project = { sheets: [], activeSheetIdx }`, `projectSwitchSheet` snapshot/restore). Sheet browser sidebar. Multi-page PDF export (`exportProjectToPDF` — hot-swaps each sheet through V15 vector path, `pdf.addPage` per sheet). Project save/load as `.sdproj` JSON. V20: Ctrl+K command palette with substring + subsequence fuzzy scoring (~40 indexed commands), floating layer-visibility panel with 10 semantic groups gated in render pipeline, keyboard help overlay (`?` key), mirror tool (`performMirror(a1, a2, viewKey)` creates new mirrored copies, preserves originals). Build-step introduction (esbuild bundler split into `src/`) consciously deferred — codebase is at ~10,500 lines but still coherent as one file; worth the conversation before committing.

---

Previously — 2026-04-18 — V19.1–4 landed. DXF export (AutoCAD R2013 ASCII) via `exportSheetToDXF` — full LAYERS + LTYPE tables, per-entity-type emission including every V16–V18 entity. Button wired. Four new entity types: `revisionTriangle` (numbered triangle tied to rev schedule), `revisionCloud` (multi-click arc-perimeter around revised work), `detailRef` (3/S-400 callout bubble), `detailCard` (heavy frame + number + scale, precursor to multi-detail layout). V19.5 (full multi-sheet project model + sheet browser + build-step introduction) deferred to next session for testing continuity. File now ~9,700 lines.

---

Previously — 2026-04-18 — V17 and V18 landed in one session. V17: deterministic sketch wobble (`_lineW` wrapping rLine/rRect via Mulberry32 hash — same line always wobbles the same way), cached paper-grain overlay for classic theme, slotted-hole entity (`drawSlot2D`, AS 1100 §7.7 default 22×40), refined centreline renderer. Wobble and grain both off by default and suppressed during PDF export. §5, §5.5, §6.8, §7.7 flipped to ✅. V18: `drawDim2D` extended with `chain` and `baseline` dim types (baseline tiers stagger 10mm per stop); new entity types `memberTag` (parametric — auto-resolves section from `memberId`), `boltCallout` (auto-counts + dominant-size logic), `sectionMark` (with `nextSectionMarkLabel` auto-lettering), `materialTag`. §3.3, §3.8 flipped to ✅. File now at ~9,100 lines. Awaiting user visual approval in browser before copy dev → root.*
