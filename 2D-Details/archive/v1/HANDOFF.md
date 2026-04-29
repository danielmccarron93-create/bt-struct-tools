# StructDraw — Handoff to next chat

**Date:** 18 April 2026
**Current version:** V22 (catalogue expansion)
**File:** `/Users/danielmccarron/Documents/GitHub/bt-struct-tools/2D-Details/dev/index.html`
**Size:** ~14,400 lines, single HTML file, no build step

> **If you're starting a new Claude Code chat: read this file first, then `CHANGELOG.md` for recent detail, then `STRUCTDRAW_PROJECT_BRIEF.md` §0 (current status) if you need the subsystem map.**

---

## 1. What StructDraw is (in one paragraph)

A browser-based 2D structural detail drawing tool for Australian structural engineers. Single HTML file (`dev/index.html`), no build step, Three.js r128 + jsPDF 2.5.1 loaded from CDN. Renders an A1 sheet (841×594mm) with four synchronised views (elevation / section A / plan B / 3D isometric) fed from one unified 3D data model. Draws to AS 1100 conventions with proper lineweight hierarchy, depth-aware occlusion, vector PDF export, DXF R2013 export, multi-sheet project support, and a full palette of structural primitives (UB/UC/PFC/SHS/RHS/CHS/EA/UA/bolt/plate/weld), connection builders (cap plate / baseplate / WSP / splice), annotations (dims, tags, section marks, revision clouds), and now drawing tools (arc, polygon, offset, fillet, chamfer, hatch, mtext, grid line, note, rev schedule).

---

## 2. The workflow discipline (don't break this)

1. **Always edit `dev/index.html` first** — never the root `index.html`
2. **Ask Dan to test in the browser** at `file:///Users/danielmccarron/Documents/GitHub/bt-struct-tools/2D-Details/dev/index.html`
3. **Only after approval**, copy `dev/index.html` → root `index.html`
4. **Dan handles the git commit** himself — don't commit or push unless he explicitly asks

---

## 3. How to pick up where we left off

### The minimum context you need

1. Read this file (`HANDOFF.md`) — you are here
2. Read the **Current Status** section (§0) of `STRUCTDRAW_PROJECT_BRIEF.md` — tells you what subsystems exist
3. Read the top entry of `CHANGELOG.md` (V22) — tells you what just shipped
4. Read `DRAFTER_MAPPING.md` if you're working on a drafter-rule implementation

### You do NOT need to read

- The historical sections of `STRUCTDRAW_PROJECT_BRIEF.md` (§1–11 is V7-era architecture, still accurate for data model / coords / rendering but you don't need it unless debugging a subsystem)
- Old CHANGELOG entries below V22 (history is in git)

### Before making ANY code change

1. Confirm `dev/index.html` is writable + reachable (user may need to mount GitHub folder)
2. Use `wc -l` to verify line count matches what the brief says (~14,400 at V22)
3. **Grep before you write** — almost every pattern you're about to invent is already in the file. New tile? Look at `populateTilePalette`. New entity type? Look at `drawEnt2D` dispatch + `mkEnt2D` factory. New section? Look at `sectionProfile` + `drawSectionMember`.

---

## 4. Where we are (V22 complete)

### Shipped through V22
- V15 — Vector PDF, weld polish, AS 1100 lineweights
- V16 — Connection library (cap plate, baseplate, WSP, splice)
- V17 — Sketch themes, slotted holes, drawable centrelines
- V18 — Chain/baseline dims, member tags, section marks, material tags
- V19 — Multi-sheet, multi-page PDF, DXF export, revisions, detail cards
- V20 — Command palette, layer UI, keyboard chords, mirror tool
- V21 — **UI rebuild** (3 themes, custom icons, tile palette, Inspector, sheet tabs)
- V22 — **Catalogue expansion** (5 new section types + 94 profiles, 13 tool placeholders activated)

### V22 palette tiles — what's live vs what's a placeholder

**Model:** UB, UC, SHS, RHS, PFC, CHS, EA, UA (all live), Bolt, Bolt Group, Slot, Plate, 4 connection wizards (all live).

**Draw:** Select, Line, Rect, Circle, Polyline, Centreline, **Arc**, **Polygon**, **Hatch**, Text, **MText**, Break Line, **Offset**, **Fillet**, **Chamfer** — all live. Faded placeholders remaining: Spline (V23), Fill (V23).

**Annotate:** Dim H / V / **Aligned** / **Angular** / Chain / Baseline, Member Tag, Material Tag, Bolt Callout, **Note**, Section Mark, Weld, Detail Ref, **Grid Line**, Rev Triangle, Rev Cloud, **Rev Schedule**, Detail Card — all live. Faded placeholder: Ordinate dim (V23).

### What's explicitly deferred

**V22.1b** (small follow-on, high value for completeness):
- **3D Three.js builders** for PFC / RHS / CHS / EA / UA — the five new V22 section types DON'T appear in the isometric view. This is the most visible gap. Needs builders matching the existing UB/SHS Three.js code at ~line 13700+.
- **True-profile DXF** for new types — they currently export as bounding-box rectangles on `S-BEAM`. CHS should emit as `CIRCLE`/`ARC`, PFC as a proper C-profile `LWPOLYLINE`.
- **CHS through-cleat connection** (drafter §9.6) — parametric builder following the pattern of `buildCapPlate` etc.
- **Welded plate-girder** parametric connection — stretch.

**V23** (workflow polish, user-facing UX):
- Connection wizard from modal → **inline Inspector** with live on-sheet preview as the user adjusts parameters. This is the biggest UX upgrade remaining. The wizard currently blocks the canvas view; making it inline + live would put StructDraw's connection UX ahead of Revit's.
- **First-run welcome tour** — 3-step guided intro for a new user.
- **Detail template library** — save a detail as a named template, reuse across projects.
- **Preferences dialog** — firm-wide defaults (firm name, tagline, default scale, default bolt grade, sketch theme defaults, templates folder).
- **Spline tool** (3D-like curve through control points)
- **Ordinate dim**
- **Fill primitive** (solid region fill, AutoCAD SOLID equivalent)

---

## 5. What to do next — ordered options

Ask Dan which of these he wants to tackle first; don't assume. All are real work, all can be done in one focused session each.

### Option A — V22.1b (finish V22 properly)
The most honest continuation. Close the loop on V22 before moving on.
1. Add Three.js builders for PFC/RHS/CHS/EA/UA so the isometric view shows them
2. Emit true-profile DXF for the 5 new types
3. Optionally add the CHS through-cleat connection

**Size:** ~500–800 lines. One focused session.

### Option B — V23.1 Inline Connection Wizard
Highest user-facing UX impact. The flagship V23 feature.
1. Strip the connection dialog's modal overlay
2. Render its form fields into the Inspector panel's tool-active state
3. Wire each field-change to re-run the builder function in "preview mode" (returns objs + ents but doesn't commit to the model)
4. Draw preview objects at reduced opacity on the canvas
5. Commit on explicit Enter / button click

**Size:** ~600–1000 lines. One focused session.

### Option C — V23 onboarding + templates
Makes StructDraw genuinely shippable to someone other than Dan.
1. First-run welcome overlay: "Welcome — click **Model** → **UB** → click two points to place your first beam."
2. Detail template library: right-click anywhere on the sheet → "Save as template" → named JSON blob stored in localStorage
3. Template browser in the hamburger menu

**Size:** ~400–600 lines.

### Option D — Drafter-spec compliance audit
Open `DRAFTER_MAPPING.md` — walk every row, flip ⏳/⚠️/❌ rows that V21/V22 silently fixed to ✅, and identify any remaining drafter §-rule gaps worth chasing.

**Size:** Read-heavy, write-light. One session.

### Option E — Bug bash / polish
Run through a real detail end-to-end and list every papercut. Likely candidates:
- 3D view doesn't show V22 section types (known, see V22.1b above)
- Some dim tool flows may require mode-specific fixes
- Sketch-wobble may need tuning after V21 theme rebuild

**Size:** Variable.

---

## 6. Architecture invariants you must NOT break

These underpin the whole app. If you think you need to change one, stop and check with Dan.

1. **Single HTML file, no build step.** Three.js + jsPDF from CDN only. Everything inline.
2. **Four synchronised views driven by ONE 3D data model.** `objects3D = [...]` is the single source of truth for structural members. Each view projects (u, v) differently but reads from the same array.
3. **Y is UP in real-world coords, DOWN on canvas.** The flip lives in `real2px()`. Never think about Y flipping anywhere else.
4. **`drawingScale` divides real coords into sheet coords.** At 1:10, a 600mm beam renders 60mm on the sheet.
5. **`ppm()` returns 1 during `pdfExportMode`.** This is how `LW.VIS * ppm()` magically becomes sheet-mm during vector PDF export. Don't break this.
6. **AS 1100 units everywhere.** mm only, no imperial.
7. **Three.js r128 APIs only.** Don't use `CapsuleGeometry` or anything introduced after r128.
8. **Variable naming:** `u,v` = view-local 2D; `x,y,z` = world 3D; `px,py` = screen pixels; `sx,sy` = sheet-mm. Functions prefixed `r` (`rLine`) draw in real-world coords. Functions prefixed `v3d` are Three.js 3D engine.
9. **Every legacy DOM ID is preserved.** V21 moved UI around but kept every ID — hidden shim elements exist for the IDs that no longer have visible chrome. Don't delete those.
10. **Active sheet = globals snapshot.** `project.sheets[activeSheetIdx]` is mirrored into `objects3D`, `entities2D`, `sheetInfo`, `secCutX`, `planCutY`, `objIdN`, `ent2dIdN` via `_projectSnapshotActive` / `_projectLoadSheet`. Render pipeline knows nothing about multi-sheet — it always reads the globals.

---

## 7. Fast-reference map

```
V22 file structure (~14,400 lines)

<head>
  CSS tokens (3 themes: Classic B&W / Dark / BT-Red)          lines 12–330
  V21 component CSS (tile / palette / inspector / picker)     330–690
  V17 wobble + paper grain CSS                                —
  Canvas container                                            670–690

<body>
  SVG icon bank (~50 symbols)                                 790–1139
  Top bar (brand / sheet tabs / mode switcher / actions)      1140–1200
  Hamburger menu dropdown                                     1200–1260
  Main grid (palette / canvas / inspector)                    1260–1290
  Hidden shim elements (preserves legacy DOM IDs)             1290–1350
  Status bar (coord readout / toggles / chips)                1350–1430
  Modal dialogs                                               1430–1600
  Command palette / kbd help / layer panel                    1600–1750

  <script>
    Section DBs (UB/UC/SHS/PFC/RHS/CHS/EA/UA/BOLT)            1850–2200
    V22 sectionProfile() + isMemberType()                     ~4000
    3D object model, undo/redo, project model                 2230–2550
    Hit testing / bounds / occlusion                          2560–3900
    LW / DASH / CONN_DEFAULTS constants                       3900–4100
    PDF export (vector shim + raster fallback)                4400–4800
    DXF writer (LAYERS, LTYPE, per-entity emitters)           4800–5100
    Render pipeline, drawBlockContent, drawSheet              5200–6500
    Object renderers (UB, SHS, drawSectionMember, Plate, Bolt) 6500–7100
    2D entity renderers (dim, weld, tag, hatch, mtext, etc.)  7300–8700
    Connection builders + placeConnection                     8800–9100
    Event handlers (mousedown dispatch per tool)              9200–9950
    initKeyboard (+ V21 chord layer)                          9950–10200
    Multi-sheet (projectInit / switch / save / load / PDF-all)10200–11000
    V21 Inspector / Size picker / Favourites / Cmd palette    11000–12000
    Tile palette (mode-filtered, placeholders)                12000–13300
    Library + toolbar + theme cycle init                      13300–14000
    Three.js 3D engine                                        14000–14400
```

---

## 8. Promotion checklist (dev → production)

Before Dan commits and pushes:

1. Dan has tested the dev file in the browser
2. Every new feature has been click-exercised
3. Exports (PDF / DXF / .sdproj save-load) work on at least one sheet
4. `dev/CHANGELOG.md` has a top-level entry for the new version
5. `dev/DRAFTER_MAPPING.md` rows updated for any drafter-rule changes
6. `<title>` tag bumped to the new version
7. Copy `dev/index.html` → root `index.html`
8. Dan commits + pushes himself

---

## 9. If a new chat asks "what should I do next?"

The answer is *"ask Dan"*. Not a meta-answer — there are 5 legitimate next directions listed in §5, and they're different kinds of work (engine completion / UX polish / onboarding / audit / polish). Dan's priorities shift with what he's trying to use StructDraw for that week. Offer the 5 options, let him pick.

If he says "you pick", **recommend Option A (V22.1b: 3D builders + DXF profiles)**. Rationale:
- It finishes V22 honestly — leaving new section types invisible in the iso view is a visible gap
- It's self-contained — doesn't touch UX, doesn't require design decisions
- It completes the "a structural engineer who places a CHS actually sees a CHS everywhere" promise

---

*End of handoff.*
