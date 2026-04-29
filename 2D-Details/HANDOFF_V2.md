# StructDraw V2 — Handoff Document

**Created:** 2026-04-29
**Author:** Cowork session (handing off to Claude Code for active development)
**Audience:** Whichever Claude session picks this up next

---

## TL;DR

StructDraw is a single-file browser app for producing AS 1100 / AS 4100 / AS 3600 compliant 2D structural details. V1 grew organically over many manual versions (`StructDraw_V1` … `V13`) and reached a point where the engineering features were strong but the application shell (sheets, layers, properties, selection model, scale handling) was clunky compared to Bluebeam Revu — Dan's daily reference.

In the previous Cowork session we **rebuilt V2 from scratch with a Bluebeam-style shell first**, deliberately leaving the engineering features for a follow-up phase. V2 is now a **shell-only working prototype**: the chrome is right, but most of the engineering value V1 had is gone or stubbed. Several V2 features are also broken or only half-implemented in subtle ways the surface doesn't reveal.

This document is the bridge. Read it end-to-end before changing anything.

---

## Repository layout

All paths are absolute on Dan's Mac. The mounted-folder paths in remote sessions vary — request access if needed.

```
/Users/danielmccarron/Documents/GitHub/bt-struct-tools/2D-Details/
├── HANDOFF_V2.md                ← THIS FILE
├── index.html                   ← V1 production (untouched, still serves users)
├── dev/
│   └── index.html               ← V2 dev (Bluebeam-style shell, the new active codebase)
└── archive/v1/
    ├── index.html               ← The last V1 dev file (V24.A3 — full feature set)
    ├── CHANGELOG.md             ← V1 version history
    ├── STRUCTDRAW_PROJECT_BRIEF.md
    ├── DRAFTER_MAPPING.md
    └── HANDOFF.md               ← Old handoff doc
```

**Important:** The current `dev/index.html` is V2 (~128 KB, ~2,985 lines).
The current root `index.html` is V1 (~635 KB) and remains in production until V2 reaches feature parity for steel detailing.

---

## V2 — what's done (current state)

### Shell (Bluebeam-style)
- Top menubar: File / Edit / View / Tools / Help
- Ribbon with 8 tabs: Home / Draw / Annotate / Modify / Members (stub) / View / Layers / Tools
- **Single left rail** with four tabs: **Pages / Layers / Properties / Tools** (Tool Chest)
- Left rail is **drag-resizable** by the right-edge handle (180–600 px)
- Drag below ~110 px **collapses fully**; a thin edge tab on the canvas edge re-opens it. Double-click handle toggles
- Status bar: snap, ortho, polar, grid toggles, current tool, page scale, cursor coords (sheet-mm), zoom %

### Pages (Multi-sheet)
- Multi-page A1 landscape system with thumbnails
- Click to jump, drag-to-reorder (no drop indicator yet), add / duplicate / delete (Ctrl+N / Ctrl+D)
- Per-page scale: presets 1:5–1:200 plus custom
- Optional scaled zones (drawn but **not yet applied to elements**, see "Half-implemented")

### Layers
- 12 default layers seeded per page (Defpoints, Construction, Members, Bolts, Plates, Welds, Hatch, Hidden, Centre, Dimensions, Text, Title)
- Per-layer: visibility, lock, print, colour swatch
- Add / delete / rename layers
- Flatten all and flatten selection
- Right-click element → "Move to layer ▶"

### Drawing tools (basic primitives)
- Line, Rectangle, Circle, Polyline, Polygon, Arc (3-point), Revision Cloud
- Text, Leader, Callout
- Dimension (horizontal only — see regressions)

### Selection
- Click / box / lasso (window vs crossing by drag direction), Shift to add/toggle
- Drag selection to translate
- Group / ungroup (tag only — see bugs)
- Mirror, Array, Copy, Cut, Paste, Paste-in-Place
- **Snip (G)** captures all overlapping elements with sheet coords; Ctrl+V or Ctrl+Shift+V pastes on any page at exact sheet coords (cross-sheet detail re-use workflow)

### Snap engine
- Endpoint, midpoint, intersection, centre, nearest with visual snap markers
- F3 toggle. Perpendicular flag is on by default but **not implemented** (see bugs)

### Constraints
- Ortho (F8) — horizontal/vertical lock
- Polar (F10) — 15° increments

### Measurement tools (Bluebeam-equivalents)
- Linear, Aligned, Area, Perimeter, Angle (3-point), Radius
- Calibrate from a known length (sets page scale)
- All scale-aware, format real-world m / mm / m²

### Properties palette
- Context-sensitive on selection
- General: type, layer, lock (lock flag is **broken**, see bugs)
- Appearance: stroke colour / weight / style (solid, dashed, hidden, centre, phantom), fill, opacity
- Geometry: numeric X/Y/W/H per element type
- Text section for text/leader/callout
- Match Properties tool (cursor doesn't revert — see bugs)

### File handling
- Save → JSON `.sd2.json`
- Open → reads same format
- PDF / DXF export are buttons that toast "coming soon"

### Keyboard shortcuts
- V/L/R/C/P/A/T/D/G + Esc, Del, Ctrl+A/C/X/V (+Shift for paste-in-place), Ctrl+Z/Y, Ctrl+G/Shift+G
- F3/F7/F8/F10 for snap/grid/ortho/polar
- F1 or Tools→Shortcuts opens reference modal
- Z then F = fit page. Wheel zoom is cursor-anchored. Alt-drag or middle-drag pans

### Undo/redo
- 100-deep snapshot stack on every state-changing command

---

## V1 → V2 REGRESSION LIST (the most important section)

V1 had a lot of structural-engineering value that V2 has not yet re-implemented. **The new session should test these in V1 before deciding what to bring back.** Open `archive/v1/index.html` in a browser to see the full V1 feature set in action.

### Major missing features (vs V1)

| V1 feature | Status in V2 | V1 reference |
|---|---|---|
| **2D-detail mode ↔ 3D mode toggle** (the "really cool" toggle Dan mentions) | **Missing** | V1 had `entities2D = { elevation, sectionA, planB }` plus a 3D iso view. Toggle changed the workspace from a single elevation pane (for clean 2D detailing) to the full 4-pane projection layout. |
| **Three orthographic views simultaneously** (Elevation X,Y · Section A Z,Y · Plan B X,Z) | **Missing** | `Block` class with `viewKey` — `projElev`, `projSecA`, `projPlanB`. World-coordinate objects projected to each view. |
| **3D isometric view** (Three.js r128) | **Missing** | `v3dInit`, `v3dRebuildScene`, `v3dBuildUB`, `v3dBuildSHS`, `v3dBuildPlate`, `v3dBuildBolt`, orbit controls (`v3dStartOrbit`, etc.). |
| **Parametric steel members**: UB, UC, PFC, SHS, RHS, CHS, EA, UA | **Missing** | Full SVG icon library (`icon-ub`, `icon-uc`, `icon-pfc`, `icon-shs`, `icon-rhs`, `icon-chs`, `icon-ea`, `icon-ua`). `isMemberType`, `memberFrame`, `memberViewMode`, `memberViewAngle`, `memberEndPoint`, `memberExtentOnAxis`. |
| **Bolts (AS 1252)** with bolt groups, slots | **Missing** | `icon-bolt`, `icon-bolt-group`, `icon-slot`. |
| **Plates** with cap-plate, baseplate, web side plate, splice variants | **Missing** | `icon-plate`, `icon-cap-plate`, `icon-baseplate`, `icon-wsp`, `icon-splice`. |
| **AS 1100 lineweight hierarchy** as a meaningful render system (0.70 / 0.35 / 0.18 / 0.13 mm) | **Cosmetic only** | V2 stores `strokeWidth` in mm but doesn't drive any of the visibility/print logic V1 had. |
| **Depth-aware occlusion** (objects in front hide objects behind in elevation/section views) | **Missing** | `getOcclusionRects`, `getDepthValue`, `get2DFootprint`, `sectionProfile`. |
| **Section cut lines** (draggable in elevation, define section view position) | **Missing** | `draggingCutLine = 'secA' / 'planB'`, `hitTestCutLine`, `drawSectionCutLines`. |
| **Edge snapping** between objects (snap-to-edge, with visual indicator lines) | **Missing** | `getSnapEdges`, `getDraggedEdges`, `applyEdgeSnap`, `drawEdgeSnapLines`. |
| **Cross-hatching** (concrete · steel · earth · timber · hidden) | **Missing** | V1 had a hatch tool (`icon-hatch`, `icon-fill`). V2 has the field in the data model but no renderer. |
| **Hidden lines, centrelines, break lines** as first-class entities (not just stroke styles) | **Stroke-style only** | V1 had `icon-break-line`. V2 only changes the line dash pattern. |
| **Fillet & chamfer** on polygon-plate vertices | **Missing** | `icon-fillet`, `icon-chamfer`. V2 ribbon has Fillet button but no implementation. |
| **Spline tool** | **Missing** | `icon-spline`. |
| **Full dimension family**: horizontal, vertical, aligned, angular, chain, baseline, ordinate | **Horizontal only** | V1 icons: `icon-dim-h`, `icon-dim-v`, `icon-dim-aligned`, `icon-dim-angular`, `icon-dim-chain`, `icon-dim-baseline`, `icon-dim-ordinate`. |
| **MText** (multi-line text with wrap) | **Missing** | `icon-mtext`. V2 only does single-line text. |
| **Section marks** | **Missing** | `icon-section-mark`. |
| **Weld symbols (AS 1101)** with full leader-and-symbol composition | **Missing** | V1 had a Weld dialog (`wdOkBtn` "Place Weld") and `icon-weld`. |
| **Title block** with editable fields (project, drawing number, scale, sheet, drawn by, checked by, date, rev) | **Missing** | V1: title block dialog (`titleBlockDialog`), persistent metadata, auto-rendered at sheet bottom. |
| **Revision schedule** (tabular revision history block) | **Missing** | V1: `icon-rev-schedule`, dedicated symbol. |
| **Member and bolt schedules** auto-generated from drawing | **Missing** | Was in the V1 brief; built or stubbed depending on V13's state. |
| **PDF export** (both raster and vector) | **Stub button** | V1: `exportSheetToPDF`, `exportSheetToPDFRaster`, `exportSheetToPDFVector`, `createPdfCanvasShim`. |
| **DXF export** | **Stub button** | V1: `exportSheetToDXF`. |
| **Block resize handles** (resize a view block on the sheet) | **Missing** | V1: `hitTestResizeHandle`, `hitTestBlockBorder`, `resizeHandleCursor`. |
| **Grip-handle drag editing** (resize a polygon-plate by dragging its grips) | **Visual only** | V1: `getGrips`, `hitTestGrip`, `applyGripDrag`. V2 draws grips but they're not pickable. |
| **Member rotation across views** | **Missing** | V1: `memberFrame`, `frameFromPreset`, `applyRotation`. V2 stores `meta.rotation` but no UI writes to it. |
| **Favourites panel** (saved tool presets, library items) | **Stubbed** | V1: `renderFavourites`. V2 has Tool Chest UI but most items are placeholders. |
| **Stamps** | **Stub** | V2 redirects stamps to Callouts. |

### What V1 had that V2 carried forward correctly

- Multi-sheet system (V1: `projectInit`, `projectAddSheet`, `projectSwitchSheet` — V2: re-implemented cleanly)
- Layer panel (V1: `renderLayerPanel`)
- Sheet browser / page thumbnails (V1: `renderSheetBrowser`)
- Undo/redo
- Edge-snap visual indicator (V1) → V2 has snap markers

---

## V2 bugs (broken / claimed-as-done but doesn't actually work)

These are not "missing features" — they're **misleading**. The UI suggests they work; they don't.

1. **Grip-handle resize and rotate are decorative.** Selection draws 8 corner/edge grips, but `toolSelectDown` always begins a translate. There's no hit-test for grips. `gripsForBBox` exists but isn't picked up by the drag handler.
2. **Element-level Lock does nothing.** Properties has a Lock checkbox that writes `el.locked`. Nothing reads it. (Layer-lock works — element-lock is the dead one.)
3. **Group / Ungroup is just a tag.** `cmd_group` writes `meta.group = id` to selected elements. Nothing reads it back. Selecting one member does not pick up the rest.
4. **Match Properties leaves the cursor stuck.** `beginMatchProps` sets `cv.style.cursor = 'copy'`. After applying, nothing reverts it.
5. **Trim / Extend / Fillet / Offset are dead buttons.** Ribbon clicks call `setTool('trim')` etc., but `handleToolDown` has no case for those tool ids. Toast looks like flow started; nothing happens on canvas.
6. **`m` keyboard shortcut maps to `select`.** Pointless. Should be `move`-within-select or removed.
7. **Layers can't be reordered.** Only add / delete / rename / recolour. They render in creation order forever.
8. **`state.snap.modes.perpendicular` is on by default but no perpendicular geometry is implemented.** Toggle does nothing.
9. **`cmd_addZone` leaks state.** If user hits Escape after invoking Add Scaled Zone, `state.pendingZone` stays `true`. The next rect they ever draw — possibly minutes later — gets hijacked into a zone prompt.

## V2 half-implemented

The data model exists, the UI hints at it, but the behaviour isn't wired through.

10. **Scaled zones are visual-only.** The zone outline draws and the label shows `1:N`, but `zone.scale` is never used. Elements drawn inside a zone render at page scale. Measurements inside a zone use page scale.
11. **Hatching:** `el.props.hatch` field exists, renderer ignores it, Tool Chest hatch entries toast "apply from Properties" but Properties has no hatch picker.
12. **Stamps redirect to Callouts.**
13. **Drag-reorder pages works** but has no visual drop-indicator line.
14. **Snap-to-grid does not snap.** Grid renders when toggled, no snapping logic exists.

## V2 intentional placeholders (flagged in the original spec)

- PDF import (the Calibrate-from-PDF flow assumes this exists)
- PDF / DXF export
- Members ribbon (UB / SHS / Plate / Bolt) — explicitly Step 2
- Print-toggle on layers writes the flag but no print export uses it
- Double-click-to-edit text in place

## V2 polish gaps worth tightening

- Status bar coordinates show **sheet-mm**. An engineer drawing at 1:10 usually wants real-world mm. Toggle or show both.
- Stroke-width Properties dropdown is a fixed list. No custom value entry.
- Properties panel doesn't update X/Y/W/H during a drag — only after mouseup.
- No visual indicator when ortho/polar lock is active.
- Fill `None` checkbox interaction is fiddly with the colour input.
- Rotation has no UI at all (`meta.rotation` exists but unused).
- `state.prevTool` is dead code.
- Title bar shows `Untitled.sd2` forever — Save doesn't update it.
- Submenu hover-leave logic in context menu can be jittery.

## V2 architectural flags (not bugs, future considerations)

- Every mousemove triggers a full canvas redraw (RAF-throttle when needed).
- History is full deep-cloned snapshots per change (delta-based later).
- No autosave, no unsaved-indicator. Browser close = lost work.
- `.sd2.json` save format has no schema version field — schema changes will silently break old files.

---

## What an Australian structural engineer actually needs from this tool

This is the lens for refinement. Dan is a Senior Structural Engineer at Bligh Tanner producing details to AS 1100 / AS 4100 / AS 3600 / AS 1101 / AS 1252. The user is **not** an architect, not a renderer, not a generic CAD operator. The tool should be **ruthlessly focused** on producing connection details and concrete details that look at least as clean as the STP Typical Structural Details PDF (page 85, details 6011.1–6011.6).

### Core daily activities (must be excellent)

1. **Detail steel connections** — cap plates, baseplates, splices, web side plates, moment connections, brace gussets, tilt-up connections. Need parametric UB/UC/PFC/SHS/RHS/CHS/EA/UA, plates of arbitrary polygon, AS 1252 bolts in plan and elevation, AS 1101 weld symbols.
2. **Detail concrete** — slab edges, beam stirrups, reinforcement callouts (N12, N16, N20, etc.), cover, lap lengths, starter bars, lift core penetrations.
3. **Annotate properly** — full dimension family (horizontal, vertical, aligned, angular, chain, baseline, ordinate, radial), leaders with arrow styles, MText with wrap, callout balloons, section marks, view labels.
4. **Hatch correctly** — concrete (45°-stipple or solid), steel (solid black for sections), timber (grain), earth (zigzag), hidden (no hatch but with hidden lines).
5. **Linework matters.** Visible, hidden, centre, phantom — with strict AS 1100 lineweight hierarchy: thick (0.70 mm cuts and prominent edges), medium (0.35 mm visible edges), thin (0.18 mm hidden / centre / dimensions).
6. **Multi-page detail sets.** A typical detail sheet has 4–8 details on one A1 page, then a project might have 5–20 sheets.
7. **Title block compliant with the office standard.** Bligh Tanner format with project name / drawing number / scale / drawn / checked / date / revision history.
8. **PDF export that prints true to scale.** When you print the PDF on A1 paper, dimensions match.
9. **DXF export** for handing to the drafter who'll incorporate into the project's main DWG.
10. **Cross-sheet detail re-use** (V2's snip-paste workflow).

### High-value differentiators (vs Bluebeam, vs AutoCAD)

- **3D isometric view** alongside the 2D — draw in 2D, sanity-check the geometry in 3D. V1 had this; V2 needs it back.
- **Parametric members that auto-render** in elevation, section, and plan with correct depth/flange/web — eliminates manual redrawing.
- **AS-aware defaults** — drop in a UB beam, get the right depth/flange/web/root radius. Drop in an M20 bolt, get the right head/shank/thread length.
- **Snap-to-edge + depth-aware occlusion** — so when you place a plate against a beam flange, it actually meets the flange edge cleanly and beam material in front hides material behind.

### What to cut or de-prioritise (this is **not** a generic CAD)

- Spline tool (engineers don't draw splines in steel detailing)
- Architectural-style annotations
- Heavy 3D rendering / lighting / textures (the iso view should be flat shaded with line edges, like a structural model)
- Ribbon tabs that mimic Office for the sake of it — only what an engineer reaches for
- Live collaboration (single-user is fine)
- Cloud sync (file-based works)
- Stamps (engineers sign drawings via title block revision, not stamps)
- Algorithmic / generative anything

### Mental model for refinement

Treat V2's Bluebeam-style shell as the **chassis**. It got the multi-page sheet system, the snip-paste workflow, the layer panel, the ribbon, the snap engine, the measurement tools right. Now bolt back V1's engineering organs (parametric members, 3D iso, depth occlusion, weld symbols, dimension family, schedules, title block, PDF/DXF export) — but only the ones an engineer actually uses, refactored cleanly into the V2 architecture.

---

## Suggested priorities (ranked)

### Tier 1 — fix or restore before V2 reaches V1 parity for steel detailing

1. **Grip-handle resize and rotate** (the single biggest "this isn't really CAD yet" gap)
2. **Trim / Extend / Fillet / Offset** working implementations (basic versions)
3. **Group behaviour** (selecting one selects siblings, moving moves all)
4. **Element-level Lock honoured by selection / drag**
5. **Match Properties cursor revert**
6. **Layer drag-to-reorder**
7. **Vertical, Aligned, Angular, Radius, Chain, Baseline dimensions** (full family)
8. **Hatching** (concrete, steel, earth, timber, hidden) — Tool Chest → Properties hatch picker → renderer
9. **MText** with wrap and edit-in-place

### Tier 2 — restore V1's engineering value

10. **Parametric members** — UB / UC / PFC / SHS / RHS / CHS / EA / UA library with auto-drawing in elevation, plan, section
11. **Bolts (AS 1252)** in elevation and section, with bolt groups
12. **Plates** with polygon edit + fillet/chamfer on vertices
13. **Weld symbols (AS 1101)** with leader composition
14. **Section cut lines** with section view auto-projection
15. **Depth-aware occlusion** between members in 2D views
16. **Edge snapping** between objects with visual indicator lines
17. **Title block** with editable fields, AS 1100 layout, Bligh Tanner default
18. **Revision schedule** (tabular block)
19. **Multi-view layout** — bring back optional 4-pane (Elevation / Section A / Plan B / 3D Iso) within a sheet, AND keep the single-elevation-only mode for clean detail sheets
20. **3D isometric view** with Three.js r128 — flat-shaded, line-edged, orbit controls
21. **2D ↔ 3D toggle** that switches the whole workspace context (Dan's "really cool toggle")

### Tier 3 — outputs and polish

22. **PDF export** — vector first, raster fallback
23. **DXF export** for drafter handoff
24. **Scaled zones** actually applying their scale to elements within
25. **Snap-to-grid** when grid is on
26. **Status bar real-world coords toggle**
27. **Custom stroke width entry**
28. **Schema-versioned save format**
29. **Autosave to localStorage** (with unsaved-indicator in title bar)

### Tier 4 — schedules / reports

30. **Member schedule** auto-generated
31. **Bolt schedule** auto-generated
32. **Reinforcement schedule** for concrete details

---

## Workflow rules (non-negotiable)

1. **Edit `dev/index.html` only.** Never touch root `index.html` directly.
2. **Dan tests the dev file in a browser** at `file:///Users/danielmccarron/Documents/GitHub/bt-struct-tools/2D-Details/dev/index.html`. After Dan approves, copy `dev/index.html` → `index.html` (root). That's the parity-merge step.
3. **Dan handles all git commits and pushes himself.** Don't stage, commit, or push.
4. **Single HTML file.** No build step. Three.js r128 via CDN allowed when 3D returns.
5. **Metric only.** mm for everything. Y is up in world coordinates, down on canvas (flip in `real2px`).
6. **Australian Standards only:** AS 1100 (drawing), AS 4100 (steel), AS 3600 (concrete), AS 1101 (welds), AS 1252 (bolts), NCC.
7. **Variable conventions** (carry over from V1 brief): `u,v` = view-local 2D · `x,y,z` = world 3D · `px,py` = screen pixels · `sx,sy` = sheet-mm · functions prefixed `r` (e.g. `rLine`) draw in real-world coords · `v3d` prefix = Three.js 3D engine. V2 uses `s2px` / `px2s` for sheet ↔ screen conversion.
8. **AS 1100 lineweights:** 0.70 / 0.35 / 0.18 / 0.13 mm (`LINE_WEIGHTS` constant in V2).
9. **Three.js r128 only** — do not use APIs introduced after r128 (e.g. `CapsuleGeometry`).

---

## How to verify changes

- The current V2 file passes a JSDOM smoke test that exercises every command, ribbon tab, panel switch, and geometry helper without a single runtime error. Add to it as new behaviour lands.
- Open `dev/index.html` in a real browser. Click everything. The DevTools console should be clean.
- Compare against V1 (`archive/v1/index.html`) for any feature claimed to be parity.
- Quality bar: STP Typical Structural Details PDF, page 85, details 6011.1–6011.6. Any cap-plate, baseplate, splice, portal frame, web side plate, or tilt-up detail produced in StructDraw should look at least as clean.

---

## Final note for the next session

Don't trust this document blindly — re-read the V2 source (`dev/index.html`) and verify before each significant change. Memory snapshots go stale fast.

Be ruthless. The user is one engineer producing one kind of output. Every feature should serve that. If something doesn't help Dan produce a Bligh Tanner steel or concrete detail to AS standards, cut it.

— End of handoff —
