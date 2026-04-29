# StructDraw — Project Brief

**Author:** Dan McCarron, Senior Structural Engineer, Bligh Tanner (Brisbane)
**Current Version:** **V22 — Catalogue expansion** (`dev/index.html`)
**Last Updated:** 18 April 2026
**File:** Single-file HTML application (~14,400 lines)
**Dependencies:** Three.js r128 + jsPDF 2.5.1 via CDN

> **Historical note:** Sections 1–11 below were written during V7 and describe the original architecture. They remain the definitive source for the **coordinate system**, **data model**, and **rendering pipeline** — all of which are unchanged through V22. The top section below summarises the V15–V22 evolution layered on top.

---

## 0. Current Status (as at V22, 18 April 2026)

### What's been shipped through V22

| Version | Theme | Headline |
|---|---|---|
| V15 | Output you can issue | Vector PDF, weld polish, AS 1100 lineweights, dash table |
| V16 | 3-minute detail | Parametric connection library (cap plate, baseplate, WSP, splice) |
| V17 | Both themes beautiful | Sketch wobble + paper grain, slotted holes, centreline primitive |
| V18 | Everything labelled | Chain/baseline dims, parametric tags, section marks, material tags |
| V19 | A set, not a sheet | Multi-sheet project model, multi-page PDF, DXF export, revisions, detail cards |
| V20 | Fast as thought | Command palette (Ctrl+K), layer visibility UI, keyboard help, mirror tool |
| **V21** | **The masterpiece UI rebuild** | 3 themes (Classic/Dark/BT), ~50 hand-drawn SVG icons, tile palette with mode-switching (Model/Draw/Annotate), Inspector panel, sheet tabs, chord keyboard layer, favourites |
| **V22** | **Catalogue expansion** | 5 new section types (PFC/RHS/CHS/EA/UA, 94 profiles), 13 palette placeholders activated (Arc/Polygon/Offset/Fillet/Chamfer/Aligned/Angular/Hatch/MText/Grid Line/Note/Rev Schedule) |

### Live subsystems (all functional in V22)

- **Rendering**: 4 synchronised views (elevation, sectionA, planB, isometric), depth-aware occlusion, AS 1100 lineweight hierarchy (`CUT:1.2, VIS_HEAVY:0.7, VIS:0.65, MW:0.5, DIM:0.4, HID/CL:0.3, HATCH:0.18`)
- **3D engine**: Three.js r128 offscreen renderer for the isometric block (note: does NOT yet support V22 section types — they don't appear in the 3D view)
- **Exports**: vector PDF via `createPdfCanvasShim` (per-entity primitives, not raster), multi-page PDF for whole project, DXF R2013 with 13 AS 1100 layers
- **Connection builders**: `buildCapPlate`, `buildBaseplate`, `buildWSP`, `buildSplice` — atomic undo via `'connection'` action type
- **Multi-sheet**: `project = { sheets[], activeSheetIdx }`, hot-swap via `projectSwitchSheet`, `.sdproj` JSON save/load
- **Parametric tagging**: `memberTag` with `memberId` lookup, auto-numbered section marks, `revSchedule` auto-populates from `revisionTriangle` entities

### Known deferrals (V22.1b)

1. Three.js 3D builders for PFC/RHS/CHS/EA/UA — currently these section types don't render in the isometric block
2. True profile-specific DXF geometry for new section types (CHS as ARC primitives, PFC as proper C-profile) — currently they export as bounding-box polylines on `S-BEAM`
3. CHS through-cleat parametric connection (drafter §9.6)
4. Welded plate-girder parametric connection

### Deferred to V23

- Connection wizard migrated from modal → inline Inspector with live preview
- First-run welcome / 3-step tour
- Detail template library (save personal standard details for reuse)
- Preferences dialog (firm defaults, templates folder, default scale, sketch-theme defaults)
- Spline tool, Ordinate dim, Fill primitive

### Active code layout (V22, ~14,400 lines)

Single HTML file. Major sections, by approximate line range:

| Lines | Section |
|---|---|
| 1–789 | `<head>` — CSS (V21 design tokens, 3 themes, component styles) |
| 790–1139 | `<svg>` icon bank — ~50 custom symbols |
| 1140–1435 | `<body>` top-bar, main layout, status bar, inspector, dialogs |
| 1850–2200 | Section databases (UB, UC, SHS, PFC, RHS, CHS, EA, UA, BOLT) |
| 2230–2550 | 3D object model, undo/redo, project model |
| 2560–3900 | Hit-testing, bounds, snap, occlusion, coordinate transforms |
| 3900–4100 | DASH + LW constants, V22 `sectionProfile` helper |
| 4400–5100 | PDF export (vector + raster paths), DXF writer |
| 5200–6500 | Render pipeline, drawSheet, drawSheetGrid, drawBlockContent |
| 6500–7100 | Object renderers: drawUB, drawSHS, **drawSectionMember** (V22 unified), drawPlate, drawBolt |
| 7300–8700 | 2D entity renderers: dim, weld, member tag, section mark, rev, grid line, note, hatch, mtext, rev schedule |
| 8800–9100 | Connection builders (§9.x), `placeConnection`, connection dialog |
| 9200–9950 | Event handlers (mousedown dispatch for every tool), keyboard, chord layer |
| 10200–11000 | Multi-sheet project, save/load, multi-page PDF, sheet browser |
| 11000–12000 | V21 Inspector, size picker, favourites, command palette, layer UI |
| 12000–13300 | Tile palette (mode-switching Model/Draw/Annotate + placeholders) |
| 13300–14400 | Library init, toolbar wiring, Three.js 3D engine, app init |

### Critical files

| Path | Role |
|---|---|
| `dev/index.html` | The only code file — edit here first |
| `index.html` | Production copy, updated after Dan approves dev changes |
| `dev/CHANGELOG.md` | Full V15–V22 history |
| `dev/DRAFTER_MAPPING.md` | Drafter house-style rule ↔ StructDraw implementation crosswalk |
| `dev/STRUCTDRAW_PROJECT_BRIEF.md` | This file |
| `dev/HANDOFF.md` | Fresh-chat handoff doc (written at end of each long session) |

---

## 1. Vision

StructDraw is a browser-based 2D structural detail drawing tool for steel connections, built as a single HTML file. The goal is to replace Bluebeam and AutoCAD for producing professional-quality 2D connection details that look as good as (or better than) what a structural drafter produces in AutoCAD. The MVP target is a full steelwork connection detail set.

All engineering work follows Australian Standards (AS 3600, AS 1720.1, AS 4100, NCC). Drawing conventions follow AS 1100. Units are metric (mm).

---

## 2. Architecture Overview

### Single A1 Sheet Canvas

The entire application renders on one HTML5 Canvas element. The canvas represents a physical A1 sheet (841 × 594 mm) with margins (L20, R10, T10, B10) and a 30mm title block. Everything is drawn in "sheet-mm" space, then transformed to screen pixels via a zoom/pan viewport.

### Four Orthographic Detail Blocks

The sheet contains four `DetailBlock` instances, each projecting the same 3D object scene from a different angle:

| Block       | View Key     | Projection         | Shows               |
|-------------|-------------|---------------------|----------------------|
| Elevation   | `elevation` | `{u: x, v: y}`     | Front face (X,Y)    |
| Section A   | `sectionA`  | `{u: z, v: y}`     | Cut section (Z,Y)   |
| Plan B      | `planB`     | `{u: x, v: z}`     | Plan view (X,Z)     |
| Isometric   | `isometric` | Three.js 3D render  | 3D orbitable view    |

Each block can be repositioned by double-clicking its label and dragging. Section A and Plan B are linked to moveable cut-line positions (`secCutX`, `planCutY`) displayed on the elevation view.

### Coordinate Transform Chain

```
Real-world mm (x,y,z)
    → View-local (u,v) via projFn
        → Sheet-mm via ÷ drawingScale + block offset
            → Screen-px via × zoom + pan
```

Key transforms: `real2px()`, `px2real()`, `s2px()`, `px2s()`

**Y-flip:** Canvas Y is down, structural Y is up. `sy = block.sheetY - v / drawingScale`

**DPR:** All rendering uses `ctx.setTransform(DPR, ...)` for retina displays.

### 3D Isometric Engine (Offscreen)

The isometric view uses Three.js with an `OrthographicCamera` (not perspective) rendering to an offscreen canvas. The camera frustum is sized to match the 3D bounding box at 1:10 scale, so members appear the same size as in the 2D views. The rendered image is blitted onto the main canvas via `ctx.drawImage()`.

Orbit interaction: double-click the "ISOMETRIC" label to enter orbit mode. Drag to rotate (theta/phi). Press Enter or Escape to lock the angle. No zoom — scale is always 1:10.

---

## 3. Data Model

### 3D Objects (`objects3D` array)

All structural members are stored as 3D objects with real-world coordinates. Each has `{id, type, x, y, z, ...}`.

**UB (Universal Beam):**
```js
{ id, type:'ub', section:'360UB 50.7', x, y, z, length:600, rot:0 }
```
Section properties from `UB_DB`: d, bf, tf, tw, r1 (20 sections, 610UB 125 → 150UB 14.0)

**SHS (Square Hollow Section):**
```js
{ id, type:'shs', section:'89x3.5', x, y, z, length:500, rot:0 }
```
Section properties from `SHS_DB`: B, t (25 sections, 89×3.5 → 250×16)

**Plate (Legacy Rectangular):**
```js
{ id, type:'plate', x, y, z, pw:200, ph:300, pt:10, rot:0 }
```

**Plate (New Polygon — V7):**
```js
{
  id, type:'plate', x, y, z,          // centroid (world coords)
  polyPts: [{dx, dy, dz}, ...],       // vertex offsets from centroid
  pt: 12,                              // thickness (mm)
  normal: 'z'                          // thickness axis ('z'|'x'|'y')
}
```
The `normal` is set by which view the plate was drawn in: elevation→'z', sectionA→'x', planB→'y'.

**Bolt:**
```js
{ id, type:'bolt', boltSize:'M20', x, y, z }
```
Bolt data from `BOLT_DB`: d, head, headH, nut, nutH (M16, M20, M24)

### 2D Entities (`entities2D` object)

Separate from 3D objects. Stored per-view: `{elevation: [], sectionA: [], planB: []}`.

Types: `line`, `rect`, `circle`, `text`, `dim` (dimension). Each has `{id, type, view, lw, ...}`.

### Undo/Redo

Action-based stacks (max 100). Actions: `addObj`, `delObj`, `moveObj`, `addEnt2D`, `delEnt2D`. Each stores before/after snapshots.

---

## 4. Drawing & Rendering

### AS 1100 Lineweight Hierarchy

| Purpose  | Lineweight (mm) | Constant |
|----------|-----------------|----------|
| Cut line | 0.70            | `LW.CUT` |
| Visible  | 0.35            | `LW.VIS` |
| Hidden   | 0.18            | `LW.HID` |
| Dimension| 0.18            | `LW.DIM` |
| Centreline| 0.18           | `LW.CL`  |

### Depth-Aware Occlusion

Objects are sorted back-to-front (painter's algorithm) using `getDepthValue()`. Each object produces occlusion rectangles for objects behind it. Lines are clipped against these rectangles — visible segments draw solid, occluded segments draw dashed (AS 1100 hidden line convention).

Key functions: `getOcclusionRects()`, `clipLineAgainstOcclusion()`, `rLineOcc()`.

### Member Rendering

Each member type (`drawUB`, `drawSHS`, `drawPlate`, `drawBolt`) renders differently in each view:
- **Elevation:** Side view showing length, depth, flanges, web, centreline
- **Section A:** Cross-section (cut profile with thick outlines)
- **Plan B:** Top view showing length and width

Rotation is applied via `withRotation()` — a canvas transform wrapper that works in elevation and plan views.

### Polygon Plate Rendering (`drawPolyPlate`)

- **Face view** (the view where the plate was drawn): Renders the full polygon outline with cut-weight lines and a light fill.
- **Edge view** (orthogonal views): Renders as a rectangle (bounding box extent × thickness) with AS 1100 steel cross-hatching at 45°.

### 3D Object Rendering

Each type has a Three.js builder:
- **UB:** `ExtrudeGeometry` from I-section `THREE.Shape` (with fillet radii via `quadraticCurveTo`), rotated so extrusion direction maps to world X-axis
- **SHS:** Four-wall construction using `BoxGeometry` panels
- **Plate (polygon):** `ExtrudeGeometry` from polygon shape, rotated per normal axis
- **Plate (rectangular):** `BoxGeometry`
- **Bolt:** Cylinder shaft + hexagonal head + nut

Edge wireframes are added as children of each mesh (`mesh.add(edgeLine)`) so they inherit transforms automatically.

---

## 5. Interaction Model

### Tools

| Tool | Key | Description |
|------|-----|-------------|
| Select | V | Click to select, drag to move, box-select |
| Line | L | Two-click line |
| Rect | R | Two-click rectangle |
| Circle | C | Centre + radius |
| Polyline | P | Multi-click, Enter/dbl-click to finish |
| Dimension | D | Three clicks: P1, P2, offset |
| Text | T | Click to place, prompt for text |
| Draw Member | (library click) | Two-click: start → end, auto-length/rotation |
| Draw Plate | (library click) | Multi-click polygon, dbl-click to close, prompt thickness |

### Two-Click Member Drawing (V7)

1. Click a UB/SHS/Bolt in the library → enters `draw-member` mode
2. **First click:** Sets start point. Shows centreline preview (AS 1100 chain-dot), faint outline showing member depth, live dimension readout (mm + angle)
3. **Second click:** Places member with computed midpoint, length, and rotation
4. Member auto-selects with grips shown
5. **Ortho:** Hold Shift (or toggle F8) to constrain to H/V from start point
6. **Dynamic input:** Type a number mid-draw, press Enter to lock exact length
7. **Bolts:** Single-click placement (no line needed)
8. **Chained drawing:** After placement, stays in draw mode for next member

### Polygon Plate Drawing (V7)

1. Click "Draw Plate" in library → enters `draw-plate` mode
2. Click corners to define plate outline. Edges are ortho-locked by default (H/V). Hold Shift to free-draw for angled edges (gusset plates)
3. Each completed edge shows its length. Rubber-band dashed line follows cursor. Light fill previews the polygon shape
4. **Dynamic input:** Type a number mid-edge → Enter locks that edge to the exact dimension
5. **Double-click or Enter** (with 3+ vertices) closes the polygon
6. **Prompt for thickness** (e.g., "12" for a 12mm plate)
7. Plate is created with polygon vertices stored as offsets from centroid
8. **Right-click** undoes the last vertex. Escape cancels the polygon
9. View-aware: drawing in elevation creates thickness in Z, section in X, plan in Y

### Selection & Grips

- Click to select, Ctrl+click to toggle, box-select for multiple
- **UB/SHS:** End grips (extend length), rotation handle
- **Plate (rect):** Edge midpoint grips (resize each edge)
- **Plate (polygon):** Vertex grips in face view (drag any vertex)
- **Bolt:** Move only

### Snap System

- **Object snap** (F3): Snaps to endpoints, midpoints, centres of all objects
- **Edge snap:** During drag, objects magnetically snap to adjacent member faces (soft-snap: accumulates offset, releases past 3mm threshold)
- **Grid snap:** Rounds to grid size (1/5/10/25/100mm)
- **Ortho** (F8): Constrains to H/V from origin

### Pan & Zoom

- **Pan:** Middle-mouse or Space+Left
- **Zoom:** Scroll wheel (centred on cursor)
- **Fit:** F key or toolbar button

---

## 6. UI Layout

```
┌─────────────────────────────────────────────┐
│  Toolbar (tools, snap/ortho/grid, scale)    │
├────────┬────────────────────────────────────┤
│Library │                                    │
│ UB     │         Canvas                     │
│ SHS    │    (A1 Sheet with 4 views)         │
│ Bolts  │                                    │
│ Plates │                    [Props Panel]   │
│ Utils  │                                    │
├────────┴────────────────────────────────────┤
│  Status Bar (view, X, Y, Z, scale, tool)   │
└─────────────────────────────────────────────┘
```

### Themes

Two themes toggled via toolbar button:
- **Dark mode** (default): Dark red background (#7B1818), white entities
- **Classic mode:** White background, black entities

---

## 7. File Inventory

| File | Lines | Description |
|------|-------|-------------|
| `StructDraw.html` | ~2,800 | V1 — Original |
| `StructDraw_V2.html` | ~2,800 | V2 — Refinements |
| `StructDraw_V3.html` | ~3,400 | V3 — Section database, occlusion |
| `StructDraw_V4.html` | ~3,400 | V4 — Bug fixes |
| `StructDraw_V5.html` | ~3,900 | V5 — Grip handles, edge snap, rotation |
| `StructDraw_V6.html` | ~3,900 | V6 — 3D isometric view (offscreen ortho) |
| `StructDraw_V7.html` | ~4,680 | **V7 — Two-click drawing, polygon plates** |

**Active working file: `StructDraw_V7.html`**

All files are standalone single-file HTML — no build step, no dependencies beyond the Three.js CDN.

---

## 8. Code Structure (V7)

### Major Sections (by line range)

| Lines | Section |
|-------|---------|
| 1–187 | CSS (themes, layout, scrollbars) |
| 188–365 | HTML (toolbar, library, canvas, status bar, dialogs) |
| 366–460 | Constants (SHEET, DA, UB_DB, SHS_DB, BOLT_DB, LW) |
| 461–530 | Object management (mkObj, addObj, delObj, undo, redo) |
| 531–645 | DetailBlock class, projections, global state variables |
| 646–810 | Coordinate transforms, snap system, cursor utilities |
| 811–1000 | Bounding boxes, hit testing, block/view utilities |
| 1000–1240 | Grip handles (getGrips, hitTestGrip, applyGripDrag) |
| 1241–1485 | Edge snap system (soft-snap, magnetic faces) |
| 1486–1578 | Render loop (requestRender, render) |
| 1579–1900 | Sheet drawing (frame, grid, projection lines, cut lines) |
| 1900–2065 | Occlusion system (depth sort, clip, hidden lines) |
| 2065–2575 | Member rendering (drawUB, drawSHS, drawPlate, drawPolyPlate, drawBolt) |
| 2575–2640 | 2D entity rendering (drawEnt2D, drawDim2D) |
| 2640–2790 | Selection highlights, view labels |
| 2790–3110 | Crosshair, click preview (draw-member, draw-plate previews) |
| 3110–3570 | Event handling (mousedown, mousemove, mouseup, dblclick, wheel) |
| 3570–3760 | Component placement (placeComponent, finishDrawMember, finishDrawPlate) |
| 3760–3920 | Keyboard handling, tool state |
| 3920–4090 | Utility functions, status bar, properties panel |
| 4090–4220 | Library population, toolbar init |
| 4220–4260 | Layout, fit, resize |
| 4260–4680 | 3D engine (Three.js init, builders, orbit) |

### Key Patterns

1. **All state is global** — no classes beyond `DetailBlock`. Functions read/write shared arrays.
2. **Render is batched** — `requestRender()` uses `requestAnimationFrame` to avoid redundant draws.
3. **Undo is action-based** — each mutation pushes to `undoStack` with snapshots.
4. **Grips are computed per-frame** — `getGrips()` returns fresh descriptors each render.
5. **Polygon plates use offset storage** — `polyPts` stores `{dx, dy, dz}` from centroid so moving updates only `x, y, z`.

---

## 9. Known Limitations & Bugs to Address

1. **Rotation only works in elevation and plan** — `withRotation()` skips sectionA. Members drawn at angles in section will render unrotated.
2. **Polygon plate edge-view rendering** — the cross-hatch algorithm uses a bounding box approximation. Complex polygons may show incorrect extent in edge views.
3. **No fillet/chamfer on polygon plates** — right-click corner to add radius was suggested but not yet implemented.
4. **No weld symbols, break lines, or centreline entities** — library items exist but are stubs ("coming in next build").
5. **No PDF/DXF export** — buttons exist but are placeholders.
6. **3D bolt orientation** — bolts are always vertical (Y-axis); no support for horizontal bolts through webs.
7. **Old plate dialog HTML still in DOM** — the `#plateDialog` div is no longer used (replaced by draw-plate workflow) but hasn't been removed from the HTML.
8. **Grip handles for SHS rotation** — SHS members can be rotated but the rotation handle isn't rendered.
9. **No copy/paste of polygon plates** — the paste function may not correctly clone `polyPts`.
10. **Dimension entities** — only horizontal dimensions implemented; no vertical, aligned, or angular.

---

## 10. Roadmap / Next Features

### Near-term (V8 candidates)
- Weld symbols (fillet, butt, site weld) per AS 1100
- Break lines for shortened members
- Centreline entities (drawable, not just member centrelines)
- Aligned and angular dimensions
- Section labels (A-A, B-B)
- Layer system (toggle visibility of layers)
- Fillet/chamfer tool for polygon plate vertices
- Clean up: remove legacy plate dialog, remove legacy `place-component` code

### Medium-term
- PDF export (true vector A1 output at correct scale)
- DXF export (AutoCAD compatibility)
- Channel sections (PFC/UPE)
- Angle sections
- Welded plate girders
- Moment/shear connections as assemblies (parametric)
- Bolt groups (rectangular grid with auto-spacing per AS 4100)
- Multiple sheets/details

### Long-term
- Connection capacity checks per AS 4100
- Auto-dimensioning
- Detail library (save/load connection types)
- Multi-user / cloud sync
- Integration with structural analysis software

---

## 11. Development Notes

### How to work on this project

1. **Active file:** Always work on `StructDraw_V7.html`. Create a new copy (V8, V9...) for major changes.
2. **Single file:** Everything lives in one HTML file. CSS at top, HTML in body, JS in a single `<script>` block.
3. **No build step:** Open the HTML file directly in a browser to test.
4. **Three.js r128:** Loaded from CDN. Do not upgrade without checking API compatibility (e.g., `CapsuleGeometry` doesn't exist in r128).
5. **Test in browser:** After changes, always open the file and verify by drawing a few members, selecting, moving, checking all four views.

### Coding conventions

- Metric units throughout (mm). No imperial.
- AS 1100 drawing conventions (lineweights, hidden lines, centrelines, section markers).
- Variable naming: `u,v` for view-local 2D coords; `x,y,z` for world 3D coords; `px,py` for screen pixels; `sx,sy` for sheet-mm.
- Functions prefixed with `r` (e.g., `rLine`, `rRect`) draw in real-world coordinates.
- Functions prefixed with `v3d` relate to the Three.js 3D engine.
- `ppm()` returns pixels-per-mm at current zoom for lineweight scaling.
- All drawing functions take `blk` (DetailBlock) as first argument.

### Critical assumptions

- Y is UP in real-world coordinates, DOWN on canvas. The flip happens in `real2px()`.
- `drawingScale` divides real-world coordinates when converting to sheet (e.g., at 1:10, a 600mm beam is 60mm on sheet).
- `viewport.zoom` is screen-pixels per sheet-mm.
- The DPR multiplier is applied to the canvas context transform, not to individual coordinates.
