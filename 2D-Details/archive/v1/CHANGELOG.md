# StructDraw Changelog

## Current: V24.A3 — Orientation UX polish (19 April 2026)

Three friction points from the V24.A2 self-test addressed in one slice: new members now pick the correct axis from the view they're drawn in, the R key cycles roll instead of prompting for a scalar angle, and the Inspector shows a live visual preview of the current orientation.

### V24.A3.1 — Context-aware placement (§A.6 closed)
- New `_placementFrameForView(viewKey, rotDeg)` helper maps the 2D drawing angle in each view to the correct 3D axis preset:
  - **elevation** (u=+X, v=+Y): horizontal draw → +X, vertical → +Y. Non-orthogonal draws (>5° off cardinal) fall back to the legacy scalar-rot path so tilted beams still work.
  - **sectionA** (u=+Z, v=+Y): horizontal → +Z, vertical → +Y. Always snaps — the legacy rot path rotates in the X-Y plane and was wrong here.
  - **planB** (u=+X, v=+Z): horizontal → +X, vertical → +Z. Always snaps for the same reason.
- `finishDrawMember` routes through the helper and injects `axis` + `up` directly into `mkObj()` for orthogonal placements (skipping the legacy `rot` field). Non-ortho elevation draws keep their `rot` and migrate through `legacyRotToFrame` as before.
- Drawing a column in sectionA now produces a Z-axis member immediately — no Inspector round-trip.

### V24.A3.2 — R-key → roll cycle (§A.8 closed)
- The `prompt('Enter rotation angle')` dialog at the old R-key handler is gone. Replaced with:
  - **R** — roll +90° (cycles through 0 → 90 → 180 → 270 → 0)
  - **Shift+R** — roll −90°
  - **Alt+R** — flip axis direction end-for-end (+X ↔ −X, etc.)
- Uses `presetFromFrame` → `setMemberFrameFromPreset` for beam-like members so the frame stays coherent. Plates and bolts keep the legacy scalar-rot behaviour but without the prompt — R just rotates by ±90°.
- Single undo entry captures the whole multi-select rotation. `v3dMarkDirty` + `invalidateWeldCache` + `updateInspector` all fire so the 3D view, weld auto-detector, and Inspector dropdowns stay in sync.

### V24.A3.3 — Inspector orientation preview
- New `_inspOrientationPreviewSVG(o)` renders two 36×36 SVGs below the Axis/Dir/Roll dropdowns:
  - **Side view** — axis-coloured glyph (X=red, Y=green, Z=blue, matching the world gizmo). Horizontal bar for X, vertical bar for Y, into/out-of-page glyph for Z (dot for +Z toward viewer, × for −Z into page).
  - **End view** — the member's section icon (`#icon-ub`, `#icon-shs`, etc.) rotated by roll°. For sections with a clear orientation axis (UB, PFC, EA, UA) roll changes are immediately visible; symmetric sections (SHS, CHS) look the same at any roll, as they should.
- Labels under each SVG: `Side · +Y` / `End · 90°` — the abstract dropdown values become concrete.
- Shortcut hint row under the preview: `R roll +90° · ⇧R -90° · ⌥R flip`.
- **Targeted refresh** — when any dropdown changes, only the `#orientPreview` container is rebuilt (not the whole Inspector). Dropdown focus is preserved so the user can keep arrow-keying through options without clicking back into the control.

### Known limitations still deferred
- **A.10 DXF export orientation-aware** — rotated members still export as legacy orientation.
- **A.11 Projection lines use memberExtentOnAxis** — inter-view projection lines still read legacy `obj.x ± length/2`.
- **A.13 Wizard auto-cancel on orientation change** — connection wizard preview doesn't cancel when anchor reorients mid-preview.
- **Cut-hatching on rotated members** — still deferred (Phase B).
- **Occlusion on rotated members** — still deferred (Phase B).
- **Preset button row** — the "quick flip to column / perpendicular / flip end" row is planned for V24.A4 once A3 is validated.

### Invariants preserved
- Single HTML file, no build step.
- Three.js r128 APIs only (no new 3D code this slice).
- `objects3D` + `entities2D` remain the single source of truth.
- Legacy `.sdproj` and `.json` files load unchanged — the legacy-rot fallback still handles tilted-beam elevation draws.
- The triple Axis/Dir/Roll dropdown API is untouched; A3 adds beside it, nothing replaces.

## V24.A2 — Grips, Snap, Welds, 3D Axis Gizmo (19 April 2026)

Closes four rotation-related gaps surfaced by V24.A1 self-testing: grips follow rotated members, edge-snap finds rotated-member faces, weld auto-detector picks up interfaces on any face of a rotated member, and a 3D axis gizmo at the origin removes ambiguity about which way the world axes point.

### V24.A2.1 — Frame-aware grips
- `getGrips()` rewritten for all beam-like members (UB/UC/SHS/PFC/RHS/CHS/EA/UA). End-grip positions computed from `memberEndPoint(obj, ±1)` projected via `_viewBasis`. Grip visibility gated by `memberProjectedAxis(obj, vk).magUV > 0.3` — hidden when the member is seen end-on.
- Grip types renamed `end-left/right/top/bot` → `end-plus/end-minus` with a `sign` field. No longer view-dependent — one grip-type per member end in the 3D frame.
- `hitTestGrip()` simplified — the rotation-un-transform branch is gone. Grip (u, v) positions are already in view-local coords.
- `applyGripDrag()` member branches collapsed from ~140 lines of per-view, per-type math into ~20 lines: anchor the opposite end via `memberEndPoint(obj, -sign)`; lift the cursor back into 3D using the view basis + anchor depth; project cursor-delta onto `f.axis` to get the new length; recompute the centre so the anchor stays put.
- Rotation grip removed — orientation is now driven by Inspector Axis/Dir/Roll dropdowns (V24.A1); the old `rot` grip only edited the legacy scalar.

### V24.A2.2 — Frame-aware edge snap
- `getSnapEdges()` unified: for each of the three frame directions × half-extents, emit a snap line in whichever view axis the face normal projects onto. Handles all 24 orthogonal orientations.
- UB-specific flange-inner snap lines preserved (drafters snap plate bolts to inner flange faces).
- `applyEdgeSnap()` delta application collapsed to `_viewBasis(viewKey)`-driven world-vector addition. No more per-view if/else chain.

### V24.A2.3 — Frame-aware weld auto-detector
- `getObjFaces()` for beam-like members reimplemented using `memberExtentOnAxis(obj, axis)`. Emits the 6 AABB faces in world coords — the existing `computeWeldInterfaces` coplanarity + overlap test works unchanged because it already operated on axis-aligned spans.
- Any member reorientation now produces correctly-located welds. The existing `invalidateWeldCache()` hook (V23.1) fires on Inspector preset changes, so the weld set refreshes automatically.
- **Deliberate trade-off**: Phase A2 uses world-AABB faces, which over-report slightly for rotated members (an AABB is conservative vs the true oriented bounding box). For orthogonal orientations the AABB IS the OBB, so this is exact. Phase C will refine with true oriented faces when non-orthogonal frames are activated.

### V24.A2.4 — 3D axis gizmo
- `THREE.AxesHelper(500)` added at world origin in `v3dInit`. Renders X=red, Y=green, Z=blue (industry standard). 500mm arms are 5cm at 1:10 sheet scale — readable but not dominating.
- `depthTest: false` + `renderOrder: 999` so the gizmo reads through meshes when it would otherwise be occluded.
- Sits in world coordinates — serves as an orientation reference for any rotated members in the scene.

### V24.A2.5 — Self-test results (in-browser)
1. Horizontal SHS (axis +X), length 600, drag plus-end grip in plan — length extends, minus end stays pinned. ✓
2. Vertical SHS (axis +Y), length 600, drag plus-end grip in elevation — length extends upward. ✓
3. Horizontal SHS + plate at right end (x=400..600) → weld auto-detected at x=400, spanning SHS right face. ✓
4. AxesHelper visible at origin in 3D iso with clear X=red / Y=green / Z=blue arrows. ✓

### Known limitations still deferred
- **A.6 Placement axis-from-view** — drawing a new member always produces an X- or Y-axis member; flipping via the Inspector still required for Z-axis placement.
- **A.8 R-key cycles roll** — rotation keyboard shortcut still prompts for scalar angle.
- **A.10 DXF export orientation-aware** — rotated members export as legacy orientation.
- **A.11 Projection lines use memberExtentOnAxis** — inter-view projection lines still read legacy `obj.x ± length/2`.
- **A.13 Wizard auto-cancel on orientation change** — connection wizard preview doesn't cancel when anchor reorients mid-preview.
- **Cut-hatching on rotated members** — still deferred (Phase B).
- **Occlusion on rotated members** — still deferred (Phase B).

## V24.A1 — Full 3D Member Orientation, Phase A slice 1 (18 April 2026)

Members now have a full 3D local frame (axis + up unit vectors) instead of a scalar rotation about Z. The 24 orthogonal orientations — any axis direction (±X/±Y/±Z) combined with any roll (0/90/180/270°) — are selectable from the Inspector and render correctly in both the 2D views and the 3D iso. Legacy files load unchanged: a UB placed horizontally in V23.1 still looks identical in V24.A1.

### V24.A0 — Foundation (pre-requisite fix + utilities)
- **Fixed blocking page-load `SyntaxError`**: three duplicate top-level `const` declarations (`PFC_DB`, `EA_DB`, `UA_DB`) left over from a half-finished V22.1 migration have been removed. The drafter-style catalogue (mass-annotated keys like `"380PFC 55.2"`) remains as the single source.
- **New data model fields**: every beam-like member (UB/UC/SHS/PFC/RHS/CHS/EA/UA) now carries `axis` and `up` unit vectors. `mkObj()` populates them on every new member; `migrateLegacyMember()` backfills them on old project loads.
- **Frame utility library** (~280 lines): `memberFrame`, `memberProjectedAxis`, `memberProjectedUp`, `memberViewMode`, `memberViewAngle`, `memberEndPoint`, `memberExtentOnAxis`, `frameFromPreset`, `presetFromFrame`, `setMemberFrameFromPreset`, plus 3D vector helpers (`_vNorm` / `_vCross` / `_vDot` / etc.) and view-basis lookups (`_viewBasis`).
- **Legacy migration** wired into `_projectLoadSheet` and the single-sheet `loadProject` path. Every member in the scene has `axis` + `up` by the time render runs.

### V24.A1 — Render + Inspector + 3D iso (updated after self-test)
- **Per-view rendering proxy** (`drawMemberProxied`) — for each member in each view, picks the right existing renderer branch (`'elevation'` / `'sectionA'` / `'planB'`) based on `memberViewMode`, remaps the member's (x, y, z) into the branch's coordinate system, and applies the in-plane 2D rotation externally via `ctx.save`/`translate`/`rotate`. Legacy-frame members take a fast-path passthrough so their rendering is byte-identical to V23.1.
- **Wired into dispatch** for `drawUB`, `drawSHS`, and `drawSectionMember` (the V22.1 unified renderer for PFC/RHS/CHS/EA/UA).
- **Frame-aware bounds** — new `_memberFrame2DBounds(obj, vk)` computes projected AABBs for any orientation by summing `|projection|` of each local half-dim onto the view's u/v basis. `getObj2DBounds` and `get2DFootprint` both use it as a fast path. This fixes click-selection on rotated members without any change to `hitTest3D` itself.
- **Frame-aware cut-class / occlusion** — `getObjAxisExtent` delegates to `memberExtentOnAxis` when a frame is present, so the live section-cut planes correctly classify rotated members as visible / hidden / cut.
- **Inspector orientation fields** — for member selections, the old "Rotation (°)" input is replaced with three dropdowns: Axis (X/Y/Z), Direction (+/−), Roll (0/90/180/270°). Together they pick one of the 24 orthogonal presets. Writing to any of them calls `setMemberFrameFromPreset()` which updates `axis` + `up` and triggers a re-render + 3D rebuild. A hidden `propRot` input is preserved for any legacy handlers that might still reference it.
- **Three.js iso sync** — new helper `_v3dApplyMemberFrame(pivot, obj)` constructs a rotation matrix from the frame's basis and applies it as a quaternion to the pivot. Used by `v3dBuildUB` and `v3dBuildSHS`. Three.js r128 APIs only (`Matrix4.makeBasis` + `Quaternion.setFromRotationMatrix`).
- **Connection builder Phase-A gate** — `_connRequireLegacyColumn` / `_connRequireLegacyBeam` helpers added and called at the top of `buildCapPlate`, `buildBaseplate`, `buildWSP`, `buildSplice`. They throw a friendly "Phase B required" error if the user tries to build a connection on a rotated anchor. Phase B removes the gates.

### V24.A1 — Self-test fix (18 April 2026)
Browser-driven self-test (flipped a horizontal UB to Z-axis, inspected all four views) uncovered a bug in the renderer proxy: the "blank occRects" cleanup was indexing the wrong element of the `rest` array via `rest.map((x,i) => i===5 ? [] : x)`, which wiped `cutClass` (index 5) instead of `occRects` (index 4). The second line `restClean[4] = []` correctly blanked occRects but never un-blanked cutClass. Result: section A of a rotated UB was receiving cutClass=cut (from the original view) and rendering an unwanted cross-section overlay.

Fix: removed the buggy map, unified to a single `restClean` that explicitly blanks both `occRects` (index 4) and `cutClass` (index 5). Cut-hatching on rotated members is now consistently deferred — renderers always get `cutClass=null` through the proxy. After the fix the three 2D views render correctly: elevation shows the end-on UB silhouette, section A shows the UB side-view, plan B shows the 171×600 plan. 3D iso matches (confirmed via quaternion frame application).

### Known limitations in V24.A1 (deferred to later slices)
- **Grips** — drag-to-extend works correctly only for members on their native axis (X-axis beams in elevation, etc.). Rotated members show grips in the wrong view or don't respond to drag. This is A.5 in the plan; one focused session.
- **Canvas placement** — drawing a new member with the "Line/UB/UC" tool always produces an X-axis or Y-axis member. Post-placement Inspector flip works; placing directly into Z requires extra logic (A.6).
- **V22.1 section types in 3D** — PFC / RHS / CHS / EA / UA still don't appear in the iso view (no `v3dBuild*` builders exist for them yet — known pre-V24 gap, see HANDOFF §4 "V22.1b").
- **DXF export** — currently emits legacy-orientation geometry. A rotated member exports as if it were on X-axis (A.10).
- **Projection lines between views** — still read legacy `obj.x ± length/2`; won't follow rotated members until A.11.
- **Keyboard R** — still prompts for scalar rotation angle; not yet routed to cycle roll (A.8).
- **Connection builders** — throw on rotated anchors. Phase B makes them orientation-aware.
- **Occlusion on rotated members** — temporarily disabled (`occRects` passed as `[]` to proxied renderers). Rotated members won't show hidden-line dashing where other members occlude them. Primary-geometry rendering is correct; hidden-line dashing re-projection is in a later slice.

### Invariants preserved
- Single HTML file, no build step.
- Three.js r128 APIs only.
- `objects3D` + `entities2D` remain the single source of truth.
- Y-up world / Y-down canvas flip unchanged.
- AS 1100 lineweight hierarchy, hatch patterns, centrelines unchanged.
- Legacy `.sdproj` and `.json` files load with members rendering byte-identical to V23.1 (verified via the fast-path passthrough for legacy frames).
- V23.1 inline connection wizard continues to work for default-orientation anchors.

## V23.1 — Inline Connection Wizard (18 April 2026)

The four connection builders (cap plate, baseplate, WSP, splice) no longer live behind a blocking modal. They now populate the right-hand Inspector panel with a live ghost preview on the sheet as the user ticks parameters.

### V23.1.1 — Inspector-driven wizard
- **New state machine** `connWizState` holds `{ kind, spec, anchor, params }` for the life of a wizard session. Null when no wizard is open.
- **`openConnectionDialog(kind)` rewritten** — same call site (tile clicks, command palette, hamburger menu, keyboard chords) but now seeds `connWizState` and routes the Inspector to render the wizard fields instead of opening a modal overlay.
- **New helpers**: `connWizClearPreview` / `connWizRebuildPreview` / `connWizCommit` / `connWizCancel` / `connWizScheduleRebuild`.
- **Inspector route** — `updateInspector()` gains a top-priority branch that calls `_inspConnectionHtml(state)` + `_wireConnectionInputs()` when a wizard is active.
- **`_inspConnectionHtml`** — two-column `.insp-row` layout per field, anchor summary / error banner / no-anchor prompt, Create + Cancel buttons. Reuses existing `.insp-field` + `.insp-btn` CSS — no new styles required.
- **`_wireConnectionInputs`** — binds `input` (numbers) or `change` (selects) events to `connWizState.params[key]`, schedules a rAF-coalesced rebuild. Every parameter tweak re-runs the builder, clears old preview objects, and injects fresh ones.

### V23.1.2 — Ghost preview rendering
- **`__preview: true` flag** on all objects/entities produced by the wizard. They live inside the real `objects3D` and `entities2D` arrays so existing renderers pick them up with zero patching.
- **Opacity wrap in `drawBlockContent`** — dispatch sites for member renderers (drawUB / drawSHS / drawPlate / drawBolt / drawSectionMember) and the 2D entity loop are wrapped in `ctx.save(); ctx.globalAlpha *= 0.5; … ctx.restore()` when `obj.__preview === true`. One localized change, no new render pass.
- **Hit-test exclusion** — `hitTest3D` + `hitTestAll3D` skip preview objects so the ghost can't be grabbed or Tab-cycled.
- **Relational exclusion** — `computeBoltGripInfo` and `computePlateHoles` segregate preview and committed items so they don't cross-contaminate bolt-through-plate calculations.

### V23.1.3 — Commit + lifecycle
- **Create button → `connWizCommit`** — strips `__preview` flags, captures snapshots, pushes ONE `{ act:'connection', objSnaps, entSnaps }` onto the undo stack. Ctrl+Z removes the whole connection in a single step, matching the old modal behaviour.
- **Cancel button / Escape key → `connWizCancel`** — splices preview items out of the arrays; no undo entry pushed.
- **Auto-cancel on export/serialize** — `exportSheetToPDF`, `exportSheetToDXF`, `exportProject`, `exportProjectToPDF`, `saveProject`, and `projectSwitchSheet` all call `connWizCancel()` at entry so ghost objects can never leak into a PDF/DXF/.sdproj file or carry across sheets.

### V23.1.4 — Modal cleanup
- Removed the `#connectionDialog` DOM block. HTML comment left in place documenting the V23.1 replacement.
- ~60 lines of modal-specific JS (dialog open/populate/bind OK handler) gone — net code delta is a slight reduction despite the new Inspector functions.

### Known limitations
- **3D iso view** does NOT show the ghost preview — the Three.js engine only re-syncs on `v3dMarkDirty()`, which is called on commit. Acceptable MVP scope; re-sync on every param tweak would churn.
- **No Enter-to-commit** — the Create button is mouse-only for this slice. Enter in a number input would conflict with spinner increment. Deferred to a future refinement.
- **Preview during another tool** — opening a wizard doesn't clear the active tool. Tool state resumes on commit/cancel.

## V22 — Catalogue expansion (18 April 2026)

Every palette placeholder from V21 now lights up. Dan can draft with the full tool kit in one session.

### V22.1 — New section catalogue (PFC / RHS / CHS / EA / UA)
- **5 new section databases**: `PFC_DB` (10 profiles, AS/NZS 3679.1), `RHS_DB` (13 profiles, AS 1163), `CHS_DB` (19 profiles, AS 1163), `EA_DB` (33 profiles AS/NZS 3679.1 equal angles), `UA_DB` (19 profiles unequal angles). ~94 new section entries total.
- **`sectionProfile(obj)` helper** — unified `{ d, bf, tf, tw, r1, t, D, shape }` return for any structural member type. Backs the geometry pipeline so UB/SHS existing codepaths stay untouched while the new types share infrastructure.
- **`isMemberType(type)`** — single predicate for beam-like types; replaces scattered `type === 'ub' || type === 'shs'` checks.
- **`drawSectionMember()`** — unified per-view renderer for all 5 new types. Elevation draws side view with section-specific inner lines (PFC flange tips, CHS hidden back-edge, angle leg thickness). Section A draws the true profile: PFC C-shape, RHS hollow box, CHS concentric circles, EA/UA L-shape. Plan B draws the length-vs-width rectangle. Cut-class rendering includes AS 1100 cross-hatching in section views.
- **Hit-testing** via `getObj2DBounds` and `getObjAxisExtent` patched with unified fallback for new types.
- **DXF export**: new `_dxfEmitGenericMember` emits bounding-box polylines on `S-BEAM` with centrelines on `S-CL`. True profile-specific DXF (arcs for CHS, C-profile for PFC) deferred to V22.1b.
- **Tile palette**: 5 faded "soon" tiles converted to live tiles with size-picker dropdowns. `_pickerItemsFor` extended with PFC (grouped by depth series), RHS / CHS / EA / UA (flat lists).
- **`finishDrawMember`**: new branch creates objects of type `pfc`/`rhs`/`chs`/`ea`/`ua` via `mkObj`.

### V22.2 — Aligned + Angular dim tools activated
- Aligned and Angular tile placeholders become live. `drawDim2D` already supported these modes from V18; V22.2 wires them to the palette tiles with proper `dimType` assignment on click.

### V22.3 — Arc, Polygon, Offset tools
- **Arc**: 3-click tool. Picks start, midpoint, end; computes circumscribed circle centre/radius via determinant formula; emits `arc` entity with `a0/a1/ccw` direction derived from midpoint sweep side.
- **Polygon**: 2-click (centre + vertex) with prompt for N sides (3–24). Emits `polygon` entity with vertex array.
- **Offset**: 2-step. Click 1 picks a source line (within 10mm tolerance). Click 2 picks the offset side + distance, deriving perpendicular unit vector × signed distance. Creates a new `line` entity parallel to the source.
- Three new dispatcher cases in `drawEnt2D` for `arc` and `polygon`.

### V22.4 — Fillet + Chamfer tools
- **Fillet**: Click near a polygon-plate vertex → prompt for radius → vertex is replaced with an 8-segment approximate arc from the bisector-offset centre. Supports elevation / sectionA / planB views.
- **Chamfer**: same interaction, but replaces the vertex with two points offset by `distance` along each adjacent edge.
- New helper `_uvToDelta(viewKey, uv, currentVertex)` converts view-local (u,v) points back into plate `polyPts` `{dx,dy,dz}` deltas, preserving the depth-axis component from the original vertex.
- Undo-stack entry uses the existing `moveObj` action with before/after snapshots of `polyPts`.

### V22.5 — Grid Line + Note
- **Grid Line**: 2-click drawable entity with chain-dash line + circular bubble at one or both ends, containing a label (A, B, 1, 2…). New `drawGridLine2D` renderer.
- **Note**: 2-click (anchor + text position) with prompt. Supports multi-line via `\n` in the text. Renders leader + filled-arrow at anchor + left/right-aligned text block based on leader direction.

### V22.6 — Hatch, MText, Rev Schedule
- **Hatch**: multi-click polygon (close on Enter/dblclick) with prompt for pattern — **steel** (AS 1100 45° diagonal), **cross** (double-hatch), **concrete** (dot array). Uses `ctx.clip()` with the polygon outline to confine the hatch to the region. Respects `--entity-color` with 0.5–0.7 alpha so it reads as secondary fill.
- **MText**: single-click placement with text + width prompts. Implements word-wrapping by measuring each candidate line against `ctx.measureText()` and breaking at word boundaries. Lines break explicitly on `\n`.
- **Rev Schedule**: single-click placement of a table anchored at top-left. Auto-scans *all* views for `revisionTriangle` entities, aggregates by rev number, and renders a 3-column table: **REV | DESCRIPTION | DATE**. Update the table by simply placing another rev triangle with the same number — the schedule re-renders live. Rev-triangle tool now also prompts for optional description + date so the schedule can auto-populate.

### Net scope
- ~2,000 lines of new code
- 94 new section-DB entries
- 9 new entity types (`arc`, `polygon`, `gridLine`, `note`, `hatch`, `mtext`, `revSchedule`, plus `pfc`/`rhs`/`chs`/`ea`/`ua` 3D member types)
- 13 tiles converted from "soon" placeholders to live tools
- All engine-level subsystems touched (render, hit-test, bounds, DXF) via minimal-footprint helpers rather than 30-site type-switch edits

### Deferred to V22.1b
- True profile-specific DXF entities (CHS circles as ARCs, PFC as proper C-profile polyline)
- Three.js 3D builders for PFC/RHS/CHS/EA/UA (currently not shown in isometric view)
- CHS through-cleat parametric connection (drafter §9.6)
- Welded plate-girder parametric connection

## V21 — The masterpiece UI rebuild (18 April 2026)

### V21.11 — Professional icon redraw
All ~50 icons rewritten from scratch to a commercial-CAD aesthetic. Single monoline language, no decorative fills, no 15%-opacity "lit" look.
- **Stroke weights**: primary outline 1.25px (set on base `.icon`), secondary detail 0.9px, accent 1.5px — overridden per element via `stroke-width` attributes.
- **Section profiles** now drawn with **root fillets** at flange–web junctions (UB/UC via quadratic-curve fillets), rounded outer corners on SHS/RHS (matching real hot-rolled sections), centreline crosshair on CHS, proper heel radius on EA/UA. Faint chain-dot web centrelines on UB/UC/CHS for engineering legibility.
- **Bolt** is now a proper side-elevation: chamfered hex head with visible chamfer line, separate washer, cylindrical shank, chamfered nut, and AS 1100 thread hash on the protrusion.
- **Bolt group** is a plan-view 2×2 with chamfered hex tops and AS 1100 chain-dot centreline crosshairs between them — instantly reads as a standard bolt group.
- **Slotted hole** extended chain-dot centreline past each cap.
- **Cap plate / Baseplate / WSP / Splice** rebuilt as proper technical drawings. Cap plate shows plate + column I-section below + bolts with visible stems. Baseplate shows column I-section + plate + HD bolts with thread hash. WSP shows column flange end + cleat plate + beam with visible flanges + 3 bolts. Splice shows two UB ends + two end-plates with mill gap + 4-bolt vertical pattern.
- **Dimensions** now draw as real AS 1100 dimension lines: witness lines + dim line + slash ticks (horiz/vert/aligned). Angular dim has a proper arc sweep with filled arrowheads. Chain has four linked segments with ticks. Baseline has three tiered dims from a datum with ticks at each stop. Ordinate shows origin circle + L-leader + arrow.
- **Section mark** is a heavy AS 1100 chain-dash cut line with **solid triangular arrows** at each end and A/A labels.
- **Weld** is a proper AS 1101.3 leader: arrow + reference line + filled fillet-triangle glyph below.
- **Detail reference** is the Australian circle-with-horizontal-divider callout — detail number above, sheet number below.
- **Grid line** is a bubble with letter on top of an AS 1100 chain-dot line.
- **Member tag / material tag** use underlined text convention ("UB", "PL") next to leader.
- **Revision triangle** is a clean stroked triangle with centred monospace number — no fill-opacity decoration.
- **UI chrome**: PDF/DXF have bottom badges with the acronym in inverse colour. Layers is a proper 3-deep isometric stack. Undo/Redo have cleaner arrow cusps. Fit shows a viewport frame with inner guide rect. Mirror shows two reflected arrows across a dashed axis. Theme is a crisp day/night half-disc.

### V21.10 — Compact palette
After visual review, tiles were too large and childish-feeling. Reduced density dramatically:
- **Tile size**: 78×78 → 48×48 (60 for tiles with a size sub-label)
- **Palette grid**: 2-column → **3-column**
- **Palette width**: 184px → 170px
- **Inspector width**: 280px → 260px
- **Favourites**: 46px 3-col strip → 32px 4-col strip (labels dropped — icon + tooltip only)
- **Group separator height**: ~22px → ~12px
- **Active tile**: accent left-border instead of 2px all-around border (no content jump)
- **Pure-icon tiles** (Arc, Line, Rect, etc.) now omit the redundant text label row entirely — icon + tooltip carry the meaning, matching Figma/Sketch convention
- Tiles with sub-sizes (UB "360UB 50.7", SHS "150x6", Bolt "M20") keep the sub-label in **monospace** for alignment
- **Inspector**: 10/14 section padding → 7/10; title margin 8 → 5; field height 28 → 24; font 12 → 11.5
- **Inspector Drawing/People rows**: 3-col equal → `triple` grid (1.6fr 0.8fr 1fr) so Drawing No gets width while Rev stays compact
- Net result: **~18 tiles visible vs V21.9's ~10**, fits complete Model or Annotate palette on one screen without scrolling



**The interface catches up with the engine.** Every screen region rewritten from scratch; zero changes to the drawing engine, exports, or data model.

### V21.1 — Design system foundation
- Three themes: **Classic B&W (new default)**, **Dark (neutral charcoal)**, **BT-Red (legacy brand)**. Legacy V20 token names aliased to new V21 semantic tokens (`--surface`, `--text`, `--accent`, `--border`, `--shadow-*`, etc.) so existing CSS rules inherit the new palette.
- Theme persists to `localStorage.structdraw_theme`; cycles via single top-bar button.
- Typography: Inter (with system-ui fallback), tabular monospace for number inputs.
- ~250 lines of new component CSS: `.tile`, `.palette`, `.palette-group`, `.sheet-tab`, `.mode-switcher`, `.icon-btn`, `.menu`, `.inspector`, `.status-bar`, `.sb-toggle`, `.picker`, `.chord-overlay`, `.soon-popover`, `.fav-tile`.

### V21.2 — Icon bank (~50 hand-drawn SVG symbols)
- Inline `<svg><defs>` bank at top of body. All UI icons as `<symbol id="icon-XXX">`. Usage: `<svg class="icon"><use href="#icon-XXX"/></svg>`.
- 20×20 viewBox, 1.5px stroke, `currentColor` — inherits theme automatically.
- **Section profiles** (UB, UC, PFC, SHS, RHS, CHS, EA, UA) drawn as technical end-views — instantly recognisable. **Fasteners** (Bolt hex-head, bolt group 2×2, slot stadium). **Connection glyphs** (cap plate, baseplate, WSP, splice). **Drawing primitives** (line, rect, circle, arc, polyline, polygon, spline, hatch, fill, break, offset, fillet, chamfer). **Text + dim variants** (H, V, aligned, angular, chain, baseline, ordinate). **Annotations** (section mark, weld, detail-ref bubble, grid-line, member tag, bolt callout, material tag, note). **Revisions** (triangle, cloud, schedule). **UI chrome** (select pointer, menu, help, theme, save, load, close, plus, model, draw, annotate, PDF, DXF, layers, undo, redo, fit, search, keyboard, mirror).

### V21.3 — Shell layout
- **Top bar (44px)**: brand ("StructDraw" + logo mark) · Chrome-style sheet tabs (replaces old sidebar sheet browser) · mode switcher **top-right** (Model / Draw / Annotate 3-segment pill) · icon-button actions (Undo, Redo, Fit, Layers, ?, Theme, ☰).
- **Hamburger menu** dropdown: File (new sheet, save, load), Export (PDF, PDF-all, DXF), Sheet (title block, 3D toggle), Help (command palette, keyboard shortcuts).
- **Status bar (28px)**: view chip · live X Y Z coord readout (monospace) · Scale / Grid / Nudge selects · **Snap / Ortho / Grid toggles moved here** (AutoCAD convention — dot-filled = on) · current tool chip · object/entity/selection counts · scale.
- **Three-column main layout**: palette (184px) · canvas · inspector (280px). Old floating props panel and left library sidebar both removed.
- Every legacy DOM ID preserved via hidden shim div — existing handlers (setTool, updateStatus, dialog wires) unchanged.

### V21.4 — Tile palette with 3-mode filtering
- **New tile component** (78×78px): 28px icon + last-used-size label + dropdown chevron. Hover, active, placeholder states. SVG icons stroke in `currentColor`.
- `populateTilePalette()` (replaces flat `populateLibrary()`) — mode-filtered, grouped. Model palette groups: **Sections / Fasteners / Plates / Connections**. Draw palette: **Primitives / Fill / Text / Construction**. Annotate palette: **Dimensions / Tags / Symbols / Revisions**.
- **Placeholder tiles** (V22/V23): RHS, PFC, CHS, EA, UA, Arc, Polygon, Spline, Hatch, Fill, MText, Offset, Fillet, Chamfer, Aligned, Angular, Ordinate, Note, Grid Line, Rev Schedule. Faded (35% opacity), "soon" badge bottom-right. Click shows a popover explaining the planned feature and its version.
- `highlightActiveTile()` keeps the tile UI in sync with the tool state.
- `selectMemberBySection` / `selectMemberByBolt` helpers replace inline library click handlers.

### V21.5 — Size-picker dropdown
- `openSizePicker(kind, tileId, anchorEl)`. Anchored next to the tile's chevron.
- Grouped by depth series for UB/UC (610 Series → 610UB 125/113/101, etc.) — matches how engineers mentally pick sections (depth first, weight second). Flat for SHS/bolts.
- Fuzzy substring search box at top.
- Last-used entry marked with ⭐.
- Keyboard: ↑↓ navigate, Enter confirms, Esc closes, outside-click closes.

### V21.6 — Inspector panel (right dock)
- Replaces the old floating `#propsPanel` with a proper context-switching right-docked panel, always visible.
- **Empty state**: inline sheet-info editor — Project, Client, Description, Drawing No, Revision, Date, Sheet of, Designer, Drawn, Checker, Firm name, Firm tagline. Editing updates the sheet live (no need to open the modal title-block dialog).
- **Single-member state**: section / length / rotation / position x/y/z / Duplicate / Delete.
- **Multi-selection state**: composition summary (counts per type) + Delete all.
- **Tool-active state**: current tool name + Layer / Lineweight / Line-style options + Esc-to-cancel hint.
- Hooks into existing `updateStatus()` so the Inspector stays in sync with every selection/tool change.

### V21.7 — Keyboard chord layer
- Press **M** / **D** / **A** → 350ms later a chord overlay appears listing mode-specific second-key options. Press the second key to jump instantly to the tool.
- Full chord map: `M U` UB, `M C` UC, `M S` SHS, `M B` bolt, `M G` bolt group, `M L` plate, `M K` cap plate. `D L` line, `D R` rect, `D C` circle, `D P` polyline, `D T` text, `D B` break. `A H` dim horiz, `A V` dim vert, `A C` dim chain, `A S` section mark, `A W` weld, `A T` member tag, `A R` rev triangle.
- Esc cancels a pending chord. Fast typing bypasses the 350ms wait.
- V20 single-key shortcuts (V=select, L=line, etc.) preserved.

### V21.8 — Favourites tracking
- LRU queue of last 6 tiles used, persisted to `localStorage.structdraw_favourites`.
- Favourites strip at the top of the palette (across all three modes — a bolt stays useful anywhere).
- Right-click any favourite → **Pin / Unpin / Remove**. Pinned favourites show a small accent corner indicator.
- Automatic tracking on every `selectMemberBySection` / `selectMemberByBolt` / tile click.

### V21.9 — Polish & migration
- Every legacy V20 DOM id preserved via hidden shim elements so no handler rewrites were needed.
- `wireSheetBrowser` neutralised (V21 wires inside `initToolbar`); `renderSheetBrowser` rewritten to emit Chrome-style tabs instead of sidebar rows.
- Connection wizard still modal — Inspector migration deferred to V23 per plan.
- Zero changes to: render pipeline, every `draw*` function, every export (vector PDF, raster PDF, DXF, multi-page PDF), every connection builder, canvas event handlers, undo/redo, 3D engine.

### What's NOT in V21 (deliberate)
- **Build step / modules** — kept single-file. File now at ~12,500 lines but organised by clear banners.
- **Real PFC/RHS/CHS/EA/UA renderers** — placeholders only. Shipped V22.
- **Inline connection wizard (live preview)** — kept modal for V21; V23 target.
- **First-run tour** — V23.
- **Drag-to-reorder sheet tabs** — stretch goal dropped.

## V20 — Fast as thought (18 April 2026)
- **Command palette** (Ctrl+K / Cmd+K). Fuzzy-matched list of every tool, library item, connection wizard, view toggle, and export action. ~40 commands indexed. Substring + subsequence scoring; Arrow keys navigate, Enter runs, Escape closes.
- **Layer visibility UI** — floating right-side panel (Layers button in toolbar). 10 semantic groups: Members, Plates, Bolts, Welds, Dimensions, Text + Tags, Centrelines, Section marks, Revisions, Construction. Each group maps to its underlying 3D/entity types via `layerVisibility` object; render pipeline gates both `drawBlockContent` (3D objects) and `drawEnt2D` (2D entities). Change once → hides across all views.
- **Keyboard help overlay** (toggle via `?` key or toolbar `?` button). Shows every shortcut grouped into Tools / Edit / View & snap / Project.
- **Mirror tool** (M key, library item, palette entry). Two clicks define the mirror axis in the active view; `performMirror(a1, a2, viewKey)` creates new mirrored copies of every selected 3D object + all 2D entities in that view. Mirrors 3D position, flips angles, reverses baseline stops.
- New toolbar buttons: Layers, ?
- New CSS: `kbd` styling for keyboard keys in help overlay, `.cmd-row` for palette rows, `.layer-row` for layer panel rows.

## V19 — A set, not a sheet (18 April 2026)
- **Multi-sheet project model (V19.5)**. Project = array of sheets; each sheet owns its own `objects3D`, `entities2D`, `sheetInfo`, cut-line positions, and id counters. Hot-swap via `projectSwitchSheet` snapshots live globals → restores target. Render pipeline remains unchanged — it always reads the same globals.
  - `projectInit`, `projectAddSheet`, `projectDeleteSheet`, `projectRenameSheet`, `projectSwitchSheet`, `_projectSnapshotActive`, `_projectLoadSheet`, `_projectMakeSheet`.
  - **Sheet browser** sidebar at top of library panel — per-sheet row with drawing number, name, rename (✎), delete (✕). "New" button creates a fresh sheet.
- **Multi-page PDF export** — new `exportProjectToPDF` hot-swaps every sheet in turn through the V15 vector path, pushing each onto the same jsPDF document via `addPage`. Original active sheet restored after.
- **Project save/load** — `exportProject` writes a `.sdproj` JSON file (full project state + metadata + version); `importProject` loads one via `<input type=file>`. Old single-sheet saves untouched.
- New toolbar buttons: PDF All, Save, Load.
- Title Block dialog now snapshots into the active sheet and refreshes the sheet browser on commit (so renaming the drawing number updates the sidebar label live).

### V19.1–4 (earlier today)
- **DXF export** (AutoCAD R2013 ASCII, `.dxf`). New `exportSheetToDXF` walks every object + 2D entity and emits a proper layered DXF: LAYERS table with 13 AS 1100 layers (`S-BEAM`, `S-PLATE`, `S-BOLT`, `S-CUT`, `S-HIDDEN`, `S-CL`, `S-DIM`, `S-TEXT`, `S-WELD`, `S-DETAIL`, `S-REVISION`, `S-NOTE`, plus `0`), LTYPE table with CONTINUOUS / HIDDEN / CENTER / DASHDOT. ENTITIES section emits LINE / LWPOLYLINE / CIRCLE / MTEXT with per-layer colour + lineweight (`370` group, 0.01mm units). Coordinate system: sheet-mm with DXF Y-up. Each of the three orthographic views (elevation / sectionA / planB) is placed at its block anchor on the A1 sheet. Isometric view is skipped (raster-only).
  - Helpers: `_dxfBuilder`, `_dxfHeader`, `_dxfTables`, `_dxfLine`, `_dxfCircle`, `_dxfText`, `_dxfPolyline`, `_dxfBlockPlace`, `_dxfEmitUB/SHS/Bolt/Plate`, `_dxfEmit2DEntity`. Every 2D entity type defined in V16–V18 has a DXF case.
  - Button `#btnExportDXF` now calls `exportSheetToDXF` (was a "coming soon" stub since V13).
- **Revision triangle** — numbered equilateral triangle entity (`revisionTriangle`), tied to the rev schedule in the title block. Click in library → click on sheet → prompt for rev number. Drawn on `S-REVISION` in DXF output.
- **Revision cloud** — perimeter of arcs around revised work (`revisionCloud`). Multi-click polygon → Enter or double-click to close. Bump radius auto-scales to segment length so short segments don't collapse to circles.
- **Detail reference callout** — "3/S-400" circle-with-divider bubble (`detailRef`). One click → prompt for detail + sheet. Standard drafter convention.
- **Detail card frame** (V19.4) — heavy border + number bubble + scale annotation (`detailCard`). Two-click rectangle → prompt for number, title, scale. Makes it possible to place multiple labelled details on one A1 today; the full multi-sheet project model (V19.5) will build on this entity.
- Title tag bumped V18 → V19. DRAFTER_MAPPING gains entries for DXF coverage + revision entities.
- **Deferred to V19.5**: full multi-sheet project model with sheet browser + build-step introduction. Scoped for the next session — benefits from you testing V19.1–4 in the browser first.

## V18 — Everything labelled (18 April 2026)
- **Chain and baseline dimensions** (drafter §3.8). `drawDim2D` extended with two new `dimType` values: `chain` (sequence of adjacent horizontal dims sharing a baseline Y) and `baseline` (each dim measured from a datum, successive tiers stagger upward by 10mm sheet-space). Data: `{ dimType:'chain'|'baseline', stops:[u0,u1,…], v, off }`.
- **Parametric member tags** — new `memberTag` entity: leader + auto-resolved label. If `memberId` is set, text pulls from the member's `section`/`boltSize`/`pt` live, so renaming a beam updates every tag pointing at it. Placed via new "Member Tag" library item.
- **Bolt group callouts** (drafter §3.3) — new `boltCallout` entity: `N/M20 8.8/S` format. Select N bolts, click tool, click text location — the tool auto-counts bolts and picks the dominant size.
- **Section marks** — new `sectionMark` entity with auto-assigned letter (A, B, C…). Two clicks set the cut line; arrows on each end point in the direction of sight; `nextSectionMarkLabel()` scans all views and picks the next unused letter.
- **Material tags** — `materialTag` entity mirrors `memberTag` but with manual text (e.g. "PL 12 THK"). Kept as its own type so V19 can route it to a different DXF/PDF layer.
- All four new label types use the `LW.DIM` / `LW.CUT` hierarchy and render through the wobble wrapper when sketch mode is on.

## V17 — Both themes beautiful (18 April 2026)
- **Sketch wobble** (drafter §5.5) — new `sketchOn` flag (toolbar toggle "Sketch wobble" in 2D Utilities). Wraps `rLine` / `rRect` via `_lineW` which subdivides screen-space segments and applies deterministic Perlin-ish jitter (Mulberry32 hash). Same entity always wobbles the same way → drawing doesn't "dance" on pan/zoom. Amplitude capped at 0.4mm sheet-space so legibility stays. Fades to zero near segment endpoints so joints look crisp.
- **Paper grain overlay** — new `sketchGrain` flag. 200×200 noise tile cached once via `_paperGrain()`, tiled across the sheet fill at 0.45 alpha. Classic theme only; suppressed during PDF export.
- **Slotted holes** (drafter §7.7) — new `slot` entity type: `{ u, v, dia, length, angle }`. Default matches AS 1100 M20 standard (22×40). Drawn as a stadium shape (two semicircular caps + parallel edges) with a chain-dash centreline along the long axis. New "Slotted Hole" library item.
- **Drawable centreline refinement** — `centreline` entity now hardcoded to `LW.CL` weight and `DASH.CL` pattern regardless of the entity's raw `lw` field. Predictable look without user having to guess values.
- Both wobble and grain are **off by default** — straight-out-of-the-box output stays AutoCAD-crisp. PDF export always runs wobble/grain off even if the toggles are on, so issue drawings remain razor-sharp.
- Title tag bumped V16 → V18. DRAFTER_MAPPING §5.5, §7.7 flipped to ✅; §3.8 flipped partial → ✅.

## V16 — 3-minute detail (17 April 2026)
- **Parametric connection library (drafter §9.x)**. New left-sidebar section "Connections" with four one-click wizards:
  - **Cap Plate** (§9.2) — `buildCapPlate(column, params)`. Derives plate size from column width + `CAP_BOLT_COL_GAP` + `CAP_AE`; places bolts in an N×M grid straddling the column face; auto-labels `CAP PL <thk> THK`; adds plate-width dimension.
  - **Column Baseplate** (§9.4) — `buildBaseplate(column, params)`. Horizontal plate at column base with holding-down bolt grid extending down for cast-in embedment.
  - **Web Side Plate** (§9.1) — `buildWSP(beam, params)`. Bolt count auto-derived from beam depth and pitch; elevation-facing plate at beam end.
  - **Moment Splice** (§9.3) — `buildSplice(beam, params)`. Pair of end plates with mill gap; bolt rows straddling the web top-to-bottom.
- **`CONN_DEFAULTS` constant** — single source for the drafter's standard dimensions (`CAP_BOLT_COL_GAP=40`, `CAP_AE=35`, `SPLICE_GAP=10`, etc.). Change a default here, every subsequent connection inherits it.
- **Atomic connection undo** — new `'connection'` undo action type. A cap plate creates ~5 objects + 2 entities; a single Ctrl+Z removes the whole connection, not bolt-by-bolt. Redo restores the whole thing.
- **`#connectionDialog`** — dynamic wizard populated per connection kind. Each field pre-filled from `CONN_DEFAULTS`. Fields auto-derived from `_connSpecs` table — adding a new connection is one table entry + one builder function.
- **`placeConnection(result)`** dispatcher commits `{objs, ents}` atomically to the model and pushes one undo entry.
- Title tag bumped V15 → V16. DRAFTER_MAPPING §9.1–9.4 all flipped ⏳ pending → ✅.

## V15 — Output you can issue (17 April 2026)
### V15.4 — Dash-pattern centralisation (drafter §3.7)
- New `const DASH` table. Named entries: `SOLID`, `CL` (centreline chain), `CL_BOLT` (bolt centreline), `SECTION` (heavy chain for cut lines), `HIDDEN` (short dash), `THREAD` (thread overlay), `SNAP` / `UI_CHAIN` / `UI_ALT` / `UI_ROT` (interaction chrome).
- ~45 inline `setLineDash([…])` call sites across members, bolts, dimensions, welds, grips, crosshair, and auto-weld pipeline now reference the table. Change once, cascade everywhere.
- Drafter §3.7 row in DRAFTER_MAPPING flipped ⚠️ scattered → ✅.

### V15.3 — Weld completeness (AS 1101.3)
- `drawWeld2D` rebuilt to the full AS 1101.3 vocabulary. New helper `drawWeldGlyph` covers six weld types on one code path: **fillet, square butt, single-V, double-V, partial-penetration, bevel**.
- Both-sides welds supported via new entity fields `otherType` / `otherSize` — renders a mirrored glyph above the reference line with its own size label.
- New modifier flags on weld entities: `siteWeld` (AS 1101.3 filled pennant at the elbow) and `tail` (forked reference tail with spec text, e.g. "AS 1554.1 SP"). `length` field for intermittent welds (e.g. "6-100").
- New `#weldDialog` HTML dialog + `openWeldDialog(onConfirm)` helper replaces the three-`prompt()` chain. Dialog is sticky across uses — last-used values repopulate.
- Auto-weld popup `#wpType` dropdown expanded from 2 options → 6 (fillet, square, single-V, double-V, partial-pen, bevel). Legacy "butt" value auto-maps to single-V.
- `drawWeldHatch` lineweight moved from hardcoded `0.25` to `LW.MW * pm` (0.50mm) — medium-weight overlay, consistent with the rest of the weld graphics.
- DRAFTER_MAPPING §3.10 (weld triangle + tick + hatch), §7.8 (weld helpers) flipped ❌ → ✅.

### V15.2 — Lineweights per drafter §3.6
- `LW` constant rewritten to the full AS 1100 / drafter §3.6 hierarchy: `CUT:1.20, VIS_HEAVY:0.70, VIS:0.65, MW:0.50, DIM:0.40, HID:0.30, CL:0.30, HATCH:0.18`. Previous values (0.70/0.35/0.18) promoted to their drafter-correct tier.
- Every `LW.* * pm` site across member/bolt/plate/weld renderers now cascades the new hierarchy — no further sweeps required.
- Hardcoded hatch lineweights (`0.12 * ppm()`, `0.15 * pm`) replaced with `LW.HATCH * pm` at `drawCrossHatch` and both `drawThreadAlongU/V` (thread minor-diameter lines).
- `Math.max(0.5, LW.HID/CL * pm)` floors lowered to `0.25` so PDF output at `pm=1` renders HID/CL at the intended 0.30mm rather than being floored to 0.5mm. Screen display is unchanged at typical zooms.

### V15.1 — Vector PDF export
- **Vector PDF export** (behind `V15_VECTOR_PDF` flag, default on). New `exportSheetToPDFVector` routes the render pipeline through `createPdfCanvasShim` — a canvas-2D-compatible shim that emits jsPDF line/rect/circle/lines/text primitives. Output is razor-sharp at any zoom; AS 1100 lineweights map 1:1 to `setLineWidth(mm)`.
- Existing raster exporter retained as `exportSheetToPDFRaster` and as fallback if the vector path throws.
- `exportSheetToPDF` dispatcher now forces classic theme during export (white sheet, black ink) — benefits both raster and vector paths.
- `ppm()` returns 1 when `pdfExportMode` is true so `LW.* * ppm()` evaluates in sheet-mm directly.
- Title tag bumped V12 → V15.

## Previous: V14 (12 April 2026)
- AS 1100 realistic bolt renderer: chamfered hex + sawtooth threads (drafter §7.3–7.5).

## V13 (2 April 2026)
- Baseline version migrated from manual version folders (V1–V13).
- Folder structure reorganised: dev/ for working, parent folder for live/GitHub.

## Previous versions (archived in git history)
- V1–V12: Legacy versions previously stored as separate HTML files, now tracked through git commits.
