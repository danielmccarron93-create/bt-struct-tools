# BT Inspection Program — Cowork Playbook

This folder is the source-of-truth for generating structural inspection programs at Bligh Tanner. When the user drops a project folder containing a combined PDF of the structural drawing set, follow the procedure below to produce a clean, on-brand, single-screen `inspection-program.html` that the engineer reviews and approves.

This file replaces ad-hoc per-project instructions. It is grounded in three reference projects:

- **52 Second Avenue, Maroochydore** — residential, hybrid concrete + CLT, 7 storeys, CFA pile foundation, 66 sheets, 31 inspections. Older BT A1 template.
- **Emmanuel College Senior School Redevelopment, Carrara** — commercial school, 3 storeys + steel roof + 4 sub-builds (lecture theatre, amphitheatre, fin frames, elevated link), bored pier foundation, 104 sheets, 33 inspections. Older BT A1 template.
- **MBC Creativity & Arts Centre, Manly West** — school arts centre, 2 storeys + 3-tier steel roof + atrium feature stair, hybrid bored pier + pad/strip footing on duricrust, PT suspended slab at L1, 62 sheets, 28 inspections, **TENDER ISSUE**. Uses the **newer BT A1 template variant** (see `skills/reading-bt-drawings/SKILL.md` Section 10).

All used a BT A1 title block (different variants); all produced clean output. Every reference value, code snippet, and design decision in this file came from those builds.

---

## 1 · The Deliverable

A single self-contained `inspection-program.html` in the project folder, ready to open from OneDrive in any browser. It contains:

- Project metadata header (job#, status, site, sheet count)
- **Sticky cover-sheet isometric on the left** with SVG hotspot polygons over the structural elements
- **Scrollable inspection program on the right** grouped into phases, with compact cards that expand on click to show drawings, rationale, and checklist
- **Two-way click highlighting** — click a card → polygons light orange on the isometric; click a polygon → corresponding card scrolls into view
- Level filter chips, stats strip, and `mailto:` approve/request-changes buttons in the footer
- Anthropic-styled visual language (Poppins headings, Lora body, cream background, orange/blue/green accents)

The HTML file is the **review artefact**. Once the engineer approves, the same data can be exported as `project-map.json` + bundled into `project.btproject` for the iPhone consumer app — but only after explicit approval.

---

## 2 · Trigger Phrases

Run the procedure when the user says any of:

- "Set up `<JobName>` for inspection"
- "Process drawings in `<JobName>`"
- "Build the inspection program for `<JobName>`"
- "Generate the program HTML for `<JobName>`"
- A new project folder appears with a single PDF and the user asks Claude to look at it

If the project name is ambiguous, list candidate folders and ask. Don't guess.

---

## 3 · Folder Convention & Architecture

This repo is the BT Inspect application backbone. It contains both the iPhone PWA itself **and** the Cowork tooling that feeds it. Layout:

```
Structural-Inspections/                ← repo root (this folder)
├── CLAUDE.md                          ← this playbook
├── README.md                          ← repo readme (existing)
├── index.html                         ← BT Inspect PWA entry
├── manifest.json, service-worker.js   ← PWA shell
├── css/, js/, assets/                 ← PWA source
├── reference/
│   ├── inspection-types.json          ← canonical catalogue, seeded into PWA's IndexedDB
│   └── project-map.schema.json        ← JSON schema validated by .btproject importer
├── skills/                            ← reusable Cowork knowledge artefacts
│   ├── README.md
│   └── <skill-name>/SKILL.md          ← e.g. reading-bt-drawings/, hotspot-authoring/
├── Structural Drawings/               ← project work folders
│   ├── _template/                     ← empty starter
│   └── <JobName>/                     ← e.g. "Mount Alvernia"
│       ├── BT <JobName>.pdf           ← combined structural set (input)
│       ├── isometric.png              ← extracted from cover (output)
│       ├── inspection-program.html    ← the review artefact (output)
│       ├── project-map.json           ← machine-readable program (post-approval)
│       ├── project.btproject          ← zipped bundle for iPhone import (post-approval)
│       └── .btinspect_scratch/        ← intermediate analysis (safe to delete)
└── test/
```

Naming rule: project folder = job name (or `<JobNumber> <JobName>` if there's a clean BT job number, e.g. `2023.0957 Emmanuel`). Use the user's preferred form — don't enforce.

The PDF is named freely (e.g. `BT Emmanuel.pdf`, `Mount Alvernia Structural.pdf`). The procedure should pick up `*.pdf` and use the first one that has the BT title block.

### How the pieces talk

```
┌────────────────────────────────────────────────────────────────────────┐
│  STRUCTURAL DRAWINGS (PDFs in 'Structural Drawings/<JobName>/')        │
│         │                                                              │
│         ▼                                                              │
│  COWORK (this chat — Claude reads CLAUDE.md + skills/)                 │
│  ─ Follows the procedure in Section 4                                  │
│  ─ Pulls reusable knowledge from skills/                               │
│  ─ Validates against reference/project-map.schema.json                 │
│  ─ Constrained to keys in reference/inspection-types.json              │
│         │                                                              │
│         ▼                                                              │
│  inspection-program.html  ────►  ENGINEER REVIEW (in browser)          │
│         │                                          │                   │
│         │ on approval                              │ requests changes  │
│         ▼                                          │                   │
│  project-map.json + project.btproject  ◄───────────┘                   │
│         │                                                              │
│         ▼                                                              │
│  BT INSPECT PWA (this repo's index.html / js/)                         │
│  ─ js/components/import-project.js consumes .btproject                 │
│  ─ js/lib/btproject.js reads + validates the bundle                    │
│  ─ js/db.js imports per project-map schema                             │
│  ─ Inspection types pre-seeded from reference/inspection-types.json    │
│         │                                                              │
│         ▼                                                              │
│  iPhone — site inspections, defect logging, report generation          │
└────────────────────────────────────────────────────────────────────────┘
```

Two integrity rules that hold the whole system together:

1. **`reference/inspection-types.json` is the contract.** Cowork must not invent new inspection-type keys; the PWA's IndexedDB is seeded from this file. New keys go through a deliberate catalogue update — see Section 6 (Catalogue Gaps).
2. **`reference/project-map.schema.json` is the wire format.** Every `project-map.json` Cowork writes must validate. Schema changes require a coordinated bump in `js/db.js` (the PWA reads `schemaVersion` and migrates).

---

## 4 · The Procedure

Six analysis steps + three render steps. Work through them in order. Do not pause between steps for the user unless something is genuinely ambiguous — collect work and produce the deliverable, then ask for review.

### 4.1 — PDF inventory

Open the PDF with PyMuPDF (`fitz`). Record:

- File size (warn if >100 MB)
- Page count
- Page 1 dimensions and `/Rotate` value
- Whether page 1 has an extractable text layer (`page.get_text().strip()`). If empty, the PDF is scanned — flag for vision fallback.

```python
import fitz
doc = fitz.open(PDF)
p = doc[0]
print(f"pages={doc.page_count} size={p.rect.width}x{p.rect.height} rot={p.rotation}")
print(f"page1 text chars: {len(p.get_text())}")
```

Expected for BT A1 set: 2384 × 1684 pt, `/Rotate=90`, page 1 text 3,500–6,000 chars.

### 4.2 — Per-page title-block extraction

The BT A1 template has a fixed positional layout. Use **anchor-based matching on (x0, y0)** — *not* center coordinates, because rotated text bbox extends downward and the center is unreliable.

Cell windows (xmin, xmax, ymin, ymax, font-size-min, font-size-max), all in unrotated PDF coords:

| Field | x range | y range | size |
|---|---|---|---|
| `sheetNumber` | 1580–1625 | 115–150 | 20–30 |
| `revision` | 1580–1625 | 50–80 | 20–30 |
| `jobNumber` | 1580–1625 | 245–280 | 20–30 |
| `drawingTitle` | 1525–1580 | 55–80 | 14–18 |
| `projectName` | 1475–1510 | 55–78 | 15–19 |
| `siteAddress` | 1505–1535 | 55–75 | 11–14 |
| `drawnBy` | 1470–1497 | 620–640 | 11–13 |
| `designBy` | 1500–1527 | 620–640 | 11–13 |
| `checkedBy` | 1530–1555 | 620–640 | 11–13 |
| `client` | 1480–1530 | 1390–1410 | 14–18 |
| `issueLine1` | 1470–1545 | 800–830 | 25–30 |
| `issueLine2` | 1505–1545 | 860–880 | 25–30 |
| `revDate` | 1485–1505 | 2050–2065 | 9–12 |
| `scale` (multi) | 1575–1625 | 905–925 | 6–8 |

Helper:

```python
def best_in_cell(spans, cell):
    xmin, xmax, ymin, ymax, smin, smax = cell
    hits = [(b, s) for b, s, _, t in spans
            if xmin <= b[0] <= xmax and ymin <= b[1] <= ymax and smin <= s <= smax]
    hits.sort(key=lambda h: (-h[1], h[0][1]))
    return hits[0] if hits else None
```

Cross-check against the Structural Drawing List on page 1 (cover). Mismatches go in `warnings`.

If a sheet doesn't extract cleanly: leave `sheetNumber` null, use `pageNumber` as the stable identifier, and add a warning. Don't make up sheet numbers.

### 4.3 — Project-level metadata

From page 1 cover (consensus across all pages for fields that should be invariant):

- `name`, `siteAddress`, `jobNumber`, `client`, `issueStatus` (CONSTRUCTION ISSUE / TENDER ISSUE / etc.)
- `engineerOfRecord` = "Bligh Tanner" for BT-authored sets; otherwise the external engineer named on the title block
- `discipline` = `structural` | `civil` | `combined`

The `client` cell in the BT A1 title block sometimes appears empty in the text layer (was the case for 52 Second Av and Emmanuel) — that's fine, leave it null and surface a warning. The `architect` cell often empty too — ask the user only if they care.

### 4.4 — General Notes parsing

Pages 2 and 3 of BT sets contain the General Notes, Design Criteria, and Concrete Cover/Bearing schedules. Extract the full text (`page.get_text()`) and pull out:

- **`designCodes`** — AS codes listed under "ALL MATERIALS AND WORKMANSHIP SHALL BE IN ACCORDANCE WITH...". Typical: AS 3600, AS 4100, AS 1720, AS 2159, AS 3700, AS 3610, AS 2269, NCC. Plus secondary codes scattered through the notes (AS 5216 anchors, AS 1554 welds, AS/NZS 5131 steel fab, AS/NZS 4680 galv, AS 1252 bolts).
- **`exposureClass`** — sometimes stated explicitly, sometimes inferred from cover values. Don't invent.
- **`concreteCover`** — the cover schedule table. Extract by **positional analysis**, not by reading text top-to-bottom. The cover table has element columns (BORED PIERS / FOOTINGS / SLAB ON GROUND / COLUMNS / WALLS / SUSPENDED SLAB / etc.) at fixed x positions, and rows BOTTOM/TOP/SIDES at distinct y positions. Match each value to its (element, face) cell. Use the same `(x0, y0)` anchor approach as the title block.
- **`concreteStrengths`** — N32, N40, N50 etc. per element type, in MPa. Found alongside the cover table.
- **`bearingCapacity`** — F1 footing notes. Foundation types vary:
  - **Pad/strip on soil** (rare in commercial): bearing values per element.
  - **CFA piles**: end-bearing values per pile family (e.g. P1=600 kPa hard clay, P2=1000 kPa medium dense sand). 52 Second Av pattern.
  - **Bored piers**: end-bearing + shaft adhesion (e.g. 1000 kPa end + 40 kPa shaft). Emmanuel pattern.
- **`windRegion`, `windVelocity` (ult + serviceability), `terrainCategory`, `importanceLevel`** — DC2 block. IL ≈ 2 for residential, 3 for school/commercial.
- **`earthquake`** — DC3 block. Z, IL, kp, sub-soil class, design category, μ, Sp.
- **`geotechReport`** — consultant name, reference number, date. Always under F1.
- **`certifiedByOthers`** — table near the top of page 2 listing items the contractor's other RPEQs certify (typically: precast lifting/inserts, steel temp props/bracing, steel stud framing, roof safety systems). These items get **excluded** from the BT inspection plan.
- **`specialNotes`** — project-specific constraints worth surfacing in checklists (block wall joint spacings, anchor load-test percentages, CLT inspection holds, curing requirements, etc.).

If a field isn't on the notes pages, leave it null. **Don't hallucinate AS clause numbers.** If unsure, write `(ref. project documents)` instead of a specific clause.

### 4.5 — Drawing classification

For each sheet, assign tags from these facets:

- **`kind:`** (one of) `cover · notes · plan · section · elevation · detail · typical-detail · schedule · loading-plan · schematic · perspective · marking-plan · na`
- **`level:`** (one of) `basement · ground · l1 · l2 · l3 · l4 · l5 · roof · lower-roof · multi · na`. The schema enum stops at L5 — for taller buildings (52 Second Av went to L7), use `multi` and add a warning.
- **`element:`** (one or more of) `pad-footing · strip-footing · bored-pier · pile · raft-slab · slab-on-ground · slab-suspended · concrete-wall · block-wall · masonry · concrete-column · steel-column · steel-beam · steel-frame · timber-framing · mass-timber · stair · lift-pit · retaining-wall · truss · bracing · connection · baseplate-anchor · waterproofing · post-tension · roof-framing · roof-purlin · head-framing · fin-frame`
- **`zone:`** (optional) `north · south · elevated-link · amphitheatre · lecture-theatre · external-works`, etc.

**Classification heuristics that need explicit rules** (learned from 52 Second Av — avoid re-discovering):

- "GENERAL ARRANGEMENT PLAN" or "REINFORCEMENT PLAN" without "SLAB" in the title is *still a slab plan*. Add the slab tag based on level.
- "STAIR PLANS / SECTIONS" → tag `element:stair` AND inspection-type `slab-prepour-suspended` (no dedicated stair key in the catalogue).
- "TIMBER CONNECTION DETAILS" should NOT inherit `steel-connection` from the generic `connection` tag. If `mass-timber` is present, suppress `steel-connection`.
- "DRYING ROOM" or other precast slabs: add a `scopeNote` flagging that precast lifting/inserts/propping is certified by others, and the BT scope is the in-situ portion only.
- "LECTURE THEATRE" / "AMPHITHEATRE" / "FIN FRAME" / "ELEVATED LINK" / similar named features → these are **sub-builds**, treat as their own phase.

Then derive `applicableInspectionTypes[]` per drawing — the inspection-type `key` values from `reference/inspection-types.json` that the drawing informs.

### 4.6 — Inspection plan synthesis

Build an ordered list of inspections in expected build sequence. Reason like a senior engineer: what does the builder actually do, in what order, when is BT inspection on the critical path?

**Canonical phase structure:**

1. **Substructure** — founding (geotech), pad/pile cap pre-pour, basement walls (if any), waterproofing pre-cover (if any).
2. **Ground Floor** — slab on ground, concrete walls, block walls (vert reo + core fill), concrete columns, stairs G→L1.
3. **Per Level (L1, L2, ...)** — repeat the same template per storey: suspended slab pre-pour reo, post-tension strand placement (if PT), concrete walls, block walls (vert reo + core fill), concrete columns, stairs to next level.
4. **Steel transition (if applicable)** — at the level where steel columns pick up to support roof: baseplate / anchor bolt inspection (S212-style hold-down bolt plan).
5. **Roof** — lower roof framing + purlins, main roof framing + connections + purlins, bracing.
6. **Sub-Builds** — separate cards per named feature: lecture theatre, amphitheatre, fin frames, elevated link, plant rooms, etc. Sequenced inside their own work front (often parallel to main).
7. **Specialty** — balcony slabs (per-level repeat), drying rooms, transfer slabs, etc.

**Foundation-specific substructure cards:**

- **Pad/strip footings on soil** (rare): one `subgrade` (geotech hold) + `pad-footing-prepour`.
- **CFA piles** (e.g. 52 Second Av): one `subgrade` proxy card titled "CFA Pile Installation — Founding Depth, Torque & Socket" (geotech-witnessed, hold point) + `pad-footing-prepour` for pile caps.
- **Bored piers** (e.g. Emmanuel): one `subgrade` proxy card titled "Bored Pier Installation — Founding Depth & Shaft" (geotech RPEQ hold) + `pad-footing-prepour` for pile caps.

**Hold points** — mark these `holdPoint: true, defaultSeverity: hold-point`:

- Founding-level (geotech-witnessed)
- Critical pours: transfer slabs, CLT-supporting slabs (CLT22 pattern)
- Pre-cover items where remediation is destructive: CLT shear-wall anchorages, basement waterproofing, post-tension strand placement
- Anchor bolt inspections at frame transitions

**Per-card payload** — every entry must have:

```yaml
sequence: 18                                # 1-based, in build order
type: "slab-prepour-suspended"              # key from inspection-types.json
title: "Suspended Slab — Level 1 — Head Framing & Pre-Pour Reinforcement"
phase: "Level 1"                            # for layout grouping
level: "l1"                                 # for filter chips
hotspots: ["l1-slab"]                       # IDs from the hotspot definitions
holdPoint: false
defaultSeverity: "observation"
stage: "pre-pour"                           # pre-reinforcement | pre-pour | post-pour | pre-cover | post-install
drawings:
  - sheetNumber: "S100"
    reason: "L1 General Arrangement — South"
  - sheetNumber: "S101"
    reason: "L1 General Arrangement — North"
  - …                                       # include South + North zones, sections, typical details, head framing
rationale: |
  L1 suspended slab — pre-pour reo. Head framing (formwork + props) check is integral with this
  inspection. Includes stair landings cast monolithic with the slab.
expectedChecklist:
  - "Bottom/top/sides cover 30 mm; concrete N40 (suspended slab)"
  - "Bottom and top reinforcement match GA sheets S100/S101"
  - "Trimmer bars at penetrations >200 sq, top & bottom"
  - "Formwork properly propped per AS 3610"
  - "Bar chairs @ 600/800 ctrs per note C9"
  - "All reo securely tied (note C18)"
  - "Pour temperature 5–35 °C (note C19)"
scopeNote: "..."                            # optional — for proxy keys or excluded items
```

**Conservative bias.** Better to include an inspection the engineer unticks than miss one they'd have done. But use `certifiedByOthers` to **EXCLUDE** items — if steel stud framing is manufacturer-certified per the General Notes, don't schedule BT to inspect it.

**Multi-zone projects** (S010/S011 South/North split): one inspection per level with both zone drawings referenced. Don't split into "L1 Slab — South" + "L1 Slab — North" — the engineer walks both zones in one visit.

**Repeats are real.** A single L4-6 card represents three field visits. Call this out in the rationale ("treat as 3 separate field visits per storey"). The engineer can duplicate cards in-app if needed.

### 4.7 — Cover-sheet isometric extraction

The cover sheet (page 1) for BT A1 sets always contains a clean structural-only 3D cutaway. Extract it as `isometric.png` in the project folder.

```python
import fitz
from PIL import Image
doc = fitz.open(PDF)
p = doc[0]
mat = fitz.Matrix(2.5, 2.5)              # 2.5x DPI for clean output
pix = p.get_pixmap(matrix=mat)
pix.save("page1_2.5x.png")
img = Image.open("page1_2.5x.png")
# Crop the isometric region — top portion of the page above the drawing list
# Eyeball the y-band from a quick render at 1.5x, then translate to 2.5x coords
crop = img.crop((100, 130, img.width - 60, 1980))
maxw = 2400                               # web-friendly
ratio = maxw / crop.width
crop.resize((maxw, int(crop.height * ratio)), Image.LANCZOS)\
    .save(f"{project_folder}/isometric.png", optimize=True)
```

If the cover doesn't have a clean isometric (e.g. client-issued sets, residential sketches): fall back to (a) ask the user to drop in `isometric.png` themselves, or (b) generate a programmatic schematic (level-band rectangles with phase labels).

For Emmanuel-style projects with an embedded image:

- `page.get_images(full=True)` lists embedded images
- The biggest 4:5-aspect images are usually the isometric cutaway tiles
- For a single-image isometric, the rendered-page-region approach above is simpler

### 4.8 — Hotspot polygon authoring

This is the slow but high-value step. The user has been clear: **highlight the actual elements**, not just level bands. When they click "Pad Footings" they should see actual footings light up — not the entire base of the building.

**Approach:**

1. Render the isometric to four quadrants for visual inspection (`crop.crop((W*i/4, 0, W*(i+1)/4, H))` for i in 0–3).
2. For each visible structural element type, identify positions in the isometric.
3. Define polygons in **percentage coordinates** (0–100 across width, 0–100 across height) so they're resolution-independent. SVG `viewBox="0 0 100 100"` + `preserveAspectRatio="none"` lets the polygons scale with the displayed image.
4. Group multiple instances under one hotspot ID (e.g. all visible pad footings under `footings-main-building`).

**Hotspot taxonomy that worked for Emmanuel:**

| Hotspot ID | Type | Notes |
|---|---|---|
| `footings-elevated-link` | Multi-poly | 2–4 small rectangles at column bases of the link |
| `footings-main-building` | Multi-poly | 8–12 small rectangles along the building base |
| `ground-slab-on-ground` | Strip | Single horizontal band just above the footings |
| `ground-walls` | Multi-poly | 4–6 small vertical rectangles per visible wall section |
| `ground-columns` | Multi-poly | 6–8 narrow vertical rectangles |
| `l1-slab`, `l2-slab` | Strip | Horizontal band at each level |
| `l1-walls`, `l2-walls` | Multi-poly | Same pattern as ground, lifted |
| `l1-columns`, `l2-columns` | Multi-poly | Same |
| `l2-steel-baseplates` | Multi-poly | Small rectangles at top of L2 slab where steel cols pick up |
| `lower-roof`, `main-roof-framing`, `main-roof-purlins` | Strip | Horizontal bands at roof level |
| `stairs-all` | Multi-poly | Bounding rectangles around visible stair zigzag patterns |
| `lecture-theatre`, `amphitheatre`, `elevated-link` | Volume | Bounding rectangles around the distinct sub-build volumes |
| `fin-frames` | Multi-poly | Narrow vertical rectangles at each visible fin |

Polygon format in JS:

```js
HOTSPOTS = {
  "ground-walls": {
    label: "Concrete walls — Ground level",
    polygons: [
      "30,55 33,55 33,77 30,77",  // each polygon = "x1,y1 x2,y2 x3,y3 x4,y4"
      "44,55 47,55 47,77 44,77",
      ...
    ]
  },
  ...
}
```

Each card's hotspots reference one or more IDs:

```js
INSPECTION_HOTSPOTS = { 5: ["ground-walls"], 18: ["l3-slab"], ... }
```

**Iteration is expected.** The first pass is approximate — author from visual inspection at the resolution available, then refine based on the user's feedback after they see it. Plan for a v2 pass.

### 4.9 — HTML rendering

The HTML is generated by a Python build script that templates everything together. **Use a non-f-string template with `__TOKEN__` placeholders** + `.replace()` substitution — f-strings break on the doubled-brace requirement when the template contains CSS and JS.

**Required structure:**

- `<header class="top">` — full-width strip with project name + sub-text
- `<main>` — 2-column grid: left = sticky `<aside class="iso">`, right = scrollable `<section class="program">`
- `<footer class="bottom">` — full-width legend + mailto buttons

**Required interactions:**

- Click card → toggle `.active` class on card AND on linked polygons (orange highlight)
- Click polygon → find the first card whose hotspots include it, scroll into view, activate
- Hover polygon → polygons in the same group get `.hover` class (lighter highlight)
- Filter chip click → cards not matching the level get `.dim` class (32% opacity)

**Required visual style** (Anthropic brand):

- `--dark: #141413`, `--light: #faf9f5`, `--mid-gray: #b0aea5`, `--light-gray: #e8e6dc`
- `--orange: #d97757` (primary accent — holds, active states)
- `--blue: #6a9bcc` (secondary accent — info badges)
- `--green: #788c5d` (tertiary accent — advisory items)
- Headings in Poppins, body in Lora, code in JetBrains Mono (or system fallback)
- 1px subtle rules in `rgba(20,20,19,0.08)`
- Cards: 3px left border, color-coded by hold/observation, soft hover

**Required portability:**

- Embed the isometric as base64 (`data:image/png;base64,...`) so the HTML file works standalone — no asset dependencies.
- All CSS and JS inline. No external scripts beyond Google Fonts (which gracefully fall back to Arial/Georgia).

**Stats strip (top-left of isometric panel):**

- Total inspections
- Hold points
- Drawings referenced
- Sheets total

**Filter chips:**

- All / Substructure + Ground / L1 / L2 / [more levels as needed] / Roof / Sub-Builds

**Mailto buttons (footer):**

- "Request changes" — pre-fills subject and a "Card numbers needing changes:" body
- "Approve program" (primary, dark button) — pre-fills approval text triggering the next step (project-map.json + btproject)

### Summary report after rendering

Emit a concise console summary:

```
✓ Project map built: <JobName>
  Site:         <site address>
  Issue:        <issue status>
  Drawings:     <N> sheets
  Inspections:  <count>, <holds> hold points
  Hotspots:     <count> regions, <polygon-count> polygons
  Output:       inspection-program.html (<size> KB)
  Next step:    Open the HTML in your browser to review and approve.
```

---

## 5 · Reference: Standard Inspection Checklists

These are the BT-standard checklist items by inspection type, drawn from the 52 Second Av and Emmanuel General Notes (which are 95% identical across BT projects). Customize per-project for cover values and concrete grades extracted in Step 4.4.

### Pad footings / pile caps pre-pour

- Cover ≥ <bottom/sides per project schedule> mm
- Pile reo extends ≥ 75 mm into footing (note P7)
- Lap lengths at corners and T-junctions per AS 3600
- Trimmer bars at penetrations >200 sq (note C6)
- Concrete grade <project N-grade>
- Foundation excavation maintained firm & dry (note F2)
- Bar chairs and ties — reo securely tied prior to pour (note C18)

### Slab on ground pre-pour

- Bottom/top/sides cover per project schedule
- Concrete grade per project schedule
- 50 mm bedding sand + DPM in place (note C16)
- Bottom + top reinforcement matches GA layout
- Mesh laps: 2 outer-most cross bars overlap (note C7)
- Trimmer bars at re-entrant corners and penetrations
- Bar chairs @ 600/800 ctrs per mesh weight (note C9)

### Suspended slab pre-pour (incl. head framing)

- Bottom/top/sides cover per suspended slab schedule
- Bottom and top reo match GA + extra reo plans (if present)
- Trimmer bars at penetrations >200 sq, top & bottom
- Formwork properly propped per AS 3610
- All reo securely tied (note C18)
- Construction joints at engineer-approved locations
- Pour temperature 5–35 °C (note C19)
- Cast-in plates / starter bars for next-level walls or steel cols

### Concrete walls pre-pour

- Sides/top cover per walls schedule
- Vertical and horizontal reinforcement per typ detail
- Lap lengths per AS 3600
- Starter bars from below securely tied
- Formwork plumb, clean, properly propped

### Concrete columns pre-pour

- Sides cover per columns schedule
- Vertical bar count, size, lap length per typ detail
- Tie/fitment spacing — closer at top & bottom of column
- 135° hook ends on ties
- Starter bar projection from slab below
- Formwork plumb, clean, propped

### Block walls — vertical reinforcement

- Vertical reo size, spacing, position per typ detail
- Starter bars securely tied prior to pour (note CM10)
- Lap lengths to AS 3700
- Mortar mix per project spec (typically M3 1:1:6 general, M4 retaining)
- Vertical control joints at ≤ 8 m, 5 m max from corners, not within 1.2 m of corners
- Clean-out blocks at all cores to be filled (note CM3)

### Block walls — core fill pre-pour

- f'c per project blockwork schedule (typically 20 or 40 MPa)
- Max slump 230 mm, max aggregate 10 mm
- Lift height ≤ 2400 mm
- All cores swept clean via clean-out blocks
- 10% of chemical anchors into core-filled blockwork load-tested to 1.5 × SWL (note CM15)

### Steel baseplate / anchor bolts

- Anchor bolt position matches plan (set-out tolerance ≤ ±5 mm)
- Bolt projection per detail
- Bolts grade 4.6 (foundation) / 8.8 (structural) per spec
- Threads clean, two threads minimum past nut after tightening
- Templates removed; bolts plumb
- Non-shrink grout pad 30 mm @ ≥ 40 MPa for column erection

### Steel frame erection

- Member sizes match framing plan
- Bolts grade 8.8/S, M16 (≤250 mm sections) / M20 (≥250 mm) per S6
- Welds 6 mm SP fillet UNO; AS 1554 procedures
- Hot-dip galvanised exterior; touched up with WATTYL Galvit
- Members in contact with concrete passivated
- ACRS certification on file (note S15)
- FC1 / CC2 per AS/NZS 5131 (or per project category)

### Steel connection inspection

- All bolts tightened (full bearing) per AS 4100 8.8/S
- Weld lengths and sizes match drawings; visual quality (SP)
- Site-weld locations only as specified (W7); independent NATA testing per W-table
- Connection plates ≥ 10 mm thick UNO
- Galvanised cleats: damage touched up

### Roof purlins

- Purlin sections per plan
- Laps ≥ 15% span or 900 mm (PU2)
- Bolts per series (PU3 schedule)
- Bridging per PU6
- Trimming purlins to penetrations per PU5

### Mass timber install (CLT-specific)

- Panel positions match layout plan; joints per typ details
- Proprietary connectors and fasteners as scheduled
- Edge distances per typ details
- Penetration rules: <100 mm panel edge clearance ≥ 400 mm
- No panel joints in lintels or within 1 m of openings
- Manufacturer quality cert provided
- Moisture content <18%; end-grain sealer applied
- Site storage protected from weather

### Bored pier / CFA pile install

- Pile location ≤ 75 mm of designated position (note P3)
- Founding depth recorded per pile, forwarded to BT within 3 days (note P4)
- Bearing/adhesion meets project values (geotech RPEQ confirms on site)
- Pile extension ≥ 75 mm into pile cap (note P7)
- Piling Contractor RPEQ certificate confirming design loads achieved (note P8)
- No pile within 1000 mm of existing stormwater (note P5)

---

## 6 · Catalogue Gaps (current)

The `inspection-types.json` catalogue has known gaps. Workarounds:

| Real inspection | Catalogue proxy | Why |
|---|---|---|
| CFA pile install | `subgrade` | No `pile-cfa-install` key |
| Bored pier install | `subgrade` | No `pile-bored-install` key |
| Concrete wall pre-pour | `retaining-wall-prepour` | No `concrete-wall-prepour` key |
| Stair flight pre-pour | `slab-prepour-suspended` | No dedicated `stair` key |
| Head framing inspection | Folded into `slab-prepour-suspended` checklist | No separate `head-framing` key |

Always set `scopeNote` when using a proxy so the engineer knows the intent.

When proposing a new key, also flag it in `warnings[]` of the project map. A future v2 catalogue should add: `pile-cfa-install`, `pile-bored-install`, `concrete-wall-prepour`, `stair-prepour`, `head-framing`, `transfer-slab-witness`.

The schema's `level:` enum stops at `l5`. For projects with L6/L7+, use `level:multi` and add a warning. (52 Second Av went to L7.)

---

## 7 · Quality Bar

Before declaring the deliverable ready:

- [ ] Every page in the PDF has either a clean sheet number OR a warning entry
- [ ] Cross-checked against the cover-sheet Structural Drawing List — no missing/extra sheets
- [ ] General Notes parsed; cover schedule extracted by position; bearing values match foundation type
- [ ] Every inspection card has: type from catalogue, level, drawings (with reasons), rationale, checklist, hotspot IDs
- [ ] Every checklist item references either a project-specific value (cover, grade) or a verified note (e.g. "note CM15") — not invented AS clauses
- [ ] Excluded items (precast cert by others, steel stud framing, etc.) confirmed against the General Notes table
- [ ] Hotspots cover every visible major element type — pad footings AND walls AND columns AND slabs AND roofs AND sub-builds (not just level bands)
- [ ] HTML opens cleanly in a browser (no broken JS, no 404 on image, polygons render in correct positions)
- [ ] Click card → polygon highlights work both ways
- [ ] Filter chips correctly dim/show cards
- [ ] Footer mailto buttons pre-fill correctly

**Never fabricate.** Null is better than a guess. Warnings describe what you're unsure of. Cite AS clauses only where verified in the General Notes.

---

## 8 · Edge Cases

- **Scanned PDFs** (no text layer): render each page at 1024–1500 px and use vision to extract title block + General Notes. Flag the project as "vision-extracted, verify first" — much higher chance of error.
- **Residential sketch sets** (small, often A3 portrait, no job number): the BT A3 template has the title block in the bottom-right corner with `/Rotate=0`, sheet-number cell much lower on the page. Use a different cell-window dictionary. Output 3–8 inspections typically; scale the program down.
- **Client-issued sets** (Queensland Rail, Department of Transport): completely different templates. BT is reviewing someone else's work. Add `scopeNote: "BT engineering review on behalf of the client — not BT-authored"` to every plan entry. No RPEQ sign-off in the plan — that's the original engineer's responsibility.
- **Combined structural + civil** sets: use `discipline: combined`. Civil inspections (drainage, pavement) typically aren't BT primary — flag them differently in the plan.
- **Architectural / services drawings bundled in**: skip any drawing whose title clearly identifies it as architectural / services / landscape — note the skip in warnings.
- **Missing General Notes**: produce the plan from per-drawing descriptions alone. `generalNotes` becomes an empty object. Engineer fills cover/bearing values when they arrive. All checklists fall back to generic AS-clause-only references.
- **Multiple PDFs in one folder** (revision uploads): the first pass treats them as separate. Diff mode handles updates — see Section 10.
- **Foundation type ambiguous** (e.g. drawing list has both "FOOTING LAYOUT PLAN" and "PILING NOTES" referenced): assume the more conservative interpretation and include both `subgrade` (geotech hold) and `pad-footing-prepour` cards.

---

## 9 · Approval Workflow

The HTML is the **review artefact**. Don't auto-write `project-map.json` or `project.btproject` until the engineer has signed off.

When the engineer hits the "Approve program" mailto button (or just confirms verbally in chat):

1. **Validate the assembled project map** against `reference/project-map.schema.json`. Fix any validation errors before proceeding.
2. **Write `project-map.json`** in the project folder, pretty-printed (2-space indent), schema-compliant.
3. **Bundle `project.btproject`** — a ZIP archive (consumed by `js/lib/btproject.js → readBundle()` in the PWA) containing:
   - `project-map.json` at the root
   - `Drawings/<filename>.pdf` for every source PDF — the folder name MUST be `Drawings/` (not `pdfs/`); the bundle reader at line 86 of btproject.js searches that exact folder
   - `manifest.txt` — human-readable summary (project name, date generated, drawing count, inspection count)
4. **Print summary** with file sizes and the iPhone import instruction (open BT Inspect → Projects → Import → select the .btproject).

When the engineer hits "Request changes":

- Read the email body for card numbers and reasons
- For each flagged card: regenerate just that card with the requested change
- Re-render the HTML
- Loop until approved

---

## 10 · Revisions

When a new PDF appears in an existing project folder (filename match or content hash differs):

1. Re-extract title-block metadata for the new PDF.
2. Diff against the existing `project-map.json`'s `drawings[]`:
   - **`added`** — sheets in new not in old
   - **`changed`** — sheet number matches but revision advanced
   - **`removed`** — sheets in old not in new
3. Write `revision-<timestamp>.json` with the diff (don't overwrite project-map.json).
4. Update the inspection program HTML — for each changed drawing, mark its referencing inspection cards with a "revision changed" badge.
5. Ask the engineer: accept (merge into project-map.json + bump schema version) / discard / regenerate plan.

---

## 11 · Tools

- **Python 3 with `fitz`** (PyMuPDF) — `pip install PyMuPDF --break-system-packages`. Used for all PDF reading, rendering, image extraction.
- **PIL (Pillow)** — for cropping/resizing the isometric image.
- **Python `zipfile`** — for the .btproject bundle.
- **`json` / `jsonschema`** — for validating the project-map against schema.
- **Read/Write/Bash/Grep/Glob** — standard tools.
- **Web search** — for verifying current AS clause numbers if uncertain. Don't rely on memory; AS standards do get amended.
- **Anthropic skills** — `as2870-slab-design` for residential slab inspections; `pdf-viewer` family for interactive PDF review with the engineer.

---

## 12 · Reference Files in This Repo

**Cowork-side (read these to do the job):**
- `CLAUDE.md` — this playbook (you are here)
- `reference/inspection-types.json` — canonical inspection-type catalogue. Don't invent new keys.
- `reference/project-map.schema.json` — JSON Schema for `project-map.json` output.
- `skills/README.md` — index of reusable Cowork knowledge artefacts (focused per-domain SKILL.md files).
- `Structural Drawings/_template/` — empty starter folder for new projects.

**PWA-side (the consumer of Cowork's output — don't break the contract):**
- `js/lib/btproject.js` — reads + validates the `.btproject` bundle.
- `js/components/import-project.js` — UI for importing a `.btproject`.
- `js/db.js` — IndexedDB schema (versioned). Reads `project-map.json` and seeds inspection types from `reference/inspection-types.json`.

---

## 13 · Lessons Learned (Don't Re-Discover)

A short list of things that took time to figure out and shouldn't be re-derived from scratch:

1. **Use `(x0, y0)` anchors for rotated text, not center coords.** PyMuPDF's bbox extends along the rotated baseline, so center-y is unreliable. Center-coord matching gives zero hits on a rotated BT A1 sheet.
2. **The BT A1 cover schedule is a positional table.** Reading the linearised text dump destroys the column/row structure. Extract by `(x0, y0)` of each numeric span, infer columns from x clusters, rows from y clusters.
3. **Drawing titles like "GENERAL ARRANGEMENT PLAN" don't include "SLAB"** — but they ARE slab plans. Classifier needs explicit rules for these.
4. **CLT timber connections shouldn't inherit `steel-connection`** from a generic `connection` element tag. Suppress when `mass-timber` is present.
5. **Cover schedule rows aren't always in the same order across templates.** Use header anchors (`BOTTOM`, `TOP`, `SIDES` text positions) to tag rows, not row index.
6. **f-strings break in CSS/JS templates** (single `}` is not allowed). Use a plain string template with `__TOKEN__` placeholders + `.replace()` for substitution. Saves debugging time.
7. **Embed the isometric as base64.** A standalone HTML file in OneDrive is much more useful than an HTML + PNG that the user has to keep together.
8. **The hotspot polygon authoring is the highest-value, slowest step.** Plan ~60–90 minutes for a complex project. Do it once, save the polygons in a per-project JSON, never redo.
9. **Conservative bias on the inspection plan.** Engineers prune in-app fast. Missing an inspection is worse than including one they delete.
10. **Sub-builds are real and named.** Lecture theatre, amphitheatre, fin frames, elevated link, plant rooms — they each get their own card(s) under a `Sub-Builds` phase. Don't try to fold them into the main level cycle.
11. **There are at least two BT A1 template variants.** The older one (52SA/Emmanuel pre-mid-2024) and the newer one (MBC, 2024.0230 onwards) have different title-block geometries. Detect which template applies before extracting — see the recogniser snippet in `skills/reading-bt-drawings/SKILL.md` Section 10. (Discovered on MBC, April 2026.)
12. **Cover-sheet drawing list is more reliable than per-page title cells.** Especially in the newer template variant, per-page title extraction is hit-and-miss but the cover-sheet `STRUCTURAL DRAWING LIST` always has every sheet. Cross-check; for projects where per-page is unreliable, use the cover list as authoritative.
13. **Tender Issue projects need a "provisional" caveat.** When a project is "TENDER — NOT FOR CONSTRUCTION", the inspection program is necessarily provisional. Flag this on the HTML (badge in header), in the approve mailto subject, and as a footer note. Re-run the procedure when the Construction Issue arrives.
14. **Multi-tier roofs are common in commercial.** MBC has Lower / Main / Upper roof tiers — three separate inspection sequences (frame + connections + purlins per tier). Don't lump them into one "roof" card.
15. **The bundle PDF folder MUST be `Drawings/` not `pdfs/`.** `js/lib/btproject.js → readBundle()` reads from exactly `Drawings/` (line 86). The `tools/build_btproject.py` generator emits to that path. (Discovered while building the v2.2 bundle generator.)
16. **Per-project `build_html.py` should expose data as module-level constants and wrap render in `if __name__ == '__main__':`.** The bundle generator imports the module to read `PROJECT_META`, `GENERAL_NOTES`, `EXCLUDED_FROM_BT`, `PLAN`, `HOTSPOTS`, `WARNINGS`. Without the guard, importing the module triggers the HTML write. Use `tools/build_btproject.py` as the canonical pattern. (v2.2 build.)
17. **The PWA's `seedReferenceData()` was originally `if (count === 0)` only — meaning new inspection-type keys never reach existing users.** Made idempotent in v2.2 — adds any missing keys without overwriting user-edited entries. Pattern to copy for any future seed function.

---

## 14 · How To Test This Playbook On A New Project

User workflow:

1. Create `Structural Drawings/<JobName>/` (e.g. `Structural Drawings/Mount Alvernia/`)
2. Drop in the combined structural PDF
3. Tell Cowork: "Set up Mount Alvernia"

Cowork follows the procedure in this file end-to-end and produces:

- `Structural Drawings/Mount Alvernia/inspection-program.html` (the review artefact)
- `Structural Drawings/Mount Alvernia/isometric.png` (extracted)
- `Structural Drawings/Mount Alvernia/.btinspect_scratch/` (intermediates — safe to delete)

User opens the HTML, reviews, and either:

- Approves → Cowork writes `project-map.json` + `project.btproject` per Section 9
- Requests changes → Cowork iterates per Section 9

---

## 15 · Iterative Learning Loop

This playbook + the `skills/` folder are designed to **get smarter with every project**. The architecture intent:

- **`CLAUDE.md` is the orchestration layer.** It describes the end-to-end workflow but stays at a "what to do, not how" level. It links out to skills for the deep how.
- **`skills/<skill-name>/SKILL.md` files are the deep-knowledge layer.** Each is a focused, self-contained reference for one capability — title-block extraction, general-notes parsing, hotspot authoring, inspection-plan synthesis, etc. They're written so a future Cowork session can pick them up cold and execute.
- **`reference/` is the shared schema layer** — `inspection-types.json` and `project-map.schema.json`. These are the contracts between Cowork and the PWA.
- **`Structural Drawings/<JobName>/` is the per-project work.** Each project run produces its own deliverables; the *learnings* from each project flow back into CLAUDE.md and skills/, not into the project folder.

### When to update what

| Discovery | Where it goes |
|---|---|
| New keyword/pattern needed in drawing classification | Update the relevant skill's classification rules |
| New BT title-block geometry or template variant | Add or update a skill (e.g. `reading-bt-a3-drawings`) |
| New per-foundation inspection pattern | Section 4.6 of CLAUDE.md OR a `foundation-patterns` skill |
| New hotspot-authoring trick | `skills/hotspot-authoring/SKILL.md` |
| New element type seen in the field | New entry in `reference/inspection-types.json` (and PWA `js/db.js` migration) |
| New schema field needed in `project-map.json` | `reference/project-map.schema.json` (and coordinate with `js/db.js`) |
| One-off lesson, not yet a pattern | Section 13 (Lessons Learned) of CLAUDE.md |

### Update protocol

When the engineer says "we just figured out X", do this:

1. **Capture it in the right place** per the table above. Don't put deep knowledge in CLAUDE.md if there's a focused skill for it — keep CLAUDE.md scannable.
2. **Cross-reference.** If a skill is updated, add a line to Section 11 (Tools) or a relevant procedure step in CLAUDE.md pointing to it.
3. **Update the date stamp** at the bottom of the file.
4. **Surface it in next conversation.** The user expects breakthroughs to compound; the next session should benefit from this one.

### Skills index

See `skills/README.md` for the canonical list. As of last revision:

- `reading-bt-drawings` — BT A1 title-block geometry (older AND newer variants), General Notes parsing, cover-schedule positional extraction, cover-sheet drawing-list extraction.

---

*This playbook is the source of truth for the BT inspection workflow. Update it when patterns emerge that aren't yet documented. Date last revised: 2026-04-29 — v2.2 build: bundle generator (`tools/build_btproject.py`), bundle verifier (`tools/verify_btproject.py`), revision diff (`tools/diff_revision.py`), PWA: outstanding rectifications register, rich pre-inspection brief, Form 12 progress tracker, drawing revision diff import. See `BUILD-PLAN.md` for the full delta.*
