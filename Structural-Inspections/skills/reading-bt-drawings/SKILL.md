# Skill: reading-bt-drawings

**Purpose.** Extract structured data from a Bligh Tanner A1 structural drawing PDF — per-page title-block metadata, project-level metadata, and the General Notes (design codes, concrete cover schedule, bearing capacity, certified-by-others, special notes).

**When Cowork uses this.** Always, on any BT-authored project. Step 4.2 through Step 4.4 of `../../CLAUDE.md`.

**Proven on.**
- 52 Second Av (66 sheets, residential, hybrid concrete + CLT) — older BT A1 template
- Emmanuel College (104 sheets, commercial school + sub-builds) — older BT A1 template
- MBC Creativity & Arts Centre (62 sheets, school arts centre, PT slabs, multi-tier steel roof) — **newer BT A1 template variant** (see Section 10)

All gave 100% clean extraction once the correct cell dictionary was applied.

**Doesn't apply to.** BT A3 sketch sets (different geometry — `/Rotate=0`, title block bottom-right). Client-issued sets (entirely different templates). Scanned PDFs with no text layer (need vision fallback).

---

## 1 · Recognise the BT A1 template

A PDF is a BT A1 set if all of:

- Page 1 dimensions = **2384 × 1684 pt** (A1 landscape)
- `/Rotate = 90` on every page (title block displays down the right edge)
- Page 1 text layer present (≥ 3,000 chars)
- Page 1 contains the strings `BLIGH TANNER`, `STRUCTURAL DRAWING LIST`, `CLIENT`, `DRAWING TITLE`, `DRAWING NUMBER`, `REVISION`, `JOB NO`

Quick check:

```python
import fitz
doc = fitz.open(PDF)
p = doc[0]
print(p.rect.width, p.rect.height, p.rotation)   # expect 2384.0 1684.0 90
print(len(p.get_text()))                          # expect > 3000
```

If any of these fail, this skill doesn't apply — fall back to a different skill or ask the engineer.

---

## 2 · The critical insight: rotated text needs (x0, y0) anchors

PyMuPDF returns a bounding box `(x0, y0, x1, y1)` for every text span. For text rotated 90° (which is everything on a `/Rotate=90` page), the bbox **height grows along the rotated baseline**, so `(y0+y1)/2` (center y) is *much larger* than where the text actually starts. Center-coord matching gives zero hits.

**Always match on `(x0, y0)`** — the top-left anchor. That's the consistent point where the text begins.

```python
# WRONG — center coords don't work for rotated text
cx = (bbox[0] + bbox[2]) / 2
cy = (bbox[1] + bbox[3]) / 2
if xmin <= cx <= xmax and ymin <= cy <= ymax: ...

# RIGHT — use the (x0, y0) anchor
x = bbox[0]
y = bbox[1]
if xmin <= x <= xmax and ymin <= y <= ymax: ...
```

This single fix takes the title-block extractor from 0% hit rate to 100%.

---

## 3 · Title-block cell geometry (BT A1)

Each cell is a `(xmin, xmax, ymin, ymax, size_min, size_max)` window over the unrotated PDF coords. The font-size filter is essential — it disambiguates cells that overlap in position.

```python
CELLS = {
    "sheetNumber":   (1580, 1625,  115, 150,  20, 30),  # e.g. "S005"
    "revision":      (1580, 1625,   50,  80,  20, 30),  # e.g. "C1"
    "jobNumber":     (1580, 1625,  245, 280,  20, 30),  # e.g. "2023.0957"
    "drawingTitle":  (1525, 1580,   55,  80,  14, 18),  # e.g. "FOOTING LAYOUT PLAN - SOUTH"
    "projectName":   (1475, 1510,   55,  78,  15, 19),  # e.g. "EMMANUEL COLLEGE..."
    "siteAddress":   (1505, 1535,   55,  75,  11, 14),  # e.g. "BIRMINGHAM ROAD, CARRARA"
    "drawnBy":       (1470, 1497,  620, 640,  11, 13),  # initials
    "designBy":      (1500, 1527,  620, 640,  11, 13),
    "checkedBy":     (1530, 1555,  620, 640,  11, 13),
    "client":        (1480, 1530, 1390, 1410, 14, 18),  # may be empty in text layer
    "issueLine1":    (1470, 1545,  800, 830,  25, 30),  # e.g. "CONSTRUCTION"
    "issueLine2":    (1505, 1545,  860, 880,  25, 30),  # e.g. "ISSUE"
    "revDate":       (1485, 1505, 2050, 2065,  9, 12),  # e.g. "17.02.25"
}
SCALE_CELL =        (1575, 1625,  905, 925,   6,  8)    # may have multiple values
```

Helper to pick the best span in a cell (largest font wins, ties broken by lower y):

```python
def best_in_cell(spans, cell):
    xmin, xmax, ymin, ymax, smin, smax = cell
    hits = [(b, sz, fn, t) for b, sz, fn, t in spans
            if xmin <= b[0] <= xmax and ymin <= b[1] <= ymax and smin <= sz <= smax]
    if not hits:
        return None
    hits.sort(key=lambda h: (-h[1], h[0][1]))
    return hits[0][3]
```

Where `spans` is `[(bbox, size, font, text), ...]` from `page.get_text("dict")`.

Multi-value cells (like `scale` when the sheet has `1:100` and `1:50` together) use `all_in_cell()` and join with `" / "`.

---

## 4 · Cross-check against the cover sheet's drawing list

Page 1 has a `STRUCTURAL DRAWING LIST` table on the **left half** (x < 1000), with sheet numbers at one column and drawing names at another. The list is rotated text too (each entry is one rotated label).

Use this to verify your per-page extraction:

- Every sheet number in your per-page extraction should appear in the drawing list.
- Every drawing list entry should have a corresponding extracted page.
- Mismatches go in `warnings[]` of the project map.

This is mostly a sanity check — both projects so far have given perfect agreement.

---

## 5 · Project-level metadata (consensus across pages)

For fields that should be invariant (project name, site address, job#, client, EOR, issue status), take the consensus across all 60–100+ pages. If a field varies wildly, take page 1's value and add a warning.

Specific gotchas:

- **`client` is often empty in the text layer** even when visible in the rendered drawing. Both 52 Second Av and Emmanuel had this. Leave `null` and warn — the engineer fills it.
- **`architect` cell has a label but rarely a value.** Same treatment.
- **`issueStatus` is split across two text spans** (`CONSTRUCTION` and `ISSUE`). Concatenate with a space.
- **`engineerOfRecord`** = `"Bligh Tanner"` for BT-authored sets (the BLIGH TANNER block at the bottom of the title block confirms it). For client-issued review sets, the original engineer's name appears here — use that, and tag the project as a review.

---

## 6 · General Notes — pages 2 & 3 (S001 and S002)

Two pages of dense general notes. Extract the full text per page (`page.get_text()`) then parse:

### 6.1 — Design codes

Listed under `ALL MATERIALS AND WORKMANSHIP SHALL BE IN ACCORDANCE WITH...`. Typical baseline:

- AS 1720 (Timber Structures)
- AS 2159 (Piling Code)
- AS 3600 (Concrete Structures)
- AS 3610 (Formwork for Concrete)
- AS 3700 (Masonry Structures)
- AS 4100 (Steel Structures)
- AS 2269 (Structural Plywood)
- NCC

Plus secondary codes scattered through the notes (mine these too): AS 5216 (anchors), AS 1554 (welds), AS 1252 (bolts), AS/NZS 5131 (steel fab), AS/NZS 4680 (galv), AS/NZS 3679 (steel sections), AS/NZS 1163 (hollow sections), AS 1684 (timber framing), AS 1170.4 (earthquake), AS 4671 (reinforcement), AS 4600 (cold-formed steel — if ST notes present).

### 6.2 — Design Criteria block (DC1, DC2, DC3)

- **DC1 — Live Loads**: Per-area loads (roof, floors, bathroom, plant room, stairs, handrails). Floors typically `REFER LOADING PLAN`.
- **DC2 — Wind Loads**: Region (B1 typical), Vu, Vs, terrain category, importance level (IL2 residential, IL3 school/commercial), internal pressure coefficient.
- **DC3 — Earthquake**: Hazard Z, IL, kp, sub-soil class, design category, μ, Sp.

### 6.3 — Foundations (F1...F8)

Foundation type drives the substructure inspection plan. Three patterns we've seen:

- **CFA piles** (52 Second Av): F1 lists pile families with end-bearing values per family (e.g. `P1 = 600 kPa hard clay`, `P2 = 1000 kPa medium dense sand`).
- **Bored piers** (Emmanuel): F1 lists `BORED PIER 1000 kPa` end bearing + `MINIMUM ALLOWABLE SHAFT ADHESION 40 kPa`.
- **Pad/strip on soil** (rare): F1 lists allowable bearing capacity per element.

Always extract the **geotech consultant + report number + date** — usually appears under or next to F1 (`REFER TO GEOTECHNICAL REPORT BY: <consultant> <ref> <date>`).

### 6.4 — Concrete cover schedule (positional extraction required)

This is a tabular block on page 2, NOT readable from linearised text. The visible (rotated) layout is:

```
COVER (mm)   CFA PILES  FOOTINGS  SLAB ON GROUND  COLUMNS  WALLS  SUSPENDED SLAB  ...
SIDES        65         50        45              45       45     35              ...
TOP          -          50        45              45       45     35              ...
BOTTOM       65         50        45              45       -      35              ...
```

In unrotated PDF coords:

- Each element occupies a **column** at a specific x value. Element labels are at varying y. Extract by sorting visible element-name spans by their x position.
- Each row (BOTTOM / TOP / SIDES) sits at a specific y value. Header labels (`BOTTOM`, `TOP`, `SIDES`) are at distinct y positions you can grep for.
- Numeric values fill the (x, y) intersections. `-` means N/A for that face (e.g. piles have no top, walls have no bottom).

```python
# Find the header row y positions
labels = {"BOTTOM": None, "TOP": None, "SIDES": None}
for bbox, sz, fn, t in spans:
    if t in labels and labels[t] is None and 200 < bbox[0] < 230:
        labels[t] = bbox[1]

# Find element column x positions (the row of N32/N40 grade labels right above the element row)
# Then for each (element_x, row_y) intersection, find the numeric span at that anchor
```

The element list and column count vary per project — Emmanuel has 9 columns including `SUSPENDED SLAB - PT`; 52 Second Av has 8 including `BALCONY SLAB` and `DRYING ROOM PRECAST SLAB`. Don't hardcode the column list — discover it from the spans.

### 6.5 — Concrete strengths

Per element, in MPa. Listed alongside the cover table (one row of `N32`, `N40`, `N50` etc. above or below the cover values). Map by x position the same way as cover.

### 6.6 — Certified by Others (G-block)

A summary table near the top of page 2. Lists structural elements **excluded** from BT inspection scope because another party (manufacturer / contractor's RPEQ) certifies them. Typical entries:

- **Precast panel** — additional reo + lifting inserts (PP8 & PP9), temporary propping (PP5)
- **Structural steelwork** — temporary propping & bracing (S18 & G2)
- **Steel stud framing** — sizing, spacing, all connections (ST1)
- **Roof safety systems** — system + connections + load paths

If an element appears here, **delete the corresponding inspection from the BT plan**. Always cite the note reference (e.g. "PP8 & PP9") in the deleted item's `scopeNote` so the engineer knows why.

### 6.7 — Special notes worth extracting

These are the project-specific values the inspection checklists need:

- Curing: `surfaces continuously wet for 3 days, prevent moisture loss for 7 days` (note C5)
- Trimmer bars: `at re-entrant corners and penetrations >200 sq, 2-N12 × 1200 long @ 100 ctrs` (note C6)
- Mesh laps: `2 outer-most cross bars overlap` (note C7)
- Bar chairs: `@ 600 ctrs SL72/82, 800 ctrs SL92+` (note C9)
- Pour temperature: `5–35 °C` (note C19)
- Block walls: `vertical control joints @ 8m max, 5m from corners, not within 1.2m of corners` (note CM6 or M9 depending on template)
- Block wall core fill: `f'c per blockwork schedule, 230 mm slump, 10 mm agg, 2400 mm lift max`
- Anchors into core-filled blockwork: `10% load-tested to 1.5 × SWL` (note CM15 or M14)
- Post-installed anchor load testing: `5% of all anchors, 100% if any failure` (note A8)
- 24-hour notice required for inspections per Form 12 certification clause
- For CLT projects: `48-hour notice` per CLT22 plus 5 explicit hold-point stages

---

## 7 · Output shape

This skill produces three structured objects:

```python
project_metadata = {
    "name": "...", "siteAddress": "...", "jobNumber": "...",
    "client": None,        # often empty
    "issueStatus": "CONSTRUCTION ISSUE",
    "engineerOfRecord": "Bligh Tanner",
    "discipline": "structural",
    "rpeqSignatory": None,  # not in title block
}

drawings = [
    {"pageNumber": 1, "sheetNumber": "S000", "revision": "C1",
     "revisionDate": "17.02.25", "description": "COVER SHEET",
     "scale": None, "drawnBy": "SR", "checkedBy": "TM", "approvedBy": "AY",
     "confidence": 1.0},
    # ... one per page
]

general_notes = {
    "designCodes": [...],
    "exposureClass": None,    # rarely stated explicitly
    "concreteCover": {"footings": {"bottom": 50, "top": 50, "sides": 50}, ...},
    "concreteStrengths": {"footings": 32, "suspendedSlab": 40, ...},
    "bearingCapacity": {...},   # shape depends on foundation type
    "windRegion": "B1", "windVelocity": {"ultimate": 55.3, "serviceability": 35.9, ...},
    "earthquake": {...},
    "geotechReport": {"consultant": "Douglas Partners", "reportNumber": "228091.00", "date": "2024-04-30"},
    "certifiedByOthers": [...],
    "specialNotes": [...],
}
```

These feed directly into the next steps (drawing classification, inspection plan synthesis).

---

## 8 · Things that break this skill

A growing list — add to it whenever you discover a new failure mode.

1. **Center-coord matching** — gives zero hits because the bbox height of rotated text is much larger than the cell window. Use `(x0, y0)` anchors. (Discovered on 52 Second Av.)
2. **f-strings around CSS/JS templates** — single `}` is invalid in f-strings. Use a plain template + `.replace("__TOKEN__", value)` for any code generation. (Discovered building the Emmanuel HTML.)
3. **Linearised cover-table text** — `page.get_text()` destroys the row/column structure. Always extract the cover schedule by `(x0, y0)` of each numeric span. (Discovered on 52 Second Av.)
4. **Architect / client cells empty in text layer** — the cells exist visually but have no text spans. Don't error; leave null and warn. (Both projects so far.)
5. **Issue status split across two spans** — `CONSTRUCTION` (issueLine1) + `ISSUE` (issueLine2). Concatenate.
6. **Cover-table column order varies** — don't hardcode the element list; discover it from x positions. Emmanuel has `SUSPENDED SLAB - PT` (9 cols), 52SA has `BALCONY SLAB` and `DRYING ROOM PRECAST SLAB` (8 cols).
7. **Revisions on individual sheets** — most sheets are at the project's master revision (e.g. `C1`) but individual sheets can be ahead (e.g. Emmanuel's `S023` was at `C2`). Capture per-page; don't assume project-level.
8. **`/Rotate` value is per-page, not per-document.** It happens to be 90 on every page of every BT A1 set we've seen, but verify per page.
9. **There are TWO BT A1 template variants — older (52SA/Emmanuel) and newer (MBC).** They have different cell positions. The recogniser must determine which template applies before picking a cell dictionary. See Section 10. (Discovered on MBC, April 2026.)
10. **Long drawing titles wrap across two lines** in the MBC template (e.g. "FOOTING PLAN NOTES, LEGENDS," + "& SCHEDULES"). Try both title cells and concatenate when both have content. Cover sheet titles like "COVER SHEET" only use the second line.
11. **Cover-sheet drawing list is always more reliable than per-page title cells for descriptions.** When the per-page title cell is hit-and-miss (as in MBC), use the cover-sheet `STRUCTURAL DRAWING LIST` extraction as the authoritative sheet→description map. Pattern: each name span sits directly above its sheet number in the rotated layout (same x, smaller y).

---

## 9 · Reference scripts

Working implementations of this skill:

- 52 Second Av extractor: `Structural Drawings/2023.0092 52 Second Avenue/.btinspect_scratch/extract_titleblock.py` (if not deleted).
- Emmanuel extractor (same approach, different PDF path): `Structural Drawings/Emmanuel/.btinspect_scratch/extract_titleblock.py`.
- MBC extractor (NEWER template — different cell positions): `Structural Drawings/MBC/.btinspect_scratch/extract_titleblock.py`.
- MBC drawing-list extractor (cover-sheet S### → name mapping): `Structural Drawings/MBC/.btinspect_scratch/extract_drawing_list.py`.

All produce ~100% extraction on their respective sets.

---

## 10 · BT A1 Template Variants

Two BT A1 template variants encountered to date. Cowork must detect which one applies and use the matching cell dictionary.

### Older variant (52SA + Emmanuel, 2023–2024)

Title block runs down the right edge in a tall, single-column stack. Issue stamp ("CONSTRUCTION ISSUE") at right-edge bottom (y≈812).

Cell dictionary in Section 3 above (`CELLS`). Use this for projects up to ~mid-2024.

### Newer variant (MBC, 2024.0230 onwards)

Restructured title block — most cells moved to lower y values (above the printing-requirements area). Issue stamp moved to **upper-left of page** (around x≈810, y≈100-160), well away from the title block. Site address split across **two lines** (line 1 at y≈63, line 2 at y≈189). Drawing title can wrap across **two lines** (line A at y≈220 always present for short titles, line B at y≈66 for the start of long titles). Engineer initials moved to y≈201 (was y≈629).

Cell dictionary for the newer variant:

```python
CELLS_MBC = {
    "sheetNumber":      (1580, 1620,   95, 130,  20, 30),   # was y 115-150
    "revision":         (1580, 1620,   55,  75,  20, 30),   # similar to old
    "jobNumber":        (1580, 1620,  195, 225,  20, 30),   # was y 245-280
    "drawingTitleA":    (1520, 1580,  215, 230,  14, 18),   # primary slot (always used)
    "drawingTitleB":    (1520, 1580,   58,  80,  14, 18),   # wrap line 1 (long titles only)
    "projectName":      (1430, 1455,  110, 130,  15, 19),   # smaller font, lower position
    "siteAddressLine1": (1455, 1475,   58,  80,  13, 16),   # SPLIT — line 1
    "siteAddressLine2": (1475, 1495,  180, 200,  13, 16),   # SPLIT — line 2
    "issueLine1":       (795, 875,   140, 165,  25, 30),    # "TENDER" or "CONSTRUCTION" — UPPER LEFT
    "issueLine2":       (845, 880,    95, 115,  16, 20),    # "NOT FOR CONSTRUCTION" or "ISSUE"
    "drawnBy":          (1340, 1360,  195, 210,  11, 13),   # initials cluster at y≈201
    "designBy":         (1368, 1385,  195, 210,  11, 13),
    "checkedBy":        (1395, 1410,  195, 210,  11, 13),
}
```

**No client cell** in the newer variant title-block area we've seen. Site address is the school name + street (e.g. "MORETON BAY COLLEGE, 450 WONDALL ROAD, MANLY WEST, QLD 4179") so the school doubles as both client and site.

### How to detect the variant

Quick heuristic — check whether the older-variant `sheetNumber` cell (y=115-150) has a hit on page 5+:

```python
older_variant_hit = any(
    1580 <= bbox[0] <= 1625 and 115 <= bbox[1] <= 150 and 20 <= sz <= 30
    for bbox, sz, _, _ in extract_spans(doc[4])
)
template = "older" if older_variant_hit else "newer"
```

Then load the matching cell dictionary.

### When per-page title extraction is unreliable: use the cover-sheet drawing list

In the MBC variant, only ~20% of pages had clean drawing-title extraction (titles vary in position, sometimes split, sometimes not present in the expected cell). The cover-sheet `STRUCTURAL DRAWING LIST` is a much more reliable source — it always contains every sheet number paired with its full name.

Extraction approach (from `extract_drawing_list.py`):

1. Find every span on page 1 matching `S\d{3}` at sz≈10 with x < 1700 (drawing-list region, not title block).
2. Find every other sz≈10 text span in the same x range.
3. For each sheet number, find the closest name span at the **same x** (within ±2.5 pt) but **smaller y** (the name sits "above" the sheet number in the rotated layout).
4. If a second name span exists just above the first (within 25 pt), it's a wrapped 2-line name — concatenate.

This gave 62/62 sheet→name mappings on MBC. Always run it as a cross-check; for newer-template projects, treat it as authoritative.

---

*Last revised: 2026-04-28. Added MBC template variant.*
