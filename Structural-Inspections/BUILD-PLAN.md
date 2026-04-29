# BT Inspect — v2.2 Build Plan

This document is the roadmap for the next build of BT Inspect. It captures the *why*, the *what*, and the *order of operations* — written so a future Cowork session (or a different engineer) can pick this up cold and finish.

**Created:** 2026-04-29.
**Origin:** A scoping conversation with Dan (Senior Structural Engineer at BT) after the MBC test project. We had built the full Cowork → HTML review → inspection-program flow on three projects (52SA, Emmanuel, MBC) and reviewed the existing PWA to identify what was missing.

The conclusion: the PWA already implements ~80% of the smart workflow we want, but a few high-leverage features and a generator script are missing. Once added, BT Inspect goes from a clever prototype to a real production tool.

---

## 1 · Vision (the engineer's day this build serves)

> **Tuesday, builder rings.** "We're pouring the L1 slab Friday morning, can you inspect Thursday afternoon?" Dan opens BT Inspect, sees the inspection on his project plan, taps it. **The pre-inspection brief loads** with the cover schedule for L1 slab, concrete grade N40, AS clauses he's working to, what's certified by others (so he knows not to inspect the steel stud framing), and a site-readiness check that pings the builder for confirmation. Dan books the visit.
>
> **Thursday 2pm, on site.** Walks the slab. Drops pins for two trimmer bar locations (defects), one cover concern (observation), takes photos, captures the concrete delivery docket with one tap (auto-OCR'd against the spec). Signs the hold-point release for the rest. Driving home, the report PDF is already drafted on his phone.
>
> **Monday morning.** Opens BT Inspect. **Outstanding rectification register** shows 14 items across three projects. Two are marked rectified by the builder over the weekend (with photos attached) — Dan reviews and one-taps to close. Twelve remain. The MBC project shows "28 of 28 inspections pending start, 0 outstanding rectifications, Form 12 ready when 12 close."

What success looks like: every step above is one or two taps, no manual data entry that the system already knows.

---

## 2 · Where the system stands today

### What's built and working

- **Cowork procedure** (`CLAUDE.md` + `skills/reading-bt-drawings/SKILL.md`) — extracts title-block metadata, General Notes (cover schedule, bearing, certified-by-others), classifies drawings, builds inspection plans. Proven on 52SA, Emmanuel, and MBC. Two BT A1 template variants supported.
- **HTML review artefact** (`Structural Drawings/<JobName>/inspection-program.html`) — Anthropic-styled, sticky isometric, click-to-highlight hotspots over visible structural elements. Engineer reviews here before bundle generation.
- **PWA architecture** (`js/`) — IndexedDB schema v5 designed for the rich project-map data; bundle reader (`js/lib/btproject.js`) validates and parses .btproject files; atomic import via `createProjectFromBundle`; project view shows inspection plan with status; tap a plan entry → `startInspectionFromPlanEntry` resolves drawing refs, prefills checklist, sets hold-point flag, links inspection back to plan; inspection-detail surfaces "From the project plan" card with rationale + suggested checks; smart drawing auto-pick from `applicableInspectionTypes[]`.
- **Reference data** (`reference/inspection-types.json`, `reference/project-map.schema.json`) — the contracts between Cowork and PWA.

### What's missing

1. **No `.btproject` generator.** Cowork can produce the HTML review artefact but doesn't yet write the `project-map.json` + zip. **The single biggest gap** — without this, nothing flows from Cowork to PWA.
2. **CLAUDE.md path bug.** Section 9 says bundle layout is `pdfs/<filename>.pdf`. PWA reads from `Drawings/<filename>.pdf`. Bundles will fail if not fixed.
3. **Outstanding rectification register** — no project-wide "what's still open across all my inspections" view. The single most valuable new screen for the senior engineer.
4. **Pre-inspection brief is bare-bones.** Tapping a plan entry just creates the inspection; the rich brief (cover values, AS clauses, certified-by-others exclusions, site-readiness check) doesn't exist.
5. **Form 12 / RPEQ certification tracking.** No top-level "Form 12 ready when these close" view.
6. **Drawing revision diff workflow** described in CLAUDE.md Section 10 but not implemented end-to-end.
7. **Inspection-types catalogue gaps.** Using proxies for CFA piles, bored piers, concrete walls, stairs, head-framing. Not blocking but worth fixing.
8. **Schema doesn't yet have:** `siteReadinessCheck[]` per plan entry, `asClauseRefs[]` per checklist item, `excludedFromBT[]` at project level (the certified-by-others summary), `postPourRecord` schema for delivery dockets / cert capture.

---

## 3 · Goals for this build (in priority order)

These come from putting on the structural-engineer hat and asking *what would actually change my workflow*.

| # | Feature | Why it matters | Effort |
|---|---|---|---|
| 1 | `.btproject` generator script | Closes the Cowork → PWA loop. Without it nothing else is testable. | M |
| 2 | Outstanding rectification register | The one screen the senior engineer would open every Monday | M |
| 3 | Rich pre-inspection brief | Engineer arrives on site sharp, not scrambling | S |
| 4 | Form 12 / progress tracker | The legal endpoint of the inspection program — currently invisible | S |
| 5 | Drawing revision diff (Cowork side + PWA badge) | Tender → Construction transition is a real risk; this de-risks it | M |
| 6 | Site-readiness check (advisory modal) | Stops wasted site trips | S |
| 7 | Add proper inspection-type keys (CFA, bored pier, etc.) | Removes ugly proxies; small coordinated change | S |
| 8 | Schema additions for the above | Makes future-me's life easier | S |

Effort: S = under an hour, M = a couple of hours.

### What's deliberately NOT in this build

- **AI comment suggestion from photos.** Tempting; legal liability + 10% error rate kills adoption. Defer.
- **BIM model viewer.** PDF drawings are the universal contract; not every project has a clean Revit model.
- **Calendar / Outlook integration.** Tangential to "what good looks like".
- **Builder-facing portal / login.** Email a PDF — that's what builders actually want.
- **Form 12 PDF generation + signature workflow.** Track status this build, generate PDF in a follow-up build.

---

## 4 · Architecture decisions for this build

### 4.1 — Schema versioning

Bumping `project-map.schema.json` to **v2** would break all existing imports (PWA hard-checks v1 in `btproject.js`). Instead: **keep schemaVersion = 1, add new fields as OPTIONAL**, document them in the schema with `default` values so older bundles still validate. PWA reads new fields if present, falls back gracefully if not.

### 4.2 — Where the rectification register lives

A new top-level view in the PWA (route `#/rectifications`), accessible from the existing nav. It's a *derived* view over `db.inspections.items` — no new tables. Items are filtered to `severity in (defect, hold-point) AND (status == 'open' OR status undefined)`. Adding a `closedAt` + `closedNotes` field to items is the only schema delta.

### 4.3 — Pre-inspection brief

Inserts a step between "tap plan entry" and "inspection created". Modal opens showing the brief, with "Start inspection" as the primary action. Drawn entirely from data we already have in `project.projectMap.generalNotes` + the plan entry itself; no new data needed. Site-readiness check is also in this modal as a simple checklist.

### 4.4 — Form 12 tracker

Top of project view: a stats strip showing total inspections, % complete, hold points outstanding, rectifications outstanding, Form 12 readiness. Plus a "Mark Form 12 issued" action when ready. Pure UI; no PDF generation in this build.

### 4.5 — Drawing revision diff

**Cowork side:** when a new PDF appears in an existing project folder, write `revision-<timestamp>.json` per CLAUDE.md Section 10. Don't overwrite project-map.json.

**PWA side:** add an "Import revision" entry to project menu. Reads the diff JSON, marks affected `drawings` rows with `revisionChangedAt`, marks affected plan entries / inspections with a "revision changed since last review" badge. Engineer reviews and one-taps to acknowledge.

### 4.6 — Schema additions (all v1-compatible, optional)

```json
"inspectionPlan[].siteReadinessCheck": {
  "type": "array",
  "items": { "type": "string" },
  "description": "Advisory checklist for the builder before booking"
}
"inspectionPlan[].expectedChecklist[].asClauseRef": {
  "type": "string",
  "description": "e.g. 'AS 3600 Cl 4.10.3' — surfaces in the brief"
}
"project.excludedFromBT": {
  "type": "array",
  "items": {
    "type": "object",
    "properties": {
      "element": { "type": "string" },
      "responsibility": { "type": "string" },
      "noteRef": { "type": "string" }
    }
  },
  "description": "Mirror of certifiedByOthers — surfaced in pre-inspection brief"
}
"inspectionPlan[].postPourRecord": {
  "type": "object",
  "properties": {
    "concreteGrade": { "type": "string" },
    "expectedSlump": { "type": "number" },
    "tempRange": { "type": "string" }
  },
  "description": "What to capture from delivery dockets"
}
```

### 4.7 — New inspection-type keys

If approved (see Q2 in the questions section): add to `reference/inspection-types.json`:

- `pile-cfa-install` — CFA pile installation (founding depth + torque + socket)
- `pile-bored-install` — Bored pier installation (founding depth + shaft adhesion)
- `concrete-wall-prepour` — Concrete wall pre-pour reinforcement
- `stair-prepour` — Stair flight pre-pour
- `head-framing` — Suspended slab head framing inspection (separate from pre-pour reo)
- `transfer-slab-witness` — Critical transfer slab pour witness

Coordinated bump in `js/db.js` (the seed array around line 175) — pure data addition, no schema migration needed because the table structure stays the same.

---

## 5 · Build phases (execution order)

### Phase A — Schema + reference data (low risk, foundational)

A1. Fix CLAUDE.md Section 9 path bug (`pdfs/` → `Drawings/`).
A2. Add new optional fields to `reference/project-map.schema.json` per 4.6.
A3. Add new inspection-type keys to `reference/inspection-types.json` per 4.7 (pending approval).
A4. Update `js/db.js` seed and add migration if any new keys.
A5. Update `skills/reading-bt-drawings/SKILL.md` with newer-template cell dictionary references (already partially done).

### Phase B — Bundle generator + MBC end-to-end

B1. Write `tools/build-btproject.py` — reusable script that takes a project folder + analysed data and produces `project-map.json` + `project.btproject`.
B2. Adapt the per-project Cowork build script (`build_html.py` style) to also call the bundle generator after writing the HTML.
B3. Generate `Structural Drawings/MBC/MBC.btproject`.
B4. Validate the bundle against schema using `jsonschema`.
B5. Manually exercise the bundle reader (`js/lib/btproject.js → readBundle`) by loading it in a Node test harness or directly in browser console.
B6. End-to-end import test in the PWA: import → see plan → tap entry → confirm prefill works.

### Phase C — PWA enhancements

#### C1. Outstanding rectification register

C1.1. Add a new view file `js/views/rectifications.js`.
C1.2. Add the `#/rectifications` route in `js/router.js`.
C1.3. Add nav item in `index.html`.
C1.4. Query: across all projects, items where severity ∈ (defect, hold-point) and status ≠ closed.
C1.5. Per-item card: shows project + inspection + first photo + comment preview + age in days.
C1.6. One-tap close-out modal: confirm + optional rectification photo + close.
C1.7. Schema delta on `db.items`: add `closedAt`, `closedNotes`, `closedByPhotoId` fields.

#### C2. Rich pre-inspection brief

C2.1. New component `js/components/pre-inspection-brief.js`.
C2.2. Modal trigger from project-detail.js when tapping a pending plan entry.
C2.3. Pulls cover values + grade for the relevant element from `projectMap.generalNotes.concreteCover` + `concreteStrengths`.
C2.4. Pulls `excludedFromBT` to surface "you're not inspecting these — here's why".
C2.5. Renders site-readiness checklist (from `inspectionPlan[].siteReadinessCheck` if present, else a sensible default per type).
C2.6. Renders AS clause refs for the type from a bundled lookup (or just shows the existing `expectedChecklist`).
C2.7. "Start inspection" → existing `startInspectionFromPlanEntry` flow.
C2.8. "Cancel" / "Not yet ready" → no inspection created.

#### C3. Form 12 progress tracker

C3.1. Add to the top of `js/views/project-detail.js` — a stats strip.
C3.2. Compute: total inspections, complete, in-progress, pending, hold points outstanding, rectifications outstanding (cross-reference items table).
C3.3. "Form 12 ready" badge when complete count == total AND outstanding rectifications == 0.
C3.4. "Issue Form 12" action — for now just sets `project.form12IssuedAt` timestamp.

#### C4. Drawing revision diff

C4.1. Cowork-side script: `tools/diff-revision.py` — takes old + new PDF, writes `revision-<ts>.json`.
C4.2. Document the file format — list of `{sheetNumber, oldRevision, newRevision, changeType: added|changed|removed, affectedInspectionPlanIndices}`.
C4.3. PWA-side: add "Import revision" button on project page, opens file picker for `revision-*.json`.
C4.4. On import: stamp affected drawings with `revisionChangedAt`, mark plan entries with `revisionPending: true`.
C4.5. UI: badge on plan entries that have unreviewed revisions; modal to review the diff.

### Phase D — Validation + documentation

D1. Full end-to-end test on MBC.
D2. Update CLAUDE.md to reflect the new bundle generator workflow + new features.
D3. Update `skills/` with anything we learned during the build.
D4. Mark BUILD-PLAN.md sections as complete.
D5. Decision: which features to ship in v2.3 (the deferred items: AI comment suggestion, Form 12 PDF gen, BIM viewer).

---

## 6 · Status checklist

Use this to track progress. Update as each item is completed. If a future session is picking this up cold, this is the first thing to read.

### Phase A — Schema + reference data

- [x] A1. CLAUDE.md path bug fixed (`pdfs/` → `Drawings/`)
- [x] A2. project-map.schema.json updated with new optional fields (`excludedFromBT`, `siteReadinessCheck`, `asClauseRef`/`noteRef` on checklist items, `postPourRecord`, `phase`)
- [x] A3. inspection-types.json updated with 6 new keys (`pile-cfa-install`, `pile-bored-install`, `concrete-wall-prepour`, `stair-prepour`, `head-framing`, `transfer-slab-witness`); catalogue version bumped to v2
- [x] A4. js/db.js seed updated to be idempotent (safely adds new keys to existing DBs); items table extended with `closedAt`/`closedNotes`/`closedByPhotoId`; new functions: `closeItem`, `reopenItem`, `listOutstandingRectifications`, `getProjectProgress`, `markForm12Issued`, `unmarkForm12Issued`, `applyRevisionDiff`, `acknowledgeRevisionChange`
- [x] A5. skills/reading-bt-drawings/SKILL.md cross-reference confirmed (already up-to-date with newer-template variant from previous session)

### Phase B — Bundle generator + MBC

- [x] B1. `tools/build_btproject.py` written (250 lines, reusable across all projects)
- [x] B2. MBC `build_html.py` refactored — exposes `PROJECT_META`, `GENERAL_NOTES`, `EXCLUDED_FROM_BT`, `WARNINGS`, `PLAN`, `HOTSPOTS` as module-level constants; HTML render wrapped in `if __name__ == '__main__'` guard. New inspection-type keys used (`pile-bored-install`, `concrete-wall-prepour`, `stair-prepour`).
- [x] B3. `Structural Drawings/MBC/project.btproject` generated (86.5 MB, 28 inspections, 62 drawings, 3 hold points, 0 soft validation issues)
- [x] B4. Bundle validates against project-map.schema.json (jsonschema confirmed)
- [x] B5. `tools/verify_btproject.py` written — mirrors `js/lib/btproject.js → readBundle()` validation; MBC.btproject passes all checks including bundle path layout (`Drawings/`)
- [ ] B6. End-to-end import in PWA — pending engineer testing in browser

### Phase C — PWA enhancements

- [x] C1. Outstanding rectification register — new view `js/views/rectifications.js`, route `#/rectifications`, nav item, age sorting, project filter, close-out modal with optional rectification photo
- [x] C2. Rich pre-inspection brief — new component `js/components/pre-inspection-brief.js`, replaces the bare `confirmDialog`, surfaces cover values + grade for the relevant element, site-readiness check (per-entry override else sensible defaults per type), expected checklist with structured AS clause refs, certified-by-others exclusions, drawings to take, post-pour record requirements
- [x] C3. Form 12 progress tracker — `renderProgressStrip` shows %, complete/pending, hold-points outstanding, rectifications outstanding, "Form 12 ready when…" status; mark-issued/undo actions wired; `renderExcludedFromBT` shows the certified-by-others summary
- [x] C4. Drawing revision diff — `tools/diff_revision.py` produces revision-{ts}.json by comparing existing project-map vs new PDF; `js/components/import-revision.js` reads + previews + applies; PWA shows pending-revisions banner + per-plan-entry "revision changed" badge with acknowledge action; round-trip diff (MBC vs MBC.pdf) returns 0/0/0 as expected

### Phase D — Validation

- [x] D1. Bundle validates against schema; round-trip diff round-trips clean; all new JS files pass `node --check`
- [x] D2. CLAUDE.md updated — Section 9 bundle path bug fixed; lessons updated for v2.2
- [x] D3. skills/ already up-to-date (newer-template variant added previously); no further updates needed this build
- [x] D4. BUILD-PLAN.md marked complete (this section)
- [ ] D5. v2.3 backlog drafted (see Section 7 — already populated with deferred items)

### v2.2.1 patch — BT Standard Checks library

After v2.2 ship, Dan asked Cowork to extract every checkable item from the BT General Notes for MBC. Result: a comprehensive per-type checklist library that should be baked into the PWA (not just produced as Markdown each time). This patch did that.

- [x] P1. New file `js/lib/bt-standard-checks.js` (971 lines) — 27 inspection types fully covered with site-readiness, on-site checks (with AS clause + note refs + critical markers), pour-day records, certifications expected, standard comments
- [x] P2. `js/components/pre-inspection-brief.js` — uses `mergeWithProjectChecks(type, entry)` from the library; renders project + standard items with provenance tags; adds dedicated "Pour-day records" and "Certifications expected" sections
- [x] P3. `js/views/inspection-detail.js` — `renderPlanContextCard` now surfaces BT standard checks alongside project-specific. Card renders even for ad-hoc inspections (no plan-entry origin) — the standard checks are always available.
- [x] P4. `db.js → GENERAL_COMMENTS_SEED` — added 17 new standard comments covering all the new inspection types so the comment library is populated when the engineer pins items
- [x] P5. CSS — source tags (project / standard), critical markers (red exclamation), capture-type chips (photo / measure / verify) on pour-day records
- [x] P6. Cache-bust bumped: `index.html` script src `v=2.2.1-standard-checks`
- [x] P7. Verified: all 27 catalogue keys present in standard library; no orphans either direction; node syntax clean

---

## 7 · Deferred to a later build (the v2.3+ backlog)

So we don't lose track of anything from the engineer's-perspective review:

- **Form 12 PDF generation** with signature capture.
- **AI comment suggestion** from photo + inspection-type context.
- **Pour-day docket capture** with OCR (delivery docket → mix ID + slump + temp).
- **Anchor / cover-meter test register** (5%/10% sample tracking).
- **"Last time I did this" memory** — pattern surfacing per inspection type from past inspections.
- **Side-by-side photo compare** — current vs previous inspection of same area.
- **Voice-to-text** for hands-free site annotation.
- **Calendar integration** (Outlook / Google).
- **Email-the-report** one-tap to the project distribution list.
- **BIM model viewer** (Revit IFC import) — replace or augment PDF drawings.
- **Multi-engineer assignment** — junior on ground floor, senior on L1+.
- **Pattern recognition across projects** — "this builder commonly has cover issues".

---

## 8 · For a future session picking this up cold

If you're a fresh Cowork session and Dan asks you to "continue the build":

1. Read `CLAUDE.md` (the orchestration playbook)
2. Read `BUILD-PLAN.md` (this file)
3. Check section 6 (status checklist) for what's done and what's next
4. Each Phase has self-contained sub-steps — just work down the list
5. Update the checkboxes as you go
6. If you discover something new, capture it in `skills/` (focused) or as a `Lessons Learned` entry in CLAUDE.md (one-off)
7. Don't break the integrity rules in CLAUDE.md Section 3:
   - `reference/inspection-types.json` is the catalogue contract — coordinated changes only
   - `reference/project-map.schema.json` is the wire format — backward-compatible changes only

Open questions that may need re-asking the engineer:

- Form 12 PDF generation — full template or just status tracking? (Current build: status only.)
- New inspection-type keys — confirmed approved, see this build for `pile-cfa-install`, `pile-bored-install`, `concrete-wall-prepour`, `stair-prepour`, `head-framing`, `transfer-slab-witness`.
- Schema bump — kept as v1 with new optional fields; bumping to v2 means coordinated PWA migration.

---

*This plan is the source of truth for the v2.2 build. Update as work progresses. Date last revised: 2026-04-29.*
