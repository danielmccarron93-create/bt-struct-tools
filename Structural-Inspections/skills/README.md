# Cowork Skills — BT Structural Inspections

This folder holds **focused, reusable knowledge artefacts** that Cowork uses to do its job. Each subfolder is one skill with a `SKILL.md` at its root.

A skill is the deep-knowledge layer for one specific capability. The orchestration layer (`../CLAUDE.md`) describes the end-to-end workflow at a high level and points here for the deep how.

## Why skills (separate from CLAUDE.md)?

- **CLAUDE.md must stay scannable.** If every lesson learned bloated it, future Cowork sessions would scroll past the important bits.
- **Skills compose.** A new project might need `reading-bt-drawings` + `hotspot-authoring`, but not `client-issued-sets`. Future Cowork pulls only what's relevant.
- **Skills evolve independently.** A breakthrough in title-block extraction shouldn't require a CLAUDE.md edit — just a SKILL.md update in `reading-bt-drawings/`.
- **Skills capture failure modes.** Each SKILL.md ends with "things that break this skill" so we don't re-learn them.

## How to use a skill (Cowork-side)

When Cowork is mid-procedure and hits a step that maps to a skill (per the table in CLAUDE.md Section 15), read the SKILL.md before executing. They're written so a fresh Cowork session can pick them up cold.

## How to add or update a skill

When the engineer says "we just figured out X":

1. Pick the right home — see CLAUDE.md Section 15 for the routing table.
2. If new pattern, create `skills/<name>/SKILL.md`. Use the existing skills as a template.
3. If updating: edit in place, add to the "Things that break this skill" list at the bottom.
4. Update this index (`README.md`) — one line per skill in the table below.
5. Update CLAUDE.md Section 15 if the workflow shape changed.

## Index

| Skill | Purpose | Status |
|---|---|---|
| [`reading-bt-drawings`](reading-bt-drawings/SKILL.md) | BT A1 title-block geometry, General Notes parsing, cover-schedule positional extraction. The first thing Cowork does on any new project. | v1 — proven on 52 Second Av + Emmanuel |

## Future skills to add (in priority order)

These are skills we know we'll need based on patterns we've seen but haven't yet formalised:

- **`drawing-classification`** — kind/level/element/zone tagging rules; the heuristics we've already discovered (GA plans = slab plans even without "SLAB" in title; CLT connections ≠ steel-connections; etc.)
- **`inspection-plan-synthesis`** — building the build-order plan from classified drawings; per-foundation patterns (CFA / bored pier / pad on soil); per-level template; sub-build handling.
- **`hotspot-authoring`** — the polygon authoring approach for the inspection-program HTML; quadrant-by-quadrant visual inspection; coordinate system; element-type taxonomy.
- **`inspection-html-rendering`** — the HTML template; layout rules; interaction model; brand tokens.
- **`reading-bt-a3-drawings`** — the A3 sketch-set variant of the BT template (different title-block geometry, smaller projects).
- **`reading-client-drawings`** — non-BT templates (Queensland Rail, DTMR, etc.) where BT is reviewing someone else's work.
- **`general-notes-by-template`** — the parsing rules per BT template variant; how to extract certified-by-others, special notes, etc.
