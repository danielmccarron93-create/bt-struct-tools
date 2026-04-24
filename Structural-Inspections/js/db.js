/**
 * IndexedDB schema via Dexie.
 *
 * All Bligh Tanner inspection data lives here, offline-first.
 *
 * Schema evolution: bump the version number and define a `.upgrade()`
 * callback. Never silently drop data.
 *
 * Phase 2 notes:
 *   - Project gains `client`, `siteAddress`, `notes` (non-indexed, no migration required).
 *   - Drawing gains `filename`, `description`, `pdfBlob` (Blob), `pageWidth`,
 *     `pageHeight`, `rotation` (0/90/180/270), `calibration` (see schema below),
 *     `uploadedAt`, `updatedAt`. All non-indexed; no migration required.
 *
 * Phase 4 notes:
 *   - New `highlights` table for extent-of-inspection rectangles. Schema bumped
 *     to v2; no data migration needed (additive only).
 *   - `items` gains pdfX/pdfY/drawingId/page/comment (all non-indexed).
 *   - `photos` gains blob/caption/width/height (all non-indexed).
 */

/* global Dexie */

export const db = new Dexie('BTInspectionReport');

db.version(1).stores({
  // Core entities
  projects:      '++id, jobNumber, name, createdAt, updatedAt',
  drawings:      '++id, projectId, sheetNumber, revision, [projectId+sheetNumber]',
  inspections:   '++id, projectId, inspectionTypeId, date, status, createdAt, updatedAt',
  items:         '++id, inspectionId, itemNumber, severity, status, [inspectionId+itemNumber]',
  photos:        '++id, itemId, createdAt',

  // Reference libraries (seeded)
  inspectionTypes:    '++id, &key, category, name',
  commentLibrary:     '++id, &key, inspectionTypeKey, severity',
  generalComments:    '++id, inspectionTypeKey',

  // User / settings
  users:         '++id, &email, rpeq, role',
  settings:      '&key'
});

// v2 — add highlights table for extent-of-inspection rectangles.
// All existing tables re-declared unchanged so Dexie keeps them.
db.version(2).stores({
  projects:      '++id, jobNumber, name, createdAt, updatedAt',
  drawings:      '++id, projectId, sheetNumber, revision, [projectId+sheetNumber]',
  inspections:   '++id, projectId, inspectionTypeId, date, status, createdAt, updatedAt',
  items:         '++id, inspectionId, itemNumber, severity, status, [inspectionId+itemNumber]',
  photos:        '++id, itemId, createdAt',
  highlights:    '++id, inspectionId, drawingId, [inspectionId+drawingId]',
  inspectionTypes:    '++id, &key, category, name',
  commentLibrary:     '++id, &key, inspectionTypeKey, severity',
  generalComments:    '++id, inspectionTypeKey',
  users:         '++id, &email, rpeq, role',
  settings:      '&key'
});

// v3 — add reports table for saved PDF reports. Additive, no migration needed.
db.version(3).stores({
  projects:      '++id, jobNumber, name, createdAt, updatedAt',
  drawings:      '++id, projectId, sheetNumber, revision, [projectId+sheetNumber]',
  inspections:   '++id, projectId, inspectionTypeId, date, status, createdAt, updatedAt',
  items:         '++id, inspectionId, itemNumber, severity, status, [inspectionId+itemNumber]',
  photos:        '++id, itemId, createdAt',
  highlights:    '++id, inspectionId, drawingId, [inspectionId+drawingId]',
  reports:       '++id, inspectionId, generatedAt, [inspectionId+generatedAt]',
  inspectionTypes:    '++id, &key, category, name',
  commentLibrary:     '++id, &key, inspectionTypeKey, severity',
  generalComments:    '++id, inspectionTypeKey',
  users:         '++id, &email, rpeq, role',
  settings:      '&key'
});

// v4 — multi-page PDF support.
//
// The old model stored a Blob on every drawing. That's wasteful for a 26-page
// structural set (we'd keep 26 copies of the same 6.8 MB file). Now: a single
// `pdfSources` record holds the blob; each `drawings` row references it via
// `sourcePdfId` and picks a specific `pageNumber`.
//
// Migration moves each existing drawing's pdfBlob into a fresh pdfSources
// record (one per drawing — no hash-based dedup in v4, can add later).
db.version(4).stores({
  projects:      '++id, jobNumber, name, createdAt, updatedAt',
  drawings:      '++id, projectId, sourcePdfId, pageNumber, sheetNumber, revision, [projectId+sheetNumber]',
  pdfSources:    '++id, filename, uploadedAt',
  inspections:   '++id, projectId, inspectionTypeId, date, status, createdAt, updatedAt',
  items:         '++id, inspectionId, itemNumber, severity, status, [inspectionId+itemNumber]',
  photos:        '++id, itemId, createdAt',
  highlights:    '++id, inspectionId, drawingId, [inspectionId+drawingId]',
  reports:       '++id, inspectionId, generatedAt, [inspectionId+generatedAt]',
  inspectionTypes:    '++id, &key, category, name',
  commentLibrary:     '++id, &key, inspectionTypeKey, severity',
  generalComments:    '++id, inspectionTypeKey',
  users:         '++id, &email, rpeq, role',
  settings:      '&key'
}).upgrade(async (tx) => {
  // Move each existing drawing's blob into a new pdfSources row.
  const drawings = await tx.drawings.toArray();
  for (const d of drawings) {
    if (!d.pdfBlob) continue;
    const sourceId = await tx.pdfSources.add({
      blob:       d.pdfBlob,
      filename:   d.filename || 'drawing.pdf',
      sizeBytes:  d.pdfBlob.size || 0,
      pageCount:  1,                     // best guess; viewer won't need this
      uploadedAt: d.uploadedAt || new Date().toISOString()
    });
    await tx.drawings.update(d.id, {
      sourcePdfId: sourceId,
      pageNumber:  1,
      // Keep pdfBlob as fallback during transition; subsequent reads prefer
      // sourcePdfId. Future cleanup pass can null these out.
    });
  }
});

/* --------------------------------------------------------------------------
   Seeding (idempotent)
   -------------------------------------------------------------------------- */

/** Sentinel value for library entries that apply to any inspection type. */
const ANY_TYPE = '_any_';

async function seedReferenceData() {
  const count = await db.inspectionTypes.count();
  if (count === 0) {
    const seed = [
      // Concrete / reinforcement
      { key: 'slab-prepour-ground',     category: 'Concrete', name: 'Slab on Ground Pre-Pour Reinforcement' },
      { key: 'slab-prepour-suspended',  category: 'Concrete', name: 'Suspended Slab Pre-Pour Reinforcement' },
      { key: 'pad-footing-prepour',     category: 'Concrete', name: 'Pad Footing Pre-Pour Reinforcement' },
      { key: 'strip-footing-prepour',   category: 'Concrete', name: 'Strip Footing Pre-Pour Reinforcement' },
      { key: 'raft-footing-prepour',    category: 'Concrete', name: 'Raft / Waffle Footing Pre-Pour Reinforcement' },
      { key: 'beam-prepour',            category: 'Concrete', name: 'Beam Pre-Pour Reinforcement' },
      { key: 'column-prepour',          category: 'Concrete', name: 'Column Pre-Pour Reinforcement' },
      { key: 'retaining-wall-prepour',  category: 'Concrete', name: 'Retaining Wall Pre-Pour Reinforcement' },
      { key: 'post-tension-strand',     category: 'Concrete', name: 'Post-Tension Strand Placement' },
      { key: 'concrete-pour-witness',   category: 'Concrete', name: 'Concrete Pour Witness' },

      // Masonry
      { key: 'blockwork-wall-reinf',    category: 'Masonry',  name: 'Blockwork Wall Vertical Reinforcement' },
      { key: 'blockwork-core-fill',     category: 'Masonry',  name: 'Blockwork Core-Fill Pre-Pour' },

      // Steel
      { key: 'steel-frame-erection',    category: 'Steel',    name: 'Steel Frame Erection' },
      { key: 'steel-connection',        category: 'Steel',    name: 'Steel Connection Inspection' },
      { key: 'baseplate-anchor',        category: 'Steel',    name: 'Baseplate / Anchor Bolt Inspection' },

      // Timber
      { key: 'timber-framing',          category: 'Timber',   name: 'Timber Framing Inspection' },
      { key: 'mass-timber-install',     category: 'Timber',   name: 'Mass Timber (GLT/CLT) Installation' },
      { key: 'bracing',                 category: 'Timber',   name: 'Bracing Inspection' },

      // Other
      { key: 'subgrade',                category: 'Other',    name: 'Subgrade / Bearing Inspection' },
      { key: 'waterproofing-precover',  category: 'Other',    name: 'Waterproofing Pre-Cover' },
      { key: 'adhoc',                   category: 'Other',    name: 'Ad-hoc (free form)' }
    ];
    await db.inspectionTypes.bulkAdd(seed);
  }

  await seedCommentLibrary();
  await seedGeneralComments();
}

/* ----- Comment library seed -----
 *
 * Idempotent per-key: any seed entry whose `key` is already in the table is
 * left alone. New entries added in a later app version get topped up on the
 * next load. User-edited entries are never overwritten.
 */
async function seedCommentLibrary() {
  const existing = await db.commentLibrary.toArray();
  const existingKeys = new Set(existing.map((e) => e.key));
  const toAdd = COMMENT_LIBRARY_SEED.filter((e) => !existingKeys.has(e.key));
  if (toAdd.length) await db.commentLibrary.bulkAdd(toAdd);
}

/** General-comments table has no unique key — dedupe by exact text match per type. */
async function seedGeneralComments() {
  const existing = await db.generalComments.toArray();
  const seen = new Set(existing.map((e) => `${e.inspectionTypeKey}::${e.text}`));
  const toAdd = GENERAL_COMMENTS_SEED.filter((e) =>
    !seen.has(`${e.inspectionTypeKey}::${e.text}`)
  );
  if (toAdd.length) await db.generalComments.bulkAdd(toAdd);
}

export async function initDb() {
  await db.open();
  await seedReferenceData();
  return db;
}

/* --------------------------------------------------------------------------
   Comment library seed content
   --------------------------------------------------------------------------
   Each entry:
     key                 unique stable identifier — safe to reference from
                          items via `libraryKey` for audit trail.
     inspectionTypeKey   specific type or '_any_' (ANY_TYPE).
     severity            'observation' | 'defect'
     category            short heading for grouping in the picker
     text                the comment body — edit freely after selection.
                         Placeholders like [X], [Y] are where the engineer
                         fills in measured values.
     asClause            the Australian Standard clause reference (if any).

   These seeds are deliberately conservative — real BT usage will diverge.
   Edits live in the seed list below; Phase 8 adds a Settings screen for
   in-app editing. Seed runs once on fresh install (guarded by a row count).
   -------------------------------------------------------------------------- */
const COMMENT_LIBRARY_SEED = [
  /* ------ Slab on Ground Pre-Pour (AS 3600, AS 2870) ------ */
  { key: 'slab-ground-rebar-conforms', inspectionTypeKey: 'slab-prepour-ground', severity: 'observation',
    category: 'General',
    text: 'Reinforcement is placed generally in accordance with the structural drawings.',
    asClause: 'AS 3600:2018 Cl. 8.1' },
  { key: 'slab-ground-cover-adequate', inspectionTypeKey: 'slab-prepour-ground', severity: 'observation',
    category: 'Cover',
    text: 'Concrete cover to bottom reinforcement confirmed at ≥ [X] mm — adequate for Exposure Class [B1].',
    asClause: 'AS 3600:2018 Cl. 4.10.3' },
  { key: 'slab-ground-chairs-ok', inspectionTypeKey: 'slab-prepour-ground', severity: 'observation',
    category: 'Supports',
    text: 'Bar supports (chairs) provided at spacings ≤ 1000 mm and hold reinforcement at correct cover.',
    asClause: 'AS 3600:2018 Cl. 17.5.3' },
  { key: 'slab-ground-vapour-barrier', inspectionTypeKey: 'slab-prepour-ground', severity: 'observation',
    category: 'Moisture',
    text: 'Vapour barrier continuous with laps ≥ 200 mm, taped and turned up at perimeter per AS 2870 Appendix E.',
    asClause: 'AS 2870:2011 Appendix E' },

  { key: 'slab-ground-cover-insufficient', inspectionTypeKey: 'slab-prepour-ground', severity: 'defect',
    category: 'Cover',
    text: 'Insufficient concrete cover — [X] mm provided where [Y] mm is required for Exposure Class [B1]. Lift / re-chair reinforcement before pour.',
    asClause: 'AS 3600:2018 Cl. 4.10.3' },
  { key: 'slab-ground-lap-short', inspectionTypeKey: 'slab-prepour-ground', severity: 'defect',
    category: 'Laps',
    text: 'Mesh lap of [X] mm is insufficient — provide minimum 500 mm or Lsy.tb per AS 3600:2018 Cl. 13.2.2.',
    asClause: 'AS 3600:2018 Cl. 13.2.2' },
  { key: 'slab-ground-chairs-missing', inspectionTypeKey: 'slab-prepour-ground', severity: 'defect',
    category: 'Supports',
    text: 'Bar supports missing / inadequate — mesh sags to sub-base. Install chairs at ≤ 1000 mm centres.',
    asClause: 'AS 3600:2018 Cl. 17.5.3' },
  { key: 'slab-ground-vapour-torn', inspectionTypeKey: 'slab-prepour-ground', severity: 'defect',
    category: 'Moisture',
    text: 'Vapour barrier torn at [location] — repair with taped patch before pour.',
    asClause: 'AS 2870:2011 Appendix E' },
  { key: 'slab-ground-termite', inspectionTypeKey: 'slab-prepour-ground', severity: 'defect',
    category: 'Termite',
    text: 'Termite management system not installed / terminated short of perimeter.',
    asClause: 'NCC Vol. 2 Part 3.1.3' },
  { key: 'slab-ground-debris', inspectionTypeKey: 'slab-prepour-ground', severity: 'defect',
    category: 'Site',
    text: 'Rebar off-cuts / tie-wire on sub-base — remove before pour.',
    asClause: '' },

  /* ------ Column Pre-Pour (AS 3600) ------ */
  { key: 'col-prepour-bars-ok', inspectionTypeKey: 'column-prepour', severity: 'observation',
    category: 'General',
    text: 'Vertical reinforcement per drawings, starter bars located per schedule.',
    asClause: 'AS 3600:2018 Cl. 10.7.1' },
  { key: 'col-prepour-ligs-ok', inspectionTypeKey: 'column-prepour', severity: 'observation',
    category: 'Ligatures',
    text: 'Ligatures at specified spacing and properly closed at standard hooks (135°).',
    asClause: 'AS 3600:2018 Cl. 10.7.3' },

  { key: 'col-prepour-cover-short', inspectionTypeKey: 'column-prepour', severity: 'defect',
    category: 'Cover',
    text: 'Insufficient cover — [X] mm to outermost ligature where [Y] mm required.',
    asClause: 'AS 3600:2018 Cl. 4.10.3' },
  { key: 'col-prepour-lig-spacing', inspectionTypeKey: 'column-prepour', severity: 'defect',
    category: 'Ligatures',
    text: 'Ligature spacing of [X] mm exceeds the maximum — reduce per AS 3600:2018 Cl. 10.7.3.',
    asClause: 'AS 3600:2018 Cl. 10.7.3' },
  { key: 'col-prepour-starter-offset', inspectionTypeKey: 'column-prepour', severity: 'defect',
    category: 'Starters',
    text: 'Starter bar location [X] mm from drawing position — straighten / cogg as directed by engineer.',
    asClause: '' },
  { key: 'col-prepour-lap-short', inspectionTypeKey: 'column-prepour', severity: 'defect',
    category: 'Laps',
    text: 'Vertical bar lap length [X] mm insufficient — provide [Y] mm per AS 3600:2018 Cl. 13.2.2.',
    asClause: 'AS 3600:2018 Cl. 13.2.2' },

  /* ------ Beam Pre-Pour (AS 3600) ------ */
  { key: 'beam-prepour-rebar-ok', inspectionTypeKey: 'beam-prepour', severity: 'observation',
    category: 'General',
    text: 'Top and bottom reinforcement arrangement per drawings.',
    asClause: 'AS 3600:2018 Cl. 8.1' },
  { key: 'beam-prepour-shear-ok', inspectionTypeKey: 'beam-prepour', severity: 'observation',
    category: 'Shear',
    text: 'Shear ligatures at support zones and midspan confirmed at specified spacing.',
    asClause: 'AS 3600:2018 Cl. 8.3.2' },

  { key: 'beam-prepour-cover-short', inspectionTypeKey: 'beam-prepour', severity: 'defect',
    category: 'Cover',
    text: 'Insufficient cover to soffit — [X] mm provided where [Y] mm required for Exposure Class [B1].',
    asClause: 'AS 3600:2018 Cl. 4.10.3' },
  { key: 'beam-prepour-shear-spacing', inspectionTypeKey: 'beam-prepour', severity: 'defect',
    category: 'Shear',
    text: 'Ligature spacing in support zone exceeds [X] mm — reduce per AS 3600:2018 Cl. 8.3.2.',
    asClause: 'AS 3600:2018 Cl. 8.3.2' },
  { key: 'beam-prepour-splice-location', inspectionTypeKey: 'beam-prepour', severity: 'defect',
    category: 'Laps',
    text: 'Top bar splice positioned in a high-moment zone — relocate to low-stress region.',
    asClause: 'AS 3600:2018 Cl. 13.2.2' },

  /* ------ Pad / Strip / Raft Footing Pre-Pour (AS 3600, AS 2870) ------ */
  { key: 'footing-cover-ground-ok', inspectionTypeKey: 'pad-footing-prepour', severity: 'observation',
    category: 'Cover',
    text: 'Cover to ground confirmed ≥ 65 mm — cage seated on chairs, not on sub-base.',
    asClause: 'AS 3600:2018 Cl. 4.10.3.5' },
  { key: 'footing-cover-ground-ok-strip', inspectionTypeKey: 'strip-footing-prepour', severity: 'observation',
    category: 'Cover',
    text: 'Cover to ground confirmed ≥ 65 mm — cage seated on chairs, not on sub-base.',
    asClause: 'AS 3600:2018 Cl. 4.10.3.5' },
  { key: 'footing-cover-ground-ok-raft', inspectionTypeKey: 'raft-footing-prepour', severity: 'observation',
    category: 'Cover',
    text: 'Cover to ground confirmed ≥ 65 mm — cage seated on chairs, not on sub-base.',
    asClause: 'AS 3600:2018 Cl. 4.10.3.5' },

  { key: 'footing-cover-ground-short', inspectionTypeKey: 'pad-footing-prepour', severity: 'defect',
    category: 'Cover',
    text: 'Insufficient cover to soil face — [X] mm provided where 65 mm required for cast-against-ground.',
    asClause: 'AS 3600:2018 Cl. 4.10.3.5' },
  { key: 'footing-depth-short', inspectionTypeKey: 'pad-footing-prepour', severity: 'defect',
    category: 'Geometry',
    text: 'Cage depth shallower than drawings — confirm excavation / founding level before pour.',
    asClause: '' },
  { key: 'footing-starter-missing', inspectionTypeKey: 'pad-footing-prepour', severity: 'defect',
    category: 'Starters',
    text: 'Starter bars for wall / column missing from cage — install before pour.',
    asClause: '' },

  /* ------ Steel Erection + Connection (AS 4100) ------ */
  { key: 'steel-tolerance-ok', inspectionTypeKey: 'steel-frame-erection', severity: 'observation',
    category: 'Tolerance',
    text: 'Erection tolerances within AS 4100:2020 Cl. 15.3.1 — members plumb and straight.',
    asClause: 'AS 4100:2020 Cl. 15.3.1' },
  { key: 'steel-bolts-snug-tight', inspectionTypeKey: 'steel-connection', severity: 'observation',
    category: 'Bolting',
    text: 'Bolts tightened to full contact (snug-tight) at bearing-type connections.',
    asClause: 'AS 4100:2020 Cl. 15.2.5.1' },
  { key: 'steel-welds-visual-ok', inspectionTypeKey: 'steel-connection', severity: 'observation',
    category: 'Welds',
    text: 'Welds visually sound — no undercut, porosity or lack of fusion observed.',
    asClause: 'AS/NZS 1554.1' },

  { key: 'steel-bolt-missing', inspectionTypeKey: 'steel-connection', severity: 'defect',
    category: 'Bolting',
    text: 'Bolt missing at connection — install and tighten to specified procedure.',
    asClause: 'AS 4100:2020 Cl. 9.3' },
  { key: 'steel-tb-cert-required', inspectionTypeKey: 'steel-connection', severity: 'defect',
    category: 'Bolting',
    text: 'Tension-bearing (TB) / tension-friction (TF) bolt tensioning not verified — provide inspection certificate or part-turn witness.',
    asClause: 'AS 4100:2020 Cl. 15.2.5.2' },
  { key: 'steel-weld-ndt', inspectionTypeKey: 'steel-connection', severity: 'defect',
    category: 'Welds',
    text: 'Weld discontinuity observed — non-destructive testing (MT / UT) required per AS/NZS 1554.1 Table 6.2.2.',
    asClause: 'AS/NZS 1554.1 Table 6.2.2' },
  { key: 'steel-member-plumb', inspectionTypeKey: 'steel-frame-erection', severity: 'defect',
    category: 'Tolerance',
    text: 'Member out of plumb by [X] mm over [Y] m height — exceeds AS 4100:2020 Cl. 15.3.1. Re-plumb before next stage.',
    asClause: 'AS 4100:2020 Cl. 15.3.1' },
  { key: 'steel-baseplate-grout', inspectionTypeKey: 'baseplate-anchor', severity: 'defect',
    category: 'Grout',
    text: 'Baseplate not grouted — install non-shrink grout per drawings and manufacturer instructions.',
    asClause: 'AS 4100:2020 Cl. 15.5' },

  /* ------ Light Timber Framing (AS 1684) ------ */
  { key: 'timber-frame-general', inspectionTypeKey: 'timber-framing', severity: 'observation',
    category: 'General',
    text: 'Framing generally in accordance with AS 1684.2 span tables / structural design.',
    asClause: 'AS 1684.2' },
  { key: 'timber-tiedown-ok', inspectionTypeKey: 'timber-framing', severity: 'observation',
    category: 'Tie-down',
    text: 'Tie-down connections match the tie-down schedule for the applicable wind classification.',
    asClause: 'AS 1684.2 Section 9' },

  { key: 'timber-plate-fixing-short', inspectionTypeKey: 'timber-framing', severity: 'defect',
    category: 'Tie-down',
    text: 'Wall plate fixing below AS 1684.2 requirements for wind classification — upgrade to specified tie-down.',
    asClause: 'AS 1684.2 Section 9' },
  { key: 'timber-brace-insufficient', inspectionTypeKey: 'timber-framing', severity: 'defect',
    category: 'Bracing',
    text: 'Bracing does not meet AS 1684.2 Appendix [A/B/C] requirements — provide additional bracing per design.',
    asClause: 'AS 1684.2 Appendix' },
  { key: 'timber-stud-spacing', inspectionTypeKey: 'timber-framing', severity: 'defect',
    category: 'Geometry',
    text: 'Stud spacing exceeds AS 1684.2 Table 7.3 for this wall type.',
    asClause: 'AS 1684.2 Table 7.3' },

  /* ------ Mass Timber (AS 1720.1) ------ */
  { key: 'mt-glt-install-ok', inspectionTypeKey: 'mass-timber-install', severity: 'observation',
    category: 'General',
    text: 'GLT beams installed per shop drawings — orientation, camber and bearing confirmed.',
    asClause: '' },
  { key: 'mt-clt-fixing-ok', inspectionTypeKey: 'mass-timber-install', severity: 'observation',
    category: 'Fixings',
    text: 'CLT panels fixed with specified screws at required pattern.',
    asClause: 'AS 1720.1:2010 Cl. 4.3' },
  { key: 'mt-moisture-ok', inspectionTypeKey: 'mass-timber-install', severity: 'observation',
    category: 'Moisture',
    text: 'Moisture content measured ≤ 15% — within AS 1720.1:2010 Cl. 2.4.3 assumptions.',
    asClause: 'AS 1720.1:2010 Cl. 2.4.3' },

  { key: 'mt-bolt-tension', inspectionTypeKey: 'mass-timber-install', severity: 'defect',
    category: 'Connections',
    text: 'Connection bolt tightening below AS 1720.1:2010 Cl. 4.4.2 — re-tension after service moisture is reached.',
    asClause: 'AS 1720.1:2010 Cl. 4.4.2' },
  { key: 'mt-clt-screw-spacing', inspectionTypeKey: 'mass-timber-install', severity: 'defect',
    category: 'Fixings',
    text: 'CLT fix-off screw spacing [X] mm exceeds specification — add intermediate fixings.',
    asClause: '' },
  { key: 'mt-moisture-high', inspectionTypeKey: 'mass-timber-install', severity: 'defect',
    category: 'Moisture',
    text: 'Moisture content [X]% exceeds the AS 1720.1:2010 Cl. 2.4.3 threshold of 15% — confirm design assumption and protect from further wetting.',
    asClause: 'AS 1720.1:2010 Cl. 2.4.3' },
  { key: 'mt-clt-split', inspectionTypeKey: 'mass-timber-install', severity: 'defect',
    category: 'Member',
    text: 'Split observed in CLT panel at [location] — requires engineer\u2019s assessment before cover-up.',
    asClause: '' },
  { key: 'mt-glt-damage', inspectionTypeKey: 'mass-timber-install', severity: 'defect',
    category: 'Member',
    text: 'GLT member damaged during handling — photograph and forward for assessment.',
    asClause: '' },

  /* ------ Masonry (AS 3700) ------ */
  { key: 'masonry-vert-bar-ok', inspectionTypeKey: 'blockwork-wall-reinf', severity: 'observation',
    category: 'Reinforcement',
    text: 'Vertical bars located in correct cores per drawings.',
    asClause: 'AS 3700:2018 Cl. 4.11' },
  { key: 'masonry-bar-wrong-core', inspectionTypeKey: 'blockwork-wall-reinf', severity: 'defect',
    category: 'Reinforcement',
    text: 'Vertical bar not in correct core — relocate before core-fill.',
    asClause: 'AS 3700:2018 Cl. 4.11' },
  { key: 'masonry-core-clean', inspectionTypeKey: 'blockwork-core-fill', severity: 'observation',
    category: 'Cleanliness',
    text: 'Cores clean, free of mortar droppings and debris — ready for core-fill.',
    asClause: 'AS 3700:2018 Cl. 11.4' },
  { key: 'masonry-core-dirty', inspectionTypeKey: 'blockwork-core-fill', severity: 'defect',
    category: 'Cleanliness',
    text: 'Mortar droppings in cores — clean through clean-out openings before fill.',
    asClause: 'AS 3700:2018 Cl. 11.4' },

  /* ------ Subgrade ------ */
  { key: 'subgrade-ok', inspectionTypeKey: 'subgrade', severity: 'observation',
    category: 'General',
    text: 'Subgrade appears consistent with the geotechnical report and drawings.',
    asClause: '' },
  { key: 'subgrade-soft', inspectionTypeKey: 'subgrade', severity: 'defect',
    category: 'Condition',
    text: 'Soft spot at [location] — remove and replace with compacted engineered fill / confirm founding level.',
    asClause: '' },
  { key: 'subgrade-water', inspectionTypeKey: 'subgrade', severity: 'defect',
    category: 'Condition',
    text: 'Groundwater present at founding level — dewater and confirm before pour.',
    asClause: '' },

  /* ------ Waterproofing Pre-Cover ------ */
  { key: 'wp-membrane-ok', inspectionTypeKey: 'waterproofing-precover', severity: 'observation',
    category: 'Membrane',
    text: 'Waterproofing membrane continuous with turn-ups at walls ≥ 150 mm.',
    asClause: 'AS 4654.2' },
  { key: 'wp-membrane-puncture', inspectionTypeKey: 'waterproofing-precover', severity: 'defect',
    category: 'Membrane',
    text: 'Membrane punctured at [location] — repair before backfill / cover.',
    asClause: 'AS 4654.2' },

  /* ------ Concrete Pour Witness ------ */
  { key: 'pour-grade-ok', inspectionTypeKey: 'concrete-pour-witness', severity: 'observation',
    category: 'Mix',
    text: 'Concrete grade and slump confirmed against specification on the delivery docket.',
    asClause: 'AS 3600:2018 Cl. 17.1' },
  { key: 'pour-compaction-ok', inspectionTypeKey: 'concrete-pour-witness', severity: 'observation',
    category: 'Placement',
    text: 'Placement method and compaction (immersion vibrator) acceptable.',
    asClause: 'AS 3600:2018 Cl. 17.1.6' },
  { key: 'pour-slump-reject', inspectionTypeKey: 'concrete-pour-witness', severity: 'defect',
    category: 'Mix',
    text: 'Slump outside tolerance — load rejected.',
    asClause: 'AS 1379 Cl. 6.3' },
  { key: 'pour-age-check', inspectionTypeKey: 'concrete-pour-witness', severity: 'defect',
    category: 'Mix',
    text: 'Concrete age on arrival exceeds 90 minutes — verify admixture retention before placement.',
    asClause: 'AS 1379 Cl. 6.4' },

  /* ------ General / any inspection type ------ */
  { key: 'any-limited-access', inspectionTypeKey: ANY_TYPE, severity: 'observation',
    category: 'Site',
    text: 'Access to [location] was limited — findings are based on visual inspection only from accessible areas.',
    asClause: '' },
  { key: 'any-weather-dry', inspectionTypeKey: ANY_TYPE, severity: 'observation',
    category: 'Site',
    text: 'Weather conditions at time of inspection were dry and did not affect findings.',
    asClause: '' },
  { key: 'any-weather-wet', inspectionTypeKey: ANY_TYPE, severity: 'observation',
    category: 'Site',
    text: 'Inspection conducted after recent rainfall — ponding / saturated conditions noted in [location].',
    asClause: '' },
  { key: 'any-refer-drawings', inspectionTypeKey: ANY_TYPE, severity: 'defect',
    category: 'Documentation',
    text: 'Refer to structural drawings for clarification / resolution — non-conformance noted.',
    asClause: '' },
  { key: 'any-eng-rfi', inspectionTypeKey: ANY_TYPE, severity: 'defect',
    category: 'Documentation',
    text: 'Requires engineer\u2019s response — RFI to be raised before works proceed at [location].',
    asClause: '' },
  { key: 'any-do-not-proceed', inspectionTypeKey: ANY_TYPE, severity: 'holdpoint',
    category: 'Hold Point',
    text: 'HOLD POINT — works are not to proceed at [location] until the listed items are rectified and re-inspected.',
    asClause: '' },

  /* ------ Suspended Slab Pre-Pour (AS 3600) — previously empty ------ */
  { key: 'slab-susp-rebar-conforms', inspectionTypeKey: 'slab-prepour-suspended', severity: 'observation',
    category: 'General',
    text: 'Top and bottom reinforcement arrangement generally in accordance with structural drawings.',
    asClause: 'AS 3600:2018 Cl. 8.1' },
  { key: 'slab-susp-cover-top-ok', inspectionTypeKey: 'slab-prepour-suspended', severity: 'observation',
    category: 'Cover',
    text: 'Top cover confirmed at \u2265 [X] mm \u2014 adequate for Exposure Class [B1] per AS 3600:2018 Cl. 4.10.3.',
    asClause: 'AS 3600:2018 Cl. 4.10.3' },
  { key: 'slab-susp-cover-bot-ok', inspectionTypeKey: 'slab-prepour-suspended', severity: 'observation',
    category: 'Cover',
    text: 'Bottom cover confirmed at \u2265 [X] mm \u2014 adequate for Exposure Class [B1].',
    asClause: 'AS 3600:2018 Cl. 4.10.3' },
  { key: 'slab-susp-chairs-ok', inspectionTypeKey: 'slab-prepour-suspended', severity: 'observation',
    category: 'Supports',
    text: 'Bar chairs installed at \u2264 1000 mm centres \u2014 reinforcement held at correct cover.',
    asClause: 'AS 3600:2018 Cl. 17.5.3' },
  { key: 'slab-susp-laps-ok', inspectionTypeKey: 'slab-prepour-suspended', severity: 'observation',
    category: 'Laps',
    text: 'Lap lengths comply with drawings and AS 3600:2018 Cl. 13.2.2.',
    asClause: 'AS 3600:2018 Cl. 13.2.2' },
  { key: 'slab-susp-trimmers-ok', inspectionTypeKey: 'slab-prepour-suspended', severity: 'observation',
    category: 'Openings',
    text: 'Trimmer bars installed around penetrations per drawings.',
    asClause: 'AS 3600:2018 Cl. 9.1.3.4' },
  { key: 'slab-susp-cover-top-short', inspectionTypeKey: 'slab-prepour-suspended', severity: 'defect',
    category: 'Cover',
    text: 'Insufficient top cover \u2014 [X] mm provided where [Y] mm is required. Adjust chairs / re-position mesh before pour.',
    asClause: 'AS 3600:2018 Cl. 4.10.3' },
  { key: 'slab-susp-cover-bot-short', inspectionTypeKey: 'slab-prepour-suspended', severity: 'defect',
    category: 'Cover',
    text: 'Insufficient bottom cover to soffit \u2014 [X] mm provided where [Y] mm required.',
    asClause: 'AS 3600:2018 Cl. 4.10.3' },
  { key: 'slab-susp-chairs-missing', inspectionTypeKey: 'slab-prepour-suspended', severity: 'defect',
    category: 'Supports',
    text: 'Bar chairs missing / inadequate \u2014 reinforcement sagging onto formwork. Install chairs at \u2264 1000 mm centres.',
    asClause: 'AS 3600:2018 Cl. 17.5.3' },
  { key: 'slab-susp-lap-short', inspectionTypeKey: 'slab-prepour-suspended', severity: 'defect',
    category: 'Laps',
    text: 'Lap length [X] mm is short of the AS 3600 Cl. 13.2.2 requirement. Extend bars or provide an engineer-approved splice.',
    asClause: 'AS 3600:2018 Cl. 13.2.2' },
  { key: 'slab-susp-trimmer-missing', inspectionTypeKey: 'slab-prepour-suspended', severity: 'defect',
    category: 'Openings',
    text: 'Trimmer bars missing around penetration at [location] \u2014 install per drawings before pour.',
    asClause: 'AS 3600:2018 Cl. 9.1.3.4' },
  { key: 'slab-susp-reo-wrong', inspectionTypeKey: 'slab-prepour-suspended', severity: 'defect',
    category: 'Reinforcement',
    text: 'Reinforcement size / spacing at [location] does not match drawings \u2014 rectify before pour.',
    asClause: 'AS 3600:2018 Cl. 8.1' },
  { key: 'slab-susp-form-debris', inspectionTypeKey: 'slab-prepour-suspended', severity: 'defect',
    category: 'Site',
    text: 'Debris, tie-wire and rebar off-cuts on formwork \u2014 remove before concrete placement.',
    asClause: '' },
  { key: 'slab-susp-penetration-unsupported', inspectionTypeKey: 'slab-prepour-suspended', severity: 'defect',
    category: 'Openings',
    text: 'Large penetration unsupported \u2014 confirm trimming steel and edge support per engineer\u2019s details.',
    asClause: '' },
  { key: 'slab-susp-pt-duct-damaged', inspectionTypeKey: 'slab-prepour-suspended', severity: 'holdpoint',
    category: 'Hold Point',
    text: 'HOLD POINT \u2014 PT duct damaged / grout leakage observed at [location]. Do not pour until rectified and re-inspected.',
    asClause: 'AS 3600:2018 Cl. 21' },

  /* ------ Slab on Ground Pre-Pour \u2014 additional entries ------ */
  { key: 'slab-ground-spacing-ok', inspectionTypeKey: 'slab-prepour-ground', severity: 'observation',
    category: 'Reinforcement',
    text: 'Mesh / reinforcement spacing checked and conforms to drawings.',
    asClause: 'AS 3600:2018 Cl. 8.1' },
  { key: 'slab-ground-stripfoot-integrated', inspectionTypeKey: 'slab-prepour-ground', severity: 'observation',
    category: 'Footings',
    text: 'Edge and internal beams tied into slab reinforcement per AS 2870 footing details.',
    asClause: 'AS 2870:2011 Section 5' },
  { key: 'slab-ground-setout-ok', inspectionTypeKey: 'slab-prepour-ground', severity: 'observation',
    category: 'Geometry',
    text: 'Slab setout and edge dimensions match architectural / structural drawings.',
    asClause: '' },
  { key: 'slab-ground-bearing-short', inspectionTypeKey: 'slab-prepour-ground', severity: 'defect',
    category: 'Geometry',
    text: 'Edge beam bearing onto foundation at [location] is less than drawings indicate \u2014 confirm founding.',
    asClause: 'AS 2870:2011 Section 5' },
  { key: 'slab-ground-slab-thin', inspectionTypeKey: 'slab-prepour-ground', severity: 'defect',
    category: 'Geometry',
    text: 'Slab depth at [location] measured [X] mm \u2014 below drawing requirement of [Y] mm.',
    asClause: '' },
  { key: 'slab-ground-pour-over-debris', inspectionTypeKey: 'slab-prepour-ground', severity: 'holdpoint',
    category: 'Hold Point',
    text: 'HOLD POINT \u2014 excessive formwork / sub-base contamination. Pour not to proceed until cleaned and re-inspected.',
    asClause: '' },

  /* ------ Raft / Waffle Footing ------ */
  { key: 'raft-pods-level', inspectionTypeKey: 'raft-footing-prepour', severity: 'observation',
    category: 'Waffle pods',
    text: 'Waffle pods placed level, supported on compacted sand, with no voids beneath.',
    asClause: 'AS 2870:2011 Section 5' },
  { key: 'raft-rib-rebar-ok', inspectionTypeKey: 'raft-footing-prepour', severity: 'observation',
    category: 'Reinforcement',
    text: 'Rib and slab reinforcement arrangement conforms to drawings \u2014 all corners and T-junctions adequately tied.',
    asClause: 'AS 2870:2011 Section 5' },
  { key: 'raft-pods-sunk', inspectionTypeKey: 'raft-footing-prepour', severity: 'defect',
    category: 'Waffle pods',
    text: 'Waffle pods sunk / tilted at [location] \u2014 re-seat before pour and confirm rib depth.',
    asClause: '' },
  { key: 'raft-articulation-missing', inspectionTypeKey: 'raft-footing-prepour', severity: 'defect',
    category: 'Articulation',
    text: 'Articulation joint at [location] not provided per AS 2870 Section 6 / drawings.',
    asClause: 'AS 2870:2011 Section 6' },

  /* ------ Pad & Strip Footings ------ */
  { key: 'footing-founding-ok', inspectionTypeKey: 'pad-footing-prepour', severity: 'observation',
    category: 'Founding',
    text: 'Founding level / bearing material consistent with geotechnical report \u2014 [X] kPa assumed.',
    asClause: 'AS 2870:2011 / AS 5100' },
  { key: 'footing-founding-soft', inspectionTypeKey: 'pad-footing-prepour', severity: 'defect',
    category: 'Founding',
    text: 'Soft / disturbed material at base of excavation at [location] \u2014 over-excavate and replace with engineered fill.',
    asClause: '' },
  { key: 'footing-excavation-dimensions', inspectionTypeKey: 'strip-footing-prepour', severity: 'defect',
    category: 'Geometry',
    text: 'Excavation width / depth at [location] less than drawing dimensions \u2014 re-excavate before pour.',
    asClause: '' },
  { key: 'footing-hold-unfavorable-ground', inspectionTypeKey: 'pad-footing-prepour', severity: 'holdpoint',
    category: 'Hold Point',
    text: 'HOLD POINT \u2014 ground conditions at [location] differ from geotech assumptions. Pour not to proceed until engineer has reviewed.',
    asClause: '' },

  /* ------ Column Pre-Pour \u2014 additional ------ */
  { key: 'col-prepour-cover-ok', inspectionTypeKey: 'column-prepour', severity: 'observation',
    category: 'Cover',
    text: 'Cover to ligatures confirmed \u2265 [X] mm on all faces.',
    asClause: 'AS 3600:2018 Cl. 4.10.3' },
  { key: 'col-prepour-tie-bent', inspectionTypeKey: 'column-prepour', severity: 'defect',
    category: 'Ligatures',
    text: 'Ligature not closed at standard 135\u00b0 hook at [location] \u2014 retie before pour.',
    asClause: 'AS 3600:2018 Cl. 10.7.3' },
  { key: 'col-prepour-plumb', inspectionTypeKey: 'column-prepour', severity: 'defect',
    category: 'Geometry',
    text: 'Column cage plumb tolerance exceeds AS 3600 acceptance \u2014 re-plumb before pour.',
    asClause: 'AS 3600:2018 Cl. 17.5' },

  /* ------ Beam Pre-Pour \u2014 additional ------ */
  { key: 'beam-prepour-bars-tied', inspectionTypeKey: 'beam-prepour', severity: 'observation',
    category: 'General',
    text: 'All bar intersections tied \u2014 no loose bars observed.',
    asClause: '' },
  { key: 'beam-prepour-cover-side', inspectionTypeKey: 'beam-prepour', severity: 'defect',
    category: 'Cover',
    text: 'Side cover to stirrups at [location] measured [X] mm \u2014 below [Y] mm required.',
    asClause: 'AS 3600:2018 Cl. 4.10.3' },
  { key: 'beam-prepour-stirrup-missing', inspectionTypeKey: 'beam-prepour', severity: 'defect',
    category: 'Shear',
    text: 'Stirrup missing at [location] \u2014 install before pour.',
    asClause: 'AS 3600:2018 Cl. 8.3' },

  /* ------ Retaining Wall Pre-Pour ------ */
  { key: 'rw-rebar-ok', inspectionTypeKey: 'retaining-wall-prepour', severity: 'observation',
    category: 'Reinforcement',
    text: 'Vertical and horizontal reinforcement installed per drawings; cover to soil face confirmed.',
    asClause: 'AS 3600:2018 Cl. 4.10.3' },
  { key: 'rw-drain-ok', inspectionTypeKey: 'retaining-wall-prepour', severity: 'observation',
    category: 'Drainage',
    text: 'Subsoil drain and weep holes detailed and installed per drawings.',
    asClause: '' },
  { key: 'rw-soil-cover-short', inspectionTypeKey: 'retaining-wall-prepour', severity: 'defect',
    category: 'Cover',
    text: 'Insufficient cover to soil-retaining face at [location] \u2014 adjust chairs before pour.',
    asClause: 'AS 3600:2018 Cl. 4.10.3.5' },
  { key: 'rw-starter-offset', inspectionTypeKey: 'retaining-wall-prepour', severity: 'defect',
    category: 'Starters',
    text: 'Starter bars out of position at [location] \u2014 straighten or supplement per engineer\u2019s detail.',
    asClause: '' },

  /* ------ Post-Tension Strand Placement ------ */
  { key: 'pt-strand-layout-ok', inspectionTypeKey: 'post-tension-strand', severity: 'observation',
    category: 'Layout',
    text: 'Strand profile and spacing conforms to shop drawings; anchorage zones clear of conflicts.',
    asClause: 'AS 3600:2018 Cl. 21' },
  { key: 'pt-stressing-pocket-clean', inspectionTypeKey: 'post-tension-strand', severity: 'observation',
    category: 'Anchorage',
    text: 'Stressing pockets clean, anchorage bursting reinforcement installed per drawings.',
    asClause: 'AS 3600:2018 Cl. 21' },
  { key: 'pt-duct-damage', inspectionTypeKey: 'post-tension-strand', severity: 'holdpoint',
    category: 'Hold Point',
    text: 'HOLD POINT \u2014 PT duct crushed / split at [location]. Stress not to proceed until engineer has approved repair.',
    asClause: 'AS 3600:2018 Cl. 21' },

  /* ------ Blockwork / Core fill \u2014 additional ------ */
  { key: 'masonry-lap-ok', inspectionTypeKey: 'blockwork-wall-reinf', severity: 'observation',
    category: 'Reinforcement',
    text: 'Vertical bar laps comply with AS 3700 requirements.',
    asClause: 'AS 3700:2018 Cl. 8.8' },
  { key: 'masonry-cavity-ok', inspectionTypeKey: 'blockwork-core-fill', severity: 'observation',
    category: 'Cleanliness',
    text: 'Cores inspected via clean-out opening \u2014 free of excess mortar, ready for core-fill.',
    asClause: 'AS 3700:2018 Cl. 11.4' },
  { key: 'masonry-bar-short', inspectionTypeKey: 'blockwork-wall-reinf', severity: 'defect',
    category: 'Reinforcement',
    text: 'Vertical bar at [location] stops short \u2014 lap into next lift or re-tie before core-fill.',
    asClause: 'AS 3700:2018 Cl. 8.8' },

  /* ------ Steel connection \u2014 additional ------ */
  { key: 'steel-weld-undercut', inspectionTypeKey: 'steel-connection', severity: 'defect',
    category: 'Welds',
    text: 'Weld undercut observed at [location] \u2014 rework per AS/NZS 1554.1 acceptance criteria.',
    asClause: 'AS/NZS 1554.1 Cl. 6.2.2' },
  { key: 'steel-bolt-not-tensioned', inspectionTypeKey: 'steel-connection', severity: 'holdpoint',
    category: 'Hold Point',
    text: 'HOLD POINT \u2014 tension-critical bolts at [location] not tightened to specified method. Next stage not to proceed until rectified and verified.',
    asClause: 'AS 4100:2020 Cl. 15.2.5.2' },

  /* ------ Timber framing \u2014 additional ------ */
  { key: 'timber-anchor-missing', inspectionTypeKey: 'timber-framing', severity: 'defect',
    category: 'Tie-down',
    text: 'Tie-down strap missing at [location] \u2014 install per drawings.',
    asClause: 'AS 1684.2 Section 9' },
  { key: 'timber-rot', inspectionTypeKey: 'timber-framing', severity: 'defect',
    category: 'Member',
    text: 'Timber member at [location] shows rot / insect damage \u2014 replace before cladding.',
    asClause: '' },

  /* ------ Subgrade / Bearing \u2014 additional ------ */
  { key: 'subgrade-proof-roll', inspectionTypeKey: 'subgrade', severity: 'observation',
    category: 'Condition',
    text: 'Proof-rolling witnessed \u2014 no soft spots detected under loaded vehicle.',
    asClause: '' },
  { key: 'subgrade-not-approved', inspectionTypeKey: 'subgrade', severity: 'holdpoint',
    category: 'Hold Point',
    text: 'HOLD POINT \u2014 bearing surface does not meet geotechnical assumptions. Footing pour not to proceed until engineer approves.',
    asClause: '' }
];

const GENERAL_COMMENTS_SEED = [
  // Universal — applicable to any inspection type
  { inspectionTypeKey: ANY_TYPE,
    text: 'The works inspected are, in general, in accordance with the structural drawings.' },
  { inspectionTypeKey: ANY_TYPE,
    text: 'This inspection is limited to the areas marked on the attached plan.' },
  { inspectionTypeKey: ANY_TYPE,
    text: 'No destructive testing was undertaken.' },
  { inspectionTypeKey: ANY_TYPE,
    text: 'This inspection does not relieve the Contractor of their obligations under the Contract or the relevant Australian Standards.' },
  { inspectionTypeKey: ANY_TYPE,
    text: 'Items listed herein require rectification before the next stage of works proceeds. Re-inspection may be required.' },
  { inspectionTypeKey: ANY_TYPE,
    text: 'Photographs taken during the inspection are attached and annotated.' },

  // Concrete pre-pour variants
  { inspectionTypeKey: 'slab-prepour-ground',
    text: 'Pre-pour inspection confirms that the reinforcement detailed is in place and ready for concrete placement, subject to rectification of the items listed above.' },
  { inspectionTypeKey: 'slab-prepour-suspended',
    text: 'Pre-pour inspection confirms that the reinforcement detailed is in place and ready for concrete placement, subject to rectification of the items listed above.' },
  { inspectionTypeKey: 'slab-prepour-suspended',
    text: "Contractor to provide minimum concrete strength f'c = [X] MPa and ensure adequate curing per AS 3600:2018 Cl. 17.1." },
  { inspectionTypeKey: 'column-prepour',
    text: 'Column reinforcement is ready for concrete placement, subject to rectification of the items listed above.' },
  { inspectionTypeKey: 'beam-prepour',
    text: 'Beam reinforcement is ready for concrete placement, subject to rectification of the items listed above.' },
  { inspectionTypeKey: 'concrete-pour-witness',
    text: 'Concrete was placed and compacted in accordance with AS 3600:2018 Cl. 17.1.' },

  // Steel
  { inspectionTypeKey: 'steel-frame-erection',
    text: 'Erection is at the stage where primary framing is plumb and laterally stable, subject to rectification of the items listed above.' },

  // Timber
  { inspectionTypeKey: 'timber-framing',
    text: 'All visible timber members are straight and free of apparent structural defects.' },
  { inspectionTypeKey: 'mass-timber-install',
    text: 'Moisture content was within acceptable limits (≤ 15%) at time of inspection. Re-tensioning of bolted connections at service moisture is recommended per AS 1720.1:2010 Cl. 4.4.2.' },

  // Retaining wall / raft / footings
  { inspectionTypeKey: 'retaining-wall-prepour',
    text: 'Wall reinforcement and formwork are ready for concrete placement, subject to rectification of the items listed above and provision of drainage per the drawings.' },
  { inspectionTypeKey: 'raft-footing-prepour',
    text: 'Raft reinforcement, waffle pods, and edge beams are ready for concrete placement, subject to rectification of the items listed above.' },
  { inspectionTypeKey: 'pad-footing-prepour',
    text: 'Footing excavation, reinforcement and starters are ready for concrete placement, subject to rectification of the items listed above.' },
  { inspectionTypeKey: 'strip-footing-prepour',
    text: 'Strip footing excavation, reinforcement and starters are ready for concrete placement, subject to rectification of the items listed above.' }
];

/* --------------------------------------------------------------------------
   Project helpers
   -------------------------------------------------------------------------- */

/**
 * Create a project. Returns the full project record with id.
 */
export async function createProject(data) {
  const now = new Date().toISOString();
  const record = {
    jobNumber:   (data.jobNumber || '').trim(),
    name:        (data.name || '').trim(),
    client:      (data.client || '').trim(),
    siteAddress: (data.siteAddress || '').trim(),
    notes:       (data.notes || '').trim(),
    createdAt:   now,
    updatedAt:   now
  };
  if (!record.jobNumber) throw new Error('Job number is required');
  if (!record.name)      throw new Error('Project name is required');
  const id = await db.projects.add(record);
  return { id, ...record };
}

export async function updateProject(id, patch) {
  const updatedAt = new Date().toISOString();
  const trimmed = {};
  for (const k of ['jobNumber', 'name', 'client', 'siteAddress', 'notes']) {
    if (patch[k] !== undefined) trimmed[k] = String(patch[k]).trim();
  }
  await db.projects.update(Number(id), { ...trimmed, updatedAt });
  return db.projects.get(Number(id));
}

/**
 * Cascade delete a project and everything under it.
 * Photos → Items → Highlights → Reports → Inspections → Drawings → Orphan PDFs → Project.
 */
export async function deleteProject(id) {
  const projectId = Number(id);
  await db.transaction('rw',
    [db.projects, db.drawings, db.pdfSources, db.inspections,
     db.items, db.photos, db.highlights, db.reports],
    async () => {
      const inspections = await db.inspections.where('projectId').equals(projectId).toArray();
      for (const insp of inspections) {
        const items = await db.items.where('inspectionId').equals(insp.id).toArray();
        for (const item of items) {
          await db.photos.where('itemId').equals(item.id).delete();
        }
        await db.items.where('inspectionId').equals(insp.id).delete();
        await db.highlights.where('inspectionId').equals(insp.id).delete();
        await db.reports.where('inspectionId').equals(insp.id).delete();
      }

      // Collect source IDs referenced by this project's drawings
      const projectDrawings = await db.drawings.where('projectId').equals(projectId).toArray();
      const sourceIds = new Set(projectDrawings.map((d) => d.sourcePdfId).filter(Boolean));

      await db.inspections.where('projectId').equals(projectId).delete();
      await db.drawings.where('projectId').equals(projectId).delete();

      // Reclaim any pdfSources no longer referenced by anything.
      for (const sid of sourceIds) {
        const stillReferenced = await db.drawings.where('sourcePdfId').equals(sid).count();
        if (stillReferenced === 0) {
          await db.pdfSources.delete(sid);
        }
      }

      await db.projects.delete(projectId);
    }
  );
}

export function getProject(id) {
  return db.projects.get(Number(id));
}

export function listProjects() {
  return db.projects.orderBy('updatedAt').reverse().toArray();
}

/* --------------------------------------------------------------------------
   Drawing helpers
   -------------------------------------------------------------------------- */

/**
 * Add a drawing to a project — single page, optionally backed by an existing
 * pdfSource. For multi-page PDFs use addDrawingsFromSource() instead.
 *
 * data = {
 *   sheetNumber, revision, description, filename,
 *   pdfBlob                  // new upload; will be moved into pdfSources
 *   sourcePdfId, pageNumber  // existing source; preferred for Phase 8+
 *   pageWidth, pageHeight
 * }
 *
 * Calibration is null until the user calibrates. Rotation defaults to 0.
 */
export async function addDrawing(projectId, data) {
  const now = new Date().toISOString();

  // Figure out which source this drawing references.
  let sourcePdfId = data.sourcePdfId ? Number(data.sourcePdfId) : null;
  if (!sourcePdfId && data.pdfBlob) {
    sourcePdfId = await db.pdfSources.add({
      blob:       data.pdfBlob,
      filename:   (data.filename || 'drawing.pdf').trim(),
      sizeBytes:  data.pdfBlob.size || 0,
      pageCount:  1,
      uploadedAt: now
    });
  }

  const record = {
    projectId:   Number(projectId),
    sourcePdfId,
    pageNumber:  Number(data.pageNumber || 1),
    sheetNumber: (data.sheetNumber || '').trim(),
    revision:    (data.revision || '').trim(),
    description: (data.description || '').trim(),
    filename:    (data.filename || '').trim(),
    pageWidth:   data.pageWidth || null,     // in PDF points (1pt = 1/72 in)
    pageHeight:  data.pageHeight || null,
    rotation:    0,                          // 0 | 90 | 180 | 270 — on top of intrinsic
    calibration: null,
    uploadedAt:  now,
    updatedAt:   now
  };
  const id = await db.drawings.add(record);
  await touchProject(record.projectId);
  return { id, ...record };
}

/**
 * Add an N-page PDF as N drawings sharing one pdfSources record.
 *
 * pages[] = [
 *   { pageNumber, sheetNumber?, revision?, description?, pageWidth?, pageHeight? }
 * ]
 *
 * Returns the array of created drawings.
 */
export async function addDrawingsFromSource(projectId, { pdfBlob, filename, pageCount }, pages) {
  if (!pdfBlob) throw new Error('PDF blob required');
  if (!Array.isArray(pages) || pages.length === 0) throw new Error('Pages list required');
  const now = new Date().toISOString();

  // Wrap in a transaction so a failure midway doesn't leave orphan state.
  // If the drawing inserts throw, Dexie rolls back the pdfSources insert too.
  return db.transaction('rw', [db.pdfSources, db.drawings, db.projects], async () => {
    const sourcePdfId = await db.pdfSources.add({
      blob:       pdfBlob,
      filename:   (filename || 'drawing.pdf').trim(),
      sizeBytes:  pdfBlob.size || 0,
      pageCount:  pageCount || pages.length,
      uploadedAt: now
    });

    const created = [];
    for (const p of pages) {
      const record = {
        projectId:   Number(projectId),
        sourcePdfId,
        pageNumber:  Number(p.pageNumber || 1),
        sheetNumber: (p.sheetNumber || '').trim(),
        revision:    (p.revision || '').trim(),
        description: (p.description || '').trim(),
        filename:    (filename || '').trim(),
        pageWidth:   p.pageWidth || null,
        pageHeight:  p.pageHeight || null,
        rotation:    0,
        calibration: null,
        uploadedAt:  now,
        updatedAt:   now
      };
      const id = await db.drawings.add(record);
      created.push({ id, ...record });
    }
    await db.projects.update(Number(projectId), { updatedAt: now });
    return created;
  });
}

/**
 * Fetch the Blob behind a drawing. Prefers the new pdfSources record but
 * falls back to a legacy drawing.pdfBlob (pre-v4 data that survived
 * migration without a source row for any reason).
 */
export async function getDrawingBlob(drawing) {
  if (drawing?.sourcePdfId) {
    const src = await db.pdfSources.get(Number(drawing.sourcePdfId));
    if (src?.blob) return src.blob;
  }
  // Legacy fallback
  return drawing?.pdfBlob || null;
}

export function getPdfSource(id) {
  return db.pdfSources.get(Number(id));
}

export async function updateDrawing(id, patch) {
  const allowed = ['sheetNumber', 'revision', 'description', 'rotation'];
  const record = {};
  for (const k of allowed) {
    if (patch[k] !== undefined) {
      record[k] = typeof patch[k] === 'string' ? patch[k].trim() : patch[k];
    }
  }
  record.updatedAt = new Date().toISOString();
  await db.drawings.update(Number(id), record);
  const d = await db.drawings.get(Number(id));
  if (d) await touchProject(d.projectId);
  return d;
}

/**
 * Store calibration result.
 *
 * calibration = {
 *   p1: { x, y },              // PDF coordinates (points), origin bottom-left per PDF spec
 *   p2: { x, y },
 *   realDistanceMm: number,    // user entered
 *   mmPerPoint: number,        // derived scale
 *   calibratedAt: ISO string
 * }
 */
export async function setDrawingCalibration(id, calibration) {
  const drawingId = Number(id);
  const record = {
    calibration,
    updatedAt: new Date().toISOString()
  };
  await db.drawings.update(drawingId, record);
  const d = await db.drawings.get(drawingId);
  if (d) await touchProject(d.projectId);
  return d;
}

/**
 * Store grid-intersection calibration — two known grid points on the plan
 * that let the app compute "Grid 3/B" style references for any pin drop.
 *
 * gridCalibration = {
 *   p1: { pdfX, pdfY, col, row },
 *   p2: { pdfX, pdfY, col, row },
 *   calibratedAt
 * }
 */
export async function setDrawingGridCalibration(id, gridCalibration) {
  const drawingId = Number(id);
  await db.drawings.update(drawingId, {
    gridCalibration,
    updatedAt: new Date().toISOString()
  });
  const d = await db.drawings.get(drawingId);
  if (d) await touchProject(d.projectId);
  return d;
}

export async function deleteDrawing(id) {
  const drawingId = Number(id);
  const d = await db.drawings.get(drawingId);
  if (!d) return;
  await db.drawings.delete(drawingId);
  // If no remaining drawings reference this source PDF, reclaim the blob.
  if (d.sourcePdfId) {
    const remaining = await db.drawings
      .where('sourcePdfId').equals(Number(d.sourcePdfId)).count();
    if (remaining === 0) {
      await db.pdfSources.delete(Number(d.sourcePdfId));
    }
  }
  await touchProject(d.projectId);
}

export function getDrawing(id) {
  return db.drawings.get(Number(id));
}

export function listDrawingsForProject(projectId) {
  return db.drawings
    .where('projectId').equals(Number(projectId))
    .toArray()
    .then((arr) =>
      // Sort by sheet number (numeric-aware) with page-number tie-break.
      // Drawings with no sheet number go to the bottom, ordered by page.
      arr.sort((a, b) => {
        const sa = (a.sheetNumber || '').trim();
        const sb = (b.sheetNumber || '').trim();
        if (sa && !sb) return -1;
        if (!sa && sb) return 1;
        if (sa && sb) {
          const cmp = sa.localeCompare(sb, undefined, { numeric: true });
          if (cmp !== 0) return cmp;
        }
        return (a.pageNumber || 0) - (b.pageNumber || 0);
      })
    );
}

async function touchProject(projectId) {
  await db.projects.update(Number(projectId), { updatedAt: new Date().toISOString() });
}

/* --------------------------------------------------------------------------
   Inspection type helpers
   -------------------------------------------------------------------------- */

export function listInspectionTypes() {
  return db.inspectionTypes.toArray();
}

export function getInspectionType(key) {
  return db.inspectionTypes.where('key').equals(String(key)).first();
}

/* --------------------------------------------------------------------------
   Inspection helpers
   -------------------------------------------------------------------------- */

/**
 * Create an inspection.
 *
 * data = {
 *   projectId, inspectionTypeKey, primaryDrawingId,
 *   date (ISO yyyy-mm-dd), inspectorName, notes
 * }
 *
 * The inspection type's display name and category are denormalised onto the
 * inspection record so the audit trail keeps its label even if the seed
 * list is later edited.
 */
export async function createInspection(data) {
  const now = new Date().toISOString();

  const projectId = Number(data.projectId);
  if (!projectId) throw new Error('Project is required');

  const type = await getInspectionType(data.inspectionTypeKey);
  if (!type) throw new Error('Inspection type is required');

  const primaryDrawingId = data.primaryDrawingId ? Number(data.primaryDrawingId) : null;
  if (!primaryDrawingId) throw new Error('Primary drawing is required');

  if (!data.date) throw new Error('Date is required');

  // Pre-populate the inspection's general comments with TYPE-SPECIFIC boilerplate
  // only — north star §7.2.5. Universal defensive statements ("limited to areas
  // inspected", "in accordance with the drawings") are already baked into the
  // scope paragraph on the PDF cover, so pre-populating them would double up.
  // The engineer can add universal ones via "Add from library" if they want.
  const typeSpecific = await db.generalComments
    .where('inspectionTypeKey').equals(type.key)
    .toArray();
  const generalComments = typeSpecific.map((g) => ({
    text:       g.text,
    libraryKey: null,
    addedAt:    now
  }));

  const record = {
    projectId,
    inspectionTypeId:       type.id,               // kept for schema-index compatibility
    inspectionTypeKey:      type.key,
    inspectionTypeName:     type.name,
    inspectionTypeCategory: type.category,
    primaryDrawingId,
    date:                   String(data.date),
    inspectorName:          (data.inspectorName || '').trim(),
    attendees:              (data.attendees || '').trim(),
    weather:                (data.weather || '').trim(),
    notes:                  (data.notes || '').trim(),
    generalComments,                               // prefilled from boilerplate
    status:                 'draft',
    createdAt:              now,
    updatedAt:              now
  };

  const id = await db.inspections.add(record);
  await touchProject(projectId);
  return { id, ...record };
}

export async function updateInspection(id, patch) {
  const inspectionId = Number(id);
  const record = {};
  const stringFields = ['inspectorName', 'attendees', 'weather', 'notes', 'status', 'date'];
  for (const k of stringFields) {
    if (patch[k] !== undefined) record[k] = String(patch[k]).trim();
  }
  if (patch.primaryDrawingId !== undefined) {
    record.primaryDrawingId = Number(patch.primaryDrawingId);
  }
  if (patch.inspectionTypeKey !== undefined) {
    const type = await getInspectionType(patch.inspectionTypeKey);
    if (type) {
      record.inspectionTypeId       = type.id;
      record.inspectionTypeKey      = type.key;
      record.inspectionTypeName     = type.name;
      record.inspectionTypeCategory = type.category;
    }
  }
  if (patch.generalComments !== undefined) {
    // Trust the caller to have shaped the array correctly.
    record.generalComments = Array.isArray(patch.generalComments) ? patch.generalComments : [];
  }
  record.updatedAt = new Date().toISOString();
  await db.inspections.update(inspectionId, record);
  const i = await db.inspections.get(inspectionId);
  if (i) await touchProject(i.projectId);
  return i;
}

/**
 * Convenience — append or replace the full generalComments array.
 *
 * comments[] = [{ text, libraryKey?, addedAt }]
 */
export async function setInspectionGeneralComments(id, comments) {
  return updateInspection(id, { generalComments: comments });
}

/**
 * Cascade delete an inspection and its items, photos, highlights, and reports.
 */
export async function deleteInspection(id) {
  const inspectionId = Number(id);
  await db.transaction('rw',
    [db.inspections, db.items, db.photos, db.highlights, db.reports],
    async () => {
      const items = await db.items.where('inspectionId').equals(inspectionId).toArray();
      for (const item of items) {
        await db.photos.where('itemId').equals(item.id).delete();
      }
      await db.items.where('inspectionId').equals(inspectionId).delete();
      await db.highlights.where('inspectionId').equals(inspectionId).delete();
      await db.reports.where('inspectionId').equals(inspectionId).delete();
      await db.inspections.delete(inspectionId);
    }
  );
}

export function getInspection(id) {
  return db.inspections.get(Number(id));
}

export function listInspectionsForProject(projectId) {
  return db.inspections
    .where('projectId').equals(Number(projectId))
    .reverse().sortBy('updatedAt');
}

export function listInspections() {
  return db.inspections.reverse().sortBy('updatedAt');
}

/* --------------------------------------------------------------------------
   Item helpers (Phase 4)
   -------------------------------------------------------------------------- */

/**
 * Create a numbered item (pin) against an inspection.
 *
 * data = {
 *   drawingId, page, pdfX, pdfY,   // pin location in PDF space
 *   comment, severity, status
 * }
 *
 * itemNumber auto-increments per inspection (1..n). Concurrency: because we
 * run inside a 'rw' transaction, the max-lookup + insert is atomic per tab.
 */
export async function createItem(inspectionId, data = {}) {
  const now = new Date().toISOString();
  const inspId = Number(inspectionId);
  if (!inspId) throw new Error('Inspection is required');

  return db.transaction('rw', [db.items, db.inspections], async () => {
    const existing = await db.items.where('inspectionId').equals(inspId).toArray();
    const nextNumber = existing.length
      ? Math.max(...existing.map((i) => i.itemNumber || 0)) + 1
      : 1;

    const record = {
      inspectionId: inspId,
      itemNumber:   nextNumber,
      drawingId:    data.drawingId ? Number(data.drawingId) : null,
      page:         data.page || 1,
      pdfX:         data.pdfX ?? null,
      pdfY:         data.pdfY ?? null,
      comment:      (data.comment || '').trim(),
      libraryKey:   data.libraryKey || null,        // set when comment came from the library
      asClause:     (data.asClause || '').trim(),   // denormalised so edits don't drift with library
      gridRef:      (data.gridRef || '').trim(),    // optional, e.g. "3/B" or "GL1"
      severity:     data.severity || 'observation', // 'observation' | 'defect' | 'holdpoint'
      status:       data.status   || 'open',        // 'open' | 'closed'
      createdAt:    now,
      updatedAt:    now
    };

    const id = await db.items.add(record);
    // Bump the inspection's updatedAt so the list reshuffles.
    await db.inspections.update(inspId, { updatedAt: now });
    return { id, ...record };
  });
}

export async function updateItem(id, patch) {
  const itemId = Number(id);
  const record = {};
  if (patch.comment    !== undefined) record.comment    = String(patch.comment).trim();
  if (patch.severity   !== undefined) record.severity   = String(patch.severity);
  if (patch.status     !== undefined) record.status     = String(patch.status);
  if (patch.pdfX       !== undefined) record.pdfX       = patch.pdfX;
  if (patch.pdfY       !== undefined) record.pdfY       = patch.pdfY;
  if (patch.libraryKey !== undefined) record.libraryKey = patch.libraryKey || null;
  if (patch.asClause   !== undefined) record.asClause   = String(patch.asClause || '').trim();
  if (patch.gridRef    !== undefined) record.gridRef    = String(patch.gridRef || '').trim();
  record.updatedAt = new Date().toISOString();

  await db.items.update(itemId, record);
  const item = await db.items.get(itemId);
  if (item) {
    await db.inspections.update(item.inspectionId, { updatedAt: record.updatedAt });
  }
  return item;
}

/**
 * Delete an item and its photos. Item numbers are NOT renumbered —
 * once issued, a number is permanent (so the report audit trail is stable).
 */
export async function deleteItem(id) {
  const itemId = Number(id);
  await db.transaction('rw', [db.items, db.photos, db.inspections], async () => {
    const item = await db.items.get(itemId);
    if (!item) return;
    await db.photos.where('itemId').equals(itemId).delete();
    await db.items.delete(itemId);
    await db.inspections.update(item.inspectionId, { updatedAt: new Date().toISOString() });
  });
}

export function getItem(id) {
  return db.items.get(Number(id));
}

export function listItemsForInspection(inspectionId) {
  return db.items
    .where('inspectionId').equals(Number(inspectionId))
    .sortBy('itemNumber');
}

/* --------------------------------------------------------------------------
   Photo helpers (Phase 4)
   -------------------------------------------------------------------------- */

/**
 * Add a photo to an item.
 *
 * data = {
 *   blob,          // Blob (JPEG — compression handled by caller)
 *   caption,       // optional string
 *   width, height  // pixels (for display layout)
 * }
 */
export async function addItemPhoto(itemId, data) {
  const now = new Date().toISOString();
  const record = {
    itemId:    Number(itemId),
    blob:      data.blob,
    caption:   (data.caption || '').trim(),
    width:     data.width || null,
    height:    data.height || null,
    createdAt: now
  };
  const id = await db.photos.add(record);
  // Touch parent item so inspection detail thumbnails refresh ordering.
  const item = await db.items.get(record.itemId);
  if (item) await db.items.update(record.itemId, { updatedAt: now });
  return { id, ...record };
}

export async function updateItemPhoto(id, patch) {
  const photoId = Number(id);
  const record = {};
  if (patch.caption !== undefined) record.caption = String(patch.caption).trim();
  if (patch.annotations !== undefined) {
    // Array of vector ops — stored verbatim. See lib/annotate.js for shape.
    record.annotations = Array.isArray(patch.annotations) ? patch.annotations : [];
  }
  await db.photos.update(photoId, record);
  return db.photos.get(photoId);
}

export async function deleteItemPhoto(id) {
  await db.photos.delete(Number(id));
}

export function getItemPhoto(id) {
  return db.photos.get(Number(id));
}

export function listPhotosForItem(itemId) {
  return db.photos
    .where('itemId').equals(Number(itemId))
    .sortBy('createdAt');
}

/* --------------------------------------------------------------------------
   Highlight helpers (Phase 4) — extent-of-inspection rectangles
   -------------------------------------------------------------------------- */

/**
 * Add a highlight rectangle. Coordinates are PDF-space (points).
 *
 * data = { drawingId, page, pdfX, pdfY, pdfW, pdfH }
 */
export async function addHighlight(inspectionId, data) {
  const now = new Date().toISOString();
  const record = {
    inspectionId: Number(inspectionId),
    drawingId:    Number(data.drawingId),
    page:         data.page || 1,
    pdfX:         data.pdfX,
    pdfY:         data.pdfY,
    pdfW:         data.pdfW,
    pdfH:         data.pdfH,
    createdAt:    now
  };
  const id = await db.highlights.add(record);
  await db.inspections.update(record.inspectionId, { updatedAt: now });
  return { id, ...record };
}

export async function deleteHighlight(id) {
  const hlId = Number(id);
  const h = await db.highlights.get(hlId);
  await db.highlights.delete(hlId);
  if (h) {
    await db.inspections.update(h.inspectionId, { updatedAt: new Date().toISOString() });
  }
}

export function listHighlightsForInspection(inspectionId) {
  return db.highlights
    .where('inspectionId').equals(Number(inspectionId))
    .toArray();
}

export function listHighlightsForInspectionAndDrawing(inspectionId, drawingId) {
  return db.highlights
    .where('[inspectionId+drawingId]')
    .equals([Number(inspectionId), Number(drawingId)])
    .toArray();
}

/* --------------------------------------------------------------------------
   Comment library helpers (Phase 5)
   --------------------------------------------------------------------------
   Returns seed entries (and anything the user has added via Settings in a
   later phase). Universal entries — those tagged with ANY_TYPE — are always
   included after the type-specific list so the engineer sees them both.
   -------------------------------------------------------------------------- */

/**
 * List library entries relevant to an inspection type.
 *
 * opts.severity — 'observation' | 'defect' | null (all)
 */
export async function listCommentsForInspectionType(typeKey, opts = {}) {
  const typed = await db.commentLibrary
    .where('inspectionTypeKey').equals(String(typeKey))
    .toArray();
  const any = await db.commentLibrary
    .where('inspectionTypeKey').equals(ANY_TYPE)
    .toArray();
  let out = [...typed, ...any];
  if (opts.severity) {
    out = out.filter((e) => e.severity === opts.severity);
  }
  return out;
}

/**
 * List general-comment library entries for the inspection report's
 * "General observations" section. Returns universal + type-specific.
 */
export async function listGeneralCommentsForInspectionType(typeKey) {
  const typed = await db.generalComments
    .where('inspectionTypeKey').equals(String(typeKey))
    .toArray();
  const any = await db.generalComments
    .where('inspectionTypeKey').equals(ANY_TYPE)
    .toArray();
  return [...typed, ...any];
}

export function getCommentLibraryEntry(key) {
  return db.commentLibrary.where('key').equals(String(key)).first();
}

/* --------------------------------------------------------------------------
   Report helpers (Phase 7)
   --------------------------------------------------------------------------
   Each report is a point-in-time snapshot — generating again doesn't replace
   the previous file but adds a new version. Engineer can delete old ones.
   -------------------------------------------------------------------------- */

/**
 * Save a generated report against an inspection.
 *
 * data = {
 *   filename,      // suggested download filename
 *   pdfBlob,       // application/pdf Blob
 *   sizeBytes,     // blob.size (denormalised so lists don't touch the blob)
 *   pageCount,     // from pdf.getNumberOfPages()
 *   generatedBy    // inspector name at time of generation
 * }
 */
export async function saveReport(inspectionId, data) {
  const now = new Date().toISOString();
  const record = {
    inspectionId: Number(inspectionId),
    filename:     (data.filename || 'report.pdf').trim(),
    pdfBlob:      data.pdfBlob,
    sizeBytes:    data.sizeBytes ?? (data.pdfBlob?.size || 0),
    pageCount:    data.pageCount || null,
    generatedBy:  (data.generatedBy || '').trim(),
    generatedAt:  now
  };
  const id = await db.reports.add(record);
  // Touch inspection so list ordering reflects the new version.
  await db.inspections.update(record.inspectionId, { updatedAt: now });
  return { id, ...record };
}

export function listReportsForInspection(inspectionId) {
  return db.reports
    .where('inspectionId').equals(Number(inspectionId))
    .reverse().sortBy('generatedAt');
}

export function getReport(id) {
  return db.reports.get(Number(id));
}

export async function deleteReport(id) {
  const reportId = Number(id);
  const r = await db.reports.get(reportId);
  await db.reports.delete(reportId);
  if (r) {
    await db.inspections.update(r.inspectionId, { updatedAt: new Date().toISOString() });
  }
}

/* --------------------------------------------------------------------------
   User profile (Phase 8)
   --------------------------------------------------------------------------
   Stored in the settings table under key 'user.profile'. Single-user app for
   v1, so one record suffices. Phase 9 could move this to the `users` table
   if we ever multi-tenant.
   -------------------------------------------------------------------------- */
const PROFILE_KEY = 'user.profile';

const EMPTY_PROFILE = Object.freeze({
  name:    '',
  email:   '',
  rpeq:    '',
  cpeng:   '',
  company: 'Bligh Tanner Pty Ltd',
  role:    'Senior Structural Engineer',
  signatureBlob: null,   // optional: PNG / JPEG Blob of the engineer's signature
  anthropicKey:  '',     // optional: Claude API key for AI comment expansion (stored per-device only)
  builderEmail:  ''      // optional: default builder email for close-out mailto links
});

export async function getUserProfile() {
  const row = await db.settings.get(PROFILE_KEY);
  if (!row) return { ...EMPTY_PROFILE };
  return { ...EMPTY_PROFILE, ...(row.value || {}) };
}

export async function saveUserProfile(profile) {
  // Preserve the existing signature blob unless the caller explicitly passed one
  // (or explicitly cleared it by passing null).
  const existing = await getUserProfile();
  const value = {
    name:    (profile.name    || '').trim(),
    email:   (profile.email   || '').trim(),
    rpeq:    (profile.rpeq    || '').trim(),
    cpeng:   (profile.cpeng   || '').trim(),
    company: (profile.company || '').trim() || 'Bligh Tanner Pty Ltd',
    role:    (profile.role    || '').trim(),
    signatureBlob: Object.prototype.hasOwnProperty.call(profile, 'signatureBlob')
      ? profile.signatureBlob
      : existing.signatureBlob,
    anthropicKey: Object.prototype.hasOwnProperty.call(profile, 'anthropicKey')
      ? String(profile.anthropicKey || '').trim()
      : existing.anthropicKey,
    builderEmail: Object.prototype.hasOwnProperty.call(profile, 'builderEmail')
      ? String(profile.builderEmail || '').trim()
      : existing.builderEmail
  };
  await db.settings.put({ key: PROFILE_KEY, value });
  return value;
}
