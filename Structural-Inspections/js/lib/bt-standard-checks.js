/**
 * BT Standard Checks Library — v2.2.1
 *
 * Per-inspection-type catalogue of every check Bligh Tanner does as standard
 * practice on a structural inspection. Drawn directly from BT's General Notes
 * pattern (95% identical across BT-authored projects) plus field practice.
 *
 * STRUCTURE
 *   - COMMON_* arrays: reusable check sets composed into specific types
 *   - BT_STANDARD_CHECKS: keyed by inspection-type key (must match
 *     reference/inspection-types.json)
 *
 * EACH ENTRY HAS
 *   - siteReadiness:    [{text}, ...]                    — yes/no questions for builder before booking
 *   - onSiteChecks:     [{text, asClauseRef?, noteRef?, critical?}, ...] — what BT verifies on site
 *   - pourDayRecords:   [{text, captureType?, spec?}, ...]  — delivery dockets, certs, test results
 *   - certifications:   [{text, providedBy, formType?}, ...]  — what contractor must supply
 *   - standardComments: [{text, severity, kind?}, ...]   — common defect/observation comments
 *
 * USAGE
 *   import { BT_STANDARD_CHECKS, getStandardChecksForType, mergeWithProjectChecks }
 *     from '../lib/bt-standard-checks.js';
 *
 *   const checks = getStandardChecksForType('pad-footing-prepour');
 *
 * Project-specific data (cover values, project bearing, specific block-wall
 * marks) overrides/augments via the Cowork-supplied expectedChecklist[]. The
 * standard checks fill the gap for everything universally true.
 */

/* ──────────────────────────────────────────────────────────────────────────
 * COMMON CHECK SETS — composed into per-type entries
 * ────────────────────────────────────────────────────────────────────────── */

const COMMON_SITE_READINESS_PREPOUR = [
  'Reinforcement fully placed and tied',
  'Cover spacers / bar chairs in place',
  'Penetrations / cast-in items installed',
  'Trimmer bars at re-entrant corners and penetrations >200 sq',
  'Formwork clean, dry, and secure',
  'Site access for vehicle / crane / EWP if required'
];

const COMMON_REINF_CHECKS = [
  { text: 'Reinforcement layout matches GA / reo plan (size, count, spacing)', critical: true },
  { text: 'All bars securely tied', noteRef: 'C18' },
  { text: 'Lap lengths at corners and T-junctions', asClauseRef: 'AS 3600' },
  { text: 'Bar chairs / cover spacers at correct centres', noteRef: 'C9' },
  { text: 'Hooks and bends per AS 3600', asClauseRef: 'AS 3600' },
  { text: 'Trimmer bars at re-entrant corners and penetrations >200 sq (2-N12 × 1200 long @ 100 ctrs UNO)', noteRef: 'C6' },
  { text: 'No reinforcement misalignment or displacement during fixing' },
  { text: 'Tags / mill marks visible for spot-check of bar grade' }
];

const COMMON_CONCRETE_POUR_CHECKS = [
  { text: 'Concrete grade matches schedule for this element', critical: true },
  { text: 'Slump within tolerance (typ. 80 mm UNO)', noteRef: 'C1' },
  { text: 'Pour temperature 5–35 °C', noteRef: 'C19' },
  { text: 'Concrete vibrated during placement', noteRef: 'C17' },
  { text: 'Curing regime in place — wet 3 days, prevent moisture loss 7 days', noteRef: 'C5' },
  { text: 'Aliphatic alcohol barrier sprayed after initial screed (slabs)', noteRef: 'C3' }
];

const COMMON_FORMWORK_CHECKS = [
  { text: 'Formwork dimensions match drawings', critical: true },
  { text: 'Formwork clean, dry, no debris' },
  { text: 'Formwork properly propped per AS 3610', asClauseRef: 'AS 3610', critical: true },
  { text: 'Construction joints at engineer-approved locations', noteRef: 'C11' }
];

const COMMON_POUR_DAY_RECORDS = [
  { text: 'Concrete delivery docket — photograph for record', captureType: 'photo' },
  { text: 'Mix ID matches project spec', captureType: 'verify' },
  { text: 'Slump on arrival (per truck)', captureType: 'measure' },
  { text: 'Air temperature at pour', captureType: 'measure' },
  { text: 'Cube test slip — collect from contractor', captureType: 'photo' }
];

const COMMON_GEOTECH_HOLD = [
  { text: 'Excavation maintained firm and dry', noteRef: 'F2', critical: true },
  { text: 'Soft ground removed and replaced with mass concrete', noteRef: 'F2' },
  { text: 'Foundation excavation per geotech recommendations', critical: true },
  { text: 'Geotech RPEQ inspection certificate provided for this element', critical: true }
];

const COMMON_STEEL_CONNECTION_CHECKS = [
  { text: 'All bolts tightened (full bearing) per AS 4100 8.8/S', asClauseRef: 'AS 4100', critical: true },
  { text: 'Min 2 threads past nut after tightening', noteRef: 'S6' },
  { text: 'Plate washers for oversized/slotted holes per AS 4100 Cl 14.3.5.2', asClauseRef: 'AS 4100 Cl 14.3.5.2' },
  { text: 'Connection plates ≥ 10 mm thick UNO', noteRef: 'S6' },
  { text: 'Welds match drawings — length, size, category (SP / GP)', critical: true },
  { text: 'Weld 6 mm SP continuous fillet UNO; AS 1554 procedures', asClauseRef: 'AS 1554', noteRef: 'S6' },
  { text: 'NDT testing per W-table (visual / MPI / UT / RT)', critical: true },
  { text: 'Bolt holes not enlarged during erection', noteRef: 'S13' }
];

const COMMON_STEEL_CORROSION_CHECKS = [
  { text: 'External steel HDG600 hot-dip galvanised to AS/NZS 4680', noteRef: 'S10' },
  { text: 'Members in contact with concrete passivated', noteRef: 'S10' },
  { text: 'Damaged galv repaired with WATTYL Galvit (or equivalent high-organic Zn epoxy)', noteRef: 'S14' },
  { text: 'ACRS certification on file for the supplier', noteRef: 'S15', critical: true }
];

const COMMON_BLOCKWORK_VERT_REO_CHECKS = [
  { text: 'Vertical reinforcement size, spacing, position per BLOCKWALL SCHEDULE', critical: true },
  { text: 'Starter bars securely tied prior to wall footing pour', noteRef: 'CM5' },
  { text: 'Lap lengths to AS 3700', asClauseRef: 'AS 3700' },
  { text: 'Mortar mix per spec (typ. M3 1:1:6 general; M4 2:1:9 retaining wall)', noteRef: 'CM' },
  { text: 'Mortar f\'uc ≥ 15 MPa (general) / 25 MPa (retaining)', noteRef: 'CM' },
  { text: 'Wall ties: medium duty @ 600 mm centres each direction (cavity walls)', noteRef: 'CM8' },
  { text: 'Vertical control joints at junction of dissimilar materials', noteRef: 'CM7' },
  { text: 'Vertical control joints ≤ 8 m, ≤ 5 m from corners, NOT within 1.2 m of corners', noteRef: 'CM6' },
  { text: 'Clean-out blocks at base of all cores to be filled', noteRef: 'CM3' }
];

const COMMON_BLOCKWORK_CORE_FILL_CHECKS = [
  { text: 'Concrete strength f\'c = 20 MPa', noteRef: 'CM3', critical: true },
  { text: 'Max slump 230 mm', noteRef: 'CM3' },
  { text: 'Max aggregate size 10 mm', noteRef: 'CM3' },
  { text: 'Max lift height ≤ 2400 mm', noteRef: 'CM3', critical: true },
  { text: 'All cores swept clean of mortar via clean-out blocks', noteRef: 'CM3', critical: true },
  { text: 'Vertical reinforcement still in correct position before fill', critical: true },
  { text: '\'H\' units used UNO', noteRef: 'CM3' },
  { text: 'Grout fully compacted — visual on consolidation' },
  { text: 'No back filling behind retaining walls within 14 days', noteRef: 'CM10' },
  { text: '10% of chemical anchors into core-filled blockwork load-tested @ 1.5 × SWL', noteRef: 'CM15', critical: true }
];

const COMMON_ANCHOR_CHECKS = [
  { text: 'All anchors comply with AS 5216:2018', asClauseRef: 'AS 5216:2018', critical: true },
  { text: 'Installed per manufacturer specification', noteRef: 'A1' },
  { text: 'Holes hammer-drilled', noteRef: 'A6', critical: true },
  { text: 'Dust-reducing drilling system used', noteRef: 'A7' },
  { text: '5% of anchors load-tested; 100% if any fail', noteRef: 'A8', critical: true },
  { text: 'Min edge distance / spacing / embedment per A1 schedule', noteRef: 'A1', critical: true },
  { text: 'Anchors at angle > 5° fitted with tapered washer', noteRef: 'A12' },
  { text: 'Aborted holes filled with non-shrink mortar ≥ substrate (min 40 MPa)', noteRef: 'A11' },
  { text: 'Slab-soffit anchors are mechanical with Loctite UNO', noteRef: 'A4' },
  { text: 'Coatings / corrosion per steelwork notes + manufacturer spec', noteRef: 'A5' }
];

const COMMON_FORM12_CERTS = [
  { text: 'BT Form 12 — RPEQ construction certification (final, at PC)', providedBy: 'Bligh Tanner' },
  { text: 'ARCS certificate — reinforcing steel supplier', providedBy: 'Reo supplier', noteRef: 'C20' },
  { text: 'ARCS certificate — structural steel supplier (overseas-sourced steel)', providedBy: 'Steel supplier', noteRef: 'S15' }
];


/* ──────────────────────────────────────────────────────────────────────────
 * BT_STANDARD_CHECKS — per inspection type
 * ────────────────────────────────────────────────────────────────────────── */

export const BT_STANDARD_CHECKS = {

  /* ════════ FOUNDATIONS ════════ */

  'subgrade': {
    name: 'Subgrade / Bearing Inspection',
    siteReadiness: [
      'Excavation complete to design level',
      'Geotech RPEQ on site for verification',
      'Excavation maintained firm and dry',
      'Soft spots removed/replaced with mass concrete',
      'Topsoil + organic material stripped (150 mm min)',
      'CBR 15 fill placed per civil engineer (if required)'
    ],
    onSiteChecks: [
      ...COMMON_GEOTECH_HOLD,
      { text: 'Founding depth as per drawings (min depth at each footing)', critical: true },
      { text: 'Allowable bearing capacity confirmed by geotech (per project F1 schedule)', critical: true },
      { text: 'Line of influence below existing footings respected (rock 2:1, clay 1:1, sand 1:2)', noteRef: 'F4' },
      { text: 'Topsoil / organics stripped 150 mm min', noteRef: 'E2' },
      { text: 'CBR 15 fill placed in 200 mm layers @ 98% MMD if applicable', noteRef: 'E3' },
      { text: 'Field density tests on each fill layer including subgrade per AS 3798-1998 §8', asClauseRef: 'AS 3798-1998 §8', noteRef: 'E4' }
    ],
    pourDayRecords: [],
    certifications: [
      { text: 'Geotech RPEQ — bearing capacity confirmed', providedBy: 'Builder\'s geotech engineer', formType: 'inspection certificate', noteRef: 'F7' },
      { text: 'Field density test reports — every fill layer + subgrade', providedBy: 'Civil engineer / NATA lab', noteRef: 'E4' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Bearing capacity confirmed by geotech on site as adequate for the footings shown.', severity: 'observation' },
      { text: 'Subgrade is firm, dry, and free of soft material at the time of inspection.', severity: 'observation' },
      { text: 'Soft material observed at <location> — to be removed and replaced with mass concrete (25 MPa min) per note F2 prior to footing pour.', severity: 'defect' }
    ]
  },

  'pile-cfa-install': {
    name: 'CFA Pile Installation',
    siteReadiness: [
      'Pile installation log complete (per pile — depth, torque, socket)',
      'Founding depth records collated and forwarded to BT within 3 days',
      'Geotech RPEQ on site for verification',
      'Pile extensions exposed for inspection',
      'Set-out survey complete (within 75 mm tolerance)',
      'Site access for vehicle / inspection'
    ],
    onSiteChecks: [
      { text: 'Pile location ≤ 75 mm of designated position; out-of-position piles notified', noteRef: 'P3', critical: true },
      { text: 'Founding depth recorded per pile (forwarded to BT within 3 days)', noteRef: 'P4', critical: true },
      { text: 'Installation torque recorded per pile' },
      { text: 'Socket length / shaft material strength confirmed by geotech RPEQ' },
      { text: 'No pile within 1000 mm of existing stormwater (or pre-bored to below invert)', noteRef: 'P5' },
      { text: 'CFA grout strength f\'c ≥ 40 MPa', noteRef: 'P2', critical: true },
      { text: 'Pile extension ≥ 75 mm into pile cap or ground beam', noteRef: 'P7', critical: true },
      { text: 'Reinforcement cage placement and lap with starter bars' }
    ],
    pourDayRecords: [
      { text: 'Pile log per pile (depth, torque, socket length)', captureType: 'photo' },
      { text: 'CFA grout delivery dockets (f\'c, slump)', captureType: 'photo' }
    ],
    certifications: [
      { text: 'Piling Contractor RPEQ — design loads achieved', providedBy: 'Piling contractor RPEQ', noteRef: 'P8', critical: true },
      { text: 'Geotech RPEQ — founding depth + socket verification', providedBy: 'Builder\'s geotech engineer' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'CFA piles installed as per plan, with founding depths and torques recorded.', severity: 'observation' },
      { text: 'BT inspection of CFA piles is advisory only; the geotechnical engineer holds the certification of the pile-bearing capacity.', severity: 'observation' },
      { text: 'Pile <ID> located more than 75 mm out of designated position — refer to BT for review of pile cap reinforcement adjustment.', severity: 'defect' }
    ]
  },

  'pile-bored-install': {
    name: 'Bored Pier Installation',
    siteReadiness: [
      'Pier installation log complete (per pier — depth, socket)',
      'Founding depth records collated and forwarded to BT within 3 days',
      'Geotech RPEQ on site for verification',
      'Pier extensions exposed for inspection',
      'Reinforcement cage ready before pour',
      'Site access for vehicle / inspection'
    ],
    onSiteChecks: [
      { text: 'Pier location ≤ 75 mm of designated position', noteRef: 'P3', critical: true },
      { text: 'Founding depth meets project min depth (verify per project schedule)', critical: true },
      { text: 'Pier diameter matches schedule (verify reo cage clear of formwork)' },
      { text: 'Min socket depth achieved into bearing strata (per project schedule)', critical: true },
      { text: 'Founding material verified by geotech RPEQ on site', critical: true },
      { text: 'Allowable end-bearing per project (typ. 1000 kPa) + shaft adhesion (typ. 40 kPa)' },
      { text: 'Pier extension ≥ 75 mm into pile cap or ground beam', noteRef: 'P7', critical: true },
      { text: 'Reinforcement cage placement and lap with starter bars' }
    ],
    pourDayRecords: [
      { text: 'Pier log per pier (depth, socket length, founding material)', captureType: 'photo' },
      { text: 'Concrete delivery dockets', captureType: 'photo' }
    ],
    certifications: [
      { text: 'Piling Contractor RPEQ — design loads achieved', providedBy: 'Piling contractor RPEQ', noteRef: 'P8', critical: true },
      { text: 'Geotech RPEQ — founding depth + bearing verification', providedBy: 'Builder\'s geotech engineer' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Bored piers installed to founding depths recorded by the geotechnical engineer; bearing capacity confirmed adequate.', severity: 'observation' },
      { text: 'BT inspection of bored piers is advisory; certification of bearing capacity rests with the geotechnical engineer.', severity: 'observation' },
      { text: 'Pier <ID> founding depth shallower than minimum specified — geotech RPEQ to confirm acceptability or pier to be deepened.', severity: 'defect' }
    ]
  },

  'pad-footing-prepour': {
    name: 'Pad Footing Pre-Pour Reinforcement',
    siteReadiness: [
      'Excavation to founding level',
      'Geotech sign-off on bearing (if not separate inspection)',
      'Reinforcement fully placed and tied',
      'Pile / pier reo extends ≥ 75 mm into footing',
      'Cover blocks / chairs in place',
      'Site access for vehicle / photographer'
    ],
    onSiteChecks: [
      ...COMMON_REINF_CHECKS,
      { text: 'Cover ≥ project schedule (typ. 50 mm bottom & sides, N32)', critical: true },
      { text: 'Pile / pier reo extends ≥ 75 mm into footing', noteRef: 'P7', critical: true },
      { text: 'Excavation maintained firm and dry', noteRef: 'F2' },
      { text: 'Reo securely tied prior to pour', noteRef: 'C18' },
      { text: 'Footing dimensions match schedule (PF mark from S010)', critical: true }
    ],
    pourDayRecords: [...COMMON_POUR_DAY_RECORDS],
    certifications: [
      { text: 'Geotech inspection cert (per footing) — bearing confirmed', providedBy: 'Builder\'s geotech engineer', noteRef: 'F7' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Pad footing reinforcement is in accordance with the structural drawings and ready for concrete placement, subject to rectification of items listed above.', severity: 'observation' },
      { text: 'Cover spacers in place; cover values verified by spot-check.', severity: 'observation' },
      { text: 'Insufficient cover at <location> — additional cover spacers required prior to pour.', severity: 'defect' },
      { text: 'Trimmer bars missing at penetration <location> — to be installed prior to pour.', severity: 'defect' }
    ]
  },

  'strip-footing-prepour': {
    name: 'Strip Footing Pre-Pour Reinforcement',
    siteReadiness: [...COMMON_SITE_READINESS_PREPOUR],
    onSiteChecks: [
      ...COMMON_REINF_CHECKS,
      { text: 'Cover ≥ project schedule (typ. 50 mm bottom & sides, N32)', critical: true },
      { text: 'Strip width and depth match schedule (SF mark from S010)', critical: true },
      { text: 'Continuous reinforcement laps per AS 3600' },
      { text: 'Excavation maintained firm and dry', noteRef: 'F2' }
    ],
    pourDayRecords: [...COMMON_POUR_DAY_RECORDS],
    certifications: [
      { text: 'Geotech inspection cert — bearing confirmed', providedBy: 'Builder\'s geotech engineer', noteRef: 'F7' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Strip footing excavation, reinforcement and starters are ready for concrete placement, subject to rectification of items listed above.', severity: 'observation' }
    ]
  },

  'raft-footing-prepour': {
    name: 'Raft / Waffle Footing Pre-Pour Reinforcement',
    siteReadiness: [...COMMON_SITE_READINESS_PREPOUR, 'Vapour barrier / DPM intact under raft area'],
    onSiteChecks: [
      ...COMMON_REINF_CHECKS,
      { text: 'Raft designed for site reactivity class (e.g. Class M per AS 2870)', asClauseRef: 'AS 2870', noteRef: 'F6', critical: true },
      { text: 'Cover per project schedule', critical: true },
      { text: 'Beam and slab reinforcement to design (size, count, position)' },
      { text: 'Waffle pods / void formers correctly positioned (if applicable)' },
      { text: 'DPM under raft area intact, properly lapped' }
    ],
    pourDayRecords: [...COMMON_POUR_DAY_RECORDS],
    certifications: [...COMMON_FORM12_CERTS],
    standardComments: [
      { text: 'Raft slab reinforcement, edge beams and pods are ready for concrete placement, subject to rectification of items listed above.', severity: 'observation' }
    ]
  },

  /* ════════ SLABS ════════ */

  'slab-prepour-ground': {
    name: 'Slab on Ground Pre-Pour Reinforcement',
    siteReadiness: [
      '50 mm bedding sand placed and compacted',
      'Damp-proof membrane intact (no holes/tears)',
      ...COMMON_SITE_READINESS_PREPOUR
    ],
    onSiteChecks: [
      { text: '50 mm bedding sand + DPM in place under slab', noteRef: 'C16', critical: true },
      ...COMMON_REINF_CHECKS,
      { text: 'Cover per project schedule (typ. internal: bot 40 / top 30 / sides 40; external: 40/40/40)', critical: true },
      { text: 'Mesh laps: 2 outer-most cross bars overlap', noteRef: 'C7', critical: true },
      { text: 'Bar chairs @ 600 ctrs SL72/82, 800 ctrs SL92+', noteRef: 'C9' },
      { text: 'Trimmer bars at re-entrant corners and penetrations >200 sq, each layer of mesh for SOG', noteRef: 'C6' },
      { text: 'No conduits / pipes outside middle one-third of slab depth, ≥3 dia spacing', noteRef: 'C13' }
    ],
    pourDayRecords: [...COMMON_POUR_DAY_RECORDS],
    certifications: [
      { text: 'ARCS certificate — reinforcing steel supplier', providedBy: 'Reo supplier', noteRef: 'C20' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Slab on ground reinforcement is in accordance with the structural drawings and ready for concrete placement, subject to rectification of items listed above.', severity: 'observation' },
      { text: 'DPM intact and continuous over the slab area at the time of inspection.', severity: 'observation' },
      { text: 'DPM punctured at <location> — to be repaired with patch and sealing tape prior to pour.', severity: 'defect' }
    ]
  },

  'slab-prepour-suspended': {
    name: 'Suspended Slab Pre-Pour Reinforcement',
    siteReadiness: [
      'Head framing (formwork + props) complete per AS 3610',
      ...COMMON_SITE_READINESS_PREPOUR,
      'Cast-in plates / starter bars for next-level walls or steel cols installed'
    ],
    onSiteChecks: [
      ...COMMON_FORMWORK_CHECKS,
      ...COMMON_REINF_CHECKS,
      { text: 'Cover per project schedule (typ. 30/30/30 mm, N40)', critical: true },
      { text: 'Bottom reo per bottom-reo plan; top reo per top-reo plan', critical: true },
      { text: 'Extra reo plans (if present) verified', noteRef: 'C6' },
      { text: 'Trimmer bars at penetrations >200 sq, top & bottom for suspended slabs', noteRef: 'C6', critical: true },
      { text: 'Starter bars projecting for next-level walls / columns', critical: true },
      { text: 'No conduits outside middle one-third, ≥3 dia spacing', noteRef: 'C13' }
    ],
    pourDayRecords: [...COMMON_POUR_DAY_RECORDS],
    certifications: [
      { text: 'ARCS certificate — reinforcing steel supplier', providedBy: 'Reo supplier', noteRef: 'C20' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Suspended slab reinforcement, head framing and starter bars are in accordance with the structural drawings and ready for concrete placement, subject to rectification of items listed above.', severity: 'observation' },
      { text: 'Formwork is securely propped and the slab soffit is straight and true at the time of inspection.', severity: 'observation' },
      { text: 'Trimmer bars missing at penetration <location> top and bottom — to be installed prior to pour.', severity: 'defect' },
      { text: 'Top reinforcement disturbed during pour preparation at <location> — re-tie before pour.', severity: 'defect' }
    ]
  },

  'head-framing': {
    name: 'Suspended Slab Head Framing (Formwork & Props)',
    siteReadiness: [
      'Falsework / scaffolding erected to designed configuration',
      'Soffit panels installed and aligned',
      'Site access (EWP if required)'
    ],
    onSiteChecks: [
      ...COMMON_FORMWORK_CHECKS,
      { text: 'Soffit level and true within tolerance', critical: true },
      { text: 'Props at designed centres; back-propping (if required) per scheme', critical: true },
      { text: 'Falsework engineer\'s erection sketches followed' },
      { text: 'Edge formwork for slab perimeter set out and braced' },
      { text: 'Stripping period understood — formwork & propping under suspended slabs to be removed PRIOR to construction of masonry over', noteRef: 'C10' }
    ],
    pourDayRecords: [],
    certifications: [
      { text: 'Falsework / scaffolding cert (if applicable)', providedBy: 'Falsework engineer / contractor RPEQ' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Head framing, formwork and propping are in place and ready for reinforcement.', severity: 'observation' },
      { text: 'Excessive deflection / springing observed at <location> — additional propping required prior to reinforcement.', severity: 'defect' }
    ]
  },

  'beam-prepour': {
    name: 'Beam Pre-Pour Reinforcement',
    siteReadiness: [...COMMON_SITE_READINESS_PREPOUR],
    onSiteChecks: [
      ...COMMON_FORMWORK_CHECKS,
      ...COMMON_REINF_CHECKS,
      { text: 'Cover per project schedule', critical: true },
      { text: 'Stirrup spacing closer at supports', critical: true },
      { text: 'Beam camber as noted on drawings (if applicable)' },
      { text: 'Continuity reinforcement at supports' }
    ],
    pourDayRecords: [...COMMON_POUR_DAY_RECORDS],
    certifications: [...COMMON_FORM12_CERTS],
    standardComments: [
      { text: 'Beam reinforcement is in accordance with the structural drawings and ready for concrete placement, subject to rectification of items listed above.', severity: 'observation' }
    ]
  },

  'column-prepour': {
    name: 'Column Pre-Pour Reinforcement',
    siteReadiness: [...COMMON_SITE_READINESS_PREPOUR],
    onSiteChecks: [
      ...COMMON_FORMWORK_CHECKS,
      ...COMMON_REINF_CHECKS,
      { text: 'Cover per project schedule (typ. 40 mm sides, N40)', critical: true },
      { text: 'Vertical bar count, size, lap length per typ detail', critical: true },
      { text: 'Tie / fitment spacing — closer at top & bottom of column', critical: true },
      { text: '135° hook ends on ties', noteRef: 'C-tie' },
      { text: 'Starter bar projection from below matches lap length' },
      { text: 'Column dimensions match schedule (CC mark from S010)', critical: true },
      { text: 'Formwork plumb to within tolerance' }
    ],
    pourDayRecords: [...COMMON_POUR_DAY_RECORDS],
    certifications: [...COMMON_FORM12_CERTS],
    standardComments: [
      { text: 'Column reinforcement is in accordance with the structural drawings and ready for concrete placement, subject to rectification of items listed above.', severity: 'observation' },
      { text: 'Tie spacing at column ends does not match typical detail at <location> — additional ties required.', severity: 'defect' }
    ]
  },

  'concrete-wall-prepour': {
    name: 'Concrete Wall Pre-Pour Reinforcement',
    siteReadiness: [...COMMON_SITE_READINESS_PREPOUR],
    onSiteChecks: [
      ...COMMON_FORMWORK_CHECKS,
      ...COMMON_REINF_CHECKS,
      { text: 'Cover per project schedule (typ. 40 mm sides + top, N40)', critical: true },
      { text: 'Vertical and horizontal reo per typ detail', critical: true },
      { text: 'Lap lengths per AS 3600', asClauseRef: 'AS 3600' },
      { text: 'Starter bars from below securely tied' },
      { text: 'Penetrations / openings have trimmer bars' },
      { text: 'Wall mark matches schedule (CW mark from S010)', critical: true }
    ],
    pourDayRecords: [...COMMON_POUR_DAY_RECORDS],
    certifications: [...COMMON_FORM12_CERTS],
    standardComments: [
      { text: 'Concrete wall reinforcement and formwork are in accordance with the structural drawings and ready for concrete placement, subject to rectification of items listed above.', severity: 'observation' }
    ]
  },

  'retaining-wall-prepour': {
    name: 'Retaining Wall Pre-Pour Reinforcement',
    siteReadiness: [...COMMON_SITE_READINESS_PREPOUR, 'Drainage layer / weep holes set out per drawings'],
    onSiteChecks: [
      ...COMMON_FORMWORK_CHECKS,
      ...COMMON_REINF_CHECKS,
      { text: 'Cover per project schedule', critical: true },
      { text: 'Heel and toe reinforcement per typ detail', critical: true },
      { text: 'Drainage layer / ag drain provision verified' },
      { text: 'Tie-down bars at base of stem', critical: true }
    ],
    pourDayRecords: [...COMMON_POUR_DAY_RECORDS],
    certifications: [...COMMON_FORM12_CERTS],
    standardComments: [
      { text: 'Retaining wall reinforcement and formwork are ready for concrete placement, subject to rectification of items listed above and provision of drainage per the drawings.', severity: 'observation' }
    ]
  },

  'stair-prepour': {
    name: 'Stair Flight & Landing Pre-Pour Reinforcement',
    siteReadiness: [...COMMON_SITE_READINESS_PREPOUR],
    onSiteChecks: [
      ...COMMON_FORMWORK_CHECKS,
      ...COMMON_REINF_CHECKS,
      { text: 'Cover per project schedule (typ. 30/30/30 mm, N40)', critical: true },
      { text: 'Tread / riser geometry per arch', critical: true },
      { text: 'Connection to landing — starter bars projecting' },
      { text: 'Landing reinforcement integrated with adjacent slab' },
      { text: 'Going / rise consistency on each flight' }
    ],
    pourDayRecords: [...COMMON_POUR_DAY_RECORDS],
    certifications: [...COMMON_FORM12_CERTS],
    standardComments: [
      { text: 'Stair flight and landing reinforcement are ready for concrete placement, subject to rectification of items listed above.', severity: 'observation' }
    ]
  },

  /* ════════ PT + POUR WITNESS ════════ */

  'post-tension-strand': {
    name: 'Post-Tension Strand Placement',
    siteReadiness: [
      'Strand fully placed per profile (PT designer\'s layout)',
      'Anchorages installed at dead-end and live-end',
      'Stressing pockets free of obstruction',
      'Sleeves / barriers at column heads',
      'Concrete delivery scheduled within strand placement window'
    ],
    onSiteChecks: [
      { text: 'Strand profile matches PT designer\'s layout (high points, low points)', critical: true },
      { text: 'Anchorage zone reinforcement per typ detail', critical: true },
      { text: 'Tendon spacing and centres verified' },
      { text: 'Dead-end / live-end stressing pockets free of obstruction', critical: true },
      { text: 'Sleeves and barriers in place at column heads (anti-bursting)' },
      { text: 'Bonded vs unbonded system per design', noteRef: 'PS' },
      { text: 'Cover to strand maintained per project schedule (suspended slab PT)', critical: true }
    ],
    pourDayRecords: [
      { text: 'PT designer sign-off prior to pour', captureType: 'verify', critical: true },
      ...COMMON_POUR_DAY_RECORDS
    ],
    certifications: [
      { text: 'PT specialist sub — Form 15 design certification', providedBy: 'PT specialist (RPEQ)', formType: 'Form 15', noteRef: 'PS3', critical: true },
      { text: 'PT specialist sub — Form 16 construction certification', providedBy: 'PT specialist (RPEQ)', formType: 'Form 16', noteRef: 'PS3', critical: true },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Post-tension strand placement is in accordance with the PT designer\'s drawings and ready for concrete placement, subject to PT designer sign-off.', severity: 'observation' },
      { text: 'BT inspection of post-tensioned concrete is limited to the in-situ supporting works; PT design and certification rests with the PT specialist sub-contractor (Form 15 / Form 16).', severity: 'observation' }
    ]
  },

  'concrete-pour-witness': {
    name: 'Concrete Pour Witness',
    siteReadiness: [
      'Pre-pour reo inspection complete + signed off',
      'Concrete pump / placement equipment on site',
      'Curing materials (alcohol barrier, hessian / poly) on site',
      'Vibrator(s) operational',
      'Pour schedule confirmed (volume, mix design, delivery time)'
    ],
    onSiteChecks: [
      { text: 'Concrete grade matches schedule for this element', critical: true },
      { text: 'Slump on arrival — within tolerance per truck', critical: true },
      { text: 'Air temperature & concrete temp 5–35 °C', noteRef: 'C19', critical: true },
      { text: 'Vibration during placement — no honeycombing visible', noteRef: 'C17' },
      { text: 'Construction joint locations as approved', noteRef: 'C11' },
      { text: 'Aliphatic alcohol barrier sprayed after initial screed (slabs)', noteRef: 'C3' },
      { text: 'Curing regime initiated — wet 3 days, prevent moisture loss 7 days', noteRef: 'C5' },
      { text: 'No re-tempering of concrete after delivery' }
    ],
    pourDayRecords: [...COMMON_POUR_DAY_RECORDS],
    certifications: [
      { text: 'Concrete delivery dockets — every truck', providedBy: 'Concrete supplier' },
      { text: 'Cube test slips for compressive testing', providedBy: 'Concrete supplier / NATA lab' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Concrete was placed and compacted in accordance with AS 3600:2018 Cl. 17.1.', severity: 'observation' },
      { text: 'Cube test specimens were taken from <truck/location> and submitted to NATA-certified testing.', severity: 'observation' }
    ]
  },

  'transfer-slab-witness': {
    name: 'Transfer Slab Pour Witness (Critical)',
    siteReadiness: [
      'Pre-pour reo inspection complete + signed off',
      'PT designer sign-off (if PT)',
      'Adequate vibration equipment on site (multiple vibrators)',
      'Pour sequence agreed with site supervisor',
      'Curing materials on site',
      'Rest periods scheduled for placement crew (long pour)'
    ],
    onSiteChecks: [
      { text: 'Concrete grade matches schedule (typ. higher grade for transfer)', critical: true },
      { text: 'Slump tighter tolerance than typical (verify per truck)', critical: true },
      { text: 'Pour sequence followed — no cold joints in critical zones', critical: true },
      { text: 'Multiple vibrators in use — full compaction at every layer', noteRef: 'C17', critical: true },
      { text: 'No movement of reo / penetrations during pour' },
      { text: 'Pour temperature monitored 5–35 °C', noteRef: 'C19', critical: true },
      { text: 'Construction joints (if any) at engineer-approved locations only', noteRef: 'C11', critical: true },
      { text: 'Curing regime initiated immediately on each pour zone', noteRef: 'C5' }
    ],
    pourDayRecords: [
      { text: 'Concrete delivery docket — every truck', captureType: 'photo' },
      { text: 'Slump test result — every truck', captureType: 'measure' },
      { text: 'Air & concrete temperature throughout pour', captureType: 'measure' },
      { text: 'Cube test specimens — multiple sets per pour', captureType: 'photo' },
      { text: 'Photos of finished pour surface + edge formwork', captureType: 'photo' }
    ],
    certifications: [
      { text: 'Concrete delivery dockets — every truck', providedBy: 'Concrete supplier' },
      { text: 'Cube test slips — multiple sets', providedBy: 'NATA lab' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Transfer slab pour was witnessed in full; placement, vibration and curing were in accordance with AS 3600 and the project specification.', severity: 'observation' },
      { text: 'BT signed off the pour at <time> with the contractor and concrete supplier representative present.', severity: 'observation' }
    ]
  },

  /* ════════ MASONRY ════════ */

  'blockwork-wall-reinf': {
    name: 'Blockwork Wall Vertical Reinforcement',
    siteReadiness: [
      'Vertical reinforcement in place per typ detail',
      'Lap lengths verifiable',
      'Starter bars from below tied securely',
      'Clean-out blocks at base of all cores to be filled',
      'Site access'
    ],
    onSiteChecks: [...COMMON_BLOCKWORK_VERT_REO_CHECKS],
    pourDayRecords: [],
    certifications: [...COMMON_FORM12_CERTS],
    standardComments: [
      { text: 'Block wall vertical reinforcement is in accordance with the structural drawings and ready for core fill.', severity: 'observation' },
      { text: 'Vertical reo has lifted out of position at <location> — to be re-fixed prior to core fill.', severity: 'defect' }
    ]
  },

  'blockwork-core-fill': {
    name: 'Blockwork Core-Fill Pre-Pour',
    siteReadiness: [
      'Cores swept clean via clean-out blocks',
      'Vertical reo still in correct position',
      'Concrete delivery scheduled (f\'c, slump per spec)',
      'Lift height ≤ 2400 mm planned',
      'Site access'
    ],
    onSiteChecks: [...COMMON_BLOCKWORK_CORE_FILL_CHECKS],
    pourDayRecords: [
      { text: 'Concrete delivery docket', captureType: 'photo' },
      { text: 'Slump test (230 mm max)', captureType: 'measure' },
      { text: 'Anchor load test results (10% chemical anchors @ 1.5 × SWL)', captureType: 'verify' }
    ],
    certifications: [
      { text: 'Anchor manufacturer load test report (10% sample @ 1.5 × SWL)', providedBy: 'Anchor manufacturer rep', noteRef: 'CM15' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Block wall core fill placement is in accordance with the structural drawings and AS 3700; lift heights, slump and grout consolidation observed and verified.', severity: 'observation' },
      { text: 'Mortar droppings remain in cores at <location> — to be cleaned out via clean-out blocks prior to core fill.', severity: 'defect' }
    ]
  },

  /* ════════ STEEL ════════ */

  'baseplate-anchor': {
    name: 'Baseplate / Anchor Bolt Inspection',
    siteReadiness: [
      'Anchor bolts cast into position',
      'Templates removed, bolts plumb',
      'Threads clean and undamaged',
      'Levelling shims and grout pads ready',
      'Steel columns ready for erection (or just before)'
    ],
    onSiteChecks: [
      { text: 'Anchor bolt position matches plan (set-out tolerance ≤ ±5 mm)', critical: true },
      { text: 'Bolt projection per detail', critical: true },
      { text: 'Bolt grade — 4.6 (foundation) / 8.8 (structural) per spec', noteRef: 'S15' },
      { text: 'Threads clean, two threads minimum past nut after tightening', noteRef: 'S6', critical: true },
      { text: 'Templates removed; bolts plumb' },
      { text: 'Non-shrink grout pad 30 mm thick @ ≥ 40 MPa', noteRef: 'S18', critical: true },
      { text: 'Plate / anchor rod assemblies cast in concrete: HDG600 to AS 4680:2006', noteRef: 'S10', asClauseRef: 'AS 4680:2006' },
      { text: 'Members in contact with concrete passivated', noteRef: 'S10' }
    ],
    pourDayRecords: [
      { text: 'Anchor torque records (post-tightening)', captureType: 'verify' },
      { text: 'Photo of finished baseplate + grout pad', captureType: 'photo' }
    ],
    certifications: [
      { text: 'AS 1252:1996 compliance certs for all bolts', providedBy: 'Steel supplier', noteRef: 'S15', critical: true },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Baseplate position, anchor projection and grout pad are in accordance with the structural drawings.', severity: 'observation' },
      { text: 'Anchor bolt at <location> outside set-out tolerance — review for column erection acceptability.', severity: 'defect' }
    ]
  },

  'steel-frame-erection': {
    name: 'Steel Frame Erection',
    siteReadiness: [
      'Steel members on site, condition-checked',
      'Erection plan submitted',
      'Crane / lifting equipment on site',
      'Steel temporary propping installed (cert by others)',
      'Site access for inspection during/after erection'
    ],
    onSiteChecks: [
      { text: 'Member sizes match framing plan', critical: true },
      { text: 'Steelwork grades per S1 spec (HR 300 / RHS&SHS 350 / CHS 250 / cold-formed 450)', noteRef: 'S1', critical: true },
      { text: 'Beam camber as noted on drawings (if specified)', noteRef: 'S2' },
      { text: 'Hollow section ends capped + vent holes if HDG', noteRef: 'S3' },
      { text: 'Plates / cleats min 10 mm UNO', noteRef: 'S6' },
      { text: 'Min 2 bolts per steel-to-steel connection', noteRef: 'S6', critical: true },
      { text: 'Bolts: M16 8.8/S < 250 mm depth, M20 8.8/S ≥ 250 mm', noteRef: 'S6' },
      ...COMMON_STEEL_CORROSION_CHECKS,
      { text: 'Construction & fabrication categories: IL2 / SC1 / FC1 / CC2 per AS/NZS 5131', asClauseRef: 'AS/NZS 5131', noteRef: 'S20' },
      { text: 'No bolt holes enlarged during erection', noteRef: 'S13' }
    ],
    pourDayRecords: [],
    certifications: [
      { text: 'AS 1252:1996 bolt compliance certs', providedBy: 'Steel supplier', noteRef: 'S15', critical: true },
      { text: 'AS/NZS 3679.1 / 3679.2 / 1163 fabricator material certs', providedBy: 'Steel fabricator', noteRef: 'S16', critical: true },
      { text: 'ACRS cert for steel supplier (overseas-sourced steel)', providedBy: 'Steel supplier', noteRef: 'S15' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Erection is at the stage where primary framing is plumb and laterally stable, subject to rectification of items listed above.', severity: 'observation' },
      { text: 'All structural steel members verified on site against the framing plan; member sizes correct and connections per typ details.', severity: 'observation' }
    ]
  },

  'steel-connection': {
    name: 'Steel Connection Inspection',
    siteReadiness: [
      'All connections complete (bolts torqued, welds finished)',
      'Site welds gouged/ground per AS 1554 (if any)',
      'NDT testing scheduled per W-table',
      'Galv damage repair (WATTYL Galvit) ready',
      'Access to all connection locations (scaffold/EWP)'
    ],
    onSiteChecks: [
      ...COMMON_STEEL_CONNECTION_CHECKS,
      { text: 'Site welds at locations specified in drawings only (W7)', noteRef: 'W7' },
      { text: 'Welder qualifications per AS 1554 §4.12.2', asClauseRef: 'AS 1554 §4.12.2', noteRef: 'W10' },
      { text: 'Welding supervisor qualification per AS 1554 §4.12.1', asClauseRef: 'AS 1554 §4.12.1', noteRef: 'W9' },
      { text: 'Welding procedures (AS 1554.1 App C) submitted before fabrication', asClauseRef: 'AS 1554.1 App C', noteRef: 'W9' },
      { text: 'Butt welds: root layer gouged or backing strip used; run-on/off plates removed', noteRef: 'W11' }
    ],
    pourDayRecords: [
      { text: 'NDT test reports (visual, MPI, UT, RT per W-table)', captureType: 'verify' },
      { text: 'Bolt torque records', captureType: 'verify' }
    ],
    certifications: [
      { text: 'NATA-approved testing authority NDT report', providedBy: 'NATA-approved testing authority', noteRef: 'W6', critical: true },
      { text: 'AS/NZS 1554.1 weld testing compliance', providedBy: 'NATA lab', noteRef: 'W6' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'All bolted and welded connections inspected and verified to comply with AS 4100, AS 1554 and the structural drawings.', severity: 'observation' },
      { text: 'Site weld at <location> shows incomplete fusion / undersized — to be re-welded or remediated per AS 1554.1.', severity: 'defect' }
    ]
  },

  /* ════════ TIMBER ════════ */

  'timber-framing': {
    name: 'Timber Framing Inspection',
    siteReadiness: [
      'All visible framing complete',
      'Roof / wall sheathing not yet applied',
      'Site access'
    ],
    onSiteChecks: [
      { text: 'Member sizes per framing plan', critical: true },
      { text: 'Studs, plates, noggins at correct centres', critical: true },
      { text: 'Lintels per span schedule' },
      { text: 'Hold-down straps / framing anchors at every required location', critical: true },
      { text: 'Bracing per AS 1684 / engineered design', asClauseRef: 'AS 1684' },
      { text: 'Timber free of structural defects (visible splits, knots in critical zones)' },
      { text: 'Treated to required hazard class (typ. H2 / H3)' }
    ],
    pourDayRecords: [],
    certifications: [...COMMON_FORM12_CERTS],
    standardComments: [
      { text: 'All visible timber members are straight and free of apparent structural defects.', severity: 'observation' }
    ]
  },

  'mass-timber-install': {
    name: 'Mass Timber (GLT/CLT) Installation',
    siteReadiness: [
      'CLT/GLT panels on site, condition-checked',
      'Supporting slab tolerance survey complete (±10 mm level / 6 mm flatness)',
      'Proprietary connectors and fasteners delivered',
      'Manufacturer cert + moisture content readings <18%',
      'Site access (crane/EWP)',
      '48-hour notice given (CLT22)'
    ],
    onSiteChecks: [
      { text: 'Panel positions match layout plan; joints per typ details', critical: true },
      { text: 'Proprietary connectors and fasteners as scheduled' },
      { text: 'Edge distances per typ details', critical: true },
      { text: 'No unauthorised penetrations >100 mm; <100 mm penetrations ≥ 400 mm from panel edge' },
      { text: 'No panel joints in lintels or within 1 m of openings', critical: true },
      { text: 'Quality cert from manufacturer provided' },
      { text: 'Moisture content <18%; end-grain sealer applied', critical: true },
      { text: 'Site storage protected from weather' }
    ],
    pourDayRecords: [
      { text: 'Moisture meter readings (multiple locations)', captureType: 'measure' },
      { text: 'Photo of completed installation per storey', captureType: 'photo' }
    ],
    certifications: [
      { text: 'Manufacturer quality cert', providedBy: 'CLT manufacturer (e.g. NeXTimber/Timberlink)', critical: true },
      { text: 'Moisture management plan + readings', providedBy: 'Builder' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'CLT panel installation is in accordance with the manufacturer\'s drawings and the structural drawings; moisture content within acceptable limits at the time of inspection.', severity: 'observation' },
      { text: 'Re-tensioning of bolted connections at service moisture is recommended per AS 1720.1:2010 Cl. 4.4.2.', severity: 'observation' }
    ]
  },

  'bracing': {
    name: 'Bracing Inspection',
    siteReadiness: [
      'All wall + roof bracing installed',
      'Tie-down complete',
      'Bracing connections complete',
      'Site access'
    ],
    onSiteChecks: [
      { text: 'Bracing locations match bracing plan', critical: true },
      { text: 'Bracing types match schedule (steel angle, strap, ply, etc.)' },
      { text: 'Connections to top/bottom plates per typ detail', critical: true },
      { text: 'No interference with services / penetrations' },
      { text: 'Continuity of load path to foundations verified', critical: true }
    ],
    pourDayRecords: [],
    certifications: [...COMMON_FORM12_CERTS],
    standardComments: [
      { text: 'Wall and roof bracing has been installed in accordance with the structural drawings and provides the lateral load resistance required.', severity: 'observation' }
    ]
  },

  /* ════════ OTHER ════════ */

  'waterproofing-precover': {
    name: 'Waterproofing Pre-Cover',
    siteReadiness: [
      'Membrane fully installed across area',
      'Penetrations sealed',
      'No protection board / backfill placed yet',
      'Site access'
    ],
    onSiteChecks: [
      { text: 'Membrane integrity — no holes, tears, voids', critical: true },
      { text: 'Lap widths per manufacturer spec (typ. 100 mm min)', critical: true },
      { text: 'Terminations at penetrations (pipes, services) sealed', critical: true },
      { text: 'Upstand / fillet at perimeter walls per typ detail' },
      { text: 'Manufacturer\'s warranty applicable to installation' },
      { text: 'Photographic record of complete area before covering', critical: true }
    ],
    pourDayRecords: [
      { text: 'Photographic record of complete membrane area', captureType: 'photo' }
    ],
    certifications: [
      { text: 'Waterproofing applicator certificate', providedBy: 'Waterproofing applicator' },
      { text: 'Manufacturer warranty', providedBy: 'Membrane manufacturer' },
      ...COMMON_FORM12_CERTS
    ],
    standardComments: [
      { text: 'Waterproofing membrane has been installed and inspected prior to cover; no defects observed at the time of inspection.', severity: 'observation' },
      { text: 'Membrane puncture / tear at <location> — to be repaired prior to backfill / protection board.', severity: 'defect' }
    ]
  },

  'adhoc': {
    name: 'Ad-hoc (free form)',
    siteReadiness: [],
    onSiteChecks: [],
    pourDayRecords: [],
    certifications: [...COMMON_FORM12_CERTS],
    standardComments: []
  }
};


/* ──────────────────────────────────────────────────────────────────────────
 * UTILITIES
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Get the standard checks bundle for an inspection type.
 * Falls back to the 'adhoc' empty bundle for unknown types.
 */
export function getStandardChecksForType(typeKey) {
  return BT_STANDARD_CHECKS[typeKey] || BT_STANDARD_CHECKS['adhoc'];
}

/**
 * Combine the BT standard checks with project-specific checks from a Cowork
 * inspection plan entry. Returns a unified `{siteReadiness, onSiteChecks,
 * pourDayRecords, certifications}` view, with each item tagged as 'standard'
 * or 'project' so the UI can show provenance.
 *
 * Project-specific items come FIRST in each list (they're more tailored).
 *
 * @param {string} typeKey   — inspection-type catalogue key
 * @param {Object} planEntry — the inspectionPlan entry from project-map (or null)
 * @returns {Object}
 */
export function mergeWithProjectChecks(typeKey, planEntry) {
  const standard = getStandardChecksForType(typeKey);
  const project = planEntry || {};

  // Project-specific site readiness (from entry.siteReadinessCheck if present)
  const siteReadinessProject = (project.siteReadinessCheck || []).map((q) => ({
    text: typeof q === 'string' ? q : (q.text || ''),
    source: 'project'
  }));
  const siteReadinessStandard = standard.siteReadiness.map((q) => ({
    text: q,
    source: 'standard'
  }));
  // De-dupe — project supersedes standard if same text
  const seenSiteText = new Set(siteReadinessProject.map((s) => s.text.toLowerCase()));
  const siteReadiness = [
    ...siteReadinessProject,
    ...siteReadinessStandard.filter((s) => !seenSiteText.has(s.text.toLowerCase()))
  ];

  // Project expected checklist may be plain strings or {text, asClauseRef, noteRef}
  const onSiteProject = (project.expectedChecklist || []).map((c) => {
    if (typeof c === 'string') return { text: c, source: 'project' };
    return { ...c, source: 'project' };
  });
  const onSiteStandard = standard.onSiteChecks.map((c) => ({ ...c, source: 'standard' }));
  const seenOnSiteText = new Set(onSiteProject.map((c) => (c.text || '').toLowerCase()));
  const onSiteChecks = [
    ...onSiteProject,
    ...onSiteStandard.filter((c) => !seenOnSiteText.has((c.text || '').toLowerCase()))
  ];

  // Pour-day records — project entry may have postPourRecord; merge with standard
  const pourDayProject = [];
  if (project.postPourRecord) {
    if (project.postPourRecord.concreteGrade)
      pourDayProject.push({ text: `Concrete grade: ${project.postPourRecord.concreteGrade}`, source: 'project' });
    if (project.postPourRecord.expectedSlump)
      pourDayProject.push({ text: `Expected slump: ${project.postPourRecord.expectedSlump} mm`, source: 'project' });
    if (project.postPourRecord.tempRange)
      pourDayProject.push({ text: `Allowable pour temp: ${project.postPourRecord.tempRange}`, source: 'project' });
    for (const it of (project.postPourRecord.captureItems || [])) {
      pourDayProject.push({ text: it, source: 'project' });
    }
  }
  const pourDayStandard = standard.pourDayRecords.map((r) => ({ ...r, source: 'standard' }));
  const seenPourText = new Set(pourDayProject.map((r) => (r.text || '').toLowerCase()));
  const pourDayRecords = [
    ...pourDayProject,
    ...pourDayStandard.filter((r) => !seenPourText.has((r.text || '').toLowerCase()))
  ];

  return {
    siteReadiness,
    onSiteChecks,
    pourDayRecords,
    certifications: standard.certifications,
    standardComments:  standard.standardComments
  };
}
