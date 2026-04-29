"""Build inspection-program.html for MBC Creativity & Arts Centre.

Adapted from the Emmanuel build script. Project specifics:
  - 2-storey commercial school arts centre (G + L1)
  - Bored pier + pad/strip footing foundation (300 kPa) — PTG Consulting geotech
  - Class M reactive soil → raft slabs designed
  - Post-tensioned suspended slabs at L1 (and possibly GF for plant slab)
  - Steel roof in THREE tiers: Lower / Main / Upper
  - Block retaining walls (RW2/3/4) at perimeter
  - Atrium feature stair (S040/S041)
  - TENDER ISSUE — flag everything as provisional pending Construction Issue
"""
import json, base64, os, datetime, html as html_lib
from pathlib import Path

ROOT = Path("/sessions/compassionate-laughing-fermat/mnt/Structural-Inspections/Structural Drawings/MBC")
ISO_PATH = ROOT / "isometric.png"
OUT_HTML = ROOT / "inspection-program.html"

# ─────────────────────────────────────────────────────────────────────────────
# HOTSPOT POLYGONS — over the MBC isometric (2400x1404)
# Coords are PERCENTAGE of image (0-100), used in SVG viewBox="0 0 100 100"
# ─────────────────────────────────────────────────────────────────────────────

HOTSPOTS = {
    # ══════════ SUBSTRUCTURE — bored piers + pad/strip footings ══════════
    "footings-bored-piers": {
        "label": "Bored piers (under columns + retaining walls)",
        "polygons": [
            # Visible pier locations at the base — under columns
            "12,82 17,82 17,90 12,90",
            "22,82 27,82 27,90 22,90",
            "33,82 38,82 38,90 33,90",
            "42,82 47,82 47,90 42,90",
            "52,82 57,82 57,90 52,90",
            "62,82 67,82 67,90 62,90",
            "72,82 77,82 77,90 72,90",
            "82,82 87,82 87,90 82,90",
            "90,82 95,82 95,90 90,90",
        ],
    },
    "footings-pad-strip": {
        "label": "Pad / strip footings + raft slab (Class M)",
        "polygons": [
            # Wider band representing the founding-level perimeter
            "8,80 92,80 92,92 8,92",
        ],
    },

    # ══════════ GROUND FLOOR ══════════
    "ground-slab-on-ground": {
        "label": "Ground floor slab on ground + plant slab",
        "polygons": [
            "8,73 95,73 95,82 8,82",
        ],
    },
    "ground-walls-concrete": {
        "label": "Concrete walls — Ground level",
        "polygons": [
            # Visible solid walls at ground, perimeter and internal
            "75,55 78,55 78,78 75,78",
            "82,55 85,55 85,78 82,78",
            "88,55 91,55 91,78 88,78",
            "10,60 13,60 13,78 10,78",
        ],
    },
    "ground-walls-block": {
        "label": "Block retaining walls — Ground (RW2/RW3/RW4)",
        "polygons": [
            # Block retaining walls at perimeter
            "5,68 95,68 95,80 5,80",
        ],
    },
    "ground-columns": {
        "label": "Concrete columns — Ground level (red columns visible)",
        "polygons": [
            # Visible pink/red columns rising from ground to L1
            "15,55 17,55 17,75 15,75",
            "22,55 24,55 24,75 22,75",
            "30,55 32,55 32,75 30,75",
            "38,55 40,55 40,75 38,75",
            "46,55 48,55 48,75 46,75",
            "54,55 56,55 56,75 54,75",
            "62,55 64,55 64,75 62,75",
            "70,55 72,55 72,75 70,75",
        ],
    },

    # ══════════ LEVEL 1 (suspended slab — PT) ══════════
    "l1-slab": {
        "label": "Level 1 suspended slab (post-tensioned)",
        "polygons": [
            "10,40 92,40 92,55 10,55",
        ],
    },
    "l1-pt": {
        "label": "Level 1 post-tension strand zones",
        "polygons": [
            # PT regions visible by orange/red highlighting on the L1 slab
            "12,42 50,42 50,52 12,52",
            "55,42 88,42 88,52 55,52",
        ],
    },
    "l1-walls": {
        "label": "Concrete + block walls — Level 1",
        "polygons": [
            "70,28 78,28 78,42 70,42",
            "82,28 90,28 90,42 82,42",
        ],
    },
    "l1-columns": {
        "label": "Concrete columns — Level 1 (continuing to roof)",
        "polygons": [
            "15,28 17,28 17,42 15,42",
            "22,28 24,28 24,42 22,42",
            "30,28 32,28 32,42 30,42",
            "38,28 40,28 40,42 38,42",
            "46,28 48,28 48,42 46,42",
            "54,28 56,28 56,42 54,42",
            "62,28 64,28 64,42 62,42",
        ],
    },
    "l1-baseplate-anchor": {
        "label": "Steel column baseplates / hold-down bolts at L1",
        "polygons": [
            # At top of L1 slab where steel cols pick up the roof
            "20,38 24,38 24,42 20,42",
            "32,38 36,38 36,42 32,42",
            "44,38 48,38 48,42 44,42",
            "56,38 60,38 60,42 56,42",
            "68,38 72,38 72,42 68,42",
        ],
    },

    # ══════════ ROOF — THREE TIERS ══════════
    "lower-roof": {
        "label": "Lower roof framing + purlins (front/left)",
        "polygons": [
            "12,22 50,22 50,38 12,38",
        ],
    },
    "main-roof": {
        "label": "Main roof framing + purlins (centre)",
        "polygons": [
            "30,8 75,8 75,28 30,28",
        ],
    },
    "upper-roof": {
        "label": "Upper roof framing + purlins (right tier)",
        "polygons": [
            "70,3 95,3 95,22 70,22",
        ],
    },
    "all-roof-connections": {
        "label": "All roof connections (across lower/main/upper)",
        "polygons": [
            "12,3 95,3 95,30 12,30",
        ],
    },

    # ══════════ ATRIUM STAIRS (feature) ══════════
    "atrium-stairs": {
        "label": "Atrium Stairs (S040/S041 — feature stair)",
        "polygons": [
            # Distinct atrium stair location in centre-front
            "37,42 50,42 50,75 37,75",
        ],
    },

    # ══════════ MISC STEELWORK ══════════
    "misc-steelwork": {
        "label": "Misc steelwork — facade brackets, balustrades, screens",
        "polygons": [
            "5,30 100,30 100,80 5,80",
        ],
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# PROJECT METADATA + GENERAL NOTES — MBC Creativity & Arts Centre
# Module-level constants consumed by tools/build_btproject.py
# ─────────────────────────────────────────────────────────────────────────────

PROJECT_META = {
    "name":            "MBC Creativity & Arts Centre",
    "siteAddress":     "Moreton Bay College, 450 Wondall Road, Manly West, QLD 4179",
    "jobNumber":       "2024.0230",
    "client":          "Moreton Bay College",
    "issueStatus":     "TENDER NOT FOR CONSTRUCTION",
    "engineerOfRecord": "Bligh Tanner",
    "discipline":      "structural",
    "builder":         None,
    "architect":       None,
    "rpeqSignatory":   None,
}

GENERAL_NOTES = {
    "designCodes": [
        "AS 1720", "AS 2159", "AS 3600", "AS 3610", "AS 3700", "AS 4100",
        "AS 2269", "NCC", "AS 5216:2018", "AS 1554", "AS 1252",
        "AS/NZS 3679.1", "AS/NZS 3679.2", "AS/NZS 1163",
        "AS/NZS 5131", "AS/NZS 4680", "AS 1170.4", "AS 4671", "AS 2870"
    ],
    "exposureClass": None,
    "concreteCover": {
        "boredPiers":            {"bottom": 50,   "top": None, "sides": 65},
        "footings":              {"bottom": 50,   "top": 50,   "sides": 50},
        "slabOnGroundInternal":  {"bottom": 40,   "top": 30,   "sides": 40},
        "slabOnGroundExternal":  {"bottom": 40,   "top": 40,   "sides": 40},
        "columns":               {"bottom": 40,   "top": 40,   "sides": 40},
        "walls":                 {"bottom": None, "top": 40,   "sides": 40},
        "stairs":                {"bottom": 40,   "top": 30,   "sides": 40},
        "suspendedSlab":         {"bottom": 30,   "top": 30,   "sides": 30},
        "suspendedSlabPT":       {"bottom": 30,   "top": 30,   "sides": 30}
    },
    "concreteStrengths": {
        "unit": "MPa",
        "boredPiers":   32, "footings": 32, "slabOnGround": 32,
        "columns":      40, "walls": 40, "stairs": 40,
        "suspendedSlab": 40, "suspendedSlabPT": 40
    },
    "bearingCapacity": {
        "unit": "kPa",
        "padFootings":   300,
        "stripFootings": 300,
        "boredPiers":    None,    # bored piers per S010 schedule (3000 mm, Ø450)
        "shaftAdhesion": None
    },
    "windRegion":      "B",
    "windVelocity":    {"ultimate": 60, "serviceability": 39, "unit": "m/s"},
    "terrainCategory": 3,
    "importanceLevel": 3,
    "earthquake": {
        "hazardZ": 0.08, "category": "II",
        "probabilityKp": 1.3, "subSoilClass": "Ce"
    },
    "geotechReport": {
        "consultant":   "PTG Consulting",
        "reportNumber": "PTG/00371",
        "date":         "2024-07"
    },
    "specialNotes": [
        "Form 12 certification: Builder must give 24h notice for inspections; failure to notify excludes works from certification.",
        "Pad & strip footings to be founded min 500 mm below FGL, at least 300 mm into duricrust @ 300 kPa allowable bearing (note F5).",
        "Provisionally allow 25 MPa mass concrete under footings to reach bearing material (note F5).",
        "Raft slabs designed for Class M reactive soil per AS 2870 (note F6).",
        "Bored piers: 3000 mm min depth, 450 mm dia (per S010 footing schedule).",
        "10% of all chemical anchors into core-filled blockwork load-tested to 1.5 × SWL (note CM15).",
        "Curing: keep concrete surfaces continuously wet for 3 days, prevent moisture loss for 7 days (note C5).",
        "Trimmer bars at re-entrant corners and penetrations >200 sq: 2-N12 × 1200 long @ 100 ctrs, top & bottom (note C6).",
        "Pour temperature 5–35 °C (note C19).",
        "Block walls: vertical control joints @ 8 m max, 5 m max from corners, not within 1.2 m of corners (note CM6).",
        "Block walls: no back filling behind retaining walls until 14 days after core fill (note CM10).",
        "ACRS certification required for all structural steel (note S15)."
    ]
}

EXCLUDED_FROM_BT = [
    {
        "element": "Structural steelwork — temporary propping & bracing",
        "responsibility": "Contractor RPEQ (temporary works engineer)",
        "noteRef": "S18 & G2",
        "scope": "Temporary works during erection"
    },
    {
        "element": "Light-gauge steel stud framing",
        "responsibility": "Manufacturer RPEQ (Form 16 on completion)",
        "noteRef": "ST1",
        "scope": "Sizing, spacing, and ALL connections to permanent structure"
    },
    {
        "element": "Roof safety systems",
        "responsibility": "Manufacturer / supplier",
        "noteRef": "(general note table)",
        "scope": "System and connections to roof, including verification of load paths to permanent bracing"
    }
]

WARNINGS = [
    "TENDER ISSUE — inspection program is provisional pending Construction Issue PDF. Re-run when Construction Issue arrives.",
    "MBC uses NEWER BT A1 template variant (per skills/reading-bt-drawings/SKILL.md Section 10) — extractor cells differ from older 52SA/Emmanuel template.",
    "Bored pier bearing capacity not explicitly stated in F1 schedule extracted — verify on site by geotech RPEQ. Pad/strip footing bearing 300 kPa per F5.",
    "Steel column hold-down bolts: extracted plan shows them at L1 slab level (slab-on-ground supports the steel roof structure). Verify location plan with engineer.",
]

TEMPLATES_DETECTED = ["BT A1 (newer variant, 2024+)"]

# ─────────────────────────────────────────────────────────────────────────────
# INSPECTION PLAN — MBC Creativity & Arts Centre
# ─────────────────────────────────────────────────────────────────────────────

PLAN = []
seq = 0

def add(**kw):
    global seq
    seq += 1
    kw["sequence"] = seq
    PLAN.append(kw)

# ── Phase 1: Substructure ──────────────────────────────────────────────────
add(
    type="pile-bored-install",
    title="Bored Pier Installation — Founding Depth & Shaft",
    phase="Substructure", level="ground",
    hotspots=["footings-bored-piers"],
    holdPoint=True, severity="hold-point", stage="post-install",
    drawings=[
        ("S010", "Footing Plan Notes, Legends, & Schedules"),
        ("S011", "Footing and Plant Slab Plan"),
        ("S015", "Footing Details — Sheet 1"),
        ("S016", "Footing Details — Sheet 2"),
    ],
    rationale=(
        "Bored piers per AS 2159. Founding depth (min 3000 mm), socket depth (min 300 mm) "
        "and 450 mm pier diameter confirmed on site by RPEQ-certified geotech (PTG Consulting "
        "ref PTG/00371, July 2024). Hold point with the geotech RPEQ; BT advisory."
    ),
    checklist=[
        "Pile location ≤ 75 mm of designated position (note P3)",
        "Founding depth recorded per pier; min 3000 mm depth, 300 mm socket into duricrust",
        "Pier diameter 450 mm; verify reo cage clear of formwork",
        "Founding material verified by geotech RPEQ on site",
        "Pier extension ≥ 75 mm into pile cap or ground beam (note P7)",
        "Piling Contractor RPEQ certificate confirming design loads achieved",
    ],
    scopeNote="Geotech RPEQ holds the certification hold point. BT advisory only.",
)

add(
    type="subgrade",
    title="Pad & Strip Footing Subgrade — Bearing Verification",
    phase="Substructure", level="ground",
    hotspots=["footings-pad-strip"],
    holdPoint=True, severity="hold-point", stage="pre-reinforcement",
    drawings=[
        ("S010", "Footing Plan Notes, Legends, & Schedules"),
        ("S011", "Footing and Plant Slab Plan"),
        ("S015", "Footing Details — Sheet 1"),
        ("S016", "Footing Details — Sheet 2"),
    ],
    rationale=(
        "Pad and strip footings to be founded min 500 mm below FGL, at least 300 mm into "
        "duricrusted material with 300 kPa allowable bearing (note F5). RPEQ geotech to confirm "
        "on site. Provisionally allow for 25 MPa mass concrete under footings to reach bearing."
    ),
    checklist=[
        "Excavation min 500 mm below FGL",
        "Min 300 mm penetration into duricrusted material",
        "Allowable bearing 300 kPa confirmed by site geotech (note F5)",
        "Mass concrete (25 MPa) added if required to reach bearing material",
        "Excavation maintained firm and dry (note F2)",
        "Soft ground replaced with mass concrete (note F2)",
    ],
)

add(
    type="pad-footing-prepour",
    title="Pad & Strip Footings — Pre-Pour Reinforcement",
    phase="Substructure", level="ground",
    hotspots=["footings-pad-strip", "footings-bored-piers"],
    drawings=[
        ("S010", "Footing Plan Notes, Legends, & Schedules"),
        ("S011", "Footing and Plant Slab Plan"),
        ("S015", "Footing Details — Sheet 1"),
        ("S016", "Footing Details — Sheet 2"),
    ],
    rationale=(
        "Pad and strip footings include pile caps over bored piers AND footings on duricrust. "
        "Multiple footing types per the schedule on S010. First reinforced cast on site."
    ),
    checklist=[
        "Bottom cover ≥ 50 mm, sides ≥ 50 mm (cover schedule, footings element, N32)",
        "Pier reo extends ≥ 75 mm into pile cap (note P7)",
        "Lap lengths at corners and T-junctions per AS 3600 / typical detail",
        "Trimmer bars at penetrations >200 sq (note C6)",
        "Concrete grade N32",
        "Bar chairs and ties — reo securely tied (note C18)",
        "Foundation excavation maintained firm & dry (note F2)",
    ],
)

# ── Phase 2: Ground Floor ──────────────────────────────────────────────────
add(
    type="slab-prepour-ground",
    title="Ground Floor Slab on Ground + Plant Slab — Pre-Pour Reinforcement",
    phase="Ground Floor", level="ground",
    hotspots=["ground-slab-on-ground"],
    drawings=[
        ("S011", "Footing and Plant Slab Plan"),
        ("S100", "GF Plan Notes, Legends, & Schedules"),
        ("S101", "GF General Arrangement Plan"),
        ("S121", "GF Bottom Reinforcement Plan"),
        ("S131", "GF Top Reinforcement Plan"),
        ("S140", "GF Post Tensioning Plan"),
        ("S151", "GF Slab Sections — Sheet 1"),
        ("S152", "GF Slab Sections — Sheet 2"),
        ("S153", "GF Slab Sections — Sheet 3"),
        ("S171", "GF Loading Plan"),
        ("S050", "Typical Suspended Slab Details"),
    ],
    rationale=(
        "Ground floor slab combines slab-on-ground (for most of the building) with the "
        "plant slab area. Note: a Post-Tensioning Plan (S140) exists for Ground Floor — "
        "verify whether GF has PT regions (raft) or only the L1 slab does."
    ),
    checklist=[
        "Internal SOG: bottom 40, top 30, sides 40 mm cover (cover schedule, N32)",
        "External SOG: bottom 40, top 40, sides 40 mm cover",
        "Concrete grade N32",
        "50 mm bedding sand + DPM (note C16)",
        "Bottom reo per S121, top reo per S131",
        "Mesh laps: 2 outer-most cross bars overlap (note C7)",
        "Trimmer bars at re-entrant corners and penetrations >200 sq",
        "Bar chairs @ 600 ctrs SL72/82, 800 ctrs SL92+ (note C9)",
        "Raft slab designed for Class M reactive soil per AS 2870 (note F6)",
    ],
)

add(
    type="post-tension-strand",
    title="Ground Floor Post-Tension Strand Placement (if applicable)",
    phase="Ground Floor", level="ground",
    hotspots=["ground-slab-on-ground"],
    drawings=[
        ("S140", "GF Post Tensioning Plan"),
        ("S055", "Typical Post Tensioning Details — Sheet 1"),
        ("S056", "Typical Post Tensioning Details — Sheet 2"),
    ],
    rationale=(
        "S140 indicates a GF post-tensioning plan exists — possibly for a transfer slab "
        "or a portion of the GF slab. Inspect strand profile, anchorage zones, dead/live ends "
        "before pour."
    ),
    checklist=[
        "Strand profile matches PT designer's layout (S140)",
        "Anchorage zone reinforcement per typical detail S055/S056",
        "Tendon spacing verified",
        "Dead-end / live-end stressing pockets free of obstruction",
        "PT designer to sign off prior to pour",
    ],
    scopeNote="Conditional — confirm whether GF actually has PT regions or only L1.",
)

add(
    type="concrete-wall-prepour",
    title="Concrete Walls — Ground Pre-Pour Reinforcement",
    phase="Ground Floor", level="ground",
    hotspots=["ground-walls-concrete"],
    drawings=[
        ("S020", "Typical Concrete Wall Details — Sheet 1"),
        ("S101", "GF GA — wall locations"),
    ],
    rationale="In-situ concrete walls at ground (lift core, blade walls, etc.).",
    checklist=[
        "Sides cover ≥ 40 mm; top cover ≥ 40 mm (cover schedule, walls, N40)",
        "Concrete grade N40",
        "Vertical and horizontal reinforcement per S020 typical",
        "Lap lengths per AS 3600",
        "Starter bars from below securely tied",
        "Formwork plumb, clean, properly propped",
    ],
    scopeNote="Catalogue 'retaining-wall-prepour' used as proxy for in-situ concrete wall.",
)

add(
    type="blockwork-wall-reinf",
    title="Block Retaining Walls — Vertical Reinforcement (RW2/RW3/RW4)",
    phase="Ground Floor", level="ground",
    hotspots=["ground-walls-block"],
    drawings=[
        ("S025", "Typical Block Wall Details"),
        ("S026", "Typical Block Retaining Wall Details — Sheet 1"),
        ("S027", "Typical Block Retaining Wall Details — Sheet 2"),
    ],
    rationale=(
        "Multiple block retaining wall types (RW2/RW3/RW4 per the wall schedule on S010). "
        "Vertical reinforcement and starter bars must be inspected before core fill."
    ),
    stage="pre-cover",
    checklist=[
        "Vert reo size, spacing, position per wall type detail (RW2/RW3/RW4)",
        "Starter bars securely tied prior to wall footing pour (note CM10)",
        "Lap lengths to AS 3700",
        "Mortar mix M3 (general) or M4 (retaining walls) per spec",
        "Vertical control joints at ≤ 8 m, 5 m max from corners, not within 1.2 m of corners (note CM6)",
        "Clean-out blocks at all cores to be filled (note CM3)",
        "No back filling behind retaining walls until 14 days after core fill (note CM10)",
    ],
)

add(
    type="blockwork-core-fill",
    title="Block Retaining Walls — Core Fill Pre-Pour",
    phase="Ground Floor", level="ground",
    hotspots=["ground-walls-block"],
    drawings=[
        ("S025", "Typical Block Wall Details"),
        ("S026", "Typical Block Retaining Wall — Sheet 1"),
        ("S027", "Typical Block Retaining Wall — Sheet 2"),
    ],
    rationale="Core fill of block retaining walls after vert reo inspection.",
    checklist=[
        "Concrete strength f'c = 20 MPa, max slump 230 mm, max aggregate 10 mm",
        "Lift height ≤ 2400 mm",
        "All cores swept clean via clean-out blocks ('H' units UNO — note CM4)",
        "10% of chemical anchors load-tested to 1.5 × SWL (note CM15)",
        "Propping at top of wall @ 1800 ctrs where slab installed over (note CM12)",
    ],
)

add(
    type="column-prepour",
    title="Concrete Columns — Ground Pre-Pour Reinforcement",
    phase="Ground Floor", level="ground",
    hotspots=["ground-columns"],
    drawings=[
        ("S035", "Concrete Column Details — Sheet 1"),
        ("S101", "GF GA — column locations"),
    ],
    rationale="Ground floor columns rise from pile caps/footings to support L1 slab.",
    checklist=[
        "Sides cover ≥ 40 mm (cover schedule, columns, N40)",
        "Concrete grade N40",
        "Vertical bar count, size, lap length per S035",
        "Tie/fitment spacing — closer at top & bottom; 135° hooks",
        "Starter bar projection out of pile cap matches lap",
        "Formwork plumb, clean, propped",
    ],
)

add(
    type="stair-prepour",
    title="Atrium Stairs — G to L1 Pre-Pour Reinforcement",
    phase="Ground Floor", level="ground",
    hotspots=["atrium-stairs"],
    drawings=[
        ("S040", "Atrium Stairs Framing Layout"),
        ("S041", "Atrium Stairs Framing Sections"),
        ("S045", "Stair Details — Sheet 1"),
    ],
    rationale=(
        "Feature atrium stair from ground to L1. Likely a steel-framed stair with concrete "
        "treads, or an in-situ concrete stair — verify from S040/S041. May warrant separate "
        "framing inspection if steel."
    ),
    checklist=[
        "Cover 40/30/40 mm (cover schedule, stairs, N40)",
        "Reinforcement matches stair section detail (S041)",
        "Tread/riser geometry per architect's drawings",
        "Connection to landing per typ detail",
        "If steel-framed: separate framing inspection per S041",
    ],
    scopeNote="Mapped to slab-prepour-suspended (no dedicated 'stair' catalogue key).",
)

# ── Phase 3: Level 1 (suspended slab — PT) ─────────────────────────────────
add(
    type="slab-prepour-suspended",
    title="Suspended Slab — Level 1 Head Framing & Pre-Pour Reinforcement",
    phase="Level 1", level="l1",
    hotspots=["l1-slab"],
    drawings=[
        ("S200", "FF Plan Notes, Legends, & Schedules"),
        ("S201", "FF General Arrangement Plan"),
        ("S221", "FF Bottom Reinforcement Plan"),
        ("S231", "FF Top Reinforcement Plan"),
        ("S251", "FF Slab Sections — Sheet 1"),
        ("S271", "FF Loading Plan"),
        ("S050", "Typical Suspended Slab Details"),
    ],
    rationale=(
        "L1 suspended slab — pre-pour reo. This slab IS post-tensioned (S241 PT plan exists), "
        "so this inspection captures the standard reo PLUS sets the stage for the PT inspection "
        "that follows. Includes head framing (formwork + props per AS 3610)."
    ),
    checklist=[
        "Bottom/top/sides cover 30 mm; concrete N40 (suspended slab)",
        "Bottom reo per S221, top reo per S231",
        "Trimmer bars at penetrations >200 sq, top & bottom",
        "Formwork properly propped per AS 3610",
        "Bar chairs per note C9",
        "All reo securely tied (note C18)",
        "Cast-in plates / starter bars for steel cols per S400 framing notes",
    ],
)

add(
    type="post-tension-strand",
    title="Level 1 — Post-Tension Strand Placement — HOLD POINT",
    phase="Level 1", level="l1",
    hotspots=["l1-pt"],
    holdPoint=True, severity="hold-point",
    drawings=[
        ("S241", "FF Post Tensioning Plan"),
        ("S055", "Typical Post Tensioning Details — Sheet 1"),
        ("S056", "Typical Post Tensioning Details — Sheet 2"),
    ],
    rationale=(
        "L1 post-tension strand placement is a hold point — cover-up means destructive remediation. "
        "Inspect strand profile, anchorage zones, dead/live ends and stressing pocket setup before "
        "any concrete placement. PT designer to sign off."
    ),
    checklist=[
        "Strand profile matches PT designer's layout (S241)",
        "Anchorage zone reinforcement per typ detail (S055/S056)",
        "Tendon spacing and centres verified",
        "Dead-end / live-end stressing pockets free of obstruction",
        "Sleeves and barriers in place at column heads",
        "PT designer to sign off prior to pour",
        "48 hours notice given to BT prior to pour",
    ],
)

add(
    type="concrete-wall-prepour",
    title="Concrete Walls — Level 1 Pre-Pour Reinforcement",
    phase="Level 1", level="l1",
    hotspots=["l1-walls"],
    drawings=[("S020", "Typical Concrete Wall Details"), ("S201", "FF GA — wall locations")],
    rationale="In-situ concrete walls at L1 (lift core continued, blade walls, parapets).",
    checklist=[
        "Sides/top cover ≥ 40 mm; concrete N40",
        "Vertical and horizontal reo per S020 typ detail",
        "Starter bar projection from L1 slab to suit lap",
        "Formwork plumb, propped",
    ],
    scopeNote="Catalogue proxy for concrete-wall-prepour.",
)

add(
    type="blockwork-wall-reinf",
    title="Block Walls — Vertical Reinforcement — Level 1",
    phase="Level 1", level="l1",
    hotspots=["l1-walls"],
    drawings=[("S025", "Typical Block Wall Details")],
    rationale="Block walls at L1 — internal partition / fire-rated walls.",
    stage="pre-cover",
    checklist=["Vert reo per typ detail", "Lap lengths to AS 3700", "Starter bars from L1 slab"],
)

add(
    type="blockwork-core-fill",
    title="Block Walls — Core Fill Pre-Pour — Level 1",
    phase="Level 1", level="l1",
    hotspots=["l1-walls"],
    drawings=[("S025", "Typical Block Wall Details")],
    rationale="Core fill of L1 block walls after vert reo inspection.",
    checklist=[
        "Cores swept clean via clean-out blocks",
        "Lift height ≤ 2400 mm",
        "Concrete f'c = 20 MPa, slump 230 mm, agg 10 mm",
    ],
)

add(
    type="column-prepour",
    title="Concrete Columns — Level 1 Pre-Pour Reinforcement (where present)",
    phase="Level 1", level="l1",
    hotspots=["l1-columns"],
    drawings=[("S035", "Concrete Column Details"), ("S201", "FF GA — column locations")],
    rationale="L1 columns where concrete cols continue up to roof support level.",
    checklist=[
        "Sides cover ≥ 40 mm; concrete N40",
        "Vertical bars per S035; ties + 135° hooks",
        "Starter bar projection from L1 slab",
    ],
)

add(
    type="baseplate-anchor",
    title="Steel Column Hold-Down Bolts — At L1 Slab",
    phase="Level 1", level="l1",
    hotspots=["l1-baseplate-anchor"],
    drawings=[
        ("S400", "Framing Notes, Legends, & Schedules"),
        ("S440", "Typical Framing Details — Sheet 1"),
    ],
    rationale=(
        "Steel roof columns are anchored into the L1 slab. Anchor bolt position, projection "
        "and threading must be verified before slab pour AND again before column erection. "
        "Non-shrink grout 30 mm @ 40 MPa typical."
    ),
    checklist=[
        "Anchor bolt position matches plan (set-out tolerance ≤ ±5 mm)",
        "Bolt projection per detail",
        "Bolts grade 4.6 (foundation) / 8.8 (structural) per spec",
        "Threads clean, two threads minimum past nut after tightening",
        "Templates removed; bolts plumb",
        "Non-shrink grout pad 30 mm @ ≥ 40 MPa for column erection",
    ],
)

add(
    type="stair-prepour",
    title="Atrium Stairs — L1 to Roof Pre-Pour Reinforcement",
    phase="Level 1", level="l1",
    hotspots=["atrium-stairs"],
    drawings=[
        ("S040", "Atrium Stairs Framing Layout"),
        ("S041", "Atrium Stairs Framing Sections"),
        ("S045", "Stair Details — Sheet 1"),
    ],
    rationale="Atrium stair flight from L1 to roof / mezzanine.",
    checklist=[
        "Cover 40/30/40 mm (cover schedule, stairs, N40)",
        "Reinforcement matches stair section detail",
        "Tread/riser geometry per arch",
        "Connection to landing per typ detail",
    ],
    scopeNote="Mapped to slab-prepour-suspended.",
)

# ── Phase 4: Roof — THREE TIERS ──────────────────────────────────────────
add(
    type="steel-frame-erection",
    title="Lower Roof — Steel Frame Erection",
    phase="Roof", level="roof",
    hotspots=["lower-roof"],
    drawings=[
        ("S401", "Lower Roof Framing Plan"),
        ("S400", "Framing Notes, Legends, & Schedules"),
        ("S421", "Framing Elevations — Sheet 1"),
        ("S422", "Framing Elevations — Sheet 2"),
        ("S440", "Typical Framing Details — Sheet 1"),
        ("S451", "Framing Details — Sheet 1"),
    ],
    rationale=(
        "Lower roof structure (front/lower portion of building). Steel-framed per AS 4100 / "
        "AS/NZS 5131. Construction category CC2."
    ),
    checklist=[
        "Member sizes match S401",
        "Bolts grade 8.8/S, M16/M20 per section depth (note S6)",
        "Welds 6 mm SP fillet UNO; AS 1554 procedures",
        "Hot-dip galvanised exterior; touched up with WATTYL Galvit",
        "Members in contact with concrete passivated",
        "ACRS certification on file (note S15)",
        "FC1 / CC2 per AS/NZS 5131",
    ],
)

add(
    type="steel-connection",
    title="Lower Roof — Connection Inspection",
    phase="Roof", level="roof",
    hotspots=["lower-roof"],
    drawings=[
        ("S401", "Lower Roof Framing Plan"),
        ("S440", "Typ Framing Details — Sheet 1"),
        ("S441", "Typ Framing Details — Sheet 2"),
        ("S442", "Typ Framing Details — Sheet 3"),
        ("S451", "Framing Details — Sheet 1"),
    ],
    rationale="Bolt-by-bolt and weld-by-weld inspection of lower roof joints.",
    checklist=[
        "All bolts tightened (full bearing) per AS 4100 8.8/S",
        "Weld lengths and sizes match drawings; visual quality (SP)",
        "Connection plates ≥ 10 mm thick UNO",
        "Galvanised cleats: damage touched up",
    ],
)

add(
    type="steel-frame-erection",
    title="Lower Roof — Purlin Install",
    phase="Roof", level="roof",
    hotspots=["lower-roof"],
    drawings=[("S402", "Lower Roof Purlin Plan"), ("S440", "Typ Framing Details")],
    rationale="Lower roof purlins per S402.",
    checklist=[
        "Purlin sections per S402",
        "Laps ≥ 15% span or 900 mm (PU2)",
        "Bolts per series (PU3)",
        "Bridging per PU6",
    ],
)

add(
    type="steel-frame-erection",
    title="Main Roof — Steel Frame Erection",
    phase="Roof", level="roof",
    hotspots=["main-roof"],
    drawings=[
        ("S405", "Main Roof Framing Plan"),
        ("S400", "Framing Notes"),
        ("S423", "Framing Elevations — Sheet 3"),
        ("S424", "Framing Elevations — Sheet 4"),
        ("S425", "Framing Elevations — Sheet 5"),
        ("S452", "Framing Details — Sheet 2"),
        ("S453", "Framing Details — Sheet 3"),
    ],
    rationale="Main roof — central long-span steel frame. Per AS 4100 / AS/NZS 5131 CC2.",
    checklist=[
        "Member sizes match S405 + framing elevations",
        "Bolts grade 8.8/S, M16/M20 per section depth",
        "Welds 6 mm SP fillet UNO",
        "HDG external; passivated where in contact with concrete",
        "Plate washers for oversize/slotted holes per AS 4100 Cl 14.3.5.2",
        "ACRS cert (S15); FC1/CC2 per AS/NZS 5131",
    ],
)

add(
    type="steel-connection",
    title="Main Roof — Connection Inspection",
    phase="Roof", level="roof",
    hotspots=["main-roof"],
    drawings=[("S405", "Main Roof Framing Plan"),
              ("S440", "Typ Framing — Sheet 1"), ("S441", "Sheet 2"), ("S442", "Sheet 3"),
              ("S452", "Framing Details — Sheet 2"), ("S453", "Sheet 3"), ("S454", "Sheet 4")],
    rationale="Connection-by-connection inspection of main roof joints.",
    checklist=[
        "All bolts tightened per AS 4100 8.8/S",
        "Welds match drawings; visual quality (SP)",
        "Site-weld locations only as specified (W7); independent NATA testing per W-table",
        "Connection plates ≥ 10 mm thick UNO",
    ],
)

add(
    type="steel-frame-erection",
    title="Main Roof — Purlin Install",
    phase="Roof", level="roof",
    hotspots=["main-roof"],
    drawings=[("S406", "Main Roof Purlin Plan")],
    rationale="Main roof purlins per S406.",
    checklist=[
        "Purlin sections per S406",
        "Laps ≥ 15% span or 900 mm (PU2)",
        "Bolts per series (PU3)",
        "Bridging per PU6",
    ],
)

add(
    type="steel-frame-erection",
    title="Upper Roof — Steel Frame Erection",
    phase="Roof", level="roof",
    hotspots=["upper-roof"],
    drawings=[
        ("S411", "Upper Roof Framing Plan"),
        ("S426", "Framing Elevations — Sheet 6"),
        ("S427", "Framing Elevations — Sheet 7"),
        ("S428", "Framing Elevations — Sheet 8"),
        ("S455", "Framing Details — Sheet 5"),
        ("S456", "Framing Details — Sheet 6"),
    ],
    rationale="Upper roof — highest tier of the multi-roof system. Steel-framed.",
    checklist=[
        "Member sizes match S411",
        "Bolts grade 8.8/S, M16/M20 per section depth",
        "Welds 6 mm SP fillet UNO",
        "HDG external",
    ],
)

add(
    type="steel-connection",
    title="Upper Roof — Connection Inspection",
    phase="Roof", level="roof",
    hotspots=["upper-roof"],
    drawings=[("S411", "Upper Roof Framing Plan"),
              ("S455", "Framing Details — Sheet 5"), ("S456", "Sheet 6"),
              ("S457", "Sheet 7"), ("S458", "Sheet 8"), ("S459", "Sheet 9")],
    rationale="Upper roof connection inspection.",
    checklist=[
        "All bolts tightened per AS 4100 8.8/S",
        "Welds match drawings",
        "Connection plates per typ details",
    ],
)

add(
    type="steel-frame-erection",
    title="Upper Roof — Purlin Install",
    phase="Roof", level="roof",
    hotspots=["upper-roof"],
    drawings=[("S412", "Upper Roof Purlin Plan")],
    rationale="Upper roof purlins per S412.",
    checklist=[
        "Purlin sections per S412",
        "Laps ≥ 15% span or 900 mm",
        "Bolts per series",
        "Bridging per PU6",
    ],
)

# ── Phase 5: Misc Steelwork ───────────────────────────────────────────────
add(
    type="steel-frame-erection",
    title="Miscellaneous Steelwork — Facade Brackets, Balustrades, Screens",
    phase="Misc", level="multi",
    hotspots=["misc-steelwork"],
    drawings=[
        ("S400", "Framing Notes, Legends, & Schedules"),
        ("S440", "Typ Framing Details — Sheet 1"),
    ],
    rationale=(
        "Misc steelwork — facade brackets, balcony balustrades, sun screens, parapets, "
        "atrium balustrades. Often installed late in build. Anchored to concrete via chemical/"
        "mechanical anchors."
    ),
    checklist=[
        "Member sizes per relevant detail",
        "Anchors per anchor notes (post-installed: 5% load-tested per note A8 if applicable)",
        "HDG to AS/NZS 4680",
        "Damage touched up with WATTYL Galvit",
    ],
)


if __name__ == '__main__':
    # ─────────────────────────────────────────────────────────────────────────────
    # RENDER HTML
    # ─────────────────────────────────────────────────────────────────────────────

    img_b64 = base64.b64encode(ISO_PATH.read_bytes()).decode("ascii")

    phases = []
    phase_order = ["Substructure", "Ground Floor", "Level 1", "Roof", "Misc"]
    for ph in phase_order:
        items = [i for i in PLAN if i["phase"] == ph]
        if items:
            phases.append((ph, items))

    total = len(PLAN)
    holds = sum(1 for i in PLAN if i.get("holdPoint"))
    drawings_referenced = len({d[0] for i in PLAN for d in i["drawings"]})

    hotspots_json = json.dumps({
        h: {"label": v["label"], "polygons": v["polygons"]}
        for h, v in HOTSPOTS.items()
    })
    inspection_hotspots_json = json.dumps({
        i["sequence"]: i.get("hotspots", [])
        for i in PLAN
    })


    def render_card(ins):
        seq = ins["sequence"]
        chips = [f'<span class="chip">{html_lib.escape(ins["type"])}</span>',
                 f'<span class="pill">{html_lib.escape(ins["level"])}</span>']
        if ins.get("holdPoint"):
            chips.append('<span class="badge hold">Hold</span>')
        chips_html = "".join(chips)
        dwg_html = "".join(
            f'<span class="dwg"><code>{html_lib.escape(sn)}</code> {html_lib.escape(desc)}</span>'
            for sn, desc in ins["drawings"]
        )
        check_html = "".join(f"<li>{html_lib.escape(c)}</li>" for c in ins["checklist"])
        scope_html = (f'<div class="scope">{html_lib.escape(ins["scopeNote"])}</div>'
                      if ins.get("scopeNote") else "")
        hold_class = " hold" if ins.get("holdPoint") else ""
        return (
            f'<div class="card{hold_class}" data-seq="{seq}" data-level="{html_lib.escape(ins["level"])}">'
            f'<div class="card-row"><div class="seq">{seq:02d}</div><div class="body">'
            f'<div class="card-head">{chips_html}</div>'
            f'<h3>{html_lib.escape(ins["title"])}</h3>'
            f'<div class="detail">'
            f'<div class="row"><div class="k">Why</div><div class="v">{html_lib.escape(ins["rationale"])}</div></div>'
            f'<div class="row drawings"><div class="k">Drawings</div><div class="v">{dwg_html}</div></div>'
            f'<div class="row checklist"><div class="k">Checklist</div><div class="v"><ul>{check_html}</ul></div></div>'
            f'{scope_html}'
            f'</div></div></div></div>'
        )


    phase_blocks_html = ""
    for ph_name, items in phases:
        cards_html = "".join(render_card(i) for i in items)
        phase_blocks_html += (
            f'<div class="phase">'
            f'<div class="phase-head">'
            f'<span class="phase-num">Phase {phase_order.index(ph_name)+1:02d}</span>'
            f'<h2>{ph_name}</h2>'
            f'<span class="count">{len(items)} items</span>'
            f'</div>{cards_html}</div>'
        )

    gen_date = datetime.date.today().isoformat()

    TEMPLATE = r"""<!doctype html>
    <html lang="en">
    <head>
    <meta charset="utf-8">
    <title>MBC Creativity & Arts Centre — Inspection Program</title>
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&family=Lora:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
    <style>
      :root {
        --dark: #141413; --light: #faf9f5; --mid-gray: #b0aea5; --light-gray: #e8e6dc;
        --orange: #d97757; --blue: #6a9bcc; --green: #788c5d;
        --rule: rgba(20,20,19,0.08); --soft: rgba(20,20,19,0.04);
      }
      * { box-sizing: border-box; }
      html, body { margin: 0; padding: 0; height: 100%; }
      body {
        background: var(--light); color: var(--dark);
        font-family: 'Lora', Georgia, serif; font-size: 15px; line-height: 1.55;
        -webkit-font-smoothing: antialiased;
        overflow: hidden;
      }
      h1, h2, h3, h4, .num, .chip, .badge, .lbl, .pill, .meta, .phase-num, .seq, .k, .filter {
        font-family: 'Poppins', Arial, sans-serif; letter-spacing: -0.005em;
      }
      code, .mono { font-family: 'JetBrains Mono', ui-monospace, monospace; }
      header.top {
        padding: 18px 28px; border-bottom: 1px solid var(--rule);
        display: flex; align-items: baseline; gap: 24px; flex-wrap: wrap;
      }
      header.top h1 { font-size: 21px; font-weight: 600; margin: 0; letter-spacing: -0.01em; }
      header.top .sub {
        color: var(--mid-gray); font-family: 'Poppins'; font-size: 11.5px;
        letter-spacing: 0.06em; text-transform: uppercase;
      }
      header.top .badge-tender {
        background: var(--blue); color: var(--light);
        font-size: 10px; padding: 3px 8px; border-radius: 3px;
        letter-spacing: 0.06em; text-transform: uppercase; font-weight: 600;
      }
      main {
        display: grid;
        grid-template-columns: minmax(0, 1fr) minmax(440px, 560px);
        height: calc(100vh - 105px);
      }
      aside.iso {
        padding: 20px 24px;
        border-right: 1px solid var(--rule);
        display: flex; flex-direction: column;
        overflow: hidden;
      }
      .iso-stats {
        display: grid; grid-template-columns: repeat(4, 1fr);
        gap: 1px; background: var(--rule);
        border: 1px solid var(--rule); border-radius: 4px;
        margin-bottom: 12px; overflow: hidden;
      }
      .iso-stats .stat { background: var(--light); padding: 9px 11px; }
      .iso-stats .stat .num {
        font-size: 20px; font-weight: 600; line-height: 1;
        display: block; margin-bottom: 2px; font-family: 'Poppins';
      }
      .iso-stats .stat .num.accent { color: var(--orange); }
      .iso-stats .stat .lbl {
        font-size: 9.5px; letter-spacing: 0.06em; text-transform: uppercase;
        color: var(--mid-gray); font-weight: 500;
      }
      .filter-row { display: flex; gap: 5px; flex-wrap: wrap; margin-bottom: 12px; }
      .filter {
        font-size: 11px; letter-spacing: 0.04em;
        padding: 4px 10px; border: 1px solid var(--rule); border-radius: 999px;
        background: var(--light); color: var(--dark); cursor: pointer;
        transition: all 0.12s ease;
      }
      .filter:hover { border-color: var(--mid-gray); }
      .filter.active { background: var(--dark); color: var(--light); border-color: var(--dark); }
      .iso-wrap {
        position: relative; flex: 1;
        display: flex; align-items: center; justify-content: center;
        background: #fbfaf6;
        border: 1px solid var(--rule); border-radius: 4px;
        padding: 10px; overflow: hidden; min-height: 0;
      }
      .iso-img-container {
        position: relative; max-width: 100%; max-height: 100%; line-height: 0;
      }
      .iso-img-container img {
        max-width: 100%; max-height: 100%; width: auto; height: auto; display: block;
      }
      svg.hotspot-overlay {
        position: absolute; top: 0; left: 0;
        width: 100%; height: 100%; pointer-events: none;
      }
      svg.hotspot-overlay polygon {
        fill: transparent; stroke: transparent; stroke-width: 0.4;
        transition: all 0.18s ease; pointer-events: auto; cursor: pointer;
      }
      svg.hotspot-overlay polygon.active {
        fill: rgba(217, 119, 87, 0.36);
        stroke: #b14a26; stroke-width: 0.5;
        filter: drop-shadow(0 0 4px rgba(217, 119, 87, 0.7));
      }
      svg.hotspot-overlay polygon.hover {
        fill: rgba(217, 119, 87, 0.18); stroke: var(--orange);
      }
      .iso-legend {
        margin-top: 10px;
        font-size: 11px; color: var(--mid-gray);
        font-family: 'Poppins'; letter-spacing: 0.02em;
        display: flex; align-items: center; gap: 8px;
      }
      .iso-legend .swatch {
        width: 11px; height: 11px; border: 1px solid #b14a26;
        background: rgba(217, 119, 87, 0.36); border-radius: 2px;
      }
      section.program {
        overflow-y: auto; padding: 18px 24px; background: var(--light);
      }
      .phase { margin-bottom: 22px; }
      .phase-head {
        display: flex; align-items: baseline; gap: 12px;
        padding-bottom: 6px; margin-bottom: 10px;
        border-bottom: 1px solid var(--rule);
      }
      .phase-num {
        font-size: 10.5px; letter-spacing: 0.1em; text-transform: uppercase;
        color: var(--mid-gray); font-weight: 500;
      }
      .phase-head h2 {
        font-size: 16px; font-weight: 600; margin: 0; letter-spacing: -0.01em;
      }
      .phase-head .count {
        margin-left: auto; font-size: 10.5px; color: var(--mid-gray);
        font-family: 'Poppins'; letter-spacing: 0.04em;
      }
      .card {
        border: 1px solid var(--rule); border-left-width: 3px;
        border-left-color: var(--mid-gray); border-radius: 4px;
        background: var(--light); padding: 10px 12px;
        margin-bottom: 6px; cursor: pointer;
        transition: border-color 0.12s ease, background 0.12s ease;
        user-select: none;
      }
      .card:hover { border-color: rgba(20,20,19,0.18); border-left-color: var(--dark); }
      .card.active {
        border-color: var(--orange); border-left-color: var(--orange);
        background: rgba(217, 119, 87, 0.05);
      }
      .card.hold { border-left-color: var(--orange); }
      .card.dim { opacity: 0.32; }
      .card-row { display: flex; align-items: flex-start; gap: 10px; }
      .card .seq {
        font-size: 17px; font-weight: 600; color: var(--dark);
        line-height: 1; min-width: 26px; padding-top: 2px;
        font-variant-numeric: tabular-nums;
      }
      .card .body { flex: 1; min-width: 0; }
      .card-head { display: flex; flex-wrap: wrap; align-items: center; gap: 5px; margin-bottom: 3px; }
      .chip {
        font-size: 10px; font-family: 'JetBrains Mono', monospace;
        padding: 2px 6px; background: var(--light-gray); color: var(--dark);
        border-radius: 2px;
      }
      .pill {
        font-size: 10px; padding: 2px 7px; border-radius: 999px;
        background: var(--soft); color: var(--dark);
        letter-spacing: 0.04em; text-transform: uppercase;
      }
      .badge {
        font-size: 9.5px; letter-spacing: 0.06em; text-transform: uppercase;
        padding: 2px 6px; border-radius: 2px; font-weight: 600;
      }
      .badge.hold { background: var(--orange); color: var(--light); }
      .card h3 {
        font-size: 13px; font-weight: 600; margin: 1px 0;
        line-height: 1.35; letter-spacing: -0.005em; color: var(--dark);
      }
      .detail {
        display: none; padding-top: 8px; margin-top: 6px;
        border-top: 1px dashed var(--rule);
      }
      .card.active .detail { display: block; }
      .detail .row {
        display: grid; grid-template-columns: 80px 1fr;
        gap: 10px; font-size: 12px; margin-bottom: 5px;
      }
      .detail .row .k {
        font-size: 9.5px; letter-spacing: 0.06em; text-transform: uppercase;
        color: var(--mid-gray); padding-top: 2px; font-weight: 500;
      }
      .detail .row .v { color: rgba(20,20,19,0.85); line-height: 1.5; }
      .detail .row.checklist .v ul { margin: 0; padding: 0; list-style: none; }
      .detail .row.checklist .v li {
        padding: 1px 0 1px 14px; position: relative; font-size: 11.5px;
      }
      .detail .row.checklist .v li::before {
        content: ""; position: absolute; left: 0; top: 7px;
        width: 4px; height: 4px;
        border: 1px solid var(--mid-gray); border-radius: 1px;
      }
      .detail .row.drawings .v { font-size: 11px; line-height: 1.5; }
      .detail .row.drawings .v .dwg { display: inline-block; margin: 1px 5px 1px 0; }
      .detail .row.drawings .v .dwg code {
        background: var(--light-gray); padding: 1px 5px; border-radius: 2px;
        font-size: 10px;
      }
      .detail .scope {
        margin-top: 6px; font-size: 11px; font-style: italic;
        color: var(--mid-gray); padding-left: 10px;
        border-left: 2px solid var(--rule);
      }
      footer.bottom {
        border-top: 1px solid var(--rule); padding: 12px 28px;
        display: flex; justify-content: space-between; align-items: center;
        font-size: 11px; color: var(--mid-gray);
        font-family: 'Poppins'; letter-spacing: 0.04em; background: var(--light);
      }
      footer.bottom .actions { display: flex; gap: 8px; }
      footer.bottom a.action {
        font-size: 11px; padding: 5px 12px; border: 1px solid var(--rule);
        border-radius: 3px; color: var(--dark); text-decoration: none;
        background: var(--light); transition: all 0.12s ease;
      }
      footer.bottom a.action:hover { border-color: var(--dark); }
      footer.bottom a.action.primary {
        background: var(--dark); color: var(--light); border-color: var(--dark);
      }
      footer.bottom a.action.primary:hover { background: var(--orange); border-color: var(--orange); }
      section.program::-webkit-scrollbar { width: 8px; }
      section.program::-webkit-scrollbar-track { background: var(--soft); }
      section.program::-webkit-scrollbar-thumb { background: var(--mid-gray); border-radius: 4px; }
    </style>
    </head>
    <body>

    <header class="top">
      <h1>MBC Creativity & Arts Centre</h1>
      <span class="sub">Job 2024.0230 · Rev T2 · Moreton Bay College, Manly West</span>
      <span class="badge-tender">TENDER · Not For Construction</span>
    </header>

    <main>
      <aside class="iso">
        <div class="iso-stats">
          <div class="stat"><span class="num">__TOTAL__</span><span class="lbl">Inspections</span></div>
          <div class="stat"><span class="num accent">__HOLDS__</span><span class="lbl">Hold points</span></div>
          <div class="stat"><span class="num">__DRAWINGS__</span><span class="lbl">Drawings ref'd</span></div>
          <div class="stat"><span class="num">62</span><span class="lbl">Sheets total</span></div>
        </div>

        <div class="filter-row">
          <button class="filter active" data-level="all">All</button>
          <button class="filter" data-level="ground">Substructure + Ground</button>
          <button class="filter" data-level="l1">L1</button>
          <button class="filter" data-level="roof">Roof (3 tiers)</button>
          <button class="filter" data-level="multi">Misc</button>
        </div>

        <div class="iso-wrap">
          <div class="iso-img-container">
            <img id="iso-img" src="data:image/png;base64,__IMG_B64__" alt="MBC Creativity & Arts Centre structural isometric">
            <svg class="hotspot-overlay" id="hotspots" preserveAspectRatio="none" viewBox="0 0 100 100"></svg>
          </div>
        </div>

        <div class="iso-legend">
          <span class="swatch"></span>
          <span>Click an inspection card → element(s) light up here. TENDER ISSUE — provisional pending Construction Issue.</span>
        </div>
      </aside>

      <section class="program" id="program">
    __PHASE_BLOCKS__
      </section>
    </main>

    <footer class="bottom">
      <span>Bligh Tanner · Cowork-generated · __DATE__ · TENDER ISSUE — verify against Construction Issue when received</span>
      <div class="actions">
        <a class="action" href="mailto:?subject=MBC inspection program — request changes&body=Card numbers needing changes:%0D%0A%0D%0A">Request changes</a>
        <a class="action primary" href="mailto:?subject=MBC inspection program — APPROVED (provisional)&body=Approved as drafted (provisional, pending Construction Issue). Proceed to write project-map.json + project.btproject.">Approve program</a>
      </div>
    </footer>

    <script>
    const HOTSPOTS = __HOTSPOTS_JSON__;
    const INSPECTION_HOTSPOTS = __INSPECTION_HOTSPOTS_JSON__;

    const svg = document.getElementById('hotspots');
    const polyElems = {};
    for (const [hid, h] of Object.entries(HOTSPOTS)) {
      polyElems[hid] = [];
      for (const points of h.polygons) {
        const poly = document.createElementNS('http://www.w3.org/2000/svg', 'polygon');
        poly.setAttribute('points', points);
        poly.setAttribute('data-hotspot', hid);
        poly.setAttribute('vector-effect', 'non-scaling-stroke');
        svg.appendChild(poly);
        polyElems[hid].push(poly);
      }
    }

    const cards = document.querySelectorAll('.card');
    cards.forEach(card => {
      card.addEventListener('click', () => {
        const seq = parseInt(card.dataset.seq);
        activateInspection(seq, card);
      });
    });

    function clearActive() {
      cards.forEach(c => c.classList.remove('active'));
      Object.values(polyElems).flat().forEach(p => p.classList.remove('active'));
    }

    function activateInspection(seq, card) {
      const wasActive = card.classList.contains('active');
      clearActive();
      if (wasActive) return;
      card.classList.add('active');
      const hids = INSPECTION_HOTSPOTS[seq] || [];
      hids.forEach(hid => {
        (polyElems[hid] || []).forEach(p => p.classList.add('active'));
      });
    }

    Object.entries(polyElems).forEach(([hid, polys]) => {
      polys.forEach(p => {
        p.addEventListener('mouseenter', () => {
          polys.forEach(pp => pp.classList.add('hover'));
        });
        p.addEventListener('mouseleave', () => {
          polys.forEach(pp => pp.classList.remove('hover'));
        });
        p.addEventListener('click', () => {
          for (const card of cards) {
            const seq = parseInt(card.dataset.seq);
            const hids = INSPECTION_HOTSPOTS[seq] || [];
            if (hids.includes(hid)) {
              activateInspection(seq, card);
              card.scrollIntoView({ behavior: 'smooth', block: 'center' });
              break;
            }
          }
        });
      });
    });

    const filterButtons = document.querySelectorAll('.filter');
    filterButtons.forEach(btn => {
      btn.addEventListener('click', () => {
        filterButtons.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const lvl = btn.dataset.level;
        cards.forEach(c => {
          if (lvl === 'all') {
            c.classList.remove('dim');
          } else if (lvl === 'ground') {
            c.classList.toggle('dim', c.dataset.level !== 'ground');
          } else {
            c.classList.toggle('dim', c.dataset.level !== lvl);
          }
        });
      });
    });
    </script>
    </body>
    </html>
    """

    html_out = (
        TEMPLATE
        .replace("__TOTAL__", str(total))
        .replace("__HOLDS__", str(holds))
        .replace("__DRAWINGS__", str(drawings_referenced))
        .replace("__IMG_B64__", img_b64)
        .replace("__PHASE_BLOCKS__", phase_blocks_html)
        .replace("__DATE__", gen_date)
        .replace("__HOTSPOTS_JSON__", hotspots_json)
        .replace("__INSPECTION_HOTSPOTS_JSON__", inspection_hotspots_json)
    )

    OUT_HTML.write_text(html_out)
    print(f"Wrote {OUT_HTML}")
    print(f"  size: {len(html_out):,} chars ({len(html_out)/1024:.0f} KB)")
    print(f"  inspections: {total}")
    print(f"  hold points: {holds}")
    print(f"  hotspots: {len(HOTSPOTS)} ({sum(len(h['polygons']) for h in HOTSPOTS.values())} polygons)")
    print(f"  drawings referenced: {drawings_referenced}")
