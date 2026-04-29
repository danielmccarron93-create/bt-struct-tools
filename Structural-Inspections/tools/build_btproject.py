"""
build_btproject.py — generate project-map.json + project.btproject from a Cowork-extracted project.

Usage:
    python3 tools/build_btproject.py "Structural Drawings/MBC"

Reads the per-project build script's data (PLAN, GENERAL_NOTES, PROJECT, EXCLUDED_FROM_BT, etc.)
and produces:
    <project>/project-map.json    — schema-validated machine-readable program
    <project>/project.btproject   — zip bundle ready for PWA import

The bundle layout matches what js/lib/btproject.js expects:
    project.btproject (zip)
    ├── project-map.json
    ├── manifest.txt
    └── Drawings/
        └── <pdf-filenames>.pdf

Per-project data must be exposed in `<project>/.btinspect_scratch/build_html.py` as module-level
constants:
    PROJECT_META  = {jobNumber, name, client, siteAddress, discipline, issueStatus, ...}
    GENERAL_NOTES = {designCodes, concreteCover, concreteStrengths, bearingCapacity, ...}
    EXCLUDED_FROM_BT = [{element, responsibility, noteRef, scope}, ...]   # certified by others
    PLAN          = [{sequence, type, title, phase, level, hotspots, drawings:[(sn,reason)],
                       rationale, checklist, holdPoint, severity, scopeNote, siteReadinessCheck,
                       postPourRecord}, ...]
    HOTSPOTS      = {hotspot_id: {label, polygons:[...]}}   # for the HTML, not the JSON

If the build script doesn't expose those (older projects), this tool will reverse-engineer
them by importing build_html.py and reading whatever it can.
"""
import argparse
import importlib.util
import json
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SCHEMA_PATH = REPO_ROOT / "reference" / "project-map.schema.json"
INSPECTION_TYPES_PATH = REPO_ROOT / "reference" / "inspection-types.json"

SCHEMA_VERSION = 1
GENERATOR_ID = "cowork-btinspect/2.2"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def load_module(path: Path):
    """Dynamically import a Python file at `path` as a module."""
    spec = importlib.util.spec_from_file_location(f"_built_{path.stem}", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_inspection_types():
    """Load the canonical catalogue → set of valid keys."""
    catalogue = json.loads(INSPECTION_TYPES_PATH.read_text())
    return {t["key"] for t in catalogue["types"]}


def normalise_drawings_index(plan_entries):
    """For each plan entry, normalise drawings to [{sheetNumber, reason}, ...]."""
    out = []
    for e in plan_entries:
        # Accept either ('S010', 'reason') tuples OR {'sheetNumber','reason'} dicts
        normalised_drawings = []
        for d in e.get("drawings", []):
            if isinstance(d, (list, tuple)) and len(d) >= 2:
                normalised_drawings.append({"sheetNumber": d[0], "reason": d[1]})
            elif isinstance(d, dict):
                normalised_drawings.append({
                    "sheetNumber": d.get("sheetNumber") or d.get("sheet") or "",
                    "reason":      d.get("reason") or d.get("desc") or ""
                })
        normalised = dict(e)
        normalised["drawings"] = normalised_drawings
        out.append(normalised)
    return out


def normalise_checklist(items):
    """Allow both plain strings and {text, asClauseRef, noteRef} dicts. Returns list."""
    out = []
    for c in items:
        if isinstance(c, str):
            out.append(c)
        elif isinstance(c, dict):
            if "text" in c:
                out.append({k: v for k, v in c.items() if k in ("text", "asClauseRef", "noteRef") and v is not None})
            else:
                # Skip malformed entries
                continue
    return out


# ---------------------------------------------------------------------------
# Drawing classification — re-derive per-drawing tags + applicableInspectionTypes
# ---------------------------------------------------------------------------

def classify_drawing(sheet_number: str, description: str, plan_entries):
    """Return (tags[], applicableInspectionTypes[]) for one drawing.

    Cross-references the plan entries: any plan entry whose drawingRefs include
    this sheet adds its `type` to applicableInspectionTypes.
    """
    sn_upper = (sheet_number or "").upper()
    desc_upper = (description or "").upper()

    # Tags: kind + level + element heuristics
    tags = set()

    # ---- KIND ----
    if "COVER" in desc_upper:
        tags.add("kind:cover")
    elif "NOTES" in desc_upper or "SAFETY IN DESIGN" in desc_upper:
        tags.add("kind:notes")
    elif "LOADING PLAN" in desc_upper:
        tags.add("kind:loading-plan")
    elif "TYPICAL" in desc_upper and ("DETAIL" in desc_upper or "DETAILS" in desc_upper):
        tags.add("kind:typical-detail")
    elif "SECTIONS" in desc_upper and "DETAILS" not in desc_upper:
        tags.add("kind:section")
    elif "BUILDING SECTIONS" in desc_upper:
        tags.add("kind:section")
    elif "DETAILS" in desc_upper and "PLAN" not in desc_upper:
        tags.add("kind:detail")
    elif "ELEVATIONS" in desc_upper:
        tags.add("kind:elevation")
    elif "PERSPECTIVE" in desc_upper:
        tags.add("kind:perspective")
    elif "PLAN" in desc_upper or "LAYOUT" in desc_upper or "ARRANGEMENT" in desc_upper:
        tags.add("kind:plan")
    else:
        tags.add("kind:plan")

    # ---- LEVEL ----
    if any(k in desc_upper for k in ("GROUND FLOOR", "FOOTING", "PILING", "FOUNDATION", "PILE CAP")):
        tags.add("level:ground")
    elif "LEVEL 1 " in desc_upper or "FIRST FLOOR" in desc_upper:
        tags.add("level:l1")
    elif "LEVEL 2 " in desc_upper:
        tags.add("level:l2")
    elif "LEVEL 3 " in desc_upper:
        tags.add("level:l3")
    elif "LEVEL 4-6" in desc_upper or "LEVEL 4–6" in desc_upper:
        tags.add("level:multi")
    elif "LEVEL 7" in desc_upper:
        tags.add("level:multi")
    elif "ROOF" in desc_upper:
        tags.add("level:roof")
    elif "kind:cover" in tags or "kind:notes" in tags or "kind:typical-detail" in tags:
        tags.add("level:na")
    else:
        tags.add("level:multi")

    # ---- ELEMENT ----
    keywords = [
        ("PILING", "pile"), ("PILE", "pile"),
        ("FOUNDATION", "pad-footing"), ("PAD FOOTING", "pad-footing"),
        ("FOOTING", "pad-footing"), ("STRIP FOOTING", "strip-footing"),
        ("RAFT", "raft-slab"), ("WAFFLE", "raft-slab"),
        ("SLAB ON GROUND", "slab-on-ground"),
        ("CONCRETE WALL", "concrete-wall"),
        ("BLOCK WALL", "block-wall"), ("MASONRY", "block-wall"),
        ("CONCRETE COLUMN", "concrete-column"),
        ("STEEL DETAIL", "steel-beam"),
        ("ROOF FRAMING", "roof-framing"),
        ("STAIR", "stair"),
        ("CLT", "mass-timber"), ("TIMBER CONNECTION", "mass-timber"),
        ("TIMBER", "timber-framing"),
        ("BALCONY", "slab-suspended"),
        ("PLANT", "slab-on-ground"),
        ("POST TENSIONING", "post-tension"), ("POST-TENSIONING", "post-tension"),
        ("LECTURE THEATRE", "steel-frame"),
        ("AMPHITHEATRE", "steel-frame"),
        ("FIN FRAME", "fin-frame"),
        ("ELEVATED LINK", "steel-frame"),
        ("ATRIUM STAIR", "stair"),
    ]
    for kw, tag in keywords:
        if kw in desc_upper:
            tags.add(f"element:{tag}")

    # GA / reinforcement plans = slab plans for known floors
    is_slab_plan = any(k in desc_upper for k in (
        "GENERAL ARRANGEMENT", "BOTTOM REINFORCEMENT PLAN",
        "TOP REINFORCEMENT PLAN", "REINFORCEMENT LAYOUTS"
    ))
    if is_slab_plan:
        if "level:ground" in tags:
            tags.add("element:slab-on-ground")
        elif any(t in tags for t in ("level:l1", "level:l2", "level:l3", "level:multi")):
            tags.add("element:slab-suspended")

    # Steel framing/details adds beam + column + connection
    if any(k in desc_upper for k in ("STEEL", "FRAMING")):
        tags.add("element:steel-beam")
        if "DETAIL" in desc_upper:
            tags.add("element:steel-column")
            tags.add("element:connection")

    # Roof framing implies steel beams + connections
    if "ROOF FRAMING" in desc_upper or "ROOF PURLIN" in desc_upper:
        tags.add("element:steel-beam")
        if "ROOF FRAMING" in desc_upper:
            tags.add("element:connection")

    # ---- applicableInspectionTypes ----
    # Find plan entries that reference this sheet number
    applicable = set()
    for entry in plan_entries:
        for ref in entry.get("drawings", []):
            if isinstance(ref, dict):
                ref_sn = ref.get("sheetNumber", "")
            else:
                ref_sn = ref[0] if isinstance(ref, (list, tuple)) and len(ref) > 0 else ""
            if (ref_sn or "").upper() == sn_upper:
                applicable.add(entry["type"])

    return sorted(tags), sorted(applicable)


# ---------------------------------------------------------------------------
# Schema validation
# ---------------------------------------------------------------------------

def validate_against_schema(project_map: dict) -> list:
    """Best-effort validation. Returns a list of warnings; raises on hard fails."""
    soft = []
    if project_map.get("schemaVersion") != SCHEMA_VERSION:
        raise RuntimeError(f"schemaVersion must be {SCHEMA_VERSION}")
    if not project_map.get("project", {}).get("name"):
        raise RuntimeError("project.name is required")
    if not project_map.get("sourceFiles"):
        raise RuntimeError("sourceFiles[] is required and non-empty")
    if not isinstance(project_map.get("drawings", []), list):
        raise RuntimeError("drawings[] must be an array")
    if not isinstance(project_map.get("inspectionPlan", []), list):
        raise RuntimeError("inspectionPlan[] must be an array")

    valid_types = load_inspection_types()
    bad_types = [
        e["type"] for e in project_map["inspectionPlan"]
        if e.get("type") not in valid_types
    ]
    if bad_types:
        raise RuntimeError(
            f"inspectionPlan contains type keys not in inspection-types.json: {sorted(set(bad_types))}"
        )

    # Cross-check drawingRefs resolve
    sheet_set = {(d.get("sheetNumber") or "").upper() for d in project_map["drawings"]}
    sheet_set.discard("")
    bad_refs = 0
    for entry in project_map["inspectionPlan"]:
        for ref in entry.get("drawingRefs", []):
            sn = (ref.get("sheetNumber") or "").upper()
            if sn and sn not in sheet_set:
                bad_refs += 1
    if bad_refs:
        soft.append(f"{bad_refs} inspection-plan drawingRefs point at sheets not in drawings[]")

    return soft


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def build(project_dir: Path, dry_run: bool = False):
    project_dir = project_dir.resolve()
    if not project_dir.is_dir():
        raise SystemExit(f"Not a directory: {project_dir}")

    # 1) Load the per-project build script (same module that built the HTML)
    build_script = project_dir / ".btinspect_scratch" / "build_html.py"
    if not build_script.exists():
        raise SystemExit(
            f"No build_html.py found at {build_script}. "
            "Run the per-project Cowork procedure first to extract data."
        )

    print(f"Loading data from {build_script}")
    mod = load_module(build_script)

    # 2) Locate PDFs
    pdfs = sorted(project_dir.glob("*.pdf"))
    if not pdfs:
        raise SystemExit(f"No PDFs found in {project_dir}")

    # 3) Load drawing list (cover-sheet authoritative source if available)
    drawing_list = {}
    drawing_list_path = Path("/tmp") / f"{project_dir.name.lower().replace(' ', '_')}_drawing_list.json"
    if drawing_list_path.exists():
        try:
            drawing_list = json.loads(drawing_list_path.read_text())
            print(f"Using drawing list from {drawing_list_path} ({len(drawing_list)} sheets)")
        except Exception as e:
            print(f"Could not parse {drawing_list_path}: {e}")
    # Try MBC-specific path too as a fallback
    for cand in [Path("/tmp/mbc_drawing_list.json"), Path("/tmp/em_drawing_list.json")]:
        if not drawing_list and cand.exists():
            drawing_list = json.loads(cand.read_text())
            print(f"Using drawing list from {cand} ({len(drawing_list)} sheets)")
            break

    # 4) Load per-page title-block extraction (for sheet inventory)
    titleblock_path = None
    for cand in [
        Path(f"/tmp/{project_dir.name.lower().replace(' ', '_')}_tb.json"),
        Path("/tmp/mbc_tb.json"),
        Path("/tmp/em_tb.json"),
        Path("/tmp/tb.json"),
    ]:
        if cand.exists():
            titleblock_path = cand
            break
    if titleblock_path is None:
        raise SystemExit(
            f"Could not find a per-page title-block JSON in /tmp. "
            "Run the title-block extractor first."
        )
    titleblock_pages = json.loads(titleblock_path.read_text())
    print(f"Loaded {len(titleblock_pages)} pages from {titleblock_path}")

    # 5) Build sourceFiles[]
    import fitz
    source_files = []
    for pdf in pdfs:
        doc = fitz.open(str(pdf))
        source_files.append({
            "filename":  pdf.name,
            "pageCount": doc.page_count,
            "sizeBytes": pdf.stat().st_size,
            "sha256":    None,
        })
        doc.close()

    # 6) Build drawings[] using titleblock data + drawing-list fallback for descriptions
    plan_entries = normalise_drawings_index(getattr(mod, "PLAN", []))
    drawings = []
    for page in titleblock_pages:
        sn = page.get("sheetNumber")
        rev = page.get("revision")
        desc = page.get("description")
        # Prefer the cover-sheet drawing list for description if per-page extraction is empty
        if (not desc) and sn and sn in drawing_list:
            desc = drawing_list[sn]
        # The page might reference a different PDF — for now assume the first PDF unless multi-PDF
        source_filename = pdfs[0].name
        tags, applicable = classify_drawing(sn, desc, plan_entries)
        drawings.append({
            "sourceFile":  source_filename,
            "pageNumber":  page["pageNumber"],
            "sheetNumber": sn,
            "revision":    rev,
            "revisionDate": page.get("revisionDate"),
            "description": desc,
            "scale":       page.get("scale"),
            "drawnBy":     page.get("drawnBy"),
            "checkedBy":   page.get("checkedBy"),
            "approvedBy":  page.get("approvedBy"),
            "tags":        tags,
            "applicableInspectionTypes": applicable,
            "elements":    [],
            "confidence":  page.get("confidence", 1.0),
        })

    # 7) Build inspectionPlan[] from PLAN with normalisations
    inspection_plan = []
    for entry in plan_entries:
        normalised = {
            "sequence": entry["sequence"],
            "type":     entry["type"],
            "title":    entry["title"],
            "level":    entry.get("level", ""),
            "drawingRefs": entry.get("drawings", []),
            "rationale":   entry.get("rationale", ""),
            "stage":       entry.get("stage", "pre-pour"),
            "holdPoint":   bool(entry.get("holdPoint", False)),
            "defaultSeverity": entry.get("severity", "observation") if not entry.get("holdPoint") else "hold-point",
            "expectedChecklist": normalise_checklist(entry.get("checklist", [])),
        }
        # Optional fields — only emit when present
        for opt_key in ("building", "zone", "scopeNote", "phase", "siteReadinessCheck", "postPourRecord"):
            if opt_key in entry and entry[opt_key] is not None:
                normalised[opt_key] = entry[opt_key]
        inspection_plan.append(normalised)

    # 8) Project metadata
    project_meta = getattr(mod, "PROJECT_META", None)
    if project_meta is None:
        # Fall back to building from page-1 of the title-block extraction
        p0 = titleblock_pages[0] if titleblock_pages else {}
        project_meta = {
            "name":        p0.get("_projectName") or project_dir.name,
            "siteAddress": p0.get("_siteAddress"),
            "jobNumber":   p0.get("_jobNumber"),
            "client":      p0.get("_client"),
            "issueStatus": p0.get("_issueStatus"),
            "engineerOfRecord": "Bligh Tanner",
            "discipline":  "structural",
        }

    # 9) General Notes (optional)
    general_notes = getattr(mod, "GENERAL_NOTES", {}) or {}
    excluded_from_bt = getattr(mod, "EXCLUDED_FROM_BT", []) or []
    if excluded_from_bt:
        project_meta = {**project_meta, "excludedFromBT": excluded_from_bt}

    # 10) Warnings (optional)
    warnings = list(getattr(mod, "WARNINGS", []) or [])

    # 11) Assemble project-map.json
    now_iso = datetime.now(timezone.utc).isoformat()
    project_map = {
        "schemaVersion": SCHEMA_VERSION,
        "generatedAt":   now_iso,
        "generator":     GENERATOR_ID,
        "sourceFolder":  project_dir.name,
        "project":       project_meta,
        "sourceFiles":   source_files,
        "drawings":      drawings,
        "generalNotes":  general_notes,
        "inspectionPlan": inspection_plan,
        "warnings":      warnings,
        "metadata": {
            "drawingsAnalysed": len(drawings),
            "templatesDetected": list(getattr(mod, "TEMPLATES_DETECTED", []) or []),
        }
    }

    # 12) Validate
    soft_warnings = validate_against_schema(project_map)
    for w in soft_warnings:
        print(f"  ⚠  {w}")

    # 13) Write project-map.json
    map_path = project_dir / "project-map.json"
    if dry_run:
        print(f"DRY RUN: would write {map_path}")
    else:
        map_path.write_text(json.dumps(project_map, indent=2, ensure_ascii=False))
        print(f"✓ Wrote {map_path} ({len(json.dumps(project_map))/1024:.1f} KB)")

    # 14) Build the .btproject zip
    btproject_path = project_dir / "project.btproject"
    if dry_run:
        print(f"DRY RUN: would write {btproject_path}")
    else:
        manifest_text = (
            f"BT Inspect — Project Bundle\n"
            f"Project:        {project_meta.get('name','')}\n"
            f"Job Number:     {project_meta.get('jobNumber','') or '—'}\n"
            f"Client:         {project_meta.get('client','') or '—'}\n"
            f"Site:           {project_meta.get('siteAddress','') or '—'}\n"
            f"Issue:          {project_meta.get('issueStatus','') or '—'}\n"
            f"Drawings:       {len(drawings)} sheet(s) across {len(source_files)} PDF(s)\n"
            f"Inspections:    {len(inspection_plan)} planned, "
            f"{sum(1 for e in inspection_plan if e['holdPoint'])} hold point(s)\n"
            f"Generated:      {now_iso}\n"
            f"Generator:      {GENERATOR_ID}\n"
        )
        with zipfile.ZipFile(btproject_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("project-map.json", json.dumps(project_map, indent=2, ensure_ascii=False))
            zf.writestr("manifest.txt", manifest_text)
            for pdf in pdfs:
                # IMPORTANT: bundle layout MUST be 'Drawings/<filename>.pdf' — the PWA reads from
                # exactly this folder name (see js/lib/btproject.js line 86).
                zf.write(pdf, f"Drawings/{pdf.name}")
        size_mb = btproject_path.stat().st_size / 1024 / 1024
        print(f"✓ Wrote {btproject_path} ({size_mb:.1f} MB)")

    # 15) Summary
    print()
    print(f"=== Summary ===")
    print(f"  Project:      {project_meta.get('name','')}")
    print(f"  Job#:         {project_meta.get('jobNumber','') or '—'}")
    print(f"  Status:       {project_meta.get('issueStatus','') or '—'}")
    print(f"  Drawings:     {len(drawings)} sheets across {len(source_files)} PDF(s)")
    print(f"  Inspections:  {len(inspection_plan)} ({sum(1 for e in inspection_plan if e['holdPoint'])} hold point(s))")
    print(f"  Warnings:     {len(warnings)} schema-level + {len(soft_warnings)} soft validation")
    print()
    print(f"Next step: open the PWA at index.html and import {btproject_path.name}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("project_dir", help="Path to a project folder, e.g. 'Structural Drawings/MBC'")
    parser.add_argument("--dry-run", action="store_true", help="Compute everything but don't write files")
    args = parser.parse_args()
    build(Path(args.project_dir), dry_run=args.dry_run)


if __name__ == "__main__":
    main()
