"""
verify_btproject.py — simulate js/lib/btproject.js → readBundle() to confirm the bundle
is consumable by the PWA before you import it.

Mirrors the JS validator's logic:
  - schemaVersion must equal SUPPORTED_SCHEMA_VERSION (1)
  - project.name required and non-empty
  - sourceFiles[] non-empty array
  - drawings[] is an array
  - inspectionPlan[] is an array
  - PDFs in 'Drawings/' folder (NOT 'pdfs/'!)
  - Every sourceFiles[].filename has a matching PDF in the zip
  - Every drawings[].sourceFile resolves to a sourceFiles entry
  - Every inspectionPlan[].drawingRefs[].sheetNumber resolves to a drawings entry

Soft warnings printed; hard failures raise.
"""
import argparse
import json
import sys
import zipfile
from pathlib import Path

SUPPORTED_SCHEMA_VERSION = 1
EXPECTED_PDF_FOLDER = "Drawings/"


def verify(bundle_path: Path) -> tuple[bool, list[str]]:
    """Returns (ok, soft_warnings)."""
    if not bundle_path.is_file():
        raise SystemExit(f"Not a file: {bundle_path}")

    print(f"Verifying {bundle_path} ({bundle_path.stat().st_size/1024/1024:.1f} MB)")
    print()

    soft = []
    with zipfile.ZipFile(bundle_path, "r") as zf:
        names = zf.namelist()

        # 1) project-map.json present?
        if "project-map.json" not in names:
            raise SystemExit("FAIL: bundle is missing project-map.json")
        print("✓ project-map.json present")

        # 2) Parse it
        try:
            project_map = json.loads(zf.read("project-map.json"))
        except Exception as e:
            raise SystemExit(f"FAIL: project-map.json not valid JSON: {e}")
        print("✓ project-map.json parses as JSON")

        # 3) Schema version
        sv = project_map.get("schemaVersion")
        if sv != SUPPORTED_SCHEMA_VERSION:
            raise SystemExit(
                f"FAIL: schemaVersion is {sv}, PWA reader supports v{SUPPORTED_SCHEMA_VERSION}"
            )
        print(f"✓ schemaVersion = {sv}")

        # 4) Project name
        proj = project_map.get("project", {})
        if not str(proj.get("name", "")).strip():
            raise SystemExit("FAIL: project.name is missing or empty")
        print(f"✓ project.name = '{proj.get('name')}'")

        # 5) sourceFiles[]
        sf = project_map.get("sourceFiles")
        if not isinstance(sf, list) or not sf:
            raise SystemExit("FAIL: sourceFiles[] is missing or empty")
        print(f"✓ sourceFiles[] has {len(sf)} entries")

        # 6) drawings[]
        if not isinstance(project_map.get("drawings"), list):
            raise SystemExit("FAIL: drawings[] is missing or not an array")
        print(f"✓ drawings[] has {len(project_map['drawings'])} entries")

        # 7) inspectionPlan[]
        if not isinstance(project_map.get("inspectionPlan"), list):
            raise SystemExit("FAIL: inspectionPlan[] is missing or not an array")
        print(f"✓ inspectionPlan[] has {len(project_map['inspectionPlan'])} entries")

        # 8) PDFs in Drawings/ folder
        pdf_entries = [n for n in names if n.startswith(EXPECTED_PDF_FOLDER) and n.lower().endswith(".pdf")]
        if not pdf_entries:
            wrong_pdf = [n for n in names if n.startswith("pdfs/") and n.lower().endswith(".pdf")]
            if wrong_pdf:
                raise SystemExit(
                    f"FAIL: PDFs are in 'pdfs/' folder ({len(wrong_pdf)}) — PWA reader expects 'Drawings/'!"
                )
            raise SystemExit(f"FAIL: no PDFs found in '{EXPECTED_PDF_FOLDER}' folder of the bundle")
        print(f"✓ {len(pdf_entries)} PDF(s) in 'Drawings/' folder of the bundle")

        # 9) Every sourceFiles[].filename has a matching PDF
        expected_filenames = [s["filename"] for s in sf]
        zip_filenames = [Path(n).name for n in pdf_entries]
        missing = [f for f in expected_filenames if f not in zip_filenames]
        extras = [f for f in zip_filenames if f not in expected_filenames]
        if missing:
            raise SystemExit(f"FAIL: missing PDF(s) in zip: {missing}")
        if extras:
            soft.append(f"Extra PDFs in zip not declared in sourceFiles[]: {extras}")
        print(f"✓ All {len(expected_filenames)} declared PDFs present in zip")

        # 10) Every drawings[].sourceFile resolves
        source_set = set(expected_filenames)
        orphans = [d for d in project_map["drawings"] if d.get("sourceFile") not in source_set]
        if orphans:
            soft.append(f"{len(orphans)} drawing(s) reference a sourceFile not in sourceFiles[]")

        # 11) Every drawingRef sheetNumber resolves
        sheet_set = {(d.get("sheetNumber") or "").upper() for d in project_map["drawings"]}
        sheet_set.discard("")
        bad_refs = 0
        for entry in project_map["inspectionPlan"]:
            for ref in entry.get("drawingRefs", []):
                sn = (ref.get("sheetNumber") or "").upper()
                if sn and sn not in sheet_set:
                    bad_refs += 1
        if bad_refs:
            soft.append(
                f"{bad_refs} inspection-plan drawingRef(s) point at sheets not in drawings[]"
            )

        # 12) Every inspection-plan type is in the catalogue
        ref_path = Path(__file__).resolve().parent.parent / "reference" / "inspection-types.json"
        if ref_path.exists():
            catalogue = json.loads(ref_path.read_text())
            valid_types = {t["key"] for t in catalogue["types"]}
            unknown = sorted({e["type"] for e in project_map["inspectionPlan"] if e.get("type") not in valid_types})
            if unknown:
                raise SystemExit(
                    f"FAIL: inspection-plan uses keys not in inspection-types.json: {unknown}"
                )
            print(f"✓ All inspection types in plan are valid catalogue keys ({len(valid_types)} types in catalogue)")

        # 13) manifest.txt (optional)
        if "manifest.txt" in names:
            manifest = zf.read("manifest.txt").decode("utf-8", errors="replace")
            print(f"✓ manifest.txt present ({len(manifest)} chars)")
        else:
            soft.append("manifest.txt is missing (optional but recommended)")

        # 14) Stats
        print()
        print("=== Summary ===")
        print(f"  Schema version:    {sv}")
        print(f"  Project:           {proj.get('name')}")
        print(f"  Job#:              {proj.get('jobNumber') or '—'}")
        print(f"  Issue status:      {proj.get('issueStatus') or '—'}")
        print(f"  PDFs:              {len(pdf_entries)} ({sum(s.get('pageCount', 0) for s in sf)} pages)")
        print(f"  Drawings:          {len(project_map['drawings'])}")
        print(f"  Inspection plan:   {len(project_map['inspectionPlan'])}")
        plan = project_map['inspectionPlan']
        holds = sum(1 for e in plan if e.get('holdPoint'))
        print(f"  Hold points:       {holds}")
        print(f"  Warnings (proj):   {len(project_map.get('warnings', []))}")
        print(f"  excludedFromBT:    {len(proj.get('excludedFromBT', []))}")

        # Plan grouped by type
        from collections import Counter
        type_counts = Counter(e["type"] for e in plan)
        print(f"  Plan by type:")
        for t, c in sorted(type_counts.items()):
            print(f"    {t:30s} {c}")

    print()
    if soft:
        print("Soft warnings (non-blocking):")
        for w in soft:
            print(f"  ⚠  {w}")
    else:
        print("No soft warnings.")
    print()
    print("✓ Bundle is valid and ready for PWA import.")
    return True, soft


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("bundle_path", help="Path to a .btproject file")
    args = parser.parse_args()
    verify(Path(args.bundle_path))


if __name__ == "__main__":
    main()
