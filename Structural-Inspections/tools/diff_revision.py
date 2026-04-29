"""
diff_revision.py — compute a revision diff between an existing project-map.json
and a new PDF (typically when Tender → Construction Issue arrives, or any
mid-project rev bump).

Output: writes <project>/revision-<timestamp>.json with the diff.
The PWA imports this file via Project → Import revision.

Diff format (consumed by js/components/import-revision.js, schema doc below):

    {
      "schemaVersion": 1,
      "generatedAt":  "2026-04-29T12:34:56Z",
      "againstSourceFile": "MBC.pdf",
      "againstIssue":      "TENDER NOT FOR CONSTRUCTION",   # what the OLD map was at
      "newIssue":          "CONSTRUCTION ISSUE",            # what the NEW PDF is at
      "added":   [{ sheetNumber, revision, description, pageNumber } ...],
      "changed": [{ sheetNumber, oldRevision, newRevision, description } ...],
      "removed": [{ sheetNumber, revision, description } ...],
      "affectedInspectionPlanIndices": [int, ...],
      "summary": "5 added, 3 changed, 1 removed"
    }

Usage:
    python3 tools/diff_revision.py "Structural Drawings/MBC" path/to/new/MBC.pdf
"""
import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]


def extract_pages_from_pdf(pdf_path: Path):
    """Re-run the BT A1 title-block extractor on a fresh PDF.

    Returns a list of {sheetNumber, revision, description, pageNumber}.
    Tries newer template first, falls back to older if extraction is sparse.
    """
    import fitz

    doc = fitz.open(str(pdf_path))
    pages_data = []

    # Newer-template cells (per skills/reading-bt-drawings/SKILL.md Section 10)
    NEWER_CELLS = {
        "sheetNumber":      (1580, 1620,   95, 130,  20, 30),
        "revision":         (1580, 1620,   55,  75,  20, 30),
        "drawingTitleA":    (1520, 1580,  215, 230,  14, 18),
        "drawingTitleB":    (1520, 1580,   58,  80,  14, 18),
    }

    # Older-template cells
    OLDER_CELLS = {
        "sheetNumber":  (1580, 1625,  115, 150,  20, 30),
        "revision":     (1580, 1625,   50,  80,  20, 30),
        "drawingTitle": (1525, 1580,   55,  80,  14, 18),
    }

    def extract_spans(page):
        spans = []
        d = page.get_text("dict")
        for block in d["blocks"]:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    t = span["text"].strip()
                    if not t:
                        continue
                    spans.append((span["bbox"], span["size"], span["font"], t))
        return spans

    def best_in_cell(spans, cell):
        xmin, xmax, ymin, ymax, smin, smax = cell
        hits = [(b, s, f, t) for b, s, f, t in spans
                if xmin <= b[0] <= xmax and ymin <= b[1] <= ymax and smin <= s <= smax]
        if not hits:
            return None
        hits.sort(key=lambda h: (-h[1], h[0][1]))
        return hits[0][3]

    # Detect template: try older on a few pages
    older_hits = sum(
        1 for i in range(min(5, doc.page_count))
        if best_in_cell(extract_spans(doc[i]), OLDER_CELLS["sheetNumber"])
    )
    use_newer = older_hits == 0

    cells = NEWER_CELLS if use_newer else OLDER_CELLS

    for i in range(doc.page_count):
        spans = extract_spans(doc[i])
        sn = best_in_cell(spans, cells["sheetNumber"])
        rev = best_in_cell(spans, cells["revision"])
        if use_newer:
            # MBC template: try title B (top line) + A (bottom line) and concat
            tA = best_in_cell(spans, cells["drawingTitleA"])
            tB = best_in_cell(spans, cells["drawingTitleB"])
            desc = f"{tB} {tA}".strip() if tA and tB else (tA or tB)
        else:
            desc = best_in_cell(spans, cells["drawingTitle"])

        pages_data.append({
            "pageNumber":  i + 1,
            "sheetNumber": sn,
            "revision":    rev,
            "description": desc,
        })

    print(f"  template detected: {'newer' if use_newer else 'older'} BT A1")
    print(f"  pages extracted:   {len(pages_data)}")
    print(f"  with sheet#:       {sum(1 for p in pages_data if p['sheetNumber'])}")
    return pages_data


def diff(project_dir: Path, new_pdf: Path):
    project_dir = project_dir.resolve()
    if not project_dir.is_dir():
        raise SystemExit(f"Not a directory: {project_dir}")

    map_path = project_dir / "project-map.json"
    if not map_path.exists():
        raise SystemExit(f"No project-map.json in {project_dir} — run build_btproject.py first")

    if not new_pdf.exists():
        raise SystemExit(f"New PDF not found: {new_pdf}")

    project_map = json.loads(map_path.read_text())
    print(f"Loaded existing project-map.json")
    print(f"  Project:     {project_map['project'].get('name')}")
    print(f"  Issue:       {project_map['project'].get('issueStatus')}")
    print(f"  Drawings:    {len(project_map['drawings'])}")
    print(f"  Plan items:  {len(project_map['inspectionPlan'])}")
    print()
    print(f"Extracting new PDF: {new_pdf.name}")
    new_pages = extract_pages_from_pdf(new_pdf)
    print()

    # Build maps keyed by uppercase sheet number
    old_by_sheet = {(d.get("sheetNumber") or "").upper(): d for d in project_map["drawings"]}
    old_by_sheet.pop("", None)
    new_by_sheet = {(p.get("sheetNumber") or "").upper(): p for p in new_pages}
    new_by_sheet.pop("", None)

    added, changed, removed = [], [], []

    for sn, new_page in new_by_sheet.items():
        old = old_by_sheet.get(sn)
        if old is None:
            added.append({
                "sheetNumber": new_page["sheetNumber"],
                "revision":    new_page["revision"],
                "description": new_page["description"],
                "pageNumber":  new_page["pageNumber"],
            })
        else:
            old_rev = (old.get("revision") or "").strip()
            new_rev = (new_page.get("revision") or "").strip()
            if new_rev and old_rev != new_rev:
                changed.append({
                    "sheetNumber":  new_page["sheetNumber"],
                    "oldRevision":  old_rev or None,
                    "newRevision":  new_rev,
                    "description":  new_page["description"] or old.get("description"),
                })

    for sn, old in old_by_sheet.items():
        if sn not in new_by_sheet:
            removed.append({
                "sheetNumber": old["sheetNumber"],
                "revision":    old.get("revision"),
                "description": old.get("description"),
            })

    # Compute affected inspection-plan indices
    affected_plan = set()
    affected_sheets = (
        {a["sheetNumber"].upper() for a in added if a.get("sheetNumber")} |
        {c["sheetNumber"].upper() for c in changed if c.get("sheetNumber")} |
        {r["sheetNumber"].upper() for r in removed if r.get("sheetNumber")}
    )
    for idx, entry in enumerate(project_map["inspectionPlan"]):
        for ref in entry.get("drawingRefs", []):
            sn = (ref.get("sheetNumber") or "").upper()
            if sn in affected_sheets:
                affected_plan.add(idx)
                break

    affected_plan_list = sorted(affected_plan)

    diff_obj = {
        "schemaVersion":      1,
        "generatedAt":        datetime.now(timezone.utc).isoformat(),
        "againstSourceFile":  project_map["sourceFiles"][0]["filename"] if project_map.get("sourceFiles") else None,
        "againstIssue":       project_map["project"].get("issueStatus"),
        "newPdfFilename":     new_pdf.name,
        "newPdfSize":         new_pdf.stat().st_size,
        "added":              added,
        "changed":            changed,
        "removed":            removed,
        "affectedInspectionPlanIndices": affected_plan_list,
        "summary": (
            f"{len(added)} added, {len(changed)} changed, {len(removed)} removed; "
            f"{len(affected_plan_list)} inspection plan entr"
            f"{'y' if len(affected_plan_list) == 1 else 'ies'} affected"
        ),
    }

    ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    out_path = project_dir / f"revision-{ts}.json"
    out_path.write_text(json.dumps(diff_obj, indent=2))
    print(f"✓ Wrote {out_path}")
    print()
    print(f"=== Diff Summary ===")
    print(f"  Added:    {len(added)} sheet{'s' if len(added) != 1 else ''}")
    if added:
        for a in added[:8]:
            print(f"    + {a['sheetNumber']} rev {a['revision'] or '—'}  {a.get('description') or ''}")
        if len(added) > 8:
            print(f"    … and {len(added) - 8} more")
    print(f"  Changed:  {len(changed)} sheet{'s' if len(changed) != 1 else ''}")
    if changed:
        for c in changed[:8]:
            print(f"    Δ {c['sheetNumber']} {c['oldRevision'] or '—'} → {c['newRevision']}  {c.get('description') or ''}")
        if len(changed) > 8:
            print(f"    … and {len(changed) - 8} more")
    print(f"  Removed:  {len(removed)} sheet{'s' if len(removed) != 1 else ''}")
    if removed:
        for r in removed[:8]:
            print(f"    - {r['sheetNumber']} rev {r['revision'] or '—'}  {r.get('description') or ''}")
        if len(removed) > 8:
            print(f"    … and {len(removed) - 8} more")
    print(f"  Affected plan items: {len(affected_plan_list)}")
    if affected_plan_list:
        for idx in affected_plan_list[:6]:
            entry = project_map["inspectionPlan"][idx]
            print(f"    #{entry.get('sequence', idx + 1)}: {entry.get('title', entry['type'])}")
        if len(affected_plan_list) > 6:
            print(f"    … and {len(affected_plan_list) - 6} more")
    print()
    print(f"Next step: in the PWA, open this project → Project menu → 'Import revision' → select {out_path.name}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("project_dir", help="Path to a project folder, e.g. 'Structural Drawings/MBC'")
    parser.add_argument("new_pdf", help="Path to the new PDF to compare against the existing project-map.json")
    args = parser.parse_args()
    diff(Path(args.project_dir), Path(args.new_pdf))


if __name__ == "__main__":
    main()
