"""BT A1 title-block extractor — same pattern proven on 52SA + Emmanuel.

Per skills/reading-bt-drawings/SKILL.md: use (x0, y0) anchors, not center coords,
because rotated text bbox extends along the rotated baseline.
"""
import fitz, json, sys

PDF = "/sessions/compassionate-laughing-fermat/mnt/Structural-Inspections/Structural Drawings/MBC/MBC.pdf"

# MBC template variant cells (different from 52SA/Emmanuel — see skill notes).
#
# Differences from the older BT A1 template:
#   - sheetNumber moved from y≈132 to y≈110
#   - jobNumber moved from y≈259 to y≈209
#   - drawingTitle moved from y≈64 to y≈220 (and can wrap to a 2nd line at y≈66)
#   - projectName moved from y≈64 to y≈117 with smaller font (sz 16)
#   - siteAddress now SPLIT across 2 spans (y≈63 + y≈189)
#   - issueStatus stamp moved from right-edge bottom (y≈812) to UPPER-LEFT (x≈810, y≈100-160)
#   - drawn/design/checked initials moved from y≈629 to y≈201
#
CELLS = {
    "sheetNumber":      (1580, 1620,   95, 130,  20, 30),
    "revision":         (1580, 1620,   55,  75,  20, 30),
    "jobNumber":        (1580, 1620,  195, 225,  20, 30),
    "drawingTitleA":    (1520, 1580,  215, 230,  14, 18),  # primary slot
    "drawingTitleB":    (1520, 1580,   58,  80,  14, 18),  # wrap line 1 (when title is long)
    "projectName":      (1430, 1455,  110, 130,  15, 19),
    "siteAddressLine1": (1455, 1475,   58,  80,  13, 16),
    "siteAddressLine2": (1475, 1495,  180, 200,  13, 16),
    "issueLine1":       (795, 875,   140, 165,  25, 30),   # "TENDER" or "CONSTRUCTION"
    "issueLine2":       (845, 880,    95, 115,  16, 20),   # "NOT FOR CONSTRUCTION" or "ISSUE"
    # Initials — three side-by-side at y≈201, x clusters around 1344/1372/1400
    "drawnBy":          (1340, 1360,  195, 210,  11, 13),
    "designBy":         (1368, 1385,  195, 210,  11, 13),
    "checkedBy":        (1395, 1410,  195, 210,  11, 13),
}
SCALE_CELL = (1575, 1625, 905, 925, 6, 8)


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
    hits = []
    for bbox, sz, fn, t in spans:
        x = bbox[0]
        y = bbox[1]
        if xmin <= x <= xmax and ymin <= y <= ymax and smin <= sz <= smax:
            hits.append((bbox, sz, fn, t))
    if not hits:
        return None
    hits.sort(key=lambda h: (-h[1], h[0][1]))
    return hits[0][3]


def all_in_cell(spans, cell):
    xmin, xmax, ymin, ymax, smin, smax = cell
    hits = []
    for bbox, sz, fn, t in spans:
        x = bbox[0]
        y = bbox[1]
        if xmin <= x <= xmax and ymin <= y <= ymax and smin <= sz <= smax:
            hits.append((bbox, sz, fn, t))
    hits.sort(key=lambda h: (h[0][1], h[0][0]))
    return [h[3] for h in hits]


def confidence_score(rec):
    score = 1.0
    if not rec.get("sheetNumber"):
        score -= 0.4
    if not rec.get("revision"):
        score -= 0.15
    if not rec.get("description"):
        score -= 0.15
    return max(0.0, round(score, 2))


def extract_page(page, pn):
    spans = extract_spans(page)
    rec = {"pageNumber": pn}
    for name, cell in CELLS.items():
        rec[name] = best_in_cell(spans, cell)
    scales = all_in_cell(spans, SCALE_CELL)
    rec["scale"] = " / ".join(scales) if scales else None

    # Combine multi-line drawing title (line B + line A) when both present
    titleA = rec.get("drawingTitleA")
    titleB = rec.get("drawingTitleB")
    if titleA and titleB:
        # Line B (top) is the start, line A (bottom) is the continuation
        description = f"{titleB} {titleA}"
    else:
        description = titleA or titleB

    # Combine multi-line site address
    addrL1 = rec.get("siteAddressLine1")
    addrL2 = rec.get("siteAddressLine2")
    if addrL1 and addrL2:
        siteAddress = f"{addrL1} {addrL2}"
    else:
        siteAddress = addrL1 or addrL2

    if rec.get("issueLine1") or rec.get("issueLine2"):
        # Note: in MBC template, issueLine2 (NOT FOR CONSTRUCTION, sz 18) is ABOVE issueLine1 (TENDER, sz 27)
        issueStatus = " ".join(
            x for x in [rec.get("issueLine1"), rec.get("issueLine2")] if x
        )
    else:
        issueStatus = None

    out = {
        "pageNumber": pn,
        "sheetNumber": rec.get("sheetNumber"),
        "revision": rec.get("revision"),
        "revisionDate": None,  # MBC doesn't have a clean revision-date cell where we expected
        "description": description,
        "scale": rec.get("scale"),
        "drawnBy": rec.get("drawnBy"),
        "checkedBy": rec.get("checkedBy"),
        "approvedBy": rec.get("designBy"),
        "_projectName": rec.get("projectName"),
        "_siteAddress": siteAddress,
        "_jobNumber": rec.get("jobNumber"),
        "_client": None,  # MBC template doesn't appear to have a separate client cell — site name = school name
        "_issueStatus": issueStatus,
    }
    out["confidence"] = confidence_score(out)
    return out


def main():
    doc = fitz.open(PDF)
    pages = [extract_page(doc[i], i + 1) for i in range(doc.page_count)]
    print(json.dumps(pages, indent=2))


if __name__ == "__main__":
    main()
