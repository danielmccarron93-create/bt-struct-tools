"""Extract the cover-sheet STRUCTURAL DRAWING LIST as the authoritative sheet→description map.

The MBC cover sheet has 4 sub-tables of drawing lists (visible at y bands ~806, ~1172, ~1538, ~1926).
Each has rows of drawing names + columns of sheet numbers below.
"""
import fitz, json, re

PDF = "/sessions/compassionate-laughing-fermat/mnt/Structural-Inspections/Structural Drawings/MBC/MBC.pdf"
doc = fitz.open(PDF)
p = doc[0]
d = p.get_text("dict")
spans = []
for block in d["blocks"]:
    if block["type"] != 0:
        continue
    for line in block["lines"]:
        for span in line["spans"]:
            t = span["text"].strip()
            if not t:
                continue
            spans.append((span["bbox"], span["size"], span["font"], t))

# Sheet numbers in the drawing list are S### text at sz=10.17
# Names are also at sz=10.17, longer text
# Pattern observation: sheet numbers appear at FIXED y positions (e.g. y=2046 for one sub-table),
# names appear ABOVE them at varying y positions (because names are rotated and longer = higher y).

# Find all S### sheet number spans with sz=10.17 (in the drawing-list region, not the title block)
sheet_spans = []
for bbox, sz, fn, t in spans:
    if 9.5 < sz < 11 and re.fullmatch(r'S\d{3}', t) and bbox[0] < 1700:
        sheet_spans.append((bbox[0], bbox[1], t))

# Find all candidate name spans (sz=10.17, longer text, in same x range, NOT a sheet number)
name_spans = []
for bbox, sz, fn, t in spans:
    if 9.5 < sz < 11 and not re.fullmatch(r'S\d{3}', t) and bbox[0] < 1700 and len(t) > 4:
        name_spans.append((bbox[0], bbox[1], t))

# For each sheet, find the closest name above it (smaller y) at the same x (within tolerance)
# In rotated coords: the sheet # is at the bottom of a column, the name is above (lower y)
mapping = {}
for sx, sy, sn in sheet_spans:
    # Find name spans with same x (±2) and y less than sheet's y
    candidates = [(nx, ny, nt) for nx, ny, nt in name_spans
                  if abs(nx - sx) < 2.5 and ny < sy]
    if not candidates:
        continue
    # Closest by y (i.e. largest y still less than sy — closest above)
    candidates.sort(key=lambda c: c[1], reverse=True)
    nx, ny, nt = candidates[0]
    # Some sheets have a 2nd line (longer names wrap). Look for another name span
    # with same x just above the first name.
    second_candidates = [(nx2, ny2, nt2) for nx2, ny2, nt2 in name_spans
                         if abs(nx2 - sx) < 2.5 and ny2 < ny - 8]
    if second_candidates:
        second_candidates.sort(key=lambda c: c[1], reverse=True)
        nx2, ny2, nt2 = second_candidates[0]
        # Combine if reasonably close
        if (ny - ny2) < 25:
            mapping[sn] = f"{nt2} {nt}"
            continue
    mapping[sn] = nt

# Print as a sorted table
print(f"Found {len(mapping)} sheet → name mappings")
print()
for sn in sorted(mapping.keys()):
    print(f"  {sn}  {mapping[sn]}")

# Save for next steps
json.dump(mapping, open("/tmp/mbc_drawing_list.json", "w"), indent=2)
print()
print(f"Saved to /tmp/mbc_drawing_list.json")
