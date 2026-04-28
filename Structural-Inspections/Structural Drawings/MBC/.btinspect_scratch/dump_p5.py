"""Dump every span on page 5 of MBC.pdf — find the new template's title-block geometry."""
import fitz
PDF = "/sessions/compassionate-laughing-fermat/mnt/Structural-Inspections/Structural Drawings/MBC/MBC.pdf"
doc = fitz.open(PDF)
p = doc[4]  # page 5
print(f"page rect: {p.rect}, rotation {p.rotation}")
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
print(f"Total spans on p5: {len(spans)}")
print()
print("--- All spans with x0 > 1400 (title block area), sorted by y0 ---")
right = sorted([s for s in spans if s[0][0] > 1400], key=lambda s: s[0][1])
for bbox, sz, fn, t in right:
    print(f"  x0={bbox[0]:7.1f} y0={bbox[1]:7.1f}  sz={sz:5.2f}  {t!r}")
print()
print("--- Top 30 spans by font size ---")
for bbox, sz, fn, t in sorted(spans, key=lambda s: -s[1])[:30]:
    print(f"  sz={sz:5.2f}  x0={bbox[0]:7.1f} y0={bbox[1]:7.1f}  {t!r}")
