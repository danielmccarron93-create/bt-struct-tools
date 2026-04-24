/**
 * Grid-calibration utilities (Phase V2).
 *
 * A drawing can have a `gridCalibration` record of the form:
 *
 *   {
 *     p1: { pdfX, pdfY, col, row },   // e.g. col="1", row="A"
 *     p2: { pdfX, pdfY, col, row },   // e.g. col="5", row="E"
 *     colAxis: 'number' | 'letter',
 *     rowAxis: 'letter' | 'number',
 *     calibratedAt: ISO string
 *   }
 *
 * From p1 & p2 we derive the PDF-space step per column and per row and can
 * resolve any (pdfX, pdfY) to a nearest-intersection grid reference like
 * "3/B" or "C/7". The north star §6.2 / §7.1.4 mandates this for PI defence.
 */

const LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

/** Parse a user-entered label — returns {axis:'letter'|'number', index:0-based}. */
export function parseLabel(label) {
  const s = String(label || '').trim().toUpperCase();
  if (!s) return null;
  if (/^\d+$/.test(s)) {
    const n = parseInt(s, 10);
    if (n < 1) return null;
    return { axis: 'number', index: n - 1, raw: String(n) };
  }
  // Letter — support single- or double-letter grids (A, B,…Z, AA, AB…)
  if (/^[A-Z]+$/.test(s)) {
    let idx = 0;
    for (const c of s) idx = idx * 26 + (LETTERS.indexOf(c) + 1);
    return { axis: 'letter', index: idx - 1, raw: s };
  }
  return null;
}

/** Convert a zero-based index back to a label using the given axis. */
export function formatLabel(index, axis) {
  if (!Number.isFinite(index)) return '';
  if (axis === 'letter') {
    let i = Math.round(index) + 1;
    let out = '';
    while (i > 0) {
      const r = (i - 1) % 26;
      out = LETTERS[r] + out;
      i = Math.floor((i - 1) / 26);
    }
    return out;
  }
  return String(Math.round(index) + 1);
}

/**
 * Given a validated calibration record (raw user input already parsed into
 * axis+index on each point) and an arbitrary pdf-space point, return the
 * nearest grid-intersection label as "col/row" (e.g. "3/B").
 *
 * Returns '' if calibration is missing or degenerate.
 */
export function computeGridRef(pdfX, pdfY, cal) {
  if (!cal || !cal.p1 || !cal.p2) return '';
  const { p1, p2 } = cal;

  const p1col = parseLabel(p1.col);
  const p1row = parseLabel(p1.row);
  const p2col = parseLabel(p2.col);
  const p2row = parseLabel(p2.row);
  if (!p1col || !p1row || !p2col || !p2row) return '';

  // Column axis — assume grids run parallel to one of the PDF axes.
  // We pick whichever axis (x or y) carries the larger delta between p1 & p2
  // for columns; rows use the other.
  const dx = p2.pdfX - p1.pdfX;
  const dy = p2.pdfY - p1.pdfY;
  const colAxisIsX = Math.abs(dx) >= Math.abs(dy);

  const colSpan = (p2col.index - p1col.index);
  const rowSpan = (p2row.index - p1row.index);
  if (colSpan === 0 || rowSpan === 0) return '';

  const colStep = (colAxisIsX ? dx : dy) / colSpan;
  const rowStep = (colAxisIsX ? dy : dx) / rowSpan;
  if (!isFinite(colStep) || !isFinite(rowStep) || colStep === 0 || rowStep === 0) return '';

  const colCoord = colAxisIsX ? pdfX : pdfY;
  const rowCoord = colAxisIsX ? pdfY : pdfX;
  const colOrigin = colAxisIsX ? p1.pdfX : p1.pdfY;
  const rowOrigin = colAxisIsX ? p1.pdfY : p1.pdfX;

  const colIdx = p1col.index + Math.round((colCoord - colOrigin) / colStep);
  const rowIdx = p1row.index + Math.round((rowCoord - rowOrigin) / rowStep);

  const col = formatLabel(colIdx, p1col.axis);
  const row = formatLabel(rowIdx, p1row.axis);
  return `${col}/${row}`;
}

/** Validate a user-entered calibration before save. Throws on bad input. */
export function validateCalibrationInput({ p1col, p1row, p2col, p2row }) {
  const a = parseLabel(p1col), b = parseLabel(p1row);
  const c = parseLabel(p2col), d = parseLabel(p2row);
  if (!a) throw new Error(`Unrecognised column label "${p1col}" (use a number or letter)`);
  if (!b) throw new Error(`Unrecognised row label "${p1row}"`);
  if (!c) throw new Error(`Unrecognised column label "${p2col}"`);
  if (!d) throw new Error(`Unrecognised row label "${p2row}"`);
  if (a.axis !== c.axis) throw new Error('Both column labels must be the same kind (both numbers OR both letters)');
  if (b.axis !== d.axis) throw new Error('Both row labels must be the same kind (both numbers OR both letters)');
  if (a.axis === b.axis) {
    throw new Error('Column and row labels must be different kinds — use numbers for one axis and letters for the other');
  }
  if (a.index === c.index) throw new Error('Both points share the same column — they need to span different columns');
  if (b.index === d.index) throw new Error('Both points share the same row — they need to span different rows');
  return true;
}
