/**
 * PDF.js wrapper helpers.
 *
 * All functions here work with PDF.js loaded globally as window.pdfjsLib
 * (via the <script> tag in index.html). The worker path is also configured
 * there.
 *
 * Coordinate system notes:
 *   - PDF.js getViewport() returns a transform that converts PDF points
 *     (origin = page bottom-left, y up) to canvas pixels (origin = top-left, y down).
 *   - For consistency and to survive zoom/rotate, we store *PDF-space*
 *     coordinates in the database (p1, p2 for calibration, and later pin
 *     drops). Translation to screen happens only at render time.
 */

/* global pdfjsLib */

/**
 * Load a PDF document from a Blob.
 * Returns the pdfjsLib document proxy.
 */
export async function loadPdf(blob) {
  if (!window.pdfjsLib) throw new Error('PDF.js not loaded');
  const data = await blob.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ data });
  return loadingTask.promise;
}

/**
 * Get the first page's un-rotated dimensions in PDF points.
 */
export async function getFirstPageDimensions(blob) {
  const doc = await loadPdf(blob);
  try {
    const page = await doc.getPage(1);
    const viewport = page.getViewport({ scale: 1 });
    return { width: viewport.width, height: viewport.height };
  } finally {
    doc.destroy();
  }
}

/**
 * Inspect every page of a PDF and return per-page metadata:
 *
 *   [{ pageNumber, width, height, sheetNumber?, description?, revision? }, ...]
 *
 * Uses getTextContent() to hunt for drawing-register patterns in the title
 * block. This is best-effort — empty strings are returned when nothing
 * sensible matches, and the engineer can edit each drawing after upload.
 *
 * Patterns this tries:
 *   sheetNumber:  /\bS\d{3}\b/i  (classic Bligh Tanner structural prefix)
 *                 or fallback A\d{3}, C\d{3}, H\d{3}, M\d{3}, E\d{3}, L\d{3}, ST\d{3}
 *   description:  heuristic — the text immediately following "DRAWING TITLE"
 *                 or just above the sheet number
 *   revision:     /\bC\d+\b|\bREV[\s.-]*([A-Z]?\d+)\b/  after the sheet number
 */
export async function extractPageMetadata(blob) {
  const doc = await loadPdf(blob);
  try {
    const out = [];
    for (let i = 1; i <= doc.numPages; i += 1) {
      const page = await doc.getPage(i);
      const viewport = page.getViewport({ scale: 1 });

      let sheetNumber = '';
      let description = '';
      let revision = '';
      try {
        const content = await page.getTextContent();
        const items = (content.items || []).map((it) => ({
          str: (it.str || '').trim()
        })).filter((it) => it.str.length > 0);

        // Concatenate all text for regex-over-stream. Join with spaces so
        // adjacent items like ["S010", "C1"] become searchable as "S010 C1".
        const joined = items.map((it) => it.str).join(' ');

        // 1) Sheet + revision together — this pattern is exclusive to the
        // title block ("S010 C1"). Body-text cross-references to other sheets
        // appear bare ("see S031"), so this avoids false positives.
        const combinedRe = /\b(S[TA]?\d{3}|[ACHMEL]\d{3})\s+(C\d+|[A-Z]\d?)\b/g;
        let lastMatch = null;
        let m;
        while ((m = combinedRe.exec(joined)) !== null) {
          lastMatch = m; // Prefer the LAST occurrence — title block renders near the end
        }
        if (lastMatch) {
          sheetNumber = lastMatch[1].toUpperCase();
          revision    = lastMatch[2].toUpperCase();
        }

        // Fallback — just a sheet number (no revision matched)
        if (!sheetNumber) {
          const sheetRe = /\b(S[TA]?\d{3}|[ACHMEL]\d{3})\b/g;
          let bareLast = null;
          while ((m = sheetRe.exec(joined)) !== null) bareLast = m;
          if (bareLast) sheetNumber = bareLast[1].toUpperCase();
        }

        // 2) Description — take the text immediately after "DRAWING TITLE".
        // Ignore label tokens that belong to adjacent title-block fields,
        // and reject the licensing / copyright block that often appears right
        // next to the title in PDF text order (even if visually separated).
        const STOP = /^(JOB|REV|ISSUED|DATE|DRAWN|DESIGN|CHECKED|PROJECT|CLIENT|DRAWING\s+NUMBER|REVISION|SCALES?|STATUS|NORTH\s+POINT|ARCHITECT|ENGINEER|CONSULTANT|DISCIPLINE|STATE|CHECKER|SCALE|PRINT|SIGNED)/i;
        // Schedule codes (RW-B, PF-A, EB-1) pollute page-5-style plans — skip any
        // short token that is ALL CAPS with an embedded hyphen or slash.
        const SCHED_CODE = /^[A-Z]{1,4}[-\/][A-Z0-9]{1,4}$/;
        // Legal / copyright / commercial-terms phrases that sometimes land
        // next to DRAWING TITLE in the text stream. Any match blocks the token.
        const LEGAL_RE = /\b(ALL\s+RIGHTS\s+RESERVED|COPYRIGHT|REPRODUCED|WRITTEN\s+PERMISSION|LICENCE|LICENSE|INSTRUCTING\s+PARTY|BLIGH\s+TANNER|CONSULTING\s+ENGINEERS|PRINT\s+THIS)\b/i;
        // Generic label words that sneak in because DRAWING TITLE and adjacent
        // fields aren't laid out in reading order in the PDF text stream.
        const LABEL_RE = /^(ARCHITECT|ENGINEER|CONSULTANT|REVISION|ISSUE|DATE|DRAWN|DESIGN|CHECKED|CLIENT|JOB\s*NUMBER|SCALES?|STATUS|SIGNED|SIGNATURE|PROJECT|DISCIPLINE)$/i;
        // Lone scale tokens like "1:100", "1:50", "NTS", "N.T.S."
        const SCALE_RE = /^(\d+\s*:\s*\d+|NTS|N\.T\.S\.?)$/i;

        const idxTitle = items.findIndex((it) => /^DRAWING\s+TITLE$/i.test(it.str));
        if (idxTitle >= 0) {
          const after = items.slice(idxTitle + 1, idxTitle + 20)
            .map((it) => it.str)
            .filter((s) =>
              s
              && s.length >= 4               // "1:100" is 5 chars but already caught by SCALE_RE
              && s.length <= 80              // drops the 300-char copyright para
              && !STOP.test(s)
              && !SCHED_CODE.test(s)
              && !LEGAL_RE.test(s)
              && !LABEL_RE.test(s)
              && !SCALE_RE.test(s)
              && !/^\d{4}[./-]\d+/.test(s)   // "2024.123" style job numbers
              && !/^\d{1,4}$/.test(s)        // bare numbers (scales, dates)
              && !/^[A-Z]\d{1,3}$/.test(s)   // sheet codes like C1, S028
              // Require at least one lower-case letter OR an all-caps phrase of
              // at least 3 words — that excludes 1-2 word label fragments like
              // "ARCHITECT SIGNED" but accepts "LEVEL 2 GENERAL ARRANGEMENT".
              && (/[a-z]/.test(s) || s.trim().split(/\s+/).length >= 3)
            )
            .slice(0, 2);
          if (after.length) {
            description = after.join(' ')
              .replace(/\s+-\s+SHEET/i, ' — Sheet') // Bligh Tanner uses "DETAILS - SHEET 1"
              .trim();
            // Final belt-and-braces — if it STILL looks like legal text, drop it.
            if (LEGAL_RE.test(description) || LABEL_RE.test(description) ||
                description.length > 120 || description.split(/\s+/).length < 2) {
              description = '';
            }
          }
        }
      } catch (err) {
        // getTextContent can fail on scanned PDFs — that's fine, keep empties.
      }

      out.push({
        pageNumber: i,
        width:  viewport.width,
        height: viewport.height,
        sheetNumber,
        description,
        revision
      });
    }
    return out;
  } finally {
    doc.destroy();
  }
}

/**
 * Render a specific page to a canvas at a given scale and rotation.
 * Returns the viewport used so callers can map coordinates.
 */
export async function renderPageToCanvas(doc, pageNumber, canvas, scale, rotation = 0) {
  const page = await doc.getPage(pageNumber);
  const viewport = page.getViewport({ scale, rotation });
  const ctx = canvas.getContext('2d', { alpha: false });

  // DPR-aware rendering so it stays crisp on retina screens
  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  canvas.width  = Math.floor(viewport.width  * dpr);
  canvas.height = Math.floor(viewport.height * dpr);
  canvas.style.width  = `${Math.floor(viewport.width)}px`;
  canvas.style.height = `${Math.floor(viewport.height)}px`;

  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  await page.render({ canvasContext: ctx, viewport }).promise;
  return viewport;
}

/**
 * Render a small thumbnail of a specific page (default: 1) into a canvas
 * element and return a data-URL. Used in drawing lists.
 */
export async function renderThumbnail(blob, maxWidth = 320, pageNumber = 1) {
  const doc = await loadPdf(blob);
  try {
    const page = await doc.getPage(pageNumber);
    const baseViewport = page.getViewport({ scale: 1 });
    const scale = Math.min(maxWidth / baseViewport.width, 1.5);
    const viewport = page.getViewport({ scale });

    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    const ctx = canvas.getContext('2d', { alpha: false });
    await page.render({ canvasContext: ctx, viewport }).promise;
    return canvas.toDataURL('image/png');
  } finally {
    doc.destroy();
  }
}

/**
 * Extract per-page text spans with positions and font sizes — for feeding to
 * an LLM that has to find the title block regardless of where it sits on the
 * page (different BT templates put it in different places).
 *
 * Returns an array, one entry per requested page:
 *
 *   [{
 *     pageNumber, width, height, rotation,
 *     spans: [
 *       { text, x, y, w, h, fontSize, fontName }, ...
 *     ]
 *   }, ...]
 *
 * Coordinates are in *natural-rotation* viewport pixels at scale 1 — i.e. the
 * geometry as a human reading the rendered page would see it. Rotation
 * (`rotation`) is the page's /Rotate hint, included so the LLM can interpret
 * coords if it needs to.
 *
 * @param {Blob}  blob
 * @param {Object} [opts]
 * @param {number[]} [opts.pageNumbers]  — limit to specific pages (1-indexed).
 * @param {number} [opts.minFontSize=8]  — drop tiny annotations.
 */
export async function extractPageSpans(blob, { pageNumbers, minFontSize = 8 } = {}) {
  const doc = await loadPdf(blob);
  try {
    const pages = pageNumbers && pageNumbers.length
      ? pageNumbers.filter((n) => n >= 1 && n <= doc.numPages)
      : Array.from({ length: doc.numPages }, (_, i) => i + 1);

    const out = [];
    for (const pageNumber of pages) {
      const page = await doc.getPage(pageNumber);
      // Use the page's natural /Rotate so coords match what the LLM would
      // see if we sent it a rendered image — keeps text+vision in agreement.
      const viewport = page.getViewport({ scale: 1, rotation: page.rotate || 0 });
      const content = await page.getTextContent();

      const spans = [];
      for (const item of content.items || []) {
        const text = (item.str || '').trim();
        if (!text) continue;

        // Apply viewport transform to the text item's transform to get
        // top-left coords + font size in display space.
        // PDF.js helper: util.transform(viewport.transform, item.transform).
        const tx = pdfjsLib.Util.transform(viewport.transform, item.transform);
        // tx = [a, b, c, d, e, f]. Font size = sqrt(c^2 + d^2) (vertical scale).
        const fontSize = Math.hypot(tx[2], tx[3]) || Math.hypot(tx[0], tx[1]) || 0;
        if (fontSize < minFontSize) continue;

        // tx[4], tx[5] is the BASELINE position in display space (y down).
        // Top-left ≈ (tx[4], tx[5] - fontSize). Width is item.width scaled.
        const w = (item.width || 0) * Math.hypot(tx[0], tx[1]) / Math.hypot(item.transform[0], item.transform[1] || 1);
        const x = round(tx[4]);
        const y = round(tx[5] - fontSize);
        spans.push({
          text,
          x, y,
          w: round(w),
          h: round(fontSize),
          fontSize: round(fontSize),
          fontName: item.fontName || ''
        });
      }

      out.push({
        pageNumber,
        width:    Math.round(viewport.width),
        height:   Math.round(viewport.height),
        rotation: page.rotate || 0,
        spans
      });
    }
    return out;
  } finally {
    doc.destroy();
  }
}

/**
 * Extract concatenated text for one or more pages — used for the General
 * Notes pages when we want to feed the whole text body (not positions) to
 * the LLM.
 *
 * Returns [{ pageNumber, text, charCount }].
 */
export async function extractPageText(blob, pageNumbers) {
  const doc = await loadPdf(blob);
  try {
    const pages = pageNumbers && pageNumbers.length
      ? pageNumbers.filter((n) => n >= 1 && n <= doc.numPages)
      : Array.from({ length: doc.numPages }, (_, i) => i + 1);

    const out = [];
    for (const pageNumber of pages) {
      const page    = await doc.getPage(pageNumber);
      const content = await page.getTextContent();
      const text = (content.items || [])
        .map((it) => it.str || '')
        .filter(Boolean)
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();
      out.push({ pageNumber, text, charCount: text.length });
    }
    return out;
  } finally {
    doc.destroy();
  }
}

/**
 * Render a page to a base64 PNG/JPEG data URL — used for the vision-fallback
 * branch of the title-block extractor (when a page has no extractable text,
 * e.g. scanned PDFs).
 *
 * @param {Blob}   blob
 * @param {number} pageNumber       — 1-indexed
 * @param {Object} [opts]
 * @param {number} [opts.maxPx=1024]
 * @param {string} [opts.format='image/jpeg']
 * @param {number} [opts.quality=0.85]
 * @returns {Promise<{ base64: string, mediaType: string, width: number, height: number }>}
 */
export async function renderPageAsImage(blob, pageNumber, { maxPx = 1024, format = 'image/jpeg', quality = 0.85 } = {}) {
  const doc = await loadPdf(blob);
  try {
    const page = await doc.getPage(pageNumber);
    const baseViewport = page.getViewport({ scale: 1, rotation: page.rotate || 0 });
    const longSide = Math.max(baseViewport.width, baseViewport.height);
    const scale = Math.min(maxPx / longSide, 1.5);
    const viewport = page.getViewport({ scale, rotation: page.rotate || 0 });

    const canvas = document.createElement('canvas');
    canvas.width  = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    const ctx = canvas.getContext('2d', { alpha: false });
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport }).promise;

    const dataUrl  = canvas.toDataURL(format, quality);
    const base64   = dataUrl.split(',', 2)[1] || '';
    const mediaType = format;
    return { base64, mediaType, width: canvas.width, height: canvas.height };
  } finally {
    doc.destroy();
  }
}

function round(n) { return Math.round(n * 10) / 10; }

/**
 * Extract a likely sheet number from a filename.
 *
 * Patterns it catches:
 *   "23123 S-01 Rev B.pdf"    → "S-01"
 *   "ABC_SK01.pdf"             → "SK01"
 *   "S100.pdf"                 → "S100"
 *   "Drawing S.02 RevA.pdf"    → "S.02"
 *   "A-100.pdf"                → "A-100"
 */
export function extractSheetNumberFromFilename(filename) {
  if (!filename) return '';
  const base = filename.replace(/\.pdf$/i, '').replace(/[_]/g, ' ');
  // Look for classic drawing codes: SK, S, A, ST, C, M, E, L, H, P, Q
  const m = base.match(/\b(SK|ST|S|A|C|M|E|L|H|P|Q|T|D)[-\s.]?\d{1,4}[A-Za-z]?\b/i);
  if (m) return m[0].replace(/\s/g, '-').toUpperCase();
  return '';
}

/**
 * Extract a likely revision from a filename.
 *   "Rev A", "REV.B", "revision 2", "Rev.P1"
 */
export function extractRevisionFromFilename(filename) {
  if (!filename) return '';
  const m = filename.match(/rev(?:ision)?[-._\s]*([A-Z]?\d{0,2}[A-Z]?\d{0,2})/i);
  return m ? m[1].toUpperCase() : '';
}

/**
 * Convert a distance between two PDF-space points (in points) to mm,
 * given a drawing's calibration record.
 */
export function pdfDistanceToMm(p1, p2, calibration) {
  if (!calibration || !calibration.mmPerPoint) return null;
  const dx = p2.x - p1.x;
  const dy = p2.y - p1.y;
  const distPoints = Math.sqrt(dx * dx + dy * dy);
  return distPoints * calibration.mmPerPoint;
}
