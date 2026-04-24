/**
 * Report generation (Phase 6).
 *
 * Produces a multi-page A4 PDF in Bligh Tanner house style:
 *   Page 1 — Cover: logo, title, project metadata, general observations,
 *                  inspector sign-off.
 *   Page 2 — Marked plan: source drawing rasterised with pins + extent
 *                        highlights, legend.
 *   Page 3+ — Item schedule: numbered table of items (paginates).
 *   Page ?+ — Photo appendix: 2 per page, captioned with item number.
 *
 * Uses jsPDF (loaded via index.html). Async because we rasterise the source
 * drawing and re-compress photos for embedding.
 */

import {
  getInspection, getProject, getDrawing,
  getDrawingBlob,
  listItemsForInspection, listPhotosForItem,
  listHighlightsForInspectionAndDrawing,
  getUserProfile
} from '../db.js';
import { loadPdf } from './pdf.js';
import { compressImage } from './photos.js';
import { ensureReportFont } from './pdf-fonts.js';
import { renderAnnotatedImage } from './annotate.js';
import { buildCloseOutMailto, buildCloseOutToken, renderQrDataUrl } from './qr.js';

/* --------------------------------------------------------------------------
   Layout / style constants — all millimetres unless noted
   -------------------------------------------------------------------------- */
const PAGE = { w: 210, h: 297 };
const MARGIN = { top: 18, right: 18, bottom: 18, left: 18 };
const CONTENT_W = PAGE.w - MARGIN.left - MARGIN.right;       // 174 mm

const COLOR = {
  blue:      [0, 170, 236],     // BT brand blue
  blueDark:  [0, 136, 192],
  grey:      [74, 68, 66],      // warm grey
  greySoft:  [107, 100, 97],
  ink:       [44, 42, 41],
  line:      [227, 225, 223],
  muted:     [138, 125, 119],
  sand:      [245, 244, 242],
  red:       [212, 68, 38],
  holdPoint: [122, 10, 10],     // deep red for hold-point pins + badges
  yellow:    [255, 199, 39],
  green:     [45, 122, 79]
};

/** Pick the badge/pin colour for an item, considering severity + status. */
function severityColor(item) {
  if (item.status === 'closed') return COLOR.greySoft;
  if (item.severity === 'holdpoint') return COLOR.holdPoint;
  if (item.severity === 'defect')    return COLOR.red;
  return COLOR.blue;
}

// Live font-family string — flipped to 'LibreFranklin' once the TTFs are loaded
// into the jsPDF instance, otherwise stays 'helvetica' so a PDF still generates
// when the font assets can't be fetched (e.g. first offline boot).
let FONT = 'helvetica';

/* --------------------------------------------------------------------------
   Public entry point
   --------------------------------------------------------------------------
   Returns { blob, filename, pageCount, data } so callers (Phase 7) can
   persist the report without recomputing metadata.
   -------------------------------------------------------------------------- */
export async function generateReport(inspectionId, { onProgress } = {}) {
  const notify = (msg) => { try { onProgress?.(msg); } catch {} };

  notify('Gathering data…');
  const data = await gatherReportData(inspectionId);

  notify('Rendering marked plan…');
  const markedPlanDataUrl = data.drawing
    ? await renderMarkedPlan(data.drawing, data.items, data.highlights)
    : null;

  notify('Preparing photos…');
  const photoMap = await prepareItemPhotos(data.items);

  notify('Preparing signature…');
  const signatureDataUrl = data.profile?.signatureBlob
    ? await blobToDataUrl(data.profile.signatureBlob).catch(() => null)
    : null;
  data.signatureDataUrl = signatureDataUrl;

  notify('Building PDF…');
  const { jsPDF } = window.jspdf || {};
  if (!jsPDF) throw new Error('jsPDF not loaded — check your network');
  const pdf = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' });

  // Register Libre Franklin so the PDF matches the UI typography; falls back
  // to Helvetica transparently if the font assets can't be fetched.
  FONT = await ensureReportFont(pdf);

  buildCoverPage(pdf, data);

  // Decide marked-plan orientation based on the drawing's effective aspect.
  // Wide drawings (A1 landscape) get an A4 landscape page so the marked plan
  // uses the full available area; otherwise portrait is fine.
  const planOrientation = pickMarkedPlanOrientation(data.drawing);
  pdf.addPage('a4', planOrientation);
  if (markedPlanDataUrl) {
    buildMarkedPlanPage(pdf, data, markedPlanDataUrl);
  } else {
    buildMarkedPlanPagePlaceholder(pdf, data);
  }

  pdf.addPage('a4', 'portrait');
  buildItemSchedule(pdf, data);
  if (hasAnyPhotos(photoMap)) {
    buildPhotoAppendix(pdf, data, photoMap);
  }

  // Close-out QR page — only when at least one item is open (there's something
  // to rectify). Skipping when everything's closed avoids surprising the
  // builder with a "please send evidence" prompt on a clean report.
  const openItems = (data.items || []).filter((it) => it.status !== 'closed');
  if (openItems.length) {
    buildCloseOutPage(pdf, data);
  }

  // Add footers after all pages exist, so we know the total count.
  addFootersToAllPages(pdf, data);

  notify('Done');
  const blob = pdf.output('blob');
  return {
    blob,
    filename:  buildReportFilename(data.inspection, data.project),
    pageCount: pdf.getNumberOfPages(),
    data
  };
}

/**
 * Filename convention:
 *   {JobNumber}_{InspectionType}_{YYYY-MM-DD}.pdf
 * Characters outside [A-Za-z0-9-] are replaced with '-' to keep it
 * email/SharePoint-safe.
 */
export function buildReportFilename(inspection, project) {
  const job  = (project?.jobNumber || 'UnknownJob').replace(/[^\w\-]+/g, '-');
  const type = (inspection?.inspectionTypeName || 'Inspection').replace(/[^\w\-]+/g, '-');
  const date = (inspection?.date || new Date().toISOString().slice(0, 10))
    .replace(/[^\d\-]/g, '');
  return `${job}_${type}_${date}.pdf`;
}

/* --------------------------------------------------------------------------
   Data gathering
   -------------------------------------------------------------------------- */
async function gatherReportData(inspectionId) {
  const inspection = await getInspection(inspectionId);
  if (!inspection) throw new Error('Inspection not found');

  const [project, drawing, profile] = await Promise.all([
    getProject(inspection.projectId),
    inspection.primaryDrawingId ? getDrawing(inspection.primaryDrawingId) : Promise.resolve(null),
    getUserProfile()
  ]);

  const allItems = await listItemsForInspection(inspection.id);
  // Only include items that actually belong to the primary drawing — other
  // drawings aren't rendered in the marked plan and would confuse the report.
  const items = drawing
    ? allItems.filter((it) => it.drawingId === drawing.id)
    : allItems;

  const highlights = (drawing)
    ? await listHighlightsForInspectionAndDrawing(inspection.id, drawing.id)
    : [];

  const builderEmail = (profile?.builderEmail || '').trim();
  const token        = buildCloseOutToken(inspection);
  const closeOutUrl  = buildCloseOutMailto({ builderEmail, project, inspection, token });

  return { inspection, project, drawing, items, highlights, profile, builderEmail, token, closeOutUrl };
}

/* --------------------------------------------------------------------------
   Cover page
   -------------------------------------------------------------------------- */
function buildCoverPage(pdf, data) {
  const { inspection, project, profile, drawing, items, signatureDataUrl } = data;

  // Blue header bar
  pdf.setFillColor(...COLOR.blue);
  pdf.rect(0, 0, PAGE.w, 10, 'F');

  // Logo (typographic)
  drawTextLogo(pdf, MARGIN.left, 22);

  // Title block
  let y = 42;
  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(20);
  pdf.setTextColor(...COLOR.ink);
  pdf.text('Structural Inspection Report', MARGIN.left, y);

  y += 8;
  pdf.setFont(FONT, 'normal');
  pdf.setFontSize(13);
  pdf.setTextColor(...COLOR.blueDark);
  pdf.text(inspection.inspectionTypeName || '—', MARGIN.left, y);

  // Horizontal rule
  y += 4;
  pdf.setDrawColor(...COLOR.line);
  pdf.setLineWidth(0.3);
  pdf.line(MARGIN.left, y, MARGIN.left + CONTENT_W, y);

  // Project info table — drawing reference now included explicitly
  // (north star §9.1 + §12.2 — drawing revision is load-bearing for PI).
  y += 8;
  const rows = [
    ['Job number',   project?.jobNumber || '—'],
    ['Project',      project?.name || '—'],
    ['Client',       project?.client || '—'],
    ['Site',         project?.siteAddress || '—'],
    ['Inspection',   inspection.inspectionTypeName || '—'],
    ['Date',         formatDateLong(inspection.date)],
    ['Inspector',    inspection.inspectorName || profile?.name || '—'],
    ['Drawing',      formatDrawingReference(drawing)]
  ];
  if (inspection.attendees) rows.push(['Attendees', inspection.attendees]);
  if (inspection.weather)   rows.push(['Weather',   inspection.weather]);
  y = drawInfoTable(pdf, MARGIN.left, y, CONTENT_W, rows);

  // Hold-point banner — §12.5 says make it impossible to miss.
  const holdPointItems = (items || []).filter((it) => it.severity === 'holdpoint' && it.status !== 'closed');
  if (holdPointItems.length > 0) {
    y += 4;
    y = drawHoldPointBanner(pdf, MARGIN.left, y, CONTENT_W, holdPointItems);
  }

  // Scope paragraph — §9.1 / §12.1 defensive language.
  y += 6;
  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(11);
  pdf.setTextColor(...COLOR.ink);
  pdf.text('SCOPE OF INSPECTION', MARGIN.left, y);

  y += 5;
  pdf.setFont(FONT, 'normal');
  pdf.setFontSize(10);
  pdf.setTextColor(...COLOR.grey);
  const scopeText = buildScopeParagraph(data);
  const scopeLines = pdf.splitTextToSize(scopeText, CONTENT_W);
  pdf.text(scopeLines, MARGIN.left, y);
  y += scopeLines.length * 4.2 + 4;

  // General observations
  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(11);
  pdf.setTextColor(...COLOR.ink);
  pdf.text('GENERAL OBSERVATIONS', MARGIN.left, y);

  y += 5;
  pdf.setFont(FONT, 'normal');
  pdf.setFontSize(10);
  pdf.setTextColor(...COLOR.grey);

  const generals = inspection.generalComments || [];
  if (generals.length === 0) {
    pdf.setFont(FONT, 'italic');
    pdf.text('See marked plan and item schedule overleaf.', MARGIN.left, y);
    y += 6;
  } else {
    // Rather than truncating long lists, flow remaining bullets onto a
    // continuation cover page before the marked plan. The engineer can then
    // see every comment that was on file.
    const remaining = [...generals];
    while (remaining.length) {
      const g = remaining.shift();
      const lines = pdf.splitTextToSize(`\u2022 ${g.text}`, CONTENT_W - 6);
      const needed = lines.length * 4.4;
      if (y + needed > PAGE.h - 70) {
        // Draw the rest of the signature block on the current page will still
        // fit below; move overflow items to a dedicated page.
        // (We unshift the current item back so it lands on the new page.)
        remaining.unshift(g);
        pdf.setFont(FONT, 'italic');
        pdf.setTextColor(...COLOR.muted);
        pdf.text('Further observations continued overleaf.', MARGIN.left, y);
        y += 6;
        // Emit the continuation page(s) inline so they land BEFORE the marked
        // plan page. We add pages here; the caller's flow will continue with
        // the marked-plan addPage() after this function returns.
        emitObservationContinuationPages(pdf, remaining);
        break;
      }
      pdf.setFont(FONT, 'normal');
      pdf.setTextColor(...COLOR.grey);
      pdf.text(lines, MARGIN.left, y);
      y += lines.length * 4.4 + 1;
    }
  }

  // Sign-off block — anchored toward bottom.
  const sigBlockH = signatureDataUrl ? 52 : 42;
  const sigY = PAGE.h - MARGIN.bottom - sigBlockH;

  // Optional signature image, above the printed name.
  if (signatureDataUrl) {
    try {
      pdf.addImage(signatureDataUrl, 'PNG', MARGIN.left, sigY - 8, 60, 18, undefined, 'FAST');
    } catch (err) {
      console.warn('Signature image failed to render:', err);
    }
  }

  pdf.setDrawColor(...COLOR.line);
  pdf.line(MARGIN.left, sigY + 12, MARGIN.left + 80, sigY + 12);

  const inspectorName = inspection.inspectorName || profile?.name || 'Inspector';
  const role          = profile?.role || '';
  const rpeq          = profile?.rpeq || '';
  const cpeng         = profile?.cpeng || '';
  const company       = profile?.company || 'Bligh Tanner Pty Ltd';

  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(11);
  pdf.setTextColor(...COLOR.ink);
  pdf.text(inspectorName, MARGIN.left, sigY + 17);

  pdf.setFont(FONT, 'normal');
  pdf.setFontSize(9);
  pdf.setTextColor(...COLOR.muted);
  let yy = sigY + 22;
  if (role)  { pdf.text(role, MARGIN.left, yy); yy += 4; }
  if (rpeq)  { pdf.text(`RPEQ ${rpeq}`, MARGIN.left, yy); yy += 4; }
  if (cpeng) { pdf.text(`CPEng ${cpeng}`, MARGIN.left, yy); yy += 4; }
  pdf.text(company, MARGIN.left, yy); yy += 4;
  pdf.text(`Report prepared ${formatDateLong(new Date().toISOString().slice(0, 10))}`,
    MARGIN.left, yy);
}

/**
 * Emit one or more A4 portrait pages holding overflow general observations.
 * Called from within buildCoverPage when the list doesn't fit on the cover.
 */
function emitObservationContinuationPages(pdf, observations) {
  while (observations.length) {
    pdf.addPage('a4', 'portrait');
    drawSectionHeader(pdf, 'GENERAL OBSERVATIONS (continued)', '');
    let y = 32;
    pdf.setFont(FONT, 'normal');
    pdf.setFontSize(10);
    pdf.setTextColor(...COLOR.grey);

    while (observations.length) {
      const g = observations[0];
      const lines = pdf.splitTextToSize(`\u2022 ${g.text}`, CONTENT_W - 6);
      const needed = lines.length * 4.4 + 1;
      if (y + needed > PAGE.h - MARGIN.bottom - 10) break;
      observations.shift();
      pdf.text(lines, MARGIN.left, y);
      y += needed;
    }
  }
}

/**
 * Builds the canonical BT scope paragraph. Uses the defensive phrasing from
 * §12.1 of the north-star document — "Based on what could be seen…",
 * "…generally installed in accordance with Bligh Tanner's latest drawings…",
 * and an explicit drawing sheet + revision + date reference.
 */
function buildScopeParagraph({ inspection, project, drawing, items, highlights }) {
  const builder = (project?.client || '').trim() || 'the Builder';
  const inspectionLabel = (inspection.inspectionTypeName || 'the nominated works').toLowerCase();
  const sheetRef = drawing?.sheetNumber
    ? `Sheet ${drawing.sheetNumber}${drawing.revision ? ` Rev ${drawing.revision}` : ''}`
    : 'the current issue of the structural drawings';
  const drawingDate = drawing?.uploadedAt
    ? ` dated ${formatDateLong(drawing.uploadedAt.slice(0, 10))}`
    : '';

  const extentDescriptor = highlights && highlights.length
    ? 'the area marked on the marking plan overleaf'
    : 'the accessible portions of the works';

  const itemCount = (items || []).length;
  const defectCount = (items || []).filter((it) => it.severity === 'defect' || it.severity === 'holdpoint').length;

  const commentSuffix = defectCount > 0
    ? `, with the following comments — refer to the item schedule overleaf for ${itemCount} item${itemCount === 1 ? '' : 's'} requiring action.`
    : (itemCount > 0
      ? `. Observations and any comments are set out in the item schedule overleaf.`
      : `, with no defects observed at the time of inspection.`);

  return [
    `Upon request from ${builder}, Bligh Tanner attended site to inspect the ${inspectionLabel}.`,
    ` Based on what could be seen within ${extentDescriptor}, the inspected works have generally`,
    ` been installed in accordance with Bligh Tanner's latest drawings (${sheetRef}${drawingDate})${commentSuffix}`
  ].join('');
}

function formatDrawingReference(drawing) {
  if (!drawing) return '—';
  const parts = [];
  if (drawing.sheetNumber) parts.push(drawing.sheetNumber);
  if (drawing.revision)    parts.push(`Rev ${drawing.revision}`);
  if (drawing.description) parts.push(drawing.description);
  return parts.length ? parts.join(' · ') : (drawing.filename || '—');
}

/**
 * Red hold-point banner — deliberately aggressive so it survives a casual
 * glance at the PDF.
 */
function drawHoldPointBanner(pdf, x, y, w, holdPointItems) {
  const pad = 4;
  const bodyText = `${holdPointItems.length} hold point${holdPointItems.length === 1 ? '' : 's'} raised — works are NOT to proceed on the affected areas until the listed items are rectified and re-inspected.`;
  const bodyLines = pdf.splitTextToSize(bodyText, w - 2 * pad - 24);
  const bannerH = 10 + bodyLines.length * 4.2;

  pdf.setFillColor(122, 10, 10);   // deep red
  pdf.rect(x, y, w, bannerH, 'F');

  pdf.setFillColor(255, 255, 255);
  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(10);
  pdf.setTextColor(255, 255, 255);
  pdf.text('HOLD POINT', x + pad, y + 6.5);

  pdf.setFont(FONT, 'normal');
  pdf.setFontSize(9.5);
  pdf.text(bodyLines, x + pad + 24, y + 6.5);

  return y + bannerH;
}

/* --------------------------------------------------------------------------
   Marked plan page — orientation is chosen per drawing at page creation time
   -------------------------------------------------------------------------- */

/**
 * Given a drawing record, return 'landscape' or 'portrait' for the A4 page
 * that'll hold the marked plan. Wide drawings (A1 landscape) get landscape
 * paper so the plan doesn't shrink awkwardly.
 */
function pickMarkedPlanOrientation(drawing) {
  if (!drawing) return 'portrait';
  let w = drawing.pageWidth  || 595;
  let h = drawing.pageHeight || 842;
  const rot = (drawing.rotation || 0) % 360;
  if (rot === 90 || rot === 270) [w, h] = [h, w];
  return w > h ? 'landscape' : 'portrait';
}

function buildMarkedPlanPage(pdf, { drawing, inspection }, imageDataUrl) {
  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  const contentW = pageW - MARGIN.left - MARGIN.right;

  // Header bar
  drawSectionHeader(pdf, 'MARKED PLAN', drawing
    ? `${drawing.sheetNumber || 'Sheet'}${drawing.revision ? ' (Rev ' + drawing.revision + ')' : ''}`
    : '');

  const topY = 28;
  const bottomY = pageH - MARGIN.bottom - 8;
  const availableH = bottomY - topY - 12;  // 12mm reserved for legend
  const availableW = contentW;

  // Decide image size keeping aspect ratio
  const img = { w: drawing?.pageWidth || 595, h: drawing?.pageHeight || 842 };
  const rot = (drawing?.rotation || 0) % 360;
  if (rot === 90 || rot === 270) [img.w, img.h] = [img.h, img.w];
  const ratio = img.w / img.h;
  let w = availableW;
  let h = w / ratio;
  if (h > availableH) { h = availableH; w = h * ratio; }
  const x = MARGIN.left + (availableW - w) / 2;
  const y = topY;

  pdf.addImage(imageDataUrl, 'JPEG', x, y, w, h, undefined, 'FAST');

  // Border around image
  pdf.setDrawColor(...COLOR.line);
  pdf.setLineWidth(0.2);
  pdf.rect(x, y, w, h);

  // Legend at bottom
  const legendY = y + h + 6;
  drawPlanLegend(pdf, MARGIN.left, legendY, inspection);
}

function buildMarkedPlanPagePlaceholder(pdf, { inspection }) {
  drawSectionHeader(pdf, 'MARKED PLAN', '');
  pdf.setFont(FONT, 'italic');
  pdf.setFontSize(11);
  pdf.setTextColor(...COLOR.muted);
  pdf.text('No primary drawing attached to this inspection.',
    MARGIN.left, 60);
}

function drawPlanLegend(pdf, x, y, inspection) {
  pdf.setDrawColor(...COLOR.line);
  pdf.setLineWidth(0.2);
  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(9);
  pdf.setTextColor(...COLOR.ink);
  pdf.text('LEGEND', x, y);

  pdf.setFont(FONT, 'normal');
  pdf.setTextColor(...COLOR.grey);
  pdf.setFontSize(9);

  // Observation
  drawLegendPin(pdf, x + 22, y - 2, '', COLOR.blue);
  pdf.text('Observation', x + 27, y);

  // Defect
  drawLegendPin(pdf, x + 55, y - 2, '', COLOR.red);
  pdf.text('Defect',      x + 60, y);

  // Hold point
  drawLegendPin(pdf, x + 78, y - 2, '!', COLOR.holdPoint);
  pdf.text('Hold point',  x + 83, y);

  // Closed
  drawLegendPin(pdf, x + 110, y - 2, '', COLOR.greySoft);
  pdf.text('Closed',      x + 115, y);

  // Extent swatch
  pdf.setFillColor(...COLOR.yellow);
  pdf.setDrawColor(155, 115, 0);
  pdf.setLineWidth(0.3);
  pdf.rect(x + 135, y - 3.5, 6, 3.5, 'FD');
  pdf.setTextColor(...COLOR.grey);
  pdf.text('Extent of inspection', x + 143, y);
}

function drawLegendPin(pdf, cx, cy, label, rgb) {
  pdf.setFillColor(...rgb);
  pdf.setDrawColor(255, 255, 255);
  pdf.setLineWidth(0.4);
  pdf.circle(cx, cy - 0.8, 2, 'FD');
  pdf.setFontSize(6);
  pdf.setTextColor(255, 255, 255);
  pdf.setFont(FONT, 'bold');
  pdf.text(label, cx, cy + 0.2, { align: 'center' });
  pdf.setFontSize(9);
  pdf.setFont(FONT, 'normal');
}

/* --------------------------------------------------------------------------
   Item schedule — table, paginates across pages
   -------------------------------------------------------------------------- */
function buildItemSchedule(pdf, { items }) {
  drawSectionHeader(pdf, 'ITEM SCHEDULE', `${items.length} item${items.length === 1 ? '' : 's'}`);

  if (items.length === 0) {
    pdf.setFont(FONT, 'italic');
    pdf.setFontSize(11);
    pdf.setTextColor(...COLOR.muted);
    pdf.text('No items recorded against this inspection.', MARGIN.left, 40);
    return;
  }

  // Column widths (mm) — sum to CONTENT_W (174)
  const cols = [
    { key: 'num',      title: '#',          w: 10,  align: 'center' },
    { key: 'sev',      title: 'Severity',   w: 20,  align: 'left'   },
    { key: 'loc',      title: 'Location',   w: 20,  align: 'left'   },
    { key: 'comment',  title: 'Comment',    w: 70,  align: 'left'   },
    { key: 'clause',   title: 'AS Clause',  w: 36,  align: 'left'   },
    { key: 'status',   title: 'Status',     w: 18,  align: 'left'   }
  ];

  let y = 32;
  drawScheduleHeader(pdf, MARGIN.left, y, cols);
  y += 7;

  for (const it of items) {
    const commentLines = pdf.splitTextToSize(it.comment || '—', cols[3].w - 4);
    const clauseLines  = pdf.splitTextToSize(it.asClause || '', cols[4].w - 4);
    const locLines     = pdf.splitTextToSize(it.gridRef  || '', cols[2].w - 4);
    const rowH = Math.max(9,
      commentLines.length * 4.2 + 3,
      clauseLines.length * 4.2 + 3,
      locLines.length * 4.2 + 3
    );

    // Pagination — leave 15mm bottom for footer
    if (y + rowH > PAGE.h - MARGIN.bottom - 10) {
      pdf.addPage();
      drawSectionHeader(pdf, 'ITEM SCHEDULE (continued)', '');
      y = 32;
      drawScheduleHeader(pdf, MARGIN.left, y, cols);
      y += 7;
    }

    drawScheduleRow(pdf, MARGIN.left, y, rowH, cols, it, commentLines, clauseLines, locLines);
    y += rowH;
  }
}

function drawScheduleHeader(pdf, x, y, cols) {
  pdf.setFillColor(...COLOR.sand);
  pdf.rect(x, y, CONTENT_W, 7, 'F');
  pdf.setDrawColor(...COLOR.line);
  pdf.setLineWidth(0.2);
  pdf.line(x, y + 7, x + CONTENT_W, y + 7);

  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(9);
  pdf.setTextColor(...COLOR.ink);

  let cx = x;
  for (const col of cols) {
    const tx = col.align === 'center' ? cx + col.w / 2 : cx + 2;
    pdf.text(col.title, tx, y + 4.8, { align: col.align === 'center' ? 'center' : 'left' });
    cx += col.w;
  }
}

function drawScheduleRow(pdf, x, y, h, cols, item, commentLines, clauseLines, locLines) {
  // Row background (subtle zebra for readability — not applied here to keep clean)
  pdf.setDrawColor(...COLOR.line);
  pdf.setLineWidth(0.15);
  pdf.line(x, y + h, x + CONTENT_W, y + h);

  let cx = x;

  // Column 1: Number badge
  const badgeColor = severityColor(item);
  pdf.setFillColor(...badgeColor);
  pdf.setDrawColor(255, 255, 255);
  pdf.setLineWidth(0.3);
  pdf.circle(cx + cols[0].w / 2, y + 5, 3.5, 'FD');
  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(9);
  pdf.setTextColor(255, 255, 255);
  pdf.text(String(item.itemNumber), cx + cols[0].w / 2, y + 6.3, { align: 'center' });
  cx += cols[0].w;

  // Column 2: Severity text (hold point shown in red for scannability)
  pdf.setFont(FONT, item.severity === 'holdpoint' ? 'bold' : 'normal');
  pdf.setFontSize(9);
  pdf.setTextColor(...(item.severity === 'holdpoint' ? COLOR.holdPoint : COLOR.ink));
  pdf.text(severityLabel(item.severity), cx + 2, y + 5);
  cx += cols[1].w;

  // Column 3: Location / grid reference
  pdf.setFont(FONT, 'normal');
  pdf.setFontSize(9);
  pdf.setTextColor(...COLOR.ink);
  if (locLines && locLines.length) {
    pdf.text(locLines, cx + 2, y + 4.5);
  } else {
    pdf.setTextColor(...COLOR.muted);
    pdf.text('—', cx + 2, y + 5);
  }
  cx += cols[2].w;

  // Column 4: Comment (wraps)
  pdf.setTextColor(...COLOR.grey);
  pdf.setFontSize(9);
  pdf.text(commentLines, cx + 2, y + 4.5);
  cx += cols[3].w;

  // Column 5: AS Clause (monospaced-ish, small)
  pdf.setFont(FONT, 'normal');
  pdf.setFontSize(8.5);
  pdf.setTextColor(...COLOR.muted);
  if (clauseLines && clauseLines.length) {
    pdf.text(clauseLines, cx + 2, y + 4.5);
  }
  cx += cols[4].w;

  // Column 6: Status pill
  const statusText = capitalize(item.status || 'open');
  const pillColor = item.status === 'closed' ? COLOR.green : [138, 102, 0];
  const pillFill  = item.status === 'closed' ? [215, 236, 222] : [255, 244, 209];
  pdf.setFillColor(...pillFill);
  pdf.setDrawColor(...pillColor);
  pdf.setLineWidth(0.2);
  pdf.roundedRect(cx + 2, y + 1.5, 14, 5, 1.2, 1.2, 'FD');
  pdf.setTextColor(...pillColor);
  pdf.setFontSize(7.5);
  pdf.setFont(FONT, 'bold');
  pdf.text(statusText.toUpperCase(), cx + 9, y + 5.1, { align: 'center' });
}

/** Human-facing label for a severity string — specifically upgrades 'holdpoint' to 'HOLD POINT'. */
function severityLabel(sev) {
  if (sev === 'holdpoint') return 'HOLD POINT';
  return capitalize(sev || 'observation');
}

/* --------------------------------------------------------------------------
   Photo appendix — 2 per page
   -------------------------------------------------------------------------- */
function buildPhotoAppendix(pdf, { items }, photoMap) {
  pdf.addPage();
  drawSectionHeader(pdf, 'PHOTO APPENDIX', '');

  // Flatten into a queue of { item, photo, photoDataUrl }
  const queue = [];
  for (const it of items) {
    const list = photoMap.get(it.id) || [];
    for (const entry of list) {
      queue.push({ item: it, photo: entry.photo, dataUrl: entry.dataUrl });
    }
  }
  if (queue.length === 0) return;

  const slotH = 118;  // mm — two slots per page below the 28mm header
  const topY  = 32;
  let slotIndex = 0;

  for (const q of queue) {
    if (slotIndex === 2) {
      pdf.addPage();
      drawSectionHeader(pdf, 'PHOTO APPENDIX (continued)', '');
      slotIndex = 0;
    }
    const y = topY + slotIndex * (slotH + 6);
    drawPhotoSlot(pdf, MARGIN.left, y, CONTENT_W, slotH, q);
    slotIndex += 1;
  }
}

function drawPhotoSlot(pdf, x, y, w, h, { item, photo, dataUrl }) {
  // Caption band (top)
  pdf.setFillColor(...COLOR.sand);
  pdf.rect(x, y, w, 8, 'F');
  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(9.5);
  pdf.setTextColor(...COLOR.ink);
  pdf.text(`Item ${item.itemNumber}`, x + 3, y + 5.4);

  pdf.setFont(FONT, 'normal');
  pdf.setFontSize(9);
  pdf.setTextColor(...COLOR.grey);
  const cap = (photo.caption || '').trim() ||
    ((item.comment || '').slice(0, 80) + (((item.comment || '').length > 80) ? '…' : ''));
  if (cap) pdf.text(cap, x + 22, y + 5.4);

  // Image body
  const imgTop = y + 10;
  const imgH   = h - 12;
  // Fit the image centred while preserving aspect ratio
  const natW = photo.width || 4;
  const natH = photo.height || 3;
  const ratio = natW / natH;
  let iw = w, ih = iw / ratio;
  if (ih > imgH) { ih = imgH; iw = ih * ratio; }
  const ix = x + (w - iw) / 2;
  const iy = imgTop + (imgH - ih) / 2;

  pdf.setDrawColor(...COLOR.line);
  pdf.setLineWidth(0.2);
  pdf.rect(x, imgTop, w, imgH);
  try {
    pdf.addImage(dataUrl, 'JPEG', ix, iy, iw, ih, undefined, 'FAST');
  } catch (err) {
    // Some blobs may fail to decode — draw a placeholder instead of crashing
    pdf.setFillColor(...COLOR.sand);
    pdf.rect(ix, iy, iw, ih, 'F');
    pdf.setFont(FONT, 'italic');
    pdf.setFontSize(9);
    pdf.setTextColor(...COLOR.muted);
    pdf.text('(photo failed to render)', x + w / 2, iy + ih / 2, { align: 'center' });
  }
}

/* --------------------------------------------------------------------------
   Marked plan — rasterise source drawing + pins + highlights to a JPEG
   -------------------------------------------------------------------------- */
async function renderMarkedPlan(drawing, items, highlights) {
  // Reasonable print quality: 2x scale gives around 200 DPI at A4 print size.
  const RENDER_SCALE = 2.0;

  const blob = await getDrawingBlob(drawing);
  if (!blob) throw new Error('Could not load the drawing blob');
  const doc = await loadPdf(blob);
  try {
    const page = await doc.getPage(drawing.pageNumber || 1);
    const viewport = page.getViewport({ scale: RENDER_SCALE, rotation: drawing.rotation || 0 });

    const canvas = document.createElement('canvas');
    canvas.width  = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    const ctx = canvas.getContext('2d', { alpha: false });

    // Render the PDF page
    await page.render({ canvasContext: ctx, viewport }).promise;

    // Draw highlights (under pins) — yellow translucent + brown stroke
    ctx.save();
    for (const h of highlights) {
      const [x0, y0] = viewport.convertToViewportPoint(h.pdfX, h.pdfY);
      const [x1, y1] = viewport.convertToViewportPoint(h.pdfX + h.pdfW, h.pdfY + h.pdfH);
      const rx = Math.min(x0, x1);
      const ry = Math.min(y0, y1);
      const rw = Math.abs(x1 - x0);
      const rh = Math.abs(y1 - y0);
      ctx.fillStyle = 'rgba(255, 199, 39, 0.28)';
      ctx.strokeStyle = '#c79a00';
      ctx.lineWidth = 2;
      ctx.fillRect(rx, ry, rw, rh);
      ctx.strokeRect(rx, ry, rw, rh);
    }
    ctx.restore();

    // Draw pins
    const PIN_R = 14;  // canvas px at RENDER_SCALE 2.0 — readable when printed
    ctx.save();
    ctx.font = `bold ${Math.round(PIN_R)}px Helvetica, Arial, sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    for (const it of items) {
      if (it.pdfX == null || it.pdfY == null) continue;
      const [cx, cy] = viewport.convertToViewportPoint(it.pdfX, it.pdfY);
      const colorRgb = severityColor(it);
      ctx.fillStyle = `rgb(${colorRgb.join(',')})`;
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 3;
      ctx.beginPath();
      ctx.arc(cx, cy, PIN_R, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();
      ctx.fillStyle = '#ffffff';
      ctx.fillText(String(it.itemNumber), cx, cy + 1);
    }
    ctx.restore();

    return canvas.toDataURL('image/jpeg', 0.88);
  } finally {
    doc.destroy();
  }
}

/* --------------------------------------------------------------------------
   Photo preparation — re-compress to ~1200px long edge for embedding
   -------------------------------------------------------------------------- */
async function prepareItemPhotos(items) {
  const out = new Map();
  for (const it of items) {
    const photos = await listPhotosForItem(it.id);
    const entries = [];
    for (const p of photos) {
      try {
        let dataUrl, width, height;
        if (Array.isArray(p.annotations) && p.annotations.length) {
          // Composite annotations onto the photo at render time.
          const r = await renderAnnotatedImage(p.blob, p.annotations, { maxLongEdge: 1200 });
          dataUrl = r.dataUrl; width = r.width; height = r.height;
        } else {
          const { blob, width: w, height: h } = await compressImage(p.blob, {
            maxLongEdge: 1200, quality: 0.8
          });
          dataUrl = await blobToDataUrl(blob);
          width = w; height = h;
        }
        entries.push({ photo: { ...p, width, height }, dataUrl });
      } catch (err) {
        console.warn('Photo prep failed for photo', p.id, err);
      }
    }
    if (entries.length) out.set(it.id, entries);
  }
  return out;
}

function hasAnyPhotos(photoMap) {
  for (const v of photoMap.values()) if (v.length) return true;
  return false;
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error('Could not read photo blob'));
    reader.readAsDataURL(blob);
  });
}

/* --------------------------------------------------------------------------
   Shared drawing helpers
   -------------------------------------------------------------------------- */
function drawSectionHeader(pdf, title, right) {
  const pageW = pdf.internal.pageSize.getWidth();

  // Top blue bar
  pdf.setFillColor(...COLOR.blue);
  pdf.rect(0, 0, pageW, 6, 'F');

  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(13);
  pdf.setTextColor(...COLOR.ink);
  pdf.text(title, MARGIN.left, 18);

  if (right) {
    pdf.setFont(FONT, 'normal');
    pdf.setFontSize(10);
    pdf.setTextColor(...COLOR.muted);
    pdf.text(right, pageW - MARGIN.right, 18, { align: 'right' });
  }

  pdf.setDrawColor(...COLOR.line);
  pdf.setLineWidth(0.3);
  pdf.line(MARGIN.left, 22, pageW - MARGIN.right, 22);
}

function drawInfoTable(pdf, x, y, w, rows) {
  const labelW = 38;
  const lineH  = 5.8;
  pdf.setDrawColor(...COLOR.line);
  pdf.setLineWidth(0.15);

  for (const [label, value] of rows) {
    pdf.setFont(FONT, 'bold');
    pdf.setFontSize(9);
    pdf.setTextColor(...COLOR.muted);
    pdf.text(label.toUpperCase(), x, y);

    pdf.setFont(FONT, 'normal');
    pdf.setFontSize(10.5);
    pdf.setTextColor(...COLOR.ink);
    const valueLines = pdf.splitTextToSize(value || '—', w - labelW);
    pdf.text(valueLines, x + labelW, y);
    const used = Math.max(lineH, valueLines.length * 4.8);
    y += used;
  }
  return y;
}

function drawTextLogo(pdf, x, y) {
  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(22);
  pdf.setTextColor(...COLOR.blue);
  pdf.text('BLIGH', x, y);
  const w = pdf.getTextWidth('BLIGH');
  pdf.setTextColor(...COLOR.grey);
  pdf.text('TANNER', x + w + 2, y);
}

function addFootersToAllPages(pdf, { project, inspection }) {
  const total = pdf.getNumberOfPages();
  for (let i = 1; i <= total; i += 1) {
    pdf.setPage(i);
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();

    pdf.setDrawColor(...COLOR.line);
    pdf.setLineWidth(0.2);
    pdf.line(MARGIN.left, pageH - MARGIN.bottom + 2,
             pageW - MARGIN.right, pageH - MARGIN.bottom + 2);

    pdf.setFont(FONT, 'normal');
    pdf.setFontSize(8);
    pdf.setTextColor(...COLOR.muted);
    const left = `${project?.jobNumber || ''} · ${inspection.inspectionTypeName || ''}`;
    pdf.text(left.trim().replace(/^·\s*/, ''), MARGIN.left, pageH - MARGIN.bottom + 7);

    pdf.text(`Page ${i} of ${total}`, pageW / 2, pageH - MARGIN.bottom + 7, { align: 'center' });
    pdf.text('Bligh Tanner', pageW - MARGIN.right, pageH - MARGIN.bottom + 7, { align: 'right' });
  }
}

/* --------------------------------------------------------------------------
   Close-out page — QR + mailto for builder rectification evidence
   -------------------------------------------------------------------------- */

function buildCloseOutPage(pdf, data) {
  pdf.addPage('a4', 'portrait');
  drawSectionHeader(pdf, 'RECTIFICATION CLOSE-OUT', '');

  let y = 34;

  pdf.setFont(FONT, 'normal');
  pdf.setFontSize(10.5);
  pdf.setTextColor(...COLOR.grey);
  const intro = 'Once the items listed in this report have been rectified, please email the engineer photographic evidence. You can scan the QR code below to open a pre-filled email on your phone, or reply directly to the email this report was sent from.';
  const introLines = pdf.splitTextToSize(intro, CONTENT_W);
  pdf.text(introLines, MARGIN.left, y);
  y += introLines.length * 4.6 + 4;

  // QR block — square, centred on the left; instructions to the right.
  const qrSize = 56; // mm
  const qrX = MARGIN.left;
  const qrY = y + 2;
  try {
    const qrUrl = renderQrDataUrl(data.closeOutUrl || 'mailto:', { scale: 8, margin: 2 });
    pdf.addImage(qrUrl, 'PNG', qrX, qrY, qrSize, qrSize, undefined, 'FAST');
  } catch (err) {
    console.warn('QR render failed:', err);
    pdf.setDrawColor(...COLOR.line);
    pdf.rect(qrX, qrY, qrSize, qrSize);
    pdf.setFont(FONT, 'italic');
    pdf.setFontSize(10);
    pdf.setTextColor(...COLOR.muted);
    pdf.text('(QR code unavailable)', qrX + qrSize / 2, qrY + qrSize / 2, { align: 'center' });
  }

  // Info column right of the QR
  const infoX = qrX + qrSize + 8;
  const infoW = CONTENT_W - qrSize - 8;
  let infoY = qrY + 4;

  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(10.5);
  pdf.setTextColor(...COLOR.ink);
  pdf.text('Inspection token', infoX, infoY); infoY += 5;
  pdf.setFont(FONT, 'normal');
  pdf.setFontSize(11);
  pdf.setTextColor(...COLOR.blueDark);
  pdf.text(data.token || '', infoX, infoY); infoY += 7;

  if (data.builderEmail) {
    pdf.setFont(FONT, 'bold');
    pdf.setFontSize(10.5);
    pdf.setTextColor(...COLOR.ink);
    pdf.text('Email to', infoX, infoY); infoY += 5;
    pdf.setFont(FONT, 'normal');
    pdf.setFontSize(10.5);
    pdf.setTextColor(...COLOR.grey);
    pdf.text(data.builderEmail, infoX, infoY); infoY += 7;
  }

  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(10.5);
  pdf.setTextColor(...COLOR.ink);
  pdf.text('What to include', infoX, infoY); infoY += 5;
  pdf.setFont(FONT, 'normal');
  pdf.setFontSize(9.5);
  pdf.setTextColor(...COLOR.grey);
  const howto = [
    '\u2022 A clear photo per item — reference the item number.',
    '\u2022 Any measurement or check-sheet evidence required.',
    '\u2022 The name and date of the tradesperson who rectified.'
  ];
  for (const line of howto) {
    const wrapped = pdf.splitTextToSize(line, infoW);
    pdf.text(wrapped, infoX, infoY);
    infoY += wrapped.length * 4.3;
  }

  // Open items list — what the builder needs to close
  y = Math.max(qrY + qrSize + 8, infoY + 4);
  pdf.setFont(FONT, 'bold');
  pdf.setFontSize(11);
  pdf.setTextColor(...COLOR.ink);
  pdf.text('ITEMS AWAITING RECTIFICATION', MARGIN.left, y);
  y += 5;
  pdf.setDrawColor(...COLOR.line);
  pdf.line(MARGIN.left, y, MARGIN.left + CONTENT_W, y);
  y += 4;

  const open = (data.items || []).filter((it) => it.status !== 'closed');
  if (open.length === 0) {
    pdf.setFont(FONT, 'italic');
    pdf.setFontSize(10);
    pdf.setTextColor(...COLOR.muted);
    pdf.text('No open items at time of issue.', MARGIN.left, y);
    return;
  }
  for (const it of open) {
    if (y > PAGE.h - MARGIN.bottom - 20) break; // leave footer room
    const badgeColor = severityColor(it);
    pdf.setFillColor(...badgeColor);
    pdf.setDrawColor(255, 255, 255);
    pdf.setLineWidth(0.3);
    pdf.circle(MARGIN.left + 3, y + 2, 3, 'FD');
    pdf.setFont(FONT, 'bold');
    pdf.setFontSize(8);
    pdf.setTextColor(255, 255, 255);
    pdf.text(String(it.itemNumber), MARGIN.left + 3, y + 3, { align: 'center' });

    pdf.setFont(FONT, 'normal');
    pdf.setFontSize(9.5);
    pdf.setTextColor(...COLOR.ink);
    const text = [
      it.gridRef ? `[${it.gridRef}] ` : '',
      (it.comment || '').trim()
    ].join('');
    const lines = pdf.splitTextToSize(text || '—', CONTENT_W - 12);
    pdf.text(lines, MARGIN.left + 10, y + 3);
    y += Math.max(7, lines.length * 4.4 + 2);
  }
}

/* --------------------------------------------------------------------------
   Small utilities
   -------------------------------------------------------------------------- */
function capitalize(s) {
  const str = String(s || '');
  return str ? str.charAt(0).toUpperCase() + str.slice(1) : '';
}

function formatDateLong(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.valueOf())) return String(iso);
  return d.toLocaleDateString('en-AU', { day: '2-digit', month: 'long', year: 'numeric' });
}
