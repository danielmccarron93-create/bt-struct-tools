/**
 * Drawing viewer — full-screen PDF renderer with pan / zoom / rotate /
 * calibrate / markup.
 *
 * Architecture:
 *   - A single <canvas> shows the PDF page at a chosen scale.
 *   - Pan/zoom is implemented via CSS transform on a wrapper element so
 *     touch gestures feel native and we don't re-rasterise on every frame.
 *   - An absolutely-positioned <svg> overlay above the canvas receives
 *     pointer events for calibration and (Phase 4+) pin drops + extent
 *     highlights. Overlay coords match canvas CSS pixels at the current
 *     render scale; PDF-space coords come via viewport.convertToPdfPoint.
 *
 * Two entry points:
 *   render(root, params)        — view/calibrate only   (#/drawing/:id)
 *   renderMarkup(root, params)  — adds pin + extent     (#/markup/:inspId)
 *
 * Modes (state.activeMode):
 *   'view'        default — pan + zoom + open existing pins
 *   'calibrate'   click two points → enter distance → save mmPerPoint
 *   'pin'         tap drops a numbered pin, opens the item panel
 *   'extent'      drag draws a yellow highlight rectangle; tap removes one
 */

import {
  getDrawing, updateDrawing, setDrawingCalibration, setDrawingGridCalibration, getProject,
  getDrawingBlob,
  getInspection,
  listItemsForInspection, createItem,
  listHighlightsForInspectionAndDrawing, addHighlight, deleteHighlight
} from '../db.js';
import { loadPdf, renderPageToCanvas } from '../lib/pdf.js';
import { openModal, confirmDialog } from '../components/modal.js';
import { toast } from '../components/toast.js';
import { go } from '../router.js';
import { openItemPanel } from '../components/item-panel.js';
import { computeGridRef, validateCalibrationInput } from '../lib/grid.js';

/* Re-raster scale thresholds — below the low bound, pixelation starts to
 * show; above the high bound we're wasting GPU. Committed zoom re-renders
 * the PDF at a fresh scale to stay crisp. */
const MIN_RENDER_SCALE = 0.5;
const MAX_RENDER_SCALE = 4.0;

const SVG_NS = 'http://www.w3.org/2000/svg';

let state = null; // viewer state — see createState()
let resizeHandler = null; // tracked so we can detach on re-mount

/* --------------------------------------------------------------------------
   Entry points
   -------------------------------------------------------------------------- */

export async function render(root, params) {
  return renderInternal(root, params, { markup: false });
}

export async function renderMarkup(root, params) {
  const inspectionId = Number(params[0]);
  if (!inspectionId) { go('inspections'); return; }

  const inspection = await getInspection(inspectionId);
  if (!inspection) {
    root.innerHTML = `
      <div class="card"><h2>Inspection not found</h2>
        <p class="muted"><a href="#/inspections">Back to inspections</a></p></div>`;
    return;
  }
  if (!inspection.primaryDrawingId) {
    root.innerHTML = `
      <div class="card"><h2>No primary drawing</h2>
        <p class="muted">This inspection has no drawing attached.
          <a href="#/inspection/${inspection.id}">Back to inspection</a>.
        </p></div>`;
    return;
  }
  return renderInternal(root, [inspection.primaryDrawingId], { markup: true, inspection });
}

async function renderInternal(root, params, opts) {
  const id = Number(params[0]);
  if (!id) { go('projects'); return; }

  const drawing = await getDrawing(id);
  if (!drawing) {
    root.innerHTML = `
      <div class="card"><h2>Drawing not found</h2>
        <p class="muted"><a href="#/projects">Back to projects</a></p></div>`;
    return;
  }
  const project = await getProject(drawing.projectId);

  const inspection = opts.inspection || null;
  const isMarkup = !!opts.markup;

  // Load markup data (per-inspection, per-drawing)
  let items = [];
  let highlights = [];
  if (isMarkup && inspection) {
    [items, highlights] = await Promise.all([
      listItemsForInspection(inspection.id).then((all) =>
        all.filter((it) => it.drawingId === drawing.id)),
      listHighlightsForInspectionAndDrawing(inspection.id, drawing.id)
    ]);
  }

  // Build the toolbar — the action buttons depend on mode.
  const backHref = isMarkup && inspection
    ? `#/inspection/${inspection.id}`
    : `#/project/${drawing.projectId}`;
  const backLabel = isMarkup ? 'Back to inspection' : 'Back to project';

  const toolbarActionsHtml = isMarkup
    ? `
      <button class="icon-btn" id="btn-fit" aria-label="Fit to screen" title="Fit to screen">
        <svg viewBox="0 0 24 24"><path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/></svg>
      </button>
      <button class="btn btn--sm btn--secondary" id="btn-pin" aria-pressed="false" title="Drop numbered pins">
        <span aria-hidden="true">📍</span> Pin
      </button>
      <button class="btn btn--sm btn--secondary" id="btn-extent" aria-pressed="false" title="Mark extent of inspection">
        <span aria-hidden="true">▭</span> Extent
      </button>
      <button class="icon-btn" id="btn-grid" aria-label="Set grid" title="Set grid reference">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3h18v18H3zM3 9h18M3 15h18M9 3v18M15 3v18"/></svg>
      </button>
      <button class="icon-btn" id="btn-calibrate-small" aria-label="Calibrate distance" title="Calibrate scale">
        <svg viewBox="0 0 24 24"><path d="M3 3l18 18M8 5l11 11M5 8l11 11"/></svg>
      </button>
      <button class="icon-btn" id="btn-rotate" aria-label="Rotate 90°" title="Rotate 90°">
        <svg viewBox="0 0 24 24"><path d="M21 2v6h-6M3 12a9 9 0 0 1 15-6.7L21 8"/></svg>
      </button>
    `
    : `
      <button class="icon-btn" id="btn-fit" aria-label="Fit to screen" title="Fit to screen">
        <svg viewBox="0 0 24 24"><path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/></svg>
      </button>
      <button class="icon-btn" id="btn-rotate" aria-label="Rotate 90°" title="Rotate 90°">
        <svg viewBox="0 0 24 24"><path d="M21 2v6h-6M3 12a9 9 0 0 1 15-6.7L21 8"/></svg>
      </button>
      <button class="btn btn--sm btn--primary" id="btn-calibrate">Distance</button>
      <button class="btn btn--sm btn--secondary" id="btn-grid">Grid</button>
    `;

  const titleLine = isMarkup && inspection
    ? `${escapeHtml(drawing.sheetNumber || '(no sheet)')}${drawing.revision ? ` · Rev ${escapeHtml(drawing.revision)}` : ''} · ${escapeHtml(inspection.inspectionTypeName || '')}`
    : `${escapeHtml(drawing.sheetNumber || '(no sheet)')}${drawing.revision ? ` · Rev ${escapeHtml(drawing.revision)}` : ''}`;

  root.innerHTML = `
    <div class="drawing-viewer ${isMarkup ? 'drawing-viewer--markup' : ''}" id="dv-root">
      <header class="drawing-viewer__toolbar">
        <a class="drawing-viewer__back" href="${backHref}" aria-label="${escapeAttr(backLabel)}">
          <svg viewBox="0 0 24 24" width="22" height="22"><path d="M15 18l-6-6 6-6"/></svg>
        </a>
        <div class="drawing-viewer__title">
          <div class="drawing-viewer__sheet">${titleLine}</div>
          <div class="drawing-viewer__project">${escapeHtml(project?.jobNumber || '')}${project ? ' · ' + escapeHtml(project.name) : ''}</div>
        </div>
        <div class="drawing-viewer__spacer"></div>
        <div class="drawing-viewer__actions">${toolbarActionsHtml}</div>
      </header>

      <div class="drawing-viewer__stage" id="dv-stage" tabindex="0">
        <div class="drawing-viewer__panner" id="dv-panner">
          <canvas id="dv-canvas" class="drawing-viewer__canvas"></canvas>
          <svg id="dv-overlay" class="drawing-viewer__overlay" xmlns="http://www.w3.org/2000/svg"></svg>
        </div>
      </div>

      <div class="drawing-viewer__statusbar" id="dv-statusbar">
        <div class="drawing-viewer__scale-info" id="dv-scale-info"></div>
        <div class="drawing-viewer__zoom-info" id="dv-zoom-info">100%</div>
      </div>

      <div class="drawing-viewer__mode-bar is-hidden" id="dv-mode-bar" role="status" aria-live="polite">
        <span id="dv-mode-text"></span>
        <button class="btn btn--sm btn--secondary" id="dv-mode-cancel">Done</button>
      </div>
    </div>
  `;

  try {
    state = await createState(drawing, { isMarkup, inspection, items, highlights });
    attachGestures(state);
    updateScaleInfo(state);
    updateZoomInfo(state);
    fitToScreen(state);
    if (isMarkup) {
      renderMarkupOverlay(state);
      if (!drawing.calibration) {
        toast('Tip: calibrate this drawing before relying on measurements', { kind: 'info' });
      }
    }
  } catch (err) {
    console.error('Viewer init failed:', err);
    root.innerHTML = `
      <div class="card"><h2>Couldn't open drawing</h2>
        <p class="muted">${escapeHtml(err.message || 'Unknown error')}</p>
        <p><a href="${backHref}">Back</a></p></div>`;
    return;
  }

  // Common controls
  root.querySelector('#btn-fit').addEventListener('click', () => fitToScreen(state));
  root.querySelector('#btn-rotate').addEventListener('click', () => rotateDrawing(state));
  root.querySelector('#dv-mode-cancel').addEventListener('click', () => exitAnyMode(state));

  // View-mode only controls
  const btnCalibrate = root.querySelector('#btn-calibrate');
  if (btnCalibrate) btnCalibrate.addEventListener('click', () => enterCalibrationMode(state));

  // Markup-mode controls
  const btnCalSmall = root.querySelector('#btn-calibrate-small');
  if (btnCalSmall) btnCalSmall.addEventListener('click', () => enterCalibrationMode(state));
  const btnPin = root.querySelector('#btn-pin');
  if (btnPin) btnPin.addEventListener('click', () => togglePinMode(state));
  const btnExtent = root.querySelector('#btn-extent');
  if (btnExtent) btnExtent.addEventListener('click', () => toggleExtentMode(state));
  const btnGrid = root.querySelector('#btn-grid');
  if (btnGrid) btnGrid.addEventListener('click', () => enterGridCalibrationMode(state));
}

/* --------------------------------------------------------------------------
   State
   -------------------------------------------------------------------------- */

async function createState(drawing, { isMarkup, inspection, items, highlights }) {
  const blob = await getDrawingBlob(drawing);
  if (!blob) throw new Error('Could not load the drawing PDF');
  const doc = await loadPdf(blob);

  const renderScale = 1.5;

  const canvas  = document.getElementById('dv-canvas');
  const overlay = document.getElementById('dv-overlay');
  const panner  = document.getElementById('dv-panner');
  const stage   = document.getElementById('dv-stage');

  const pageNumber = drawing.pageNumber || 1;

  // Auto-orient on first open: if the user hasn't manually rotated this drawing
  // (drawing.rotation === 0) AND the page is landscape but the viewport is
  // portrait, render it 90° rotated so the long edge of the drawing lines up
  // with the long edge of the screen. This is what every engineer wants on a
  // portrait phone: drawing fills the screen, no squinting at a tiny landscape
  // strip in the middle.
  //
  // The user-controlled rotate button keeps working — it stores its result
  // explicitly via updateDrawing, so a manual choice always wins on next open.
  const initialRotation = await pickInitialRotation(doc, pageNumber, drawing, stage);
  const viewport = await renderPageToCanvas(doc, pageNumber, canvas, renderScale, initialRotation);

  overlay.setAttribute('viewBox', `0 0 ${viewport.width} ${viewport.height}`);
  overlay.style.width  = `${viewport.width}px`;
  overlay.style.height = `${viewport.height}px`;

  return {
    drawing,
    pageNumber,
    doc,
    viewport,
    renderScale,
    canvas,
    overlay,
    panner,
    stage,
    // Interaction transform (applied as CSS on the panner)
    tx: 0, ty: 0, scale: 1,
    // Pan/zoom internal state
    pointers: new Map(),
    lastPinchDist: 0,
    lastPinchMid: null,
    // Distance-calibration flow state
    calibration: {
      active: false,
      points: []
    },
    // Grid-calibration flow state (separate — two labelled grid intersections)
    gridCalibration: {
      active: false,
      points: []
    },
    // Cached for convenience
    rotation:        initialRotation,
    autoRotated:     initialRotation !== (drawing.rotation || 0),  // tag so the rotate button knows to start persisting

    // Markup context (Phase 4)
    isMarkup: !!isMarkup,
    inspection: inspection || null,
    items: items || [],
    highlights: highlights || [],
    activeMode: 'view',    // 'view' | 'calibrate' | 'pin' | 'extent'
    extent: {
      dragStartCanvas: null,
      tempEl: null
    }
  };
}

/**
 * Decide the rotation to render at on first open.
 *
 * Rules (in priority order):
 *   1. If the user has manually rotated this drawing before
 *      (drawing.rotation > 0), respect their choice.
 *   2. Otherwise, if the page is landscape and the viewport is portrait,
 *      auto-rotate 90° so the drawing fills the screen long edge.
 *   3. Otherwise, no rotation.
 *
 * Caveat: this only runs at viewer mount. If the user re-orients the device
 * after opening, the existing rotate button is one tap away.
 */
async function pickInitialRotation(doc, pageNumber, drawing, stage) {
  if (drawing.rotation && drawing.rotation > 0) return drawing.rotation;

  // Get the page's natural rendered dimensions (rotation = 0).
  const page = await doc.getPage(pageNumber);
  const natural = page.getViewport({ scale: 1, rotation: 0 });
  const pageIsLandscape = natural.width > natural.height;

  // Stage may not have laid out yet on the very first frame — getBoundingClientRect
  // returns 0×0. Fall back to window inner size so we still make a sensible call.
  const stageRect = stage.getBoundingClientRect();
  const vw = stageRect.width  > 10 ? stageRect.width  : window.innerWidth;
  const vh = stageRect.height > 10 ? stageRect.height : window.innerHeight;
  const viewportIsPortrait = vh > vw;

  return (pageIsLandscape && viewportIsPortrait) ? 90 : 0;
}

/* --------------------------------------------------------------------------
   View transforms — apply pan/zoom as CSS
   -------------------------------------------------------------------------- */

function applyTransform(s) {
  s.panner.style.transform = `translate(${s.tx}px, ${s.ty}px) scale(${s.scale})`;
  updateZoomInfo(s);
  updatePinSizes(s);
}

function updateZoomInfo(s) {
  const zoomPct = Math.round(s.scale * s.renderScale * 100);
  const el = document.getElementById('dv-zoom-info');
  if (el) el.textContent = `${zoomPct}%`;
}

function updateScaleInfo(s) {
  const el = document.getElementById('dv-scale-info');
  if (!el) return;
  const c = s.drawing.calibration;
  if (c && c.mmPerPoint) {
    const ratio = c.mmPerPoint / 0.3527777778;
    const rounded = niceRatio(ratio);
    el.textContent = `Scale 1:${rounded}`;
  } else {
    el.textContent = `Not calibrated`;
  }
}

function niceRatio(r) {
  const common = [1, 2, 5, 10, 20, 25, 50, 75, 100, 150, 200, 250, 500, 1000, 2000, 2500, 5000];
  let best = common[0], bestErr = Infinity;
  for (const c of common) {
    const err = Math.abs(r - c) / c;
    if (err < bestErr) { bestErr = err; best = c; }
  }
  return bestErr < 0.06 ? best : Math.round(r);
}

/* --------------------------------------------------------------------------
   Fit / rotate
   -------------------------------------------------------------------------- */

function fitToScreen(s) {
  const stage = s.stage.getBoundingClientRect();
  // If layout hasn't settled yet (stage has zero size), retry on the next frame.
  // Without this, the first paint on a landscape A1 drawing comes up at 0.05
  // scale with the drawing parked off-screen.
  if (stage.width < 10 || stage.height < 10) {
    requestAnimationFrame(() => fitToScreen(s));
    return;
  }
  const w = s.viewport.width;
  const h = s.viewport.height;
  const scale = Math.min(stage.width / w, stage.height / h) * 0.96;
  s.scale = Math.max(0.05, scale);
  s.tx = (stage.width  - w * s.scale) / 2;
  s.ty = (stage.height - h * s.scale) / 2;
  applyTransform(s);
}

async function rotateDrawing(s) {
  s.rotation = (s.rotation + 90) % 360;
  await updateDrawing(s.drawing.id, { rotation: s.rotation });
  s.viewport = await renderPageToCanvas(s.doc, s.pageNumber, s.canvas, s.renderScale, s.rotation);
  s.overlay.setAttribute('viewBox', `0 0 ${s.viewport.width} ${s.viewport.height}`);
  s.overlay.style.width  = `${s.viewport.width}px`;
  s.overlay.style.height = `${s.viewport.height}px`;
  if (s.isMarkup) renderMarkupOverlay(s);
  fitToScreen(s);
}

/* --------------------------------------------------------------------------
   Gestures — pointer events for pan / pinch-zoom, wheel for desktop
   -------------------------------------------------------------------------- */

function attachGestures(s) {
  const stage = s.stage;

  stage.addEventListener('pointerdown', (e) => {
    stage.setPointerCapture(e.pointerId);
    s.pointers.set(e.pointerId, { x: e.clientX, y: e.clientY, startTx: s.tx, startTy: s.ty });

    if (s.pointers.size === 2) {
      const [a, b] = Array.from(s.pointers.values());
      s.lastPinchDist = Math.hypot(b.x - a.x, b.y - a.y);
      s.lastPinchMid  = { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 };
    }
  });

  stage.addEventListener('pointermove', (e) => {
    if (!s.pointers.has(e.pointerId)) return;
    const p = s.pointers.get(e.pointerId);
    p.x = e.clientX; p.y = e.clientY;

    // Pinch-zoom when two fingers are down (works in every mode except calibrate)
    if (s.pointers.size === 2 && !s.calibration.active) {
      const [a, b] = Array.from(s.pointers.values());
      const dist = Math.hypot(b.x - a.x, b.y - a.y);
      const mid  = { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 };
      if (s.lastPinchDist > 0) {
        const factor = dist / s.lastPinchDist;
        zoomAt(s, mid.x, mid.y, factor);
      }
      s.lastPinchDist = dist;
      s.lastPinchMid  = mid;
    }
  });

  stage.addEventListener('pointerup',    (e) => s.pointers.delete(e.pointerId));
  stage.addEventListener('pointercancel',(e) => s.pointers.delete(e.pointerId));
  stage.addEventListener('pointerleave', (e) => s.pointers.delete(e.pointerId));

  // --- Single-finger / mouse pan. Suspended in calibrate + grid-calibrate + extent modes. ---
  let panStart = null;
  stage.addEventListener('pointerdown', (e) => {
    if (s.calibration.active || s.gridCalibration.active) return;
    if (s.activeMode === 'extent') return;
    panStart = { x: e.clientX, y: e.clientY, tx: s.tx, ty: s.ty };
  });
  stage.addEventListener('pointermove', (e) => {
    if (!panStart || s.pointers.size !== 1) return;
    if (s.calibration.active || s.gridCalibration.active) return;
    if (s.activeMode === 'extent') return;
    s.tx = panStart.tx + (e.clientX - panStart.x);
    s.ty = panStart.ty + (e.clientY - panStart.y);
    applyTransform(s);
  });
  stage.addEventListener('pointerup', () => { panStart = null; });

  // --- Extent freehand drag (single finger only) — captures a polygon outline ---
  stage.addEventListener('pointerdown', (e) => {
    if (s.activeMode !== 'extent') return;
    if (e.pointerType === 'touch' && s.pointers.size > 1) return;
    // Ignore clicks on an existing highlight — those are handled via click.
    if (e.target.closest('[data-highlight-id]')) return;
    const startPt = clientToCanvas(s, e.clientX, e.clientY);
    s.extent.points       = [startPt];
    s.extent.tempEl       = createTempExtentPolygon(s, startPt);
    s.extent.dragStartCanvas = startPt;       // kept for backwards-compat checks below
  });
  stage.addEventListener('pointermove', (e) => {
    if (s.activeMode !== 'extent' || !s.extent.dragStartCanvas) return;
    if (s.pointers.size > 1) return;
    const cur = clientToCanvas(s, e.clientX, e.clientY);
    appendPolygonPoint(s, cur);
  });
  stage.addEventListener('pointerup', async (e) => {
    if (s.activeMode !== 'extent' || !s.extent.dragStartCanvas) return;
    const points = s.extent.points || [];
    // Clean up temp shape
    if (s.extent.tempEl) { s.extent.tempEl.remove(); s.extent.tempEl = null; }
    s.extent.dragStartCanvas = null;
    s.extent.points = null;
    // Sanity: need at least 4 points and a meaningful bbox to be a closed polygon.
    if (points.length < 4) return;
    const bbox = polygonBbox(points);
    if (bbox.w < 12 || bbox.h < 12) return;     // ignore tiny taps
    await commitHighlightPolygon(s, points);
  });

  // --- Mouse wheel zoom (desktop) ---
  stage.addEventListener('wheel', (e) => {
    e.preventDefault();
    const factor = Math.exp(-e.deltaY * 0.0018);
    zoomAt(s, e.clientX, e.clientY, factor);
  }, { passive: false });

  // --- Click / tap router — calibrate, grid-calibrate, open-pin, drop-pin, delete-highlight ---
  stage.addEventListener('click', (e) => {
    // 1) Distance calibration
    if (s.calibration.active) {
      const { x, y } = clientToCanvas(s, e.clientX, e.clientY);
      handleCalibrationClick(s, x, y);
      return;
    }

    // 1b) Grid calibration — 2-point flow, labels asked after 2nd tap
    if (s.gridCalibration.active) {
      const { x, y } = clientToCanvas(s, e.clientX, e.clientY);
      handleGridCalibrationClick(s, x, y);
      return;
    }

    // 2) Click on an existing pin (any mode — view included)
    const pinEl = e.target.closest('[data-item-id]');
    if (pinEl && s.isMarkup) {
      const itemId = Number(pinEl.dataset.itemId);
      openPinItemPanel(s, itemId);
      return;
    }

    // 3) Extent mode: click an existing highlight to delete
    const hlEl = e.target.closest('[data-highlight-id]');
    if (hlEl && s.activeMode === 'extent') {
      const hid = Number(hlEl.dataset.highlightId);
      confirmAndDeleteHighlight(s, hid);
      return;
    }

    // 4) Pin mode: drop a new pin at the tap
    if (s.activeMode === 'pin') {
      const { x, y } = clientToCanvas(s, e.clientX, e.clientY);
      dropPinAt(s, x, y);
      return;
    }
  });

  stage.addEventListener('contextmenu', (e) => e.preventDefault());

  // Resize handler — track and clean up on re-mount
  if (resizeHandler) window.removeEventListener('resize', resizeHandler);
  resizeHandler = () => { try { fitToScreen(s); } catch {} };
  window.addEventListener('resize', resizeHandler, { passive: true });
}

function zoomAt(s, clientX, clientY, factor) {
  const stageRect = s.stage.getBoundingClientRect();
  const localX = clientX - stageRect.left;
  const localY = clientY - stageRect.top;

  const newScale = clamp(s.scale * factor, 0.05, 8.0);
  const k = newScale / s.scale;
  s.tx = localX - (localX - s.tx) * k;
  s.ty = localY - (localY - s.ty) * k;
  s.scale = newScale;
  applyTransform(s);
}

function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }

function clientToCanvas(s, clientX, clientY) {
  const rect = s.canvas.getBoundingClientRect();
  return {
    x: (clientX - rect.left) / s.scale,
    y: (clientY - rect.top)  / s.scale
  };
}

/* --------------------------------------------------------------------------
   Calibration
   -------------------------------------------------------------------------- */

function enterCalibrationMode(s) {
  exitAnyMode(s, { silent: true });
  s.calibration.active = true;
  s.calibration.points = [];
  clearCalibrationMarkers(s);
  showModeBar('Tap the first known point on the drawing');
  s.stage.classList.add('is-calibrating');
}

function exitCalibrationMode(s) {
  s.calibration.active = false;
  s.calibration.points = [];
  clearCalibrationMarkers(s);
  hideModeBar();
  s.stage.classList.remove('is-calibrating');
}

function clearCalibrationMarkers(s) {
  s.overlay.querySelectorAll('[data-calibration]').forEach((el) => el.remove());
}

function handleCalibrationClick(s, x, y) {
  s.calibration.points.push({ x, y });
  drawCalibrationMarkers(s);

  if (s.calibration.points.length === 1) {
    document.getElementById('dv-mode-text').textContent = 'Tap the second known point';
  } else if (s.calibration.points.length === 2) {
    promptCalibrationDistance(s);
  }
}

function drawCalibrationMarkers(s) {
  clearCalibrationMarkers(s);
  const pts = s.calibration.points;
  if (pts.length >= 1) {
    const c1 = document.createElementNS(SVG_NS, 'circle');
    c1.setAttribute('cx', pts[0].x); c1.setAttribute('cy', pts[0].y);
    c1.setAttribute('r', '8'); c1.setAttribute('class', 'cal-dot');
    c1.setAttribute('data-calibration', '');
    s.overlay.appendChild(c1);
  }
  if (pts.length >= 2) {
    const line = document.createElementNS(SVG_NS, 'line');
    line.setAttribute('x1', pts[0].x); line.setAttribute('y1', pts[0].y);
    line.setAttribute('x2', pts[1].x); line.setAttribute('y2', pts[1].y);
    line.setAttribute('class', 'cal-line');
    line.setAttribute('data-calibration', '');
    s.overlay.appendChild(line);

    const c2 = document.createElementNS(SVG_NS, 'circle');
    c2.setAttribute('cx', pts[1].x); c2.setAttribute('cy', pts[1].y);
    c2.setAttribute('r', '8'); c2.setAttribute('class', 'cal-dot');
    c2.setAttribute('data-calibration', '');
    s.overlay.appendChild(c2);
  }
}

async function promptCalibrationDistance(s) {
  const [p1, p2] = s.calibration.points;

  const result = await openModal({
    title: 'Enter true distance',
    primaryLabel: 'Set scale',
    bodyHtml: `
      <p class="muted">You tapped two points. Enter the real-world distance between them in millimetres (from the drawing's dimension annotation).</p>
      <div class="form-field">
        <label for="cal-dist">Distance</label>
        <div class="input-with-suffix">
          <input id="cal-dist" name="realDistanceMm" type="number" min="1" step="1" inputmode="numeric" required autofocus placeholder="e.g. 5000"/>
          <span class="suffix">mm</span>
        </div>
        <div class="form-field__help muted">Typical: a gridline spacing, a slab width, or a column-to-column dimension shown on the drawing.</div>
      </div>
    `,
    onConfirm: (dialog) => {
      const val = Number(dialog.querySelector('[name=realDistanceMm]').value);
      if (!val || val <= 0) throw new Error('Enter a positive distance');
      return val;
    }
  });

  if (!result) {
    exitCalibrationMode(s);
    return;
  }

  const vp = s.viewport;
  const pdf1 = vp.convertToPdfPoint(p1.x, p1.y);
  const pdf2 = vp.convertToPdfPoint(p2.x, p2.y);
  const pdfDist = Math.hypot(pdf2[0] - pdf1[0], pdf2[1] - pdf1[1]);
  const mmPerPoint = result / pdfDist;

  const calibration = {
    p1: { x: pdf1[0], y: pdf1[1] },
    p2: { x: pdf2[0], y: pdf2[1] },
    realDistanceMm: result,
    mmPerPoint,
    calibratedAt: new Date().toISOString()
  };

  await setDrawingCalibration(s.drawing.id, calibration);
  s.drawing.calibration = calibration;
  toast('Calibration saved', { kind: 'success' });
  exitCalibrationMode(s);
  updateScaleInfo(s);
}

/* --------------------------------------------------------------------------
   Mode bar helpers
   -------------------------------------------------------------------------- */

function showModeBar(text) {
  const bar = document.getElementById('dv-mode-bar');
  const label = document.getElementById('dv-mode-text');
  if (bar && label) {
    label.textContent = text;
    bar.classList.remove('is-hidden');
  }
}

function hideModeBar() {
  const bar = document.getElementById('dv-mode-bar');
  if (bar) bar.classList.add('is-hidden');
}

/* --------------------------------------------------------------------------
   Grid calibration — two labelled intersections → auto gridRef on pin drops
   -------------------------------------------------------------------------- */

function enterGridCalibrationMode(s) {
  exitAnyMode(s, { silent: true });
  s.gridCalibration.active = true;
  s.gridCalibration.points = [];
  clearGridCalibrationMarkers(s);
  s.stage.classList.add('is-calibrating');
  showModeBar('Tap the first known grid intersection (e.g. 1/A)');
}

function exitGridCalibrationMode(s) {
  s.gridCalibration.active = false;
  s.gridCalibration.points = [];
  clearGridCalibrationMarkers(s);
  hideModeBar();
  s.stage.classList.remove('is-calibrating');
}

function clearGridCalibrationMarkers(s) {
  s.overlay.querySelectorAll('[data-gridcal]').forEach((el) => el.remove());
}

async function handleGridCalibrationClick(s, x, y) {
  s.gridCalibration.points.push({ x, y });
  drawGridCalibrationMarkers(s);

  if (s.gridCalibration.points.length === 1) {
    // Ask for labels of the first point
    const labels = await promptGridLabels('First intersection', 'e.g. column 1, row A');
    if (!labels) { exitGridCalibrationMode(s); return; }
    s.gridCalibration.points[0] = { ...s.gridCalibration.points[0], ...labels };
    document.getElementById('dv-mode-text').textContent = 'Tap the second known grid intersection';
  } else if (s.gridCalibration.points.length === 2) {
    const labels = await promptGridLabels('Second intersection', 'Needs to be on a different column AND row');
    if (!labels) { exitGridCalibrationMode(s); return; }
    s.gridCalibration.points[1] = { ...s.gridCalibration.points[1], ...labels };
    await commitGridCalibration(s);
  }
}

function drawGridCalibrationMarkers(s) {
  clearGridCalibrationMarkers(s);
  for (const p of s.gridCalibration.points) {
    const c = document.createElementNS(SVG_NS, 'circle');
    c.setAttribute('cx', p.x); c.setAttribute('cy', p.y);
    c.setAttribute('r', '10'); c.setAttribute('class', 'cal-dot');
    c.setAttribute('data-gridcal', '');
    s.overlay.appendChild(c);
  }
}

async function promptGridLabels(title, help) {
  let labels = null;
  await openModal({
    title,
    primaryLabel: 'OK',
    bodyHtml: `
      <p class="muted">${help}</p>
      <div class="form" style="display:grid; grid-template-columns:1fr 1fr; gap:var(--space-3);">
        <div class="form-field">
          <label for="gc-col">Column</label>
          <input id="gc-col" name="col" type="text" placeholder="e.g. 1" autofocus autocomplete="off"/>
        </div>
        <div class="form-field">
          <label for="gc-row">Row</label>
          <input id="gc-row" name="row" type="text" placeholder="e.g. A" autocomplete="off"/>
        </div>
      </div>
      <div class="form-field__help muted">Use numbers for one axis and letters for the other, just as they appear on the drawing.</div>
    `,
    onConfirm: (dialog) => {
      const col = dialog.querySelector('[name=col]').value.trim();
      const row = dialog.querySelector('[name=row]').value.trim();
      if (!col || !row) throw new Error('Enter both a column and a row label');
      labels = { col, row };
      return true;
    }
  });
  return labels;
}

async function commitGridCalibration(s) {
  const [p1, p2] = s.gridCalibration.points;
  // Validate before save
  try {
    validateCalibrationInput({
      p1col: p1.col, p1row: p1.row,
      p2col: p2.col, p2row: p2.row
    });
  } catch (err) {
    toast(err.message, { kind: 'error' });
    exitGridCalibrationMode(s);
    return;
  }

  // Convert canvas to PDF space for storage — survives zoom/rotate.
  const vp = s.viewport;
  const [p1x, p1y] = vp.convertToPdfPoint(p1.x, p1.y);
  const [p2x, p2y] = vp.convertToPdfPoint(p2.x, p2.y);

  const gridCalibration = {
    p1: { pdfX: p1x, pdfY: p1y, col: p1.col, row: p1.row },
    p2: { pdfX: p2x, pdfY: p2y, col: p2.col, row: p2.row },
    calibratedAt: new Date().toISOString()
  };

  await setDrawingGridCalibration(s.drawing.id, gridCalibration);
  s.drawing.gridCalibration = gridCalibration;
  toast('Grid saved — new pins will auto-label their location.', { kind: 'success' });
  exitGridCalibrationMode(s);
}

/* --------------------------------------------------------------------------
   Markup — Pin mode
   -------------------------------------------------------------------------- */

function togglePinMode(s) {
  if (s.activeMode === 'pin') {
    exitAnyMode(s);
  } else {
    enterPinMode(s);
  }
}

function enterPinMode(s) {
  exitAnyMode(s, { silent: true });
  s.activeMode = 'pin';
  s.stage.classList.add('is-pinning');
  showModeBar('Tap the plan to drop a numbered pin');
  const btn = document.getElementById('btn-pin');
  if (btn) btn.setAttribute('aria-pressed', 'true');
}

function exitPinMode(s) {
  if (s.activeMode !== 'pin') return;
  s.activeMode = 'view';
  s.stage.classList.remove('is-pinning');
  hideModeBar();
  const btn = document.getElementById('btn-pin');
  if (btn) btn.setAttribute('aria-pressed', 'false');
}

async function dropPinAt(s, canvasX, canvasY) {
  if (!s.isMarkup || !s.inspection) return;

  const [pdfX, pdfY] = s.viewport.convertToPdfPoint(canvasX, canvasY);

  // If the drawing has a grid calibration, auto-fill the gridRef so the
  // engineer doesn't have to type the location.
  let gridRef = '';
  if (s.drawing.gridCalibration) {
    try { gridRef = computeGridRef(pdfX, pdfY, s.drawing.gridCalibration); }
    catch {}
  }

  let item;
  try {
    item = await createItem(s.inspection.id, {
      drawingId: s.drawing.id,
      page: 1,
      pdfX, pdfY,
      comment: '',
      gridRef,
      severity: 'observation',
      status: 'open'
    });
  } catch (err) {
    console.error(err);
    toast('Couldn\u2019t create item', { kind: 'error' });
    return;
  }

  s.items.push(item);
  renderMarkupOverlay(s);
  // Open the panel so the engineer can add comment + photos immediately.
  await openPinItemPanel(s, item.id);
}

async function openPinItemPanel(s, itemId) {
  const result = await openItemPanel({ itemId });
  // Refresh the local item list from DB in case comment/severity/status changed
  // or the item was deleted.
  s.items = (await listItemsForInspection(s.inspection.id))
    .filter((it) => it.drawingId === s.drawing.id);
  renderMarkupOverlay(s);

  if (result?.deleted) toast('Item deleted', { kind: 'info' });
  else if (result?.item) toast('Item saved', { kind: 'success' });
}

/* --------------------------------------------------------------------------
   Markup — Extent mode
   -------------------------------------------------------------------------- */

function toggleExtentMode(s) {
  if (s.activeMode === 'extent') {
    exitAnyMode(s);
  } else {
    enterExtentMode(s);
  }
}

function enterExtentMode(s) {
  exitAnyMode(s, { silent: true });
  s.activeMode = 'extent';
  s.stage.classList.add('is-extenting');
  showModeBar('Trace around the area inspected with one finger · tap an existing area to remove it');
  const btn = document.getElementById('btn-extent');
  if (btn) btn.setAttribute('aria-pressed', 'true');
}

function exitExtentMode(s) {
  if (s.activeMode !== 'extent') return;
  s.activeMode = 'view';
  s.stage.classList.remove('is-extenting');
  hideModeBar();
  if (s.extent.tempEl) { s.extent.tempEl.remove(); s.extent.tempEl = null; }
  s.extent.dragStartCanvas = null;
  const btn = document.getElementById('btn-extent');
  if (btn) btn.setAttribute('aria-pressed', 'false');
}

/**
 * Create the in-progress polygon element for the extent drag. We use SVG
 * <polygon> so the shape closes visually as the user drags, even mid-stroke.
 */
function createTempExtentPolygon(s, startCanvas) {
  const poly = document.createElementNS(SVG_NS, 'polygon');
  poly.setAttribute('class', 'highlight-poly highlight-poly--temp');
  poly.setAttribute('points', `${startCanvas.x},${startCanvas.y}`);
  s.overlay.appendChild(poly);
  return poly;
}

/**
 * Append a point to the in-progress polygon. For performance and clean
 * visuals we skip points that are very close to the previous one (sub-pixel
 * jitter from a stationary finger). The DB stores a simplified version
 * after pointerup via simplifyPolygon().
 */
function appendPolygonPoint(s, pt) {
  if (!s.extent.tempEl) return;
  const last = s.extent.points[s.extent.points.length - 1];
  if (last && Math.hypot(pt.x - last.x, pt.y - last.y) < 2) return;  // dedupe ~< 2px
  s.extent.points.push(pt);
  const attr = s.extent.points.map((p) => `${p.x},${p.y}`).join(' ');
  s.extent.tempEl.setAttribute('points', attr);
}

/**
 * Convert a canvas-space polygon to PDF space, simplify, and store. The
 * simplification keeps the visible silhouette while shrinking the JSON
 * footprint — typical 100-point freehand outline drops to ~15-20 points.
 */
async function commitHighlightPolygon(s, canvasPoints) {
  const vp = s.viewport;
  const pdfPoints = canvasPoints.map((p) => {
    const [px, py] = vp.convertToPdfPoint(p.x, p.y);
    return { x: px, y: py };
  });

  // Douglas–Peucker tolerance ≈ 1.5 PDF points (~0.5 mm at typical scales).
  const simplified = simplifyPolygon(pdfPoints, 1.5);
  if (simplified.length < 3) return;

  try {
    const hl = await addHighlight(s.inspection.id, {
      drawingId: s.drawing.id,
      page: 1,
      pdfPoints: simplified
    });
    s.highlights.push(hl);
    renderMarkupOverlay(s);
  } catch (err) {
    console.error(err);
    toast('Couldn\u2019t save area', { kind: 'error' });
  }
}

/* ----- Polygon utilities ----- */

function polygonBbox(points) {
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  for (const p of points) {
    if (p.x < minX) minX = p.x;
    if (p.y < minY) minY = p.y;
    if (p.x > maxX) maxX = p.x;
    if (p.y > maxY) maxY = p.y;
  }
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
}

/**
 * Iterative Douglas–Peucker polyline simplification. Keeps the points that
 * matter for the silhouette and drops noise. Always preserves first + last.
 */
function simplifyPolygon(points, tolerance) {
  if (points.length <= 2) return points;
  const sqTol = tolerance * tolerance;

  const keep = new Array(points.length).fill(false);
  keep[0] = true;
  keep[points.length - 1] = true;

  const stack = [[0, points.length - 1]];
  while (stack.length) {
    const [first, last] = stack.pop();
    let maxDist = 0;
    let index   = first;
    for (let i = first + 1; i < last; i++) {
      const d = sqDistToSegment(points[i], points[first], points[last]);
      if (d > maxDist) { maxDist = d; index = i; }
    }
    if (maxDist > sqTol) {
      keep[index] = true;
      stack.push([first, index]);
      stack.push([index, last]);
    }
  }
  return points.filter((_, i) => keep[i]);
}

function sqDistToSegment(p, a, b) {
  let x = a.x, y = a.y;
  let dx = b.x - x, dy = b.y - y;
  if (dx !== 0 || dy !== 0) {
    const t = ((p.x - x) * dx + (p.y - y) * dy) / (dx * dx + dy * dy);
    if (t > 1) { x = b.x; y = b.y; }
    else if (t > 0) { x += dx * t; y += dy * t; }
  }
  dx = p.x - x; dy = p.y - y;
  return dx * dx + dy * dy;
}

async function confirmAndDeleteHighlight(s, highlightId) {
  // Inline delete — no nested modal needed. Direct toast undo would be nicer,
  // but for v1 we just ask.
  const ok = await confirmDialog({
    title: 'Remove inspected area?',
    message: 'The highlight will be removed from this inspection.',
    confirmLabel: 'Remove',
    danger: true
  });
  if (!ok) return;
  try {
    await deleteHighlight(highlightId);
    s.highlights = s.highlights.filter((h) => h.id !== highlightId);
    renderMarkupOverlay(s);
  } catch (err) {
    console.error(err);
    toast('Couldn\u2019t remove area', { kind: 'error' });
  }
}

/* --------------------------------------------------------------------------
   Markup — Overlay rendering (pins + highlights)
   -------------------------------------------------------------------------- */

function renderMarkupOverlay(s) {
  // Wipe existing markup layers (keep calibration markers if present)
  s.overlay.querySelectorAll('[data-markup-layer]').forEach((el) => el.remove());

  // Highlights first — painted under pins
  const hlLayer = document.createElementNS(SVG_NS, 'g');
  hlLayer.setAttribute('data-markup-layer', 'highlights');
  s.overlay.appendChild(hlLayer);
  for (const h of s.highlights) renderHighlight(s, hlLayer, h);

  // Pins on top
  const pinLayer = document.createElementNS(SVG_NS, 'g');
  pinLayer.setAttribute('data-markup-layer', 'pins');
  s.overlay.appendChild(pinLayer);
  for (const it of s.items) renderPin(s, pinLayer, it);

  updatePinSizes(s);
}

function renderHighlight(s, layer, h) {
  const vp = s.viewport;

  // New polygon shape (Phase 11+).
  if (Array.isArray(h.pdfPoints) && h.pdfPoints.length >= 3) {
    const ptsAttr = h.pdfPoints.map((p) => {
      const [cx, cy] = vp.convertToViewportPoint(p.x, p.y);
      return `${cx},${cy}`;
    }).join(' ');
    const poly = document.createElementNS(SVG_NS, 'polygon');
    poly.setAttribute('class', 'highlight-poly');
    poly.setAttribute('data-highlight-id', h.id);
    poly.setAttribute('points', ptsAttr);
    layer.appendChild(poly);
    return;
  }

  // Legacy rectangle.
  const [cx0, cy0] = vp.convertToViewportPoint(h.pdfX, h.pdfY);
  const [cx1, cy1] = vp.convertToViewportPoint(h.pdfX + h.pdfW, h.pdfY + h.pdfH);
  const x = Math.min(cx0, cx1);
  const y = Math.min(cy0, cy1);
  const w = Math.abs(cx1 - cx0);
  const h2 = Math.abs(cy1 - cy0);

  const rect = document.createElementNS(SVG_NS, 'rect');
  rect.setAttribute('class', 'highlight-rect');
  rect.setAttribute('data-highlight-id', h.id);
  rect.setAttribute('x', x);
  rect.setAttribute('y', y);
  rect.setAttribute('width',  w);
  rect.setAttribute('height', h2);
  layer.appendChild(rect);
}

function renderPin(s, layer, item) {
  if (item.pdfX == null || item.pdfY == null) return;
  const [cx, cy] = s.viewport.convertToViewportPoint(item.pdfX, item.pdfY);

  const g = document.createElementNS(SVG_NS, 'g');
  const classes = ['pin', `pin--${item.severity || 'observation'}`];
  if (item.status === 'closed') classes.push('pin--closed');
  g.setAttribute('class', classes.join(' '));
  g.setAttribute('data-item-id', item.id);
  g.setAttribute('data-cx', cx);
  g.setAttribute('data-cy', cy);
  // transform set in updatePinSizes so it counter-scales with zoom

  const c = document.createElementNS(SVG_NS, 'circle');
  c.setAttribute('class', 'pin__circle');
  c.setAttribute('r', '14');
  g.appendChild(c);

  const t = document.createElementNS(SVG_NS, 'text');
  t.setAttribute('class', 'pin__text');
  t.setAttribute('text-anchor', 'middle');
  t.setAttribute('dominant-baseline', 'central');
  t.setAttribute('y', '1');
  t.textContent = String(item.itemNumber);
  g.appendChild(t);

  layer.appendChild(g);
}

function updatePinSizes(s) {
  // Counter-scale pins so they stay a constant screen size across zoom levels.
  const inv = 1 / Math.max(s.scale, 0.05);
  s.overlay.querySelectorAll('.pin').forEach((g) => {
    const cx = g.getAttribute('data-cx');
    const cy = g.getAttribute('data-cy');
    g.setAttribute('transform', `translate(${cx},${cy}) scale(${inv})`);
  });
}

/* --------------------------------------------------------------------------
   Mode exit helper
   -------------------------------------------------------------------------- */

function exitAnyMode(s, { silent = false } = {}) {
  if (s.calibration.active) {
    exitCalibrationMode(s);
  }
  if (s.gridCalibration.active) {
    exitGridCalibrationMode(s);
  }
  if (s.activeMode === 'pin') exitPinMode(s);
  if (s.activeMode === 'extent') exitExtentMode(s);
  if (!silent) hideModeBar();
}

/* --------------------------------------------------------------------------
   Utilities
   -------------------------------------------------------------------------- */

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
function escapeAttr(s) { return escapeHtml(s).replace(/"/g, '&quot;'); }
