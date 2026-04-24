/**
 * Photo annotation modal (Phase V2).
 *
 * Opens a full-screen overlay where the engineer marks up a photo with
 * simple field-friendly tools:
 *
 *   - Circle     (drag → ellipse)
 *   - Arrow      (drag → straight arrow from start to end)
 *   - Freehand   (drag → smoothed polyline)
 *   - Text       (tap → one-line caption)
 *   - Undo       (pop last operation)
 *   - Clear
 *
 * Annotations are stored as a plain JSON array of vector operations on the
 * photo record, NOT burned into the source JPEG. The report generator
 * composites photo + annotations onto a canvas at render time.
 *
 * Each op:
 *   { kind: 'circle', x, y, r, color }
 *   { kind: 'arrow',  x1, y1, x2, y2, color }
 *   { kind: 'path',   points: [{x,y}, ...], color }
 *   { kind: 'text',   x, y, text, color, size }
 *
 * Coordinates are expressed in the photo's native pixel space so the
 * annotations stay correct regardless of display / print scale.
 *
 * Usage:
 *   const updated = await openAnnotateModal({
 *     blob:        photoBlob,
 *     annotations: existingAnnotations || []
 *   });
 *   if (updated) { ...save updated.annotations on the photo... }
 */

export async function openAnnotateModal({ blob, annotations = [] }) {
  const { imgEl, width, height } = await loadImage(blob);

  return new Promise((resolve) => {
    // Working copy we mutate on each draw. Undo pops from this array.
    const ops = annotations.map((o) => ({ ...o }));
    const state = {
      tool:  'arrow',
      color: '#ff3333',
      // active drag state
      dragging: null   // { kind, startX, startY, lastPoints? }
    };

    const scrim = document.createElement('div');
    scrim.className = 'annotate-scrim';
    scrim.setAttribute('role', 'dialog');
    scrim.setAttribute('aria-modal', 'true');
    scrim.setAttribute('aria-label', 'Annotate photo');
    scrim.innerHTML = `
      <div class="annotate-toolbar">
        <button class="annotate-tool" data-tool="arrow"    aria-pressed="true">↗ Arrow</button>
        <button class="annotate-tool" data-tool="circle"   aria-pressed="false">◯ Circle</button>
        <button class="annotate-tool" data-tool="path"     aria-pressed="false">✎ Draw</button>
        <button class="annotate-tool" data-tool="text"     aria-pressed="false">Abc Text</button>
        <span class="annotate-toolbar__spacer"></span>
        <button class="annotate-color" data-color="#ff3333" aria-label="Red"    style="background:#ff3333"></button>
        <button class="annotate-color" data-color="#ffd500" aria-label="Yellow" style="background:#ffd500"></button>
        <button class="annotate-color" data-color="#2aa9ff" aria-label="Blue"   style="background:#2aa9ff"></button>
        <button class="annotate-color" data-color="#ffffff" aria-label="White"  style="background:#ffffff; border-color:#555"></button>
        <span class="annotate-toolbar__spacer"></span>
        <button class="annotate-btn" id="ann-undo" title="Undo">↶ Undo</button>
        <button class="annotate-btn" id="ann-clear" title="Clear all">✕ Clear</button>
      </div>
      <div class="annotate-stage" id="ann-stage">
        <canvas class="annotate-canvas" id="ann-canvas"></canvas>
      </div>
      <div class="annotate-footer">
        <button class="btn btn--secondary" id="ann-cancel">Cancel</button>
        <button class="btn btn--primary" id="ann-save">Save</button>
      </div>
    `;

    document.body.appendChild(scrim);

    const canvas = scrim.querySelector('#ann-canvas');
    const ctx = canvas.getContext('2d');
    canvas.width = width;
    canvas.height = height;

    // Fit the canvas into the viewport at its natural aspect
    const stage = scrim.querySelector('#ann-stage');
    const fitCanvas = () => {
      const rect = stage.getBoundingClientRect();
      const ratio = width / height;
      let w = rect.width, h = w / ratio;
      if (h > rect.height) { h = rect.height; w = h * ratio; }
      canvas.style.width  = `${w}px`;
      canvas.style.height = `${h}px`;
    };
    fitCanvas();
    const onResize = () => { fitCanvas(); };
    window.addEventListener('resize', onResize);

    function redraw() {
      ctx.clearRect(0, 0, width, height);
      ctx.drawImage(imgEl, 0, 0, width, height);
      for (const o of ops) drawOp(ctx, o);
      // Live preview of in-progress drag
      if (state.dragging && state.dragging.preview) {
        drawOp(ctx, state.dragging.preview);
      }
    }
    redraw();

    // --- Input routing ---
    function toImageCoords(ev) {
      const rect = canvas.getBoundingClientRect();
      const x = ((ev.clientX - rect.left) / rect.width)  * width;
      const y = ((ev.clientY - rect.top)  / rect.height) * height;
      return { x, y };
    }

    function onPointerDown(ev) {
      ev.preventDefault();
      canvas.setPointerCapture?.(ev.pointerId);
      const { x, y } = toImageCoords(ev);
      if (state.tool === 'text') {
        promptAndAddText(x, y);
        return;
      }
      state.dragging = { kind: state.tool, startX: x, startY: y, points: [{ x, y }], preview: null };
    }

    function onPointerMove(ev) {
      if (!state.dragging) return;
      const { x, y } = toImageCoords(ev);
      const d = state.dragging;
      if (d.kind === 'circle') {
        const r = Math.hypot(x - d.startX, y - d.startY);
        d.preview = { kind: 'circle', x: d.startX, y: d.startY, r, color: state.color };
      } else if (d.kind === 'arrow') {
        d.preview = { kind: 'arrow', x1: d.startX, y1: d.startY, x2: x, y2: y, color: state.color };
      } else if (d.kind === 'path') {
        d.points.push({ x, y });
        d.preview = { kind: 'path', points: d.points.slice(), color: state.color };
      }
      redraw();
    }

    function onPointerUp() {
      const d = state.dragging;
      state.dragging = null;
      if (!d || !d.preview) { redraw(); return; }
      ops.push(d.preview);
      redraw();
    }

    async function promptAndAddText(x, y) {
      // Use a prompt to stay tiny — text annotations are rare.
      const text = window.prompt('Text label:');
      if (!text) return;
      const size = Math.max(18, Math.round(height / 24));
      ops.push({ kind: 'text', x, y, text, color: state.color, size });
      redraw();
    }

    canvas.addEventListener('pointerdown', onPointerDown);
    canvas.addEventListener('pointermove', onPointerMove);
    canvas.addEventListener('pointerup',   onPointerUp);
    canvas.addEventListener('pointercancel', onPointerUp);
    canvas.addEventListener('pointerleave',  () => { if (state.dragging) onPointerUp(); });

    // --- Toolbar buttons ---
    scrim.querySelectorAll('.annotate-tool').forEach((btn) => {
      btn.addEventListener('click', () => {
        scrim.querySelectorAll('.annotate-tool').forEach((b) => b.setAttribute('aria-pressed', 'false'));
        btn.setAttribute('aria-pressed', 'true');
        state.tool = btn.dataset.tool;
      });
    });
    scrim.querySelectorAll('.annotate-color').forEach((btn) => {
      btn.addEventListener('click', () => {
        scrim.querySelectorAll('.annotate-color').forEach((b) => b.classList.remove('is-selected'));
        btn.classList.add('is-selected');
        state.color = btn.dataset.color;
      });
    });
    scrim.querySelector('.annotate-color[data-color="#ff3333"]').classList.add('is-selected');

    scrim.querySelector('#ann-undo').addEventListener('click', () => {
      if (ops.length) ops.pop();
      redraw();
    });
    scrim.querySelector('#ann-clear').addEventListener('click', () => {
      ops.length = 0;
      redraw();
    });

    function cleanup(result) {
      window.removeEventListener('resize', onResize);
      scrim.remove();
      resolve(result);
    }

    scrim.querySelector('#ann-cancel').addEventListener('click', () => cleanup(null));
    scrim.querySelector('#ann-save').addEventListener('click', () =>
      cleanup({ annotations: ops }));

    document.addEventListener('keydown', function onKey(ev) {
      if (!document.body.contains(scrim)) {
        document.removeEventListener('keydown', onKey);
        return;
      }
      if (ev.key === 'Escape') { cleanup(null); document.removeEventListener('keydown', onKey); }
    });
  });
}

/**
 * Render an ops array on top of an image into a canvas. Returns a data URL
 * suitable for pdf.addImage(). Exported so report.js can composite at render
 * time without needing to re-open the annotation UI.
 */
export async function renderAnnotatedImage(blob, annotations, { maxLongEdge = 1600 } = {}) {
  const { imgEl, width, height } = await loadImage(blob);
  const ratio = width / height;
  let outW = Math.min(maxLongEdge, width);
  let outH = Math.round(outW / ratio);
  if (outH > maxLongEdge) { outH = maxLongEdge; outW = Math.round(outH * ratio); }

  const canvas = document.createElement('canvas');
  canvas.width  = outW;
  canvas.height = outH;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(imgEl, 0, 0, outW, outH);

  // Scale ops from native pixel space to output canvas space
  const sx = outW / width;
  const sy = outH / height;
  for (const op of annotations || []) {
    drawOp(ctx, scaleOp(op, sx, sy));
  }

  return { dataUrl: canvas.toDataURL('image/jpeg', 0.85), width: outW, height: outH };
}

function scaleOp(op, sx, sy) {
  const s = Math.min(sx, sy); // For stroke widths & radii
  switch (op.kind) {
    case 'circle': return { ...op, x: op.x * sx, y: op.y * sy, r: op.r * s };
    case 'arrow':  return { ...op, x1: op.x1 * sx, y1: op.y1 * sy, x2: op.x2 * sx, y2: op.y2 * sy };
    case 'path':   return { ...op, points: op.points.map((p) => ({ x: p.x * sx, y: p.y * sy })) };
    case 'text':   return { ...op, x: op.x * sx, y: op.y * sy, size: op.size * s };
    default:       return op;
  }
}

function drawOp(ctx, op) {
  const color = op.color || '#ff3333';
  ctx.save();
  ctx.strokeStyle = color;
  ctx.fillStyle = color;
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  const w = Math.max(3, Math.round(Math.min(ctx.canvas.width, ctx.canvas.height) / 180));
  ctx.lineWidth = w;

  if (op.kind === 'circle') {
    ctx.beginPath();
    ctx.arc(op.x, op.y, op.r, 0, Math.PI * 2);
    ctx.stroke();
  } else if (op.kind === 'arrow') {
    const { x1, y1, x2, y2 } = op;
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    // Arrowhead
    const ang = Math.atan2(y2 - y1, x2 - x1);
    const head = Math.max(14, w * 5);
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    ctx.lineTo(x2 - head * Math.cos(ang - Math.PI / 7), y2 - head * Math.sin(ang - Math.PI / 7));
    ctx.lineTo(x2 - head * Math.cos(ang + Math.PI / 7), y2 - head * Math.sin(ang + Math.PI / 7));
    ctx.closePath();
    ctx.fill();
  } else if (op.kind === 'path') {
    const pts = op.points || [];
    if (pts.length < 2) { ctx.restore(); return; }
    ctx.beginPath();
    ctx.moveTo(pts[0].x, pts[0].y);
    for (let i = 1; i < pts.length; i += 1) ctx.lineTo(pts[i].x, pts[i].y);
    ctx.stroke();
  } else if (op.kind === 'text') {
    const size = op.size || 22;
    ctx.font = `bold ${size}px -apple-system, BlinkMacSystemFont, "Helvetica Neue", Arial, sans-serif`;
    // Shadow for legibility on photo backgrounds
    ctx.strokeStyle = 'rgba(0,0,0,0.75)';
    ctx.lineWidth = Math.max(3, size / 6);
    ctx.lineJoin = 'round';
    ctx.strokeText(op.text, op.x, op.y);
    ctx.fillStyle = color;
    ctx.fillText(op.text, op.x, op.y);
  }

  ctx.restore();
}

function loadImage(blob) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      const w = img.naturalWidth || img.width;
      const h = img.naturalHeight || img.height;
      // Resolve but keep the URL alive — we revoke on loaded so next call gets a fresh URL.
      resolve({ imgEl: img, width: w, height: h });
      // Revoking now is fine: the img element already holds the decoded bitmap.
      setTimeout(() => URL.revokeObjectURL(url), 0);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error('Could not load photo for annotation'));
    };
    img.src = url;
  });
}
