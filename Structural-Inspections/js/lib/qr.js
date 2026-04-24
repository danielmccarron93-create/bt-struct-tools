/**
 * QR-code generation helper (Phase V2).
 *
 * Wraps `window.qrcode` from the qrcode-generator library (loaded in
 * index.html). Renders to a canvas so we can feed a PNG data-URL into jsPDF.
 *
 * Close-out workflow: the generated report's final page carries a QR that
 * the builder scans to open a pre-filled mailto. No backend required — the
 * builder emails rectification evidence back, the engineer pastes it
 * against each item inside the app.
 */

/**
 * Build the mailto URL the QR will encode.
 *
 *   mailto:{to}?subject=Re: Inspection 24089 — rectification evidence [2026-0001]&body=...
 *
 * Returns a string suitable for QR encoding or for a clickable link in the
 * app. Omits the `to` portion if no builderEmail is set (gives the builder
 * an empty-to mailto that still pre-fills subject + body).
 */
export function buildCloseOutMailto({ builderEmail, project, inspection, token }) {
  const job    = project?.jobNumber || 'UNKNOWN';
  const type   = inspection?.inspectionTypeName || 'inspection';
  const date   = inspection?.date || new Date().toISOString().slice(0, 10);
  const subject = `Re: Inspection ${job} — rectification evidence [${token}]`;
  const body = [
    `Hi,`,
    ``,
    `Attached are rectification photos for the items raised in the ${type} inspection of ${date}.`,
    ``,
    `Please reply against each item number as listed in the report.`,
    ``,
    `Inspection token: ${token}`,
    `Job: ${job}`,
    ``,
    `Thanks,`,
    `— Builder`
  ].join('\r\n');

  const q = `subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  return builderEmail ? `mailto:${builderEmail}?${q}` : `mailto:?${q}`;
}

/**
 * Stable short token for an inspection — used in the mailto subject so replies
 * can be tracked back. Non-secret; purely a human-readable reference.
 */
export function buildCloseOutToken(inspection) {
  const d = (inspection?.date || '').slice(0, 10).replace(/-/g, '');
  return `BT-${d}-${inspection?.id || '0'}`;
}

/**
 * Render a QR encoding `text` to a PNG data URL. `scale` is the pixel-size
 * of each QR module; 6-8 reads well in a 30-40mm print square.
 */
export function renderQrDataUrl(text, { scale = 7, margin = 2 } = {}) {
  if (!window.qrcode) throw new Error('QR library not loaded');
  // Type 0 = auto-choose ECC version. Error correction "M" is a good default.
  const qr = window.qrcode(0, 'M');
  qr.addData(text);
  qr.make();
  const modules = qr.getModuleCount();
  const size = (modules + margin * 2) * scale;

  const canvas = document.createElement('canvas');
  canvas.width = canvas.height = size;
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, size, size);
  ctx.fillStyle = '#000000';
  for (let r = 0; r < modules; r += 1) {
    for (let c = 0; c < modules; c += 1) {
      if (qr.isDark(r, c)) {
        ctx.fillRect((c + margin) * scale, (r + margin) * scale, scale, scale);
      }
    }
  }
  return canvas.toDataURL('image/png');
}
