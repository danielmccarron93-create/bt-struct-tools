/**
 * Photo capture + compression (Phase 4).
 *
 * Two helpers:
 *   - openPhotoCapture()    → prompts the device camera (or file picker on desktop),
 *                             returns the raw File, or null if cancelled.
 *   - compressImage(file)   → resizes + re-encodes as JPEG, returns
 *                             { blob, width, height } ready to persist.
 *
 * Storage target: max 2400px long edge, JPEG q=0.85. That keeps each photo
 * around ~0.8–1.2 MB for a modern phone camera — good enough for an A4 report
 * and forgiving on IndexedDB.
 *
 * EXIF orientation: createImageBitmap honours the image's EXIF rotation when
 * given `imageOrientation: 'from-image'` (Safari 14+, all Chromium). This
 * means photos taken in portrait render upright without extra rotation code.
 */

const DEFAULT_MAX_EDGE = 2400;
const DEFAULT_QUALITY  = 0.85;

/**
 * Open a file picker that prefers the rear camera on mobile.
 *
 * Returns a Promise<File | null>.
 */
export function openPhotoCapture({ preferCamera = true } = {}) {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type   = 'file';
    input.accept = 'image/*';
    if (preferCamera) input.setAttribute('capture', 'environment');
    input.style.position = 'fixed';
    input.style.left = '-9999px';
    input.style.opacity = '0';
    document.body.appendChild(input);

    let settled = false;
    const finish = (file) => {
      if (settled) return;
      settled = true;
      input.remove();
      resolve(file || null);
    };

    input.addEventListener('change', () => {
      finish(input.files && input.files[0] ? input.files[0] : null);
    });
    // `focus` returning without `change` means the user cancelled.
    // Safari/iOS is unreliable here so we also resolve after a long grace.
    window.addEventListener('focus', () => {
      setTimeout(() => {
        if (!input.files || input.files.length === 0) finish(null);
      }, 1500);
    }, { once: true });

    input.click();
  });
}

/**
 * Open a plain file picker (no camera hint) — useful for "Choose from library"
 * on mobile or selecting an existing image on desktop.
 */
export function openPhotoLibrary() {
  return openPhotoCapture({ preferCamera: false });
}

/**
 * Resize + re-encode an image File/Blob.
 *
 * Returns { blob, width, height } where blob is a fresh image/jpeg blob.
 *
 * Throws if the file cannot be decoded (e.g. HEIC on old Safari, corrupt file).
 */
export async function compressImage(fileOrBlob, {
  maxLongEdge = DEFAULT_MAX_EDGE,
  quality     = DEFAULT_QUALITY,
  mimeType    = 'image/jpeg'
} = {}) {
  if (!fileOrBlob) throw new Error('No image supplied');

  // Prefer createImageBitmap — faster and honours EXIF orientation natively.
  let bitmap = null;
  try {
    bitmap = await createImageBitmap(fileOrBlob, { imageOrientation: 'from-image' });
  } catch {
    // Fallback: <img> → canvas (ignores EXIF orientation, but works on older Safari)
    bitmap = await loadImageElement(fileOrBlob);
  }

  const srcW = bitmap.width  || bitmap.naturalWidth;
  const srcH = bitmap.height || bitmap.naturalHeight;
  if (!srcW || !srcH) throw new Error('Could not read image dimensions');

  const longEdge = Math.max(srcW, srcH);
  const scale    = longEdge > maxLongEdge ? maxLongEdge / longEdge : 1;
  const outW     = Math.round(srcW * scale);
  const outH     = Math.round(srcH * scale);

  const canvas = document.createElement('canvas');
  canvas.width  = outW;
  canvas.height = outH;
  const ctx = canvas.getContext('2d');
  // White background for JPEG (avoids black fringes on transparent PNGs)
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, outW, outH);
  ctx.drawImage(bitmap, 0, 0, outW, outH);

  // Clean up the bitmap if it supports it
  if (bitmap.close) bitmap.close();

  const blob = await new Promise((resolve, reject) => {
    canvas.toBlob(
      (b) => (b ? resolve(b) : reject(new Error('Image encoding failed'))),
      mimeType,
      quality
    );
  });

  return { blob, width: outW, height: outH };
}

/**
 * Fallback loader for browsers where createImageBitmap(blob) fails.
 */
function loadImageElement(fileOrBlob) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(fileOrBlob);
    const img = new Image();
    img.decoding = 'async';
    img.onload  = () => { URL.revokeObjectURL(url); resolve(img); };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Image decode failed')); };
    img.src = url;
  });
}

/* --------------------------------------------------------------------------
   Object URL lifecycle helper
   -------------------------------------------------------------------------- */

/**
 * Create an object URL for a Blob and schedule it for revocation once the
 * provided cleanup signal fires. Use via:
 *
 *   const url = objectUrlFor(blob, cleanup);
 *   imgEl.src = url;
 *   // ...later
 *   cleanup.revokeAll();
 *
 * This keeps memory tidy when rendering photo grids that get re-drawn.
 */
export function createUrlRegistry() {
  const urls = [];
  return {
    urlFor(blob) {
      const u = URL.createObjectURL(blob);
      urls.push(u);
      return u;
    },
    revokeAll() {
      while (urls.length) URL.revokeObjectURL(urls.pop());
    }
  };
}
