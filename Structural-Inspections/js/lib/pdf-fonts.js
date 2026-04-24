/**
 * Libre Franklin font loader for jsPDF (Phase V2).
 *
 * jsPDF ships with the 14 standard PDF fonts only (Helvetica, Times, Courier).
 * To match the Bligh Tanner house style we embed Libre Franklin — the web
 * equivalent of the Franklin Gothic brand typeface — as TTF bytes.
 *
 * Three weights are enough for the report:
 *   Regular (400), Medium (500), Bold (700)
 *
 * Loaded lazily the first time generateReport() runs. Cached between calls.
 *
 * On failure (network error, corrupt asset, etc.) the caller falls back to
 * Helvetica so the report still generates — never let typography block an
 * engineer leaving site with a PDF.
 */

const FONT_URLS = {
  regular: 'assets/fonts/LibreFranklin-Regular.ttf',
  medium:  'assets/fonts/LibreFranklin-Medium.ttf',
  bold:    'assets/fonts/LibreFranklin-Bold.ttf',
  italic:  'assets/fonts/LibreFranklin-Italic.ttf'
};

const FONT_FAMILY = 'LibreFranklin';

let registrationPromise = null;

async function fetchAsBase64(url) {
  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`Font fetch failed (${resp.status}) for ${url}`);
  const buf = await resp.arrayBuffer();
  const bytes = new Uint8Array(buf);
  // chunk to avoid hitting the call-stack limit on String.fromCharCode for ~113KB TTFs
  let bin = '';
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  }
  return btoa(bin);
}

async function registerFonts(pdf) {
  const [reg, med, bold, italic] = await Promise.all([
    fetchAsBase64(FONT_URLS.regular),
    fetchAsBase64(FONT_URLS.medium),
    fetchAsBase64(FONT_URLS.bold),
    fetchAsBase64(FONT_URLS.italic)
  ]);

  pdf.addFileToVFS('LibreFranklin-Regular.ttf', reg);
  pdf.addFont('LibreFranklin-Regular.ttf', FONT_FAMILY, 'normal');
  pdf.addFileToVFS('LibreFranklin-Medium.ttf', med);
  pdf.addFont('LibreFranklin-Medium.ttf', FONT_FAMILY, 'normal', 500);
  pdf.addFileToVFS('LibreFranklin-Bold.ttf', bold);
  pdf.addFont('LibreFranklin-Bold.ttf', FONT_FAMILY, 'bold');
  pdf.addFileToVFS('LibreFranklin-Italic.ttf', italic);
  pdf.addFont('LibreFranklin-Italic.ttf', FONT_FAMILY, 'italic');
}

/**
 * Register Libre Franklin with a jsPDF instance. Returns the font family
 * string to use going forward, or 'helvetica' if loading failed.
 *
 * Caller pattern:
 *   const family = await ensureReportFont(pdf);
 *   pdf.setFont(family, 'bold');
 */
export async function ensureReportFont(pdf) {
  try {
    if (!registrationPromise) {
      registrationPromise = registerFonts(pdf);
    } else {
      // Re-register on new pdf instances — font cache lives on the instance.
      await registerFonts(pdf);
    }
    await registrationPromise;
    // Sanity check: can we set it?
    pdf.setFont(FONT_FAMILY, 'normal');
    return FONT_FAMILY;
  } catch (err) {
    console.warn('Libre Franklin load failed, falling back to Helvetica:', err);
    registrationPromise = null;
    return 'helvetica';
  }
}
