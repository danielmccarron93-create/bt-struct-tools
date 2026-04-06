/**
 * 18a-steel-sections.js
 * Liberty/OneSteel Australian steel section property lookup table
 * and improved drawing functions with root fillets and accurate geometry.
 *
 * All functions declared globally (no IIFE) for shared scope with other files.
 * Loads BEFORE 18-details-system.js to provide lookup data.
 */

// ============================================================================
// STEEL SECTIONS CATALOGUE — Liberty/OneSteel Australian Sections
// ============================================================================

const STEEL_CATALOGUE = {
  UB: {
    '610UB125': { d: 612, bf: 229, tf: 19.6, tw: 11.9, r1: 14 },
    '610UB113': { d: 607, bf: 228, tf: 17.3, tw: 11.2, r1: 14 },
    '610UB101': { d: 602, bf: 228, tf: 14.8, tw: 10.6, r1: 14 },
    '530UB92.4': { d: 533, bf: 209, tf: 15.6, tw: 10.2, r1: 14 },
    '530UB82.0': { d: 528, bf: 209, tf: 13.2, tw: 9.6, r1: 14 },
    '460UB82.1': { d: 460, bf: 191, tf: 16.0, tw: 9.9, r1: 11.4 },
    '460UB74.6': { d: 457, bf: 190, tf: 14.5, tw: 9.1, r1: 11.4 },
    '460UB67.1': { d: 454, bf: 190, tf: 12.7, tw: 8.5, r1: 11.4 },
    '410UB59.7': { d: 406, bf: 178, tf: 12.8, tw: 7.8, r1: 11.4 },
    '410UB53.7': { d: 403, bf: 178, tf: 10.9, tw: 7.6, r1: 11.4 },
    '360UB56.7': { d: 359, bf: 172, tf: 13.0, tw: 8.0, r1: 11.4 },
    '360UB50.7': { d: 356, bf: 171, tf: 11.5, tw: 7.3, r1: 11.4 },
    '360UB44.7': { d: 352, bf: 171, tf: 9.7, tw: 6.9, r1: 11.4 },
    '310UB46.2': { d: 307, bf: 166, tf: 11.8, tw: 6.7, r1: 11.4 },
    '310UB40.4': { d: 304, bf: 165, tf: 10.2, tw: 6.1, r1: 11.4 },
    '310UB32.0': { d: 298, bf: 149, tf: 8.0, tw: 5.5, r1: 13 },
    '250UB37.3': { d: 256, bf: 146, tf: 10.9, tw: 6.4, r1: 8.9 },
    '250UB31.4': { d: 252, bf: 146, tf: 8.6, tw: 6.1, r1: 8.9 },
    '250UB25.7': { d: 248, bf: 124, tf: 8.0, tw: 5.0, r1: 12 },
    '200UB29.8': { d: 207, bf: 134, tf: 9.6, tw: 6.3, r1: 8.9 },
    '200UB25.4': { d: 203, bf: 133, tf: 7.8, tw: 5.8, r1: 8.9 },
    '200UB22.3': { d: 202, bf: 133, tf: 7.0, tw: 5.0, r1: 8.9 },
    '200UB18.2': { d: 198, bf: 99, tf: 7.0, tw: 4.5, r1: 11.4 },
    '180UB22.2': { d: 179, bf: 90, tf: 10.0, tw: 6.0, r1: 8.9 },
    '180UB18.1': { d: 175, bf: 90, tf: 7.9, tw: 4.9, r1: 8.9 },
    '180UB16.1': { d: 173, bf: 90, tf: 7.0, tw: 4.5, r1: 8.9 },
    '150UB18.0': { d: 155, bf: 75, tf: 9.5, tw: 6.0, r1: 8 },
    '150UB14.0': { d: 150, bf: 75, tf: 7.0, tw: 5.0, r1: 8 }
  },
  UC: {
    '310UC158': { d: 327, bf: 311, tf: 25.0, tw: 15.7, r1: 16.5 },
    '310UC137': { d: 321, bf: 309, tf: 21.7, tw: 13.8, r1: 16.5 },
    '310UC118': { d: 315, bf: 307, tf: 18.7, tw: 11.9, r1: 16.5 },
    '310UC96.8': { d: 308, bf: 305, tf: 15.4, tw: 9.9, r1: 16.5 },
    '250UC89.5': { d: 260, bf: 256, tf: 17.3, tw: 10.5, r1: 12.7 },
    '250UC72.9': { d: 254, bf: 254, tf: 14.2, tw: 8.6, r1: 12.7 },
    '200UC59.5': { d: 210, bf: 205, tf: 14.2, tw: 9.3, r1: 11.4 },
    '200UC52.2': { d: 206, bf: 204, tf: 12.5, tw: 8.0, r1: 11.4 },
    '200UC46.2': { d: 203, bf: 203, tf: 11.0, tw: 7.3, r1: 11.4 },
    '150UC37.2': { d: 162, bf: 154, tf: 11.5, tw: 8.1, r1: 8.9 },
    '150UC30.0': { d: 158, bf: 153, tf: 9.4, tw: 6.6, r1: 8.9 },
    '150UC23.4': { d: 152, bf: 152, tf: 6.8, tw: 6.1, r1: 8.9 },
    '100UC14.8': { d: 97, bf: 99, tf: 7.0, tw: 5.0, r1: 8.9 }
  },
  PFC: {
    '380PFC': { d: 380, bf: 100, tf: 17.5, tw: 10.0, r1: 14 },
    '300PFC': { d: 300, bf: 90, tf: 16.0, tw: 8.0, r1: 14 },
    '250PFC': { d: 250, bf: 90, tf: 15.0, tw: 8.0, r1: 12 },
    '200PFC': { d: 200, bf: 75, tf: 12.0, tw: 7.0, r1: 12 },
    '150PFC': { d: 150, bf: 75, tf: 9.5, tw: 6.0, r1: 10 }
  },
  EA: {
    '200x200x26EA': { d: 200, bf: 200, t: 26, r1: 18 },
    '200x200x20EA': { d: 200, bf: 200, t: 20, r1: 18 },
    '200x200x16EA': { d: 200, bf: 200, t: 16, r1: 18 },
    '150x150x19EA': { d: 150, bf: 150, t: 19, r1: 12 },
    '150x150x16EA': { d: 150, bf: 150, t: 16, r1: 12 },
    '150x150x12EA': { d: 150, bf: 150, t: 12, r1: 12 },
    '150x150x10EA': { d: 150, bf: 150, t: 10, r1: 12 },
    '125x125x16EA': { d: 125, bf: 125, t: 16, r1: 10 },
    '125x125x12EA': { d: 125, bf: 125, t: 12, r1: 10 },
    '100x100x12EA': { d: 100, bf: 100, t: 12, r1: 8 },
    '100x100x10EA': { d: 100, bf: 100, t: 10, r1: 8 },
    '100x100x8EA': { d: 100, bf: 100, t: 8, r1: 8 },
    '75x75x10EA': { d: 75, bf: 75, t: 10, r1: 6 },
    '75x75x8EA': { d: 75, bf: 75, t: 8, r1: 6 },
    '75x75x6EA': { d: 75, bf: 75, t: 6, r1: 6 }
  },
  UA: {
    '150x100x12UA': { d: 150, bf: 100, t: 12, r1: 12 },
    '150x100x10UA': { d: 150, bf: 100, t: 10, r1: 12 },
    '125x75x12UA': { d: 125, bf: 75, t: 12, r1: 10 },
    '125x75x10UA': { d: 125, bf: 75, t: 10, r1: 10 },
    '100x75x10UA': { d: 100, bf: 75, t: 10, r1: 8 },
    '100x75x8UA': { d: 100, bf: 75, t: 8, r1: 8 },
    '75x50x8UA': { d: 75, bf: 50, t: 8, r1: 6 },
    '75x50x6UA': { d: 75, bf: 50, t: 6, r1: 6 }
  },
  SHS: {
    '400x400x16SHS': { d: 400, bf: 400, t: 16, rExt: 40 },
    '400x400x12.5SHS': { d: 400, bf: 400, t: 12.5, rExt: 31.25 },
    '350x350x16SHS': { d: 350, bf: 350, t: 16, rExt: 40 },
    '350x350x12.5SHS': { d: 350, bf: 350, t: 12.5, rExt: 31.25 },
    '300x300x16SHS': { d: 300, bf: 300, t: 16, rExt: 40 },
    '300x300x12.5SHS': { d: 300, bf: 300, t: 12.5, rExt: 31.25 },
    '300x300x10SHS': { d: 300, bf: 300, t: 10, rExt: 25 },
    '250x250x16SHS': { d: 250, bf: 250, t: 16, rExt: 40 },
    '250x250x12.5SHS': { d: 250, bf: 250, t: 12.5, rExt: 31.25 },
    '250x250x10SHS': { d: 250, bf: 250, t: 10, rExt: 25 },
    '250x250x8SHS': { d: 250, bf: 250, t: 8, rExt: 20 },
    '200x200x12.5SHS': { d: 200, bf: 200, t: 12.5, rExt: 31.25 },
    '200x200x10SHS': { d: 200, bf: 200, t: 10, rExt: 25 },
    '200x200x8SHS': { d: 200, bf: 200, t: 8, rExt: 20 },
    '200x200x6SHS': { d: 200, bf: 200, t: 6, rExt: 15 },
    '150x150x10SHS': { d: 150, bf: 150, t: 10, rExt: 25 },
    '150x150x8SHS': { d: 150, bf: 150, t: 8, rExt: 20 },
    '150x150x6SHS': { d: 150, bf: 150, t: 6, rExt: 15 },
    '150x150x5SHS': { d: 150, bf: 150, t: 5, rExt: 12.5 },
    '100x100x8SHS': { d: 100, bf: 100, t: 8, rExt: 20 },
    '100x100x6SHS': { d: 100, bf: 100, t: 6, rExt: 15 },
    '100x100x5SHS': { d: 100, bf: 100, t: 5, rExt: 12.5 }
  },
  RHS: {
    '400x300x16RHS': { d: 400, bf: 300, t: 16, rExt: 40 },
    '400x300x12.5RHS': { d: 400, bf: 300, t: 12.5, rExt: 31.25 },
    '400x200x16RHS': { d: 400, bf: 200, t: 16, rExt: 40 },
    '400x200x12.5RHS': { d: 400, bf: 200, t: 12.5, rExt: 31.25 },
    '350x250x16RHS': { d: 350, bf: 250, t: 16, rExt: 40 },
    '350x250x12.5RHS': { d: 350, bf: 250, t: 12.5, rExt: 31.25 },
    '300x200x16RHS': { d: 300, bf: 200, t: 16, rExt: 40 },
    '300x200x12.5RHS': { d: 300, bf: 200, t: 12.5, rExt: 31.25 },
    '300x200x10RHS': { d: 300, bf: 200, t: 10, rExt: 25 },
    '250x150x12.5RHS': { d: 250, bf: 150, t: 12.5, rExt: 31.25 },
    '250x150x10RHS': { d: 250, bf: 150, t: 10, rExt: 25 },
    '250x150x8RHS': { d: 250, bf: 150, t: 8, rExt: 20 },
    '200x100x10RHS': { d: 200, bf: 100, t: 10, rExt: 25 },
    '200x100x8RHS': { d: 200, bf: 100, t: 8, rExt: 20 },
    '200x100x6RHS': { d: 200, bf: 100, t: 6, rExt: 15 },
    '150x100x10RHS': { d: 150, bf: 100, t: 10, rExt: 25 },
    '150x100x8RHS': { d: 150, bf: 100, t: 8, rExt: 20 },
    '150x100x6RHS': { d: 150, bf: 100, t: 6, rExt: 15 }
  },
  CHS: {
    '457x12.7CHS': { od: 457, t: 12.7 },
    '457x9.5CHS': { od: 457, t: 9.5 },
    '406.4x12.7CHS': { od: 406.4, t: 12.7 },
    '406.4x9.5CHS': { od: 406.4, t: 9.5 },
    '355.6x12.7CHS': { od: 355.6, t: 12.7 },
    '355.6x9.5CHS': { od: 355.6, t: 9.5 },
    '323.9x12.7CHS': { od: 323.9, t: 12.7 },
    '323.9x9.5CHS': { od: 323.9, t: 9.5 },
    '273.1x9.3CHS': { od: 273.1, t: 9.3 },
    '273.1x6.4CHS': { od: 273.1, t: 6.4 },
    '219.1x8.2CHS': { od: 219.1, t: 8.2 },
    '219.1x6.4CHS': { od: 219.1, t: 6.4 },
    '168.3x7.1CHS': { od: 168.3, t: 7.1 },
    '168.3x5.0CHS': { od: 168.3, t: 5.0 },
    '139.7x5.4CHS': { od: 139.7, t: 5.4 },
    '114.3x5.4CHS': { od: 114.3, t: 5.4 },
    '114.3x3.6CHS': { od: 114.3, t: 3.6 },
    '88.9x5.5CHS': { od: 88.9, t: 5.5 },
    '88.9x3.2CHS': { od: 88.9, t: 3.2 }
  }
};

// ============================================================================
// SECTION LOOKUP FUNCTION
// ============================================================================

/**
 * lookupSteelSection(sizeStr)
 * Parses a size string and returns exact properties from the catalogue,
 * or falls back to ratio-based estimation if not found.
 *
 * @param {string} sizeStr - e.g., "360UB56.7", "250x250x10SHS", "457x12.7CHS"
 * @returns {object} { type, d, bf, tf?, tw?, t?, r1?, rExt?, od?, source }
 */
function lookupSteelSection(sizeStr) {
  if (!sizeStr) return null;

  // Try each section type
  for (const [sectionType, sections] of Object.entries(STEEL_CATALOGUE)) {
    if (sections[sizeStr]) {
      const props = sections[sizeStr];
      return {
        type: sectionType,
        ...props,
        source: 'catalogue'
      };
    }
  }

  // Fallback: estimate based on ratios for common section types
  // UB/UC pattern: "360UB56.7" or "310UC96.8"
  const ubucMatch = sizeStr.match(/^(\d+)(UB|UC)([\d.]+)?$/i);
  if (ubucMatch) {
    const d = parseInt(ubucMatch[1]);
    const typeCode = ubucMatch[2].toUpperCase();
    if (typeCode === 'UC') {
      return { type: 'UC', d, bf: d * 0.95, tf: d * 0.055, tw: d * 0.035, r1: d * 0.04, source: 'estimated' };
    }
    return { type: 'UB', d, bf: d * 0.48, tf: d * 0.035, tw: d * 0.022, r1: d * 0.03, source: 'estimated' };
  }

  // PFC pattern: "250PFC"
  const pfcMatch = sizeStr.match(/^(\d+)PFC$/i);
  if (pfcMatch) {
    const d = parseInt(pfcMatch[1]);
    return { type: 'PFC', d, bf: d * 0.35, tf: d * 0.05, tw: d * 0.03, r1: d * 0.04, source: 'estimated' };
  }

  // SHS pattern: "200x200x10SHS"
  const shsMatch = sizeStr.match(/^(\d+)x(\d+)x([\d.]+)SHS$/i);
  if (shsMatch) {
    const d = parseInt(shsMatch[1]);
    const t = parseFloat(shsMatch[3]);
    return { type: 'SHS', d, bf: d, t, rExt: t * 2.5, source: 'estimated' };
  }

  // RHS pattern: "300x200x10RHS"
  const rhsMatch = sizeStr.match(/^(\d+)x(\d+)x([\d.]+)RHS$/i);
  if (rhsMatch) {
    const d = parseInt(rhsMatch[1]);
    const bf = parseInt(rhsMatch[2]);
    const t = parseFloat(rhsMatch[3]);
    return { type: 'RHS', d, bf, t, rExt: t * 2.5, source: 'estimated' };
  }

  // CHS pattern: "219.1x6.4CHS"
  const chsMatch = sizeStr.match(/^([\d.]+)x([\d.]+)CHS$/i);
  if (chsMatch) {
    return { type: 'CHS', od: parseFloat(chsMatch[1]), t: parseFloat(chsMatch[2]), source: 'estimated' };
  }

  // EA pattern: "100x100x8EA"
  const eaMatch = sizeStr.match(/^(\d+)x(\d+)x([\d.]+)EA$/i);
  if (eaMatch) {
    return { type: 'EA', d: parseInt(eaMatch[1]), bf: parseInt(eaMatch[2]), t: parseFloat(eaMatch[3]), r1: parseFloat(eaMatch[3]) * 1.2, source: 'estimated' };
  }

  // UA pattern: "150x100x10UA"
  const uaMatch = sizeStr.match(/^(\d+)x(\d+)x([\d.]+)UA$/i);
  if (uaMatch) {
    return { type: 'UA', d: parseInt(uaMatch[1]), bf: parseInt(uaMatch[2]), t: parseFloat(uaMatch[3]), r1: parseFloat(uaMatch[3]) * 1.2, source: 'estimated' };
  }

  return null;
}

// ============================================================================
// DRAWING FUNCTIONS
// ============================================================================

/**
 * drawSteelSection(ctx, cx, cy, profile, scale)
 * Dispatcher that routes to the correct drawing function.
 *
 * @param {CanvasRenderingContext2D} ctx
 * @param {number} cx - centre X coordinate (px)
 * @param {number} cy - centre Y coordinate at BOTTOM of section (beam soffit)
 * @param {object} profile - section properties from lookupSteelSection()
 * @param {number} scale - mm to px conversion factor
 */
function drawSteelSection(ctx, cx, cy, profile, scale) {
  if (!profile || !profile.type) return;

  switch (profile.type) {
    case 'UB':
    case 'UC':
      drawUBUCSection(ctx, cx, cy, profile, scale);
      break;
    case 'PFC':
      drawPFCSection(ctx, cx, cy, profile, scale);
      break;
    case 'EA':
      drawEASection(ctx, cx, cy, profile, scale);
      break;
    case 'UA':
      drawUASection(ctx, cx, cy, profile, scale);
      break;
    case 'SHS':
    case 'RHS':
      drawSHSRHSSection(ctx, cx, cy, profile, scale);
      break;
    case 'CHS':
      drawCHSSection(ctx, cx, cy, profile, scale);
      break;
    default:
      console.warn('Unknown section type:', profile.type);
  }
}

/**
 * drawUBUCSection(ctx, cx, cy, profile, scale)
 * Draws I-section (Universal Beam or Universal Column) with root fillets.
 * Section centred horizontally and vertically at (cx, cy).
 */
function drawUBUCSection(ctx, cx, cy, profile, scale) {
  const { d, bf, tf, tw, r1 } = profile;
  if (!d || !bf || !tf || !tw) return;

  // Convert to pixels with minimum sizes for visibility
  const dPx = Math.max(d * scale, 8);
  const bfPx = Math.max(bf * scale, 6);
  const tfPx = Math.max(tf * scale, 1.5);
  const twPx = Math.max(tw * scale, 2);     // Min 2px web so it's always visible
  const r1Px = Math.min(
    Math.max((r1 || 0) * scale, 0),
    (bfPx / 2 - twPx / 2) * 0.8,
    tfPx * 0.8
  );

  // cy = centre of section
  const top = cy - dPx / 2;
  const bot = cy + dPx / 2;
  const tfi = top + tfPx;           // inner face of top flange
  const bfi = bot - tfPx;           // inner face of bottom flange
  const wl = cx - twPx / 2;         // web left face
  const wr = cx + twPx / 2;         // web right face
  const fl = cx - bfPx / 2;         // flange left edge
  const fr = cx + bfPx / 2;         // flange right edge

  ctx.save();

  // ── Fill: draw three rectangles (flanges + web) for reliable fill ──
  ctx.fillStyle = 'rgba(200, 210, 220, 0.35)';

  // Top flange
  ctx.fillRect(fl, top, bfPx, tfPx);
  // Web
  ctx.fillRect(wl, tfi, twPx, bfi - tfi);
  // Bottom flange
  ctx.fillRect(fl, bfi, bfPx, tfPx);

  // ── Stroke: single outline path with root fillets ──
  ctx.strokeStyle = '#333333';
  ctx.lineWidth = 1.5;
  ctx.beginPath();

  // Clockwise from top-left of top flange
  ctx.moveTo(fl, top);
  ctx.lineTo(fr, top);                         // top edge
  ctx.lineTo(fr, tfi);                         // right side of top flange down

  if (r1Px > 0.5) {
    // Inner face of top flange toward web
    ctx.lineTo(wr + r1Px, tfi);
    // Root fillet: top-right
    ctx.arc(wr + r1Px, tfi + r1Px, r1Px, -Math.PI / 2, Math.PI, true);
  } else {
    ctx.lineTo(wr, tfi);
  }

  // Down right side of web
  if (r1Px > 0.5) {
    ctx.lineTo(wr, bfi - r1Px);
    // Root fillet: bottom-right
    ctx.arc(wr + r1Px, bfi - r1Px, r1Px, Math.PI, Math.PI / 2, true);
  } else {
    ctx.lineTo(wr, bfi);
  }

  ctx.lineTo(fr, bfi);                         // inner face of bottom flange right
  ctx.lineTo(fr, bot);                         // right side of bottom flange down
  ctx.lineTo(fl, bot);                         // bottom edge
  ctx.lineTo(fl, bfi);                         // left side of bottom flange up

  if (r1Px > 0.5) {
    ctx.lineTo(wl - r1Px, bfi);
    // Root fillet: bottom-left
    ctx.arc(wl - r1Px, bfi - r1Px, r1Px, Math.PI / 2, 0, true);
  } else {
    ctx.lineTo(wl, bfi);
  }

  // Up left side of web
  if (r1Px > 0.5) {
    ctx.lineTo(wl, tfi + r1Px);
    // Root fillet: top-left
    ctx.arc(wl - r1Px, tfi + r1Px, r1Px, 0, -Math.PI / 2, true);
  } else {
    ctx.lineTo(wl, tfi);
  }

  ctx.lineTo(fl, tfi);                         // inner face of top flange left
  ctx.closePath();
  ctx.stroke();

  // ── Root fillet fills (darker triangular wedges) ──
  if (r1Px > 0.5) {
    ctx.fillStyle = 'rgba(180, 190, 200, 0.4)';
    // Top-right fillet fill
    ctx.beginPath();
    ctx.moveTo(wr, tfi);
    ctx.lineTo(wr + r1Px, tfi);
    ctx.arc(wr + r1Px, tfi + r1Px, r1Px, -Math.PI / 2, Math.PI, true);
    ctx.lineTo(wr, tfi);
    ctx.fill();
    // Top-left fillet fill
    ctx.beginPath();
    ctx.moveTo(wl, tfi);
    ctx.lineTo(wl - r1Px, tfi);
    ctx.arc(wl - r1Px, tfi + r1Px, r1Px, -Math.PI / 2, 0, false);
    ctx.lineTo(wl, tfi);
    ctx.fill();
    // Bottom-right fillet fill
    ctx.beginPath();
    ctx.moveTo(wr, bfi);
    ctx.lineTo(wr + r1Px, bfi);
    ctx.arc(wr + r1Px, bfi - r1Px, r1Px, Math.PI / 2, Math.PI, false);
    ctx.lineTo(wr, bfi);
    ctx.fill();
    // Bottom-left fillet fill
    ctx.beginPath();
    ctx.moveTo(wl, bfi);
    ctx.lineTo(wl - r1Px, bfi);
    ctx.arc(wl - r1Px, bfi - r1Px, r1Px, Math.PI / 2, 0, true);
    ctx.lineTo(wl, bfi);
    ctx.fill();
  }

  ctx.restore();
}

/**
 * drawPFCSection(ctx, cx, cy, profile, scale)
 * Draws Parallel Flange Channel (C-shape).
 * Web on left side, flanges extending right, open face to the left.
 */
function drawPFCSection(ctx, cx, cy, profile, scale) {
  const { d, bf, tf, tw, r1 } = profile;
  if (!d || !bf || !tf || !tw) return;

  const dPx = Math.max(d * scale, 8);
  const bfPx = Math.max(bf * scale, 4);
  const tfPx = Math.max(tf * scale, 1.5);
  const twPx = Math.max(tw * scale, 2);
  const r1Px = Math.min((r1 || 0) * scale, (bfPx - twPx) * 0.8, tfPx * 0.8);

  // Channel centred at cx, cy. Web on LEFT, flanges extend RIGHT.
  const top = cy - dPx / 2;
  const bot = cy + dPx / 2;
  const tfi = top + tfPx;
  const bfi = bot - tfPx;
  const webL = cx - bfPx / 2;
  const webR = webL + twPx;
  const flangeR = cx + bfPx / 2;

  ctx.save();
  // Fill: three rects for reliability
  ctx.fillStyle = 'rgba(200, 210, 220, 0.35)';
  ctx.fillRect(webL, top, bfPx, tfPx);           // top flange
  ctx.fillRect(webL, tfi, twPx, bfi - tfi);      // web
  ctx.fillRect(webL, bfi, bfPx, tfPx);           // bottom flange

  // Stroke outline
  ctx.strokeStyle = '#333333';
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.moveTo(webL, top);
  ctx.lineTo(flangeR, top);
  ctx.lineTo(flangeR, tfi);
  if (r1Px > 0.5) {
    ctx.lineTo(webR + r1Px, tfi);
    ctx.arc(webR + r1Px, tfi + r1Px, r1Px, -Math.PI / 2, Math.PI, true);
  } else {
    ctx.lineTo(webR, tfi);
  }
  if (r1Px > 0.5) {
    ctx.lineTo(webR, bfi - r1Px);
    ctx.arc(webR + r1Px, bfi - r1Px, r1Px, Math.PI, Math.PI / 2, true);
  } else {
    ctx.lineTo(webR, bfi);
  }
  ctx.lineTo(flangeR, bfi);
  ctx.lineTo(flangeR, bot);
  ctx.lineTo(webL, bot);
  ctx.closePath();
  ctx.stroke();
  ctx.restore();
}

/**
 * drawEASection(ctx, cx, cy, profile, scale)
 * Draws Equal Angle (L-shape) with vertical leg = d, horizontal leg = bf.
 */
function drawEASection(ctx, cx, cy, profile, scale) {
  const { d, bf, t, r1 } = profile;
  if (!d || !bf || !t) return;

  const dPx = Math.max(d * scale, 6);
  const bfPx = Math.max(bf * scale, 6);
  const tPx = Math.max(t * scale, 2);
  const r1Px = Math.min((r1 || 0) * scale, tPx * 0.8);

  // For equal angles, legs share the same thickness
  // Position: bottom-left corner at the joint (cx, cy)
  // Vertical leg extends up (Y-), horizontal leg extends right (X+)

  ctx.save();
  ctx.fillStyle = 'rgba(200, 210, 220, 0.3)';
  ctx.strokeStyle = '#333333';
  ctx.lineWidth = 1.5;

  ctx.beginPath();

  // Start at top-left of vertical leg
  ctx.moveTo(cx, cy - dPx);
  // Top of vertical leg
  ctx.lineTo(cx + tPx, cy - dPx);
  // Right edge of vertical leg, down to the fillet
  ctx.lineTo(cx + tPx, cy - r1Px);

  // Inner fillet at bottom-right of vertical leg
  if (r1Px > 0) {
    ctx.arcTo(cx + tPx, cy - r1Px, cx + tPx + r1Px, cy, r1Px);
  }

  // Transition to horizontal leg
  ctx.lineTo(cx + bfPx, cy);
  // Right edge of horizontal leg
  ctx.lineTo(cx + bfPx, cy + tPx);
  // Bottom of horizontal leg, back toward joint
  ctx.lineTo(cx, cy + tPx);
  // Left edge back up
  ctx.lineTo(cx, cy - dPx);

  ctx.fill();
  ctx.stroke();
  ctx.restore();
}

/**
 * drawUASection(ctx, cx, cy, profile, scale)
 * Draws Unequal Angle (L-shape) with vertical leg = d, horizontal leg = bf.
 */
function drawUASection(ctx, cx, cy, profile, scale) {
  const { d, bf, t, r1 } = profile;
  if (!d || !bf || !t) return;

  const dPx = Math.max(d * scale, 6);
  const bfPx = Math.max(bf * scale, 4);
  const tPx = Math.max(t * scale, 2);
  const r1Px = Math.min((r1 || 0) * scale, tPx * 0.8);

  ctx.save();
  ctx.fillStyle = 'rgba(200, 210, 220, 0.3)';
  ctx.strokeStyle = '#333333';
  ctx.lineWidth = 1.5;

  ctx.beginPath();

  // Start at top-left of vertical leg
  ctx.moveTo(cx, cy - dPx);
  // Top of vertical leg
  ctx.lineTo(cx + tPx, cy - dPx);
  // Right edge of vertical leg down to fillet
  ctx.lineTo(cx + tPx, cy - r1Px);

  // Inner fillet at joint
  if (r1Px > 0) {
    ctx.arcTo(cx + tPx, cy - r1Px, cx + tPx + r1Px, cy, r1Px);
  }

  // Transition to horizontal leg (right edge)
  ctx.lineTo(cx + bfPx, cy);
  // Bottom of horizontal leg
  ctx.lineTo(cx + bfPx, cy + tPx);
  // Back toward vertical leg on horizontal bottom
  ctx.lineTo(cx, cy + tPx);
  // Up the left side
  ctx.lineTo(cx, cy - dPx);

  ctx.fill();
  ctx.stroke();
  ctx.restore();
}

/**
 * drawSHSRHSSection(ctx, cx, cy, profile, scale)
 * Draws Square or Rectangular Hollow Section.
 * Outer rectangle with rounded corners, inner rectangle with same corner radius.
 */
function drawSHSRHSSection(ctx, cx, cy, profile, scale) {
  const { d, bf, t, rExt } = profile;
  if (!d || !bf || !t) return;

  const dPx = Math.max(d * scale, 6);
  const bfPx = Math.max(bf * scale, 6);
  const tPx = Math.max(t * scale, 1.5);
  const rExtPx = Math.min((rExt || 0) * scale, dPx * 0.3, bfPx * 0.3);
  const rIntPx = Math.max(0, rExtPx - tPx);

  const halfDepth = dPx / 2;
  const halfWidth = bfPx / 2;

  ctx.save();
  ctx.fillStyle = 'rgba(200, 210, 220, 0.3)';
  ctx.strokeStyle = '#333333';
  ctx.lineWidth = 1.5;

  // Draw outer rounded rectangle
  ctx.beginPath();
  if (ctx.roundRect) {
    ctx.roundRect(
      cx - halfWidth, cy - halfDepth,
      bfPx, dPx,
      rExtPx
    );
  } else {
    // Fallback: manual arc approach for outer rectangle
    drawManualRoundRect(ctx, cx - halfWidth, cy - halfDepth, bfPx, dPx, rExtPx);
  }

  // Draw inner rounded rectangle (hole) as subtract
  if (ctx.roundRect) {
    ctx.roundRect(
      cx - halfWidth + tPx, cy - halfDepth + tPx,
      bfPx - 2 * tPx, dPx - 2 * tPx,
      Math.max(0, rExtPx - tPx),
      'evenodd'
    );
  } else {
    drawManualRoundRect(ctx, cx - halfWidth + tPx, cy - halfDepth + tPx, bfPx - 2 * tPx, dPx - 2 * tPx, rIntPx, true);
  }

  ctx.fill('evenodd');
  ctx.stroke();
  ctx.restore();
}

/**
 * drawCHSSection(ctx, cx, cy, profile, scale)
 * Draws Circular Hollow Section (annular ring).
 */
function drawCHSSection(ctx, cx, cy, profile, scale) {
  const { od, t } = profile;
  if (!od || !t) return;

  const odPx = Math.max(od * scale, 6);
  const tPx = Math.max(t * scale, 1.5);
  const rOuter = odPx / 2;
  const rInner = Math.max(1, rOuter - tPx);

  ctx.save();
  ctx.fillStyle = 'rgba(200, 210, 220, 0.35)';
  ctx.strokeStyle = '#333333';
  ctx.lineWidth = 1.5;

  // Draw annular ring using evenodd fill (outer CW, inner CCW)
  ctx.beginPath();
  ctx.arc(cx, cy, rOuter, 0, 2 * Math.PI, false);  // outer clockwise
  ctx.arc(cx, cy, rInner, 0, 2 * Math.PI, true);    // inner counter-clockwise
  ctx.fill('evenodd');

  // Stroke both circles
  ctx.beginPath();
  ctx.arc(cx, cy, rOuter, 0, 2 * Math.PI);
  ctx.stroke();

  ctx.beginPath();
  ctx.arc(cx, cy, rInner, 0, 2 * Math.PI);
  ctx.stroke();

  ctx.restore();
}

/**
 * Helper: Manual rounded rectangle drawing (fallback for ctx.roundRect)
 */
function drawManualRoundRect(ctx, x, y, width, height, radius, isHole = false) {
  const r = Math.min(radius, Math.min(width, height) / 2);

  if (!isHole) {
    ctx.beginPath();
  }

  ctx.moveTo(x + r, y);
  ctx.lineTo(x + width - r, y);
  ctx.arcTo(x + width, y, x + width, y + r, r);
  ctx.lineTo(x + width, y + height - r);
  ctx.arcTo(x + width, y + height, x + width - r, y + height, r);
  ctx.lineTo(x + r, y + height);
  ctx.arcTo(x, y + height, x, y + height - r, r);
  ctx.lineTo(x, y + r);
  ctx.arcTo(x, y, x + r, y, r);
  ctx.closePath();

  if (isHole) {
    ctx.closePath();
  }
}

// ============================================================================
// INITIALIZATION LOG
// ============================================================================

const _totalSections = Object.values(STEEL_CATALOGUE).reduce((sum, typeObj) => sum + Object.keys(typeObj).length, 0);
console.log(`[18a-steel-sections] LOADED — ${_totalSections} sections in catalogue`);
