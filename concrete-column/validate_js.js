// Validate the JS port of the AS 3600 column engine against the same RCB cases.
// Extracts the engine code from the HTML file and runs the same checks.

const fs = require('fs');
const path = require('path');

// Read the HTML file
const htmlPath = path.join(__dirname, 'index.html');
const html = fs.readFileSync(htmlPath, 'utf8');

// Extract the engine section: from BAR_AREAS to "// ════════════════════════════════════════════════════════════════════════\n// STATE + UI WIRING"
const startMarker = 'const BAR_AREAS';
const endMarker = '// STATE + UI WIRING';
const startIdx = html.indexOf(startMarker);
const endIdx = html.indexOf(endMarker);
if (startIdx === -1 || endIdx === -1) { console.error('Markers not found'); process.exit(1); }
const engineCode = html.substring(startIdx, endIdx);

// Strip 'const'/'let' declarations at top level so they leak to global, then eval indirectly.
// Actually simpler: use Function() to expose names back.
const wrapped = engineCode + `\nreturn { Section, alpha1, alpha2, gammaCo, biaxialUtil, bucklingNc, deltaB, kmFactor, hscSimplified, hscDeemed, BAR_AREAS };\n`;
const lib = new Function(wrapped)();
const { Section, alpha1, alpha2, gammaCo, biaxialUtil, bucklingNc, deltaB, kmFactor, hscSimplified, hscDeemed, BAR_AREAS } = lib;

// ──────── Validation ────────
let pass = 0, fail = 0;
function check(name, computed, expected, tol = 0.01) {
  const err = Math.abs(computed - expected) / Math.abs(expected);
  const ok = err <= tol;
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name.padEnd(38)} computed=${computed.toFixed(4).padStart(12)}  expected=${expected.toFixed(4).padStart(12)}  err=${(err*100).toFixed(2)}%`);
  if (ok) pass++; else fail++;
  return ok;
}

console.log('\n=== EXAMPLE 5.2 — Section capacity line ===');
const sec52 = new Section({
  shape: 'rect', D: 600, b: 400, fc: 40, fsy: 500,
  layers: [{ d: 74, area: 1200 }, { d: 526, area: 1200 }],
});
check('α1', alpha1(40), 0.85);
check('α2', alpha2(40), 0.79);
check('γ',  gammaCo(40), 0.87);
check('Plastic centroid', sec52.plasticCentroid(), 300);
// AS 3600 Cl 10.6.2.2: net concrete area (Ag - As) and steel stress capped at
// 0.0025*Es = 500 MPa. Engine returns 9278 kN; RCB book quotes 9360 (uses Ag).
check('Nuo (squash, AS 3600)', sec52.nuo() / 1000, 9278);
const decomp52 = sec52.decompressionPoint();
check('B: Nu @ ku=1', decomp52[0]/1000, 6385);
check('B: Mu @ ku=1', decomp52[1]/1e6, 547);
const bp52 = sec52.balancedPoint();
check('C: Nub', bp52.Nu/1000, 3090);
check('C: Mub', bp52.Mu/1e6, 809);
check('kub', bp52.kub, 0.545);
const pb52 = sec52.pureBending();
check('D: Muo', pb52.Mu/1e6, 302, 0.02);
check('E: Nuot', sec52.nuot()/1000, 1200);

console.log('\n=== EXAMPLE 5.4 — Biaxial ===');
const sec54 = new Section({
  shape:'rect', D:600, b:400, fc:40, fsy:500,
  layers: [{d:74, area:2*615.8}, {d:526, area:2*615.8}],
});
const Nuo54 = sec54.nuo()/1000;
const ut54 = biaxialUtil(280, 140, 390, 230, 4000, 9360);
check('αn (AS 3600 Cl 10.6.4)', ut54.an, 0.7 + 1.7*4000/9360);

console.log('\n=== EXAMPLE 5.5 — Slender column ===');
const sec55 = new Section({
  shape:'rect', D:600, b:400, fc:40, fsy:500,
  layers: [{d:74, area:1200}, {d:526, area:1200}],
});
const Nc55 = bucklingNc(sec55, 6340, 0.7) / 1000;
check('Nc interior (AS 3600)', Nc55, 7272, 0.01);
const Nc55ext = bucklingNc(sec55, 8640, 0.7) / 1000;
check('Nc exterior (AS 3600)', Nc55ext, 3916, 0.01);
check('km', kmFactor(-270, 360), 0.90);
check('δb (book Nc=6706)', deltaB(0.9, 2000, 6706), 1.28, 0.02);

console.log('\n=== EXAMPLE 5.6 — HSC (high axial) ===');
const inp56 = {
  shape:'rect', bc:628, dc:628, n:12, w:152, AbFit:110,
  fsyF:500, fc:80, m:4, ds:628, thetaDeg:90,
};
check('Simplified s_max', hscSimplified(inp56), 249, 0.05);
check('Deemed s_max', hscDeemed(inp56), 197, 0.02);

console.log('\n=== EXAMPLE 5.7 — HSC (moderate axial, high moment) ===');
const inp57 = {
  shape:'rect', bc:280, dc:280, n:4, w:230, AbFit:80,
  fsyF:500, fc:80, m:2, ds:280, thetaDeg:90,
};
check('Simplified s_max', hscSimplified(inp57), 121, 0.05);
check('Deemed s_max', hscDeemed(inp57), 107, 0.05);

console.log(`\n${pass} passed, ${fail} failed.`);
process.exit(fail === 0 ? 0 : 1);
