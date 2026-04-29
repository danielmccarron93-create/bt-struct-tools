#!/usr/bin/env node
/**
 * Validate that the JS engine embedded in blockwork_wall_designer.html
 * produces the same outputs as as3700_walls.py for the test battery.
 *
 * Approach: extract the JS engine code from the HTML between the markers,
 * eval it in a sandbox, and run the same checks the Python validator runs.
 */
const fs = require('fs');
const path = require('path');

const HTML = fs.readFileSync(path.join(__dirname, 'blockwork_wall_designer.html'), 'utf-8');

// Extract engine code: from "AS 3700:2018 ENGINE" comment marker through "function checkMany"
const start = HTML.indexOf('// AS 3700:2018 ENGINE');
const end = HTML.indexOf('// ════════════════════════════════════════════════════════════════════════\n// UI WIRING');
if (start < 0 || end < 0) {
  console.error('Could not find engine markers in HTML');
  process.exit(1);
}
const engineSrc = HTML.slice(start, end);

// Eval into the global scope
eval(engineSrc);

// ─────────────────────────────────────────────────────────────────────────────
// Test runner mirrors validate_engine.py
// ─────────────────────────────────────────────────────────────────────────────
const PASSES = [];
const FAILS = [];
const TOL = 1e-3;

function approxEq(a, b, rel = TOL, absT = 1e-6) {
  if (Math.abs(b) < 1e-9) return Math.abs(a) < Math.max(absT, 1e-6);
  return Math.abs(a - b) / Math.abs(b) <= rel || Math.abs(a - b) <= absT;
}

function check(label, actual, expected, rel = TOL) {
  const ok = approxEq(actual, expected, rel);
  const status = ok ? 'PASS' : 'FAIL';
  const diff = Math.abs(expected) > 1e-9 ? Math.abs(actual - expected) / Math.abs(expected) * 100 : 0;
  const line = `  [${status}] ${label.padEnd(55)} actual=${String(actual.toFixed(4)).padStart(14)}  expected=${String(expected.toFixed(4)).padStart(14)}  diff=${diff.toFixed(3)}%`;
  console.log(line);
  (ok ? PASSES : FAILS).push(line);
}

function section(title) {
  console.log('\n' + '═'.repeat(110));
  console.log('  ' + title);
  console.log('═'.repeat(110));
}

// ─────────────────────────────────────────────────────────────────────────────
// Build inputs object factory
// ─────────────────────────────────────────────────────────────────────────────
function makeInputs(overrides = {}) {
  const base = {
    geometry: { L: 2000, t: 190, H: 5000, tfs: 30, grouting: 'full' },
    material: { fuc: 15, fcg: 20, unit_type: 'concrete', bedding: 'full', mortar: 'M3', kh: 1.3, fsy: 500, unit_density_high: true, kc_override: null },
    reinforcement: {
      vert_primary_code: "N16", vert_primary_spacing: 200,
      vert_additional_code: "-", vert_additional_spacing: 400,
      horiz_primary_code: "N12", horiz_primary_spacing: 600,
      horiz_additional_code: "-", horiz_additional_spacing: 600,
    },
    slenderness: { av: 0.85, ah: 2.5, kt: 1.0, edges_supported: 'one' },
    eccentricity: { e: 32 },
    method: {
      reinforced: true, alphar: 0.40, bending_method: 'kes_in_axial', sr_rule: 'standard',
      enforce_cl_8_8_cap: true, enforce_stability_check: false, enforce_eq_detailing: false,
      enforce_concentrated_loads: false, include_oop_bending: false, no_top_restraint: false,
    },
  };
  return { ...base, ...overrides };
}

// ─────────────────────────────────────────────────────────────────────────────
// Test 1: example wall (matches Python TEST 1)
// ─────────────────────────────────────────────────────────────────────────────
section('JS TEST 1 — Example wall: 190·2000·5000');
{
  const inputs = makeInputs();
  const lc = { name: 'LC1', Nc_static: 600, Nt: 0, Mx_top: 50, Vx_top: 80 };
  const r = checkLoadCase(inputs, lc);
  check('f\'m', computeFm(inputs.material), 7.0488);
  check('Sr1', r.compression.Sr1, 22.3684);
  check('Sr2', r.compression.Sr2, 16.9834);
  check('Fo basic per m', r.compression.Fo_basic_per_m_kN, 1498.6);
  check('kes', r.compression.kes, 0.2924);
  check('φNu.c per m', r.compression.phi_Nuc_per_m_kN, 371.55, 3e-3);
  check('φNu.c total', r.compression.phi_Nuc_total_kN, 743.10, 3e-3);
  check('Compression η', r.compression.utilisation, 0.807, 3e-3);
  check('φNt total', r.tension.phi_Nt_total_kN, 750.0, 2e-3);
  check('Long Vu', r.shear.long_branch.Vu_kN, 241.67, 3e-3);
  check('Short Vu raw', r.shear.short_branch.Vu_raw_kN, 351.33, 3e-3);
  check('φVu', r.shear.phi_Vu_kN, 263.5, 3e-3);
  check('Shear η', r.shear.utilisation, 0.304, 3e-3);
}

// ─────────────────────────────────────────────────────────────────────────────
// Test 2: Long wall shear (matches Python TEST 4)
// ─────────────────────────────────────────────────────────────────────────────
section('JS TEST 2 — Long wall (H/L < 2.3)');
{
  const inputs = makeInputs({
    geometry: { L: 4000, t: 190, H: 3000, tfs: 30, grouting: 'full' },
    reinforcement: {
      vert_primary_code: "N16", vert_primary_spacing: 400,
      vert_additional_code: "-", vert_additional_spacing: 400,
      horiz_primary_code: "N12", horiz_primary_spacing: 600,
      horiz_additional_code: "-", horiz_additional_spacing: 600,
    },
    slenderness: { av: 1.0, ah: 1.0, kt: 1.0, edges_supported: 'both' },
    eccentricity: { e: 10 },
  });
  const lc = { name: 'Long', Nc_static: 400, Vx_top: 200 };
  const r = checkLoadCase(inputs, lc);
  check('H/L', r.shear.H_over_L, 0.75);
  check('fvr', r.shear.long_branch.fvr, 1.125);
  check('As (mm²)', r.shear.long_branch.As_for_shear, 550.0, 2e-3);
  check('Vu (long)', r.shear.long_branch.Vu_kN, 1075.0, 2e-3);
  check('φVu', r.shear.phi_Vu_kN, 806.25, 2e-3);
  check('Shear η', r.shear.utilisation, 0.2481, 3e-3);
}

// ─────────────────────────────────────────────────────────────────────────────
// Test 3: Cl 8.8 cap (matches Python TEST 5)
// ─────────────────────────────────────────────────────────────────────────────
section('JS TEST 3 — Cl 8.8 cap actively governing');
{
  const inputs = makeInputs({
    geometry: { L: 1500, t: 190, H: 4500, tfs: 30, grouting: 'full' },
    reinforcement: {
      vert_primary_code: "N20", vert_primary_spacing: 200,
      vert_additional_code: "-", vert_additional_spacing: 200,
      horiz_primary_code: "N20", horiz_primary_spacing: 200,
      horiz_additional_code: "-", horiz_additional_spacing: 200,
    },
  });
  const lc = { name: 'Cap', Nc_static: 400, Vx_top: 200 };
  const r = checkLoadCase(inputs, lc);
  check('H/L', r.shear.H_over_L, 3.0);
  check('Cap', r.shear.short_branch.cap_kN, 399.0);
  check('cap governs?', r.shear.short_branch.cap_governs ? 1.0 : 0.0, 1.0);
  check('Vu after cap', r.shear.short_branch.Vu_kN, 399.0);
}

// ─────────────────────────────────────────────────────────────────────────────
// Test 4: kes bounds (matches Python TEST 2)
// ─────────────────────────────────────────────────────────────────────────────
section('JS TEST 4 — kes bounds');
check('kes(0,0)', kesFn(0, 0, 190), 1.0);
check('kes(40,0)', kesFn(40, 0, 190), 0.0);
check('kes(0,e/t=0.5)', kesFn(0, 95, 190), 0.0);
check('kes(20,e=19,t=190)', kesFn(20, 19, 190), 0.4);
check('kes negative clamped', kesFn(50, 0, 190), 0.0);

// ─────────────────────────────────────────────────────────────────────────────
// Test 5: αr behaviour (matches Python TEST 7)
// ─────────────────────────────────────────────────────────────────────────────
section('JS TEST 5 — αr=0 vs αr=0.40 capacity uplift');
{
  const inputs0 = makeInputs({ method: { ...makeInputs().method, alphar: 0.0 } });
  const inputs40 = makeInputs();
  const lc = { name: 'A', Nc_static: 600, Vx_top: 80 };
  const r0 = checkLoadCase(inputs0, lc);
  const r40 = checkLoadCase(inputs40, lc);
  check('Steel term @ αr=0', r0.compression.steel_term_N_per_m / 1000, 0.0);
  check('Steel term @ αr=0.4', r40.compression.steel_term_N_per_m / 1000, 200.0, 2e-3);
  const uplift = r40.compression.phi_Nuc_total_kN / r0.compression.phi_Nuc_total_kN - 1;
  console.log(`  [INFO] Uplift αr=0 → αr=0.4: ${(uplift * 100).toFixed(1)}%`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Test 6: Multi-LC (matches Python TEST 8)
// ─────────────────────────────────────────────────────────────────────────────
section('JS TEST 6 — Multi-LC governing pick');
{
  const inputs = makeInputs();
  const lcs = [
    { name: 'LC1', Nc_static: 400, Vx_top: 50 },
    { name: 'LC2', Nc_static: 600, Vx_top: 80, Mx_top: 50 },
    { name: 'LC3', Nc_static: 200, Vx_top: 120, Mx_top: 80 },
    { name: 'LC4', Nc_static: 350, Vx_top: 100, Mx_top: 70 },
  ];
  const res = checkMany(inputs, lcs);
  console.log(`  Governing: ${res.governing.lc_name}  η = ${res.governing.max_utilisation.toFixed(3)}`);
  check('Number of LCs', res.load_cases.length, 4);
  const maxU = Math.max(...res.load_cases.map(r => r.max_utilisation));
  check('Governing util = max', res.governing.max_utilisation, maxU);
}

// ─────────────────────────────────────────────────────────────────────────────
// Test 7: Static vs Area worst case
// ─────────────────────────────────────────────────────────────────────────────
section('JS TEST 7 — Static vs Area worst-case');
{
  const inputs = makeInputs();
  const lc1 = { name: 'X', Nc_static: 300, Nc_area: 500, Vx_top: 50 };
  const r1 = checkLoadCase(inputs, lc1);
  check('Nc_worst when Area > Static', r1.Nc_kN, 500);
  const lc2 = { name: 'Y', Nc_static: 600, Nc_area: 400, Vx_top: 50 };
  const r2 = checkLoadCase(inputs, lc2);
  check('Nc_worst when Static > Area', r2.Nc_kN, 600);
}

// ─────────────────────────────────────────────────────────────────────────────
// Test 8: Cl 8.6 explicit flexure (matches Python TEST 12)
// ─────────────────────────────────────────────────────────────────────────────
section('JS TEST 8 — Cl 8.6 explicit flexure');
{
  const inputs = makeInputs({
    method: { ...makeInputs().method, bending_method: 'explicit_8_6' }
  });
  const lc = { name: 'Flex', Nc_static: 100, Mx_top: 200 };
  const r = checkLoadCase(inputs, lc);
  const f = r.flexure_inplane;
  check('Asd_used', f.Asd_used, 1817.27, 5e-3);
  check('Mu', f.Mu_kNm, 1351.4, 5e-3);
  check('φMu', f.phi_Mu_kNm, 1013.5, 5e-3);
}

// ─────────────────────────────────────────────────────────────────────────────
// Test 9: REO_TABLE — EF, additional bars, mesh
// ─────────────────────────────────────────────────────────────────────────────
section('JS TEST 9 — REO_TABLE: EF, additional bars, mesh');
{
  const baseRest = {
    geometry: { L: 2000, t: 190, H: 5000, tfs: 30, grouting: 'full' },
    material: makeInputs().material,
    slenderness: makeInputs().slenderness,
    eccentricity: makeInputs().eccentricity,
    method: makeInputs().method,
  };
  const mk = (reo) => ({ ...baseRest, reinforcement: reo });
  const r1 = checkLoadCase(mk({
    vert_primary_code: "N16", vert_primary_spacing: 200,
    vert_additional_code: "-", vert_additional_spacing: 200,
    horiz_primary_code: "N12", horiz_primary_spacing: 600,
    horiz_additional_code: "-", horiz_additional_spacing: 600,
  }), { name: 'X', Nc_static: 100 });
  check('Asv N16@200', r1.compression.As_per_m, 1000.0);
  const r2 = checkLoadCase(mk({
    vert_primary_code: "N16 EF", vert_primary_spacing: 200,
    vert_additional_code: "-", vert_additional_spacing: 200,
    horiz_primary_code: "N12", horiz_primary_spacing: 600,
    horiz_additional_code: "-", horiz_additional_spacing: 600,
  }), { name: 'X', Nc_static: 100 });
  check('Asv N16 EF doubles', r2.compression.As_per_m, 2000.0);
  const r3 = checkLoadCase(mk({
    vert_primary_code: "N16", vert_primary_spacing: 400,
    vert_additional_code: "N12", vert_additional_spacing: 400,
    horiz_primary_code: "N12", horiz_primary_spacing: 600,
    horiz_additional_code: "-", horiz_additional_spacing: 600,
  }), { name: 'X', Nc_static: 100 });
  check('Asv primary+additional', r3.compression.As_per_m, 775.0);
  const r4 = checkLoadCase(mk({
    vert_primary_code: "SL82", vert_primary_spacing: 999,
    vert_additional_code: "-", vert_additional_spacing: 200,
    horiz_primary_code: "SL82", horiz_primary_spacing: 42,
    horiz_additional_code: "-", horiz_additional_spacing: 600,
  }), { name: 'X', Nc_static: 100 });
  check('Asv SL82 (mesh, ignores spacing)', r4.compression.As_per_m, 227.0);
  const r5 = checkLoadCase(mk({
    vert_primary_code: "SL82 EF", vert_primary_spacing: 200,
    vert_additional_code: "-", vert_additional_spacing: 200,
    horiz_primary_code: "-", horiz_primary_spacing: 600,
    horiz_additional_code: "-", horiz_additional_spacing: 600,
  }), { name: 'X', Nc_static: 100 });
  check('Asv SL82 EF (corrected)', r5.compression.As_per_m, 454.0);
  const r6 = checkLoadCase(mk({
    vert_primary_code: "-", vert_primary_spacing: 200,
    vert_additional_code: "-", vert_additional_spacing: 200,
    horiz_primary_code: "-", horiz_primary_spacing: 600,
    horiz_additional_code: "-", horiz_additional_spacing: 600,
  }), { name: 'X', Nc_static: 100 });
  check('Asv "-"', r6.compression.As_per_m, 0.0);
}

// ─────────────────────────────────────────────────────────────────────────────
// Summary
// ─────────────────────────────────────────────────────────────────────────────
console.log('\n' + '═'.repeat(110));
console.log(`  JS PORT VALIDATION SUMMARY: ${PASSES.length} PASSED, ${FAILS.length} FAILED`);
console.log('═'.repeat(110));
if (FAILS.length) {
  console.log('\nFAILURES:');
  FAILS.forEach(f => console.log(f));
  process.exit(1);
}
process.exit(0);
