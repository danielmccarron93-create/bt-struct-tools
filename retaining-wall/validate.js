// Validation harness — RCB 3e Worked Example 6.3 (cantilever wall, 4 m, gravelly clay).
// Reproduces the engine's core computations standalone for verification.

function ka(phiDeg, betaDeg) {
  const phi = phiDeg * Math.PI / 180;
  const beta = (betaDeg || 0) * Math.PI / 180;
  if (Math.abs(beta) < 1e-9) {
    const t = Math.tan(Math.PI/4 - phi/2);
    return t * t;
  }
  const cb = Math.cos(beta), cp = Math.cos(phi);
  const sq = Math.sqrt(Math.max(0, cb*cb - cp*cp));
  return cb * (cb - sq) / (cb + sq);
}
function bearingPressure(N, M, L, W=1) {
  const e = Math.abs(M) / N;
  const kern = L / 6;
  if (e <= kern + 1e-9) {
    const qAvg = N / (W*L);
    const qBend = 6 * Math.abs(M) / (W*L*L);
    return { qmax: qAvg + qBend, qmin: qAvg - qBend, e, kern, mode: 'trapezoidal',
             contactLength: L, qAtX: x => qAvg + qBend - 2*qBend*x/L };
  } else {
    const a = L/2 - e;
    const contact = 3*a;
    const qmax = 2*N/(W*contact);
    return { qmax, qmin: 0, e, kern, mode: 'triangular', contactLength: contact,
             qAtX: x => x <= contact ? qmax*(1 - x/contact) : 0 };
  }
}

// ── Example 6.3 inputs ──
const H = 4.0, t = 0.3, D = 0.3, Lt = 0.6, Lh = 2.1;
const L = Lt + t + Lh;                    // 3.0 m
const beta = 0;
const gF = 17, s = 5, phiF = 30;
const phiB = 30, gB = 18, cB = 5, qa = 250;
const dRatio = 0.75;
const delta = phiB * dRatio;              // 22.5°
const gC = 25;
const fc = 25, fsy = 500, coverW = 50, coverB = 75;
const Ka = ka(phiF, beta);                // expect 0.333

const Ht = H + D;                         // 4.3 m

const W1 = t * H * gC;                    // 30
const W1_arm = Lt + t/2;                  // 0.75
const W2 = L * D * gC;                    // 22.5
const W2_arm = L/2;                       // 1.5
const W3 = Lh * H * gF;                   // 142.8
const W3_arm = Lt + t + Lh/2;             // 1.95
const Sv = s * Lh;                        // 10.5
const Sv_arm = Lt + t + Lh/2;             // 1.95
const PA1 = 0.5 * Ka * gF * Ht * Ht;
const PA2 = Ka * s * Ht;
const PA1_h = Ht / 3;
const PA2_h = Ht / 2;
const R = W1 + W2 + W3 + Sv;              // 205.8

// Moments about toe (book includes Sv in M_stb total)
const M_stb_total = W1*W1_arm + W2*W2_arm + W3*W3_arm + Sv*Sv_arm;
const M_dst_e = PA1 * PA1_h;
const M_dst_q = PA2 * PA2_h;

const gStb = 0.9, gE = 1.5, psiC = 1.0;
const M_stb_factored = gStb * M_stb_total;
const M_dst_factored = gE * M_dst_e + psiC * M_dst_q;
const overturnUtil = M_dst_factored / M_stb_factored;

// Bearing — moment about base CENTROID (sign: + tips toward toe)
const M_centroid_v = (W1*(L/2 - W1_arm) + W2*(L/2 - W2_arm)
                     + W3*(L/2 - W3_arm) + Sv*(L/2 - Sv_arm));
const M_centroid_h = PA1 * PA1_h + PA2 * PA2_h;
const M_centroid = M_centroid_v + M_centroid_h;
const bp = bearingPressure(R, M_centroid, L, 1);

// Sliding (book: F_friction uses full R including surcharge)
const F_friction = R * Math.tan(delta * Math.PI/180);
const F_cohesion = cB * Math.min(L, bp.contactLength || L);
const F_stb = F_friction + F_cohesion;
const F_dst_factored = gE * PA1 + psiC * PA2;
const F_stb_factored = gStb * F_stb;
const slidingUtil = F_dst_factored / F_stb_factored;

// Stem flexure
const M_stem_e = Ka * gF * Math.pow(H,3) / 6;
const M_stem_q = Ka * s * H * H / 2;
const M_stem_serv = M_stem_e + M_stem_q;
const M_stem_star = 1.5 * M_stem_serv;

// Toe flexure
const q_toe_end = bp.qAtX(0);
const q_face_toe = bp.qAtX(Lt);
const M_toe_up = Lt * Lt * (q_face_toe + 2*q_toe_end) / 6;
const M_toe_base_down = gC * D * Lt * Lt / 2;
const M_toe_serv = Math.max(0, M_toe_up - M_toe_base_down);
const M_toe_star = 1.5 * M_toe_serv;

// ── REPORT ──
function fmt(label, value, expected, unit, tol) {
  const pass = expected == null ? '' :
    Math.abs(value - expected) <= (tol || Math.abs(expected) * 0.05) ? 'PASS' : 'FAIL';
  const expStr = expected == null ? '' : `(expected ~${expected})`;
  console.log(`${label.padEnd(30)} ${value.toFixed(2).padStart(10)} ${unit.padEnd(8)} ${expStr.padEnd(20)} ${pass}`);
}

console.log('=================================================================');
console.log(' RCB 3e Worked Example 6.3 — engine validation');
console.log('=================================================================');
fmt('k_a', Ka, 0.33, '', 0.01);
fmt('W1 stem', W1, 30, 'kN/m');
fmt('W2 base', W2, 22.5, 'kN/m');
fmt('W3 soil over heel', W3, 142.8, 'kN/m');
fmt('S surcharge over heel', Sv, 10.5, 'kN/m');
fmt('R = ΣV', R, 205.8, 'kN/m');
fmt('P_A1', PA1, 51.9, 'kN/m');
fmt('P_A2', PA2, 7.1, 'kN/m');
console.log('');
fmt('M about centroid', M_centroid, 43.2, 'kNm/m', 1.5);
fmt('Eccentricity e', bp.e, 0.21, 'm', 0.01);
fmt('q_max', bp.qmax, 97, 'kPa', 2);
fmt('q_min', bp.qmin, 40, 'kPa', 2);
console.log('');
console.log('Overturning:');
fmt('  LHS = 0.9*M_stb', M_stb_factored, 319, 'kNm/m', 5);
fmt('  RHS = 1.5*Me + ψcMq', M_dst_factored, 127, 'kNm/m', 3);
fmt('  Util', overturnUtil, 0.40, '', 0.05);
console.log('');
console.log('Sliding:');
fmt('  F_stb = R*tanδ + c*a', F_stb, 98.1, 'kN/m', 1.5);
fmt('  LHS = 0.9*F_stb', F_stb_factored, 88.3, 'kN/m', 1.5);
fmt('  RHS = 1.5*Pe + ψcPs', F_dst_factored, 85.0, 'kN/m', 1.5);
fmt('  Util', slidingUtil, 0.96, '', 0.05);
console.log('');
console.log('Components:');
fmt('Stem M_serv', M_stem_serv, 73, 'kNm/m', 5);
fmt('Stem M*', M_stem_star, 110, 'kNm/m', 5);
fmt('Toe M*', M_toe_star, 24.7, 'kNm/m', 3);
console.log('');
console.log('Bearing distribution:');
console.log('  q at toe (x=0):     ', bp.qAtX(0).toFixed(1), 'kPa');
console.log('  q at face of stem (toe-side, x=Lt):', bp.qAtX(Lt).toFixed(1), 'kPa');
console.log('  q at face of stem (heel-side, x=Lt+t):', bp.qAtX(Lt+t).toFixed(1), 'kPa');
console.log('  q at heel end (x=L):', bp.qAtX(L).toFixed(1), 'kPa');
