// ── FOUNDING STRATA MODULE ───────────────────────────────
// ══════════════════════════════════════════════════════════

const foundingStrata = {
    active: false,
    settings: {
        method: 'idw',          // 'tin' or 'idw'
        contourInterval: 0.2,   // metres
        showHeatmap: true,
        showContours: true,
        showTIN: false,
        heatmapOpacity: 0.10,
        socketDepth: 200,       // mm — depth to socket into rock
        contingency: 0,         // mm — extra buffer
    },
    // Computed cache
    _grid: null,
    _gridCols: 0,
    _gridRows: 0,
    _triangles: [],
    _minDepth: 0,
    _maxDepth: 1,
    _coords: [],
    _depths: [],
};

// ── Borehole Element Type ──
// Boreholes stored as project.elements with type: 'borehole'
// { id, type:'borehole', layer:'S-GEO', level:'GF', x, y, tag:'BH1',
//   depthToRock: 1.2,      // depth in metres below ground
//   groundRL: 13.29,       // optional — RL of existing ground at borehole location
//   rockRL: 12.09          // optional — RL of top of rock (alternative to depthToRock)
// }

let _bhNum = 1;

// ── Founding Strata Button ──
document.getElementById('btn-founding-strata').addEventListener('click', () => {
    foundingStrata.active = !foundingStrata.active;
    document.getElementById('btn-founding-strata').classList.toggle('active', foundingStrata.active);
    document.getElementById('founding-panel').classList.toggle('hidden', !foundingStrata.active);
    if (foundingStrata.active) {
        setActiveTool('borehole');
        updateFoundingPanel();
    } else {
        if (activeTool === 'borehole') setActiveTool('select');
    }
    engine.requestRender();
});

// Stop canvas handlers from stealing clicks inside the founding panel
document.getElementById('founding-panel').addEventListener('mousedown', (e) => {
    e.stopPropagation();
});
document.getElementById('founding-panel').addEventListener('dblclick', (e) => {
    e.stopPropagation();
});

document.getElementById('fp-close').addEventListener('click', () => {
    foundingStrata.active = false;
    document.getElementById('btn-founding-strata').classList.remove('active');
    document.getElementById('founding-panel').classList.add('hidden');
    if (activeTool === 'borehole') setActiveTool('select');
    engine.requestRender();
});

// ── Borehole Placement ──
container.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'borehole') return;

    // Don't place boreholes when clicking inside the founding panel
    const fpPanel = document.getElementById('founding-panel');
    if (fpPanel && fpPanel.contains(e.target)) return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    const sheetPos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    const realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);

    // Find next available BH number
    const existingBHs = project.elements.filter(el => el.type === 'borehole');
    let nextNum = 1;
    while (existingBHs.some(bh => bh.tag === 'BH' + nextNum)) nextNum++;
    const tag = 'BH' + nextNum;

    const newBH = {
        id: generateId(),
        type: 'borehole',
        layer: 'S-GEO',
        level: 'GF',
        x: realPos.x,
        y: realPos.y,
        tag: tag,
        depthToRock: 0.5,  // default — overridden by prompt below
        groundRL: levelSystem.groundRL || undefined,  // snapshot ground RL at placement
    };

    history.execute({
        description: 'Add borehole: ' + tag,
        execute() { project.elements.push(newBH); },
        undo() {
            const i = project.elements.indexOf(newBH);
            if (i !== -1) project.elements.splice(i, 1);
        }
    });

    engine.requestRender();
    updateFoundingPanel();

    // Prompt for rock data — offer RL input when ground RL is set
    setTimeout(() => {
        const hasRL = levelSystem.groundRL !== 0;
        let promptMsg, defaultVal;
        if (hasRL) {
            promptMsg = `${tag} — Enter data for rock level.\n\n` +
                `Option 1: Depth to rock in metres (e.g. 1.5)\n` +
                `Option 2: RL of top of rock (prefix with "RL", e.g. RL11.80)\n\n` +
                `Ground RL at this location: ${(newBH.groundRL || levelSystem.groundRL).toFixed(3)}`;
            defaultVal = '0.5';
        } else {
            promptMsg = `${tag} — Enter depth to rock (m):`;
            defaultVal = '0.5';
        }
        const inputStr = prompt(promptMsg, defaultVal);
        if (inputStr !== null) {
            const trimmed = inputStr.trim();
            if (trimmed.toUpperCase().startsWith('RL')) {
                // RL-based input
                const rlVal = parseFloat(trimmed.substring(2).trim());
                if (!isNaN(rlVal)) {
                    newBH.rockRL = rlVal;
                    const grl = newBH.groundRL || levelSystem.groundRL;
                    newBH.depthToRock = Math.max(0, grl - rlVal);
                }
            } else {
                const d = parseFloat(trimmed);
                if (!isNaN(d) && d >= 0) {
                    newBH.depthToRock = d;
                    if (hasRL) {
                        const grl = newBH.groundRL || levelSystem.groundRL;
                        newBH.rockRL = grl - d;
                    }
                }
            }
        } else {
            // User cancelled — remove the borehole
            const idx = project.elements.indexOf(newBH);
            if (idx !== -1) project.elements.splice(idx, 1);
        }
        _lastGridHash = '';
        engine.requestRender();
        updateFoundingPanel();
    }, 10);
});

// ── Double-click borehole to edit depth ──
container.addEventListener('dblclick', (e) => {
    if (activeTool !== 'select' && activeTool !== 'borehole') return;
    if (!selectedElement || selectedElement.type !== 'borehole') return;

    const el = selectedElement;
    const hasRL = levelSystem.groundRL !== 0;
    const grl = el.groundRL || levelSystem.groundRL;
    let promptMsg;
    if (hasRL) {
        promptMsg = `Borehole ${el.tag}\n` +
            `Ground RL: ${grl.toFixed(3)}\n` +
            `Current: ${el.depthToRock.toFixed(2)}m deep` +
            (el.rockRL !== undefined ? ` (Rock RL ${el.rockRL.toFixed(3)})` : '') +
            `\n\nEnter depth (m) or RL of rock (prefix "RL"):`;
    } else {
        promptMsg = `Borehole ${el.tag}\nEnter depth to rock (m):`;
    }
    const inputStr = prompt(promptMsg, String(el.depthToRock));
    if (inputStr === null) return;
    const trimmed = inputStr.trim();
    if (trimmed.toUpperCase().startsWith('RL')) {
        const rlVal = parseFloat(trimmed.substring(2).trim());
        if (!isNaN(rlVal)) {
            el.rockRL = rlVal;
            el.depthToRock = Math.max(0, grl - rlVal);
        }
    } else {
        const val = parseFloat(trimmed);
        if (!isNaN(val) && val >= 0) {
            el.depthToRock = val;
            if (hasRL) el.rockRL = grl - val;
        }
    }
    _lastGridHash = '';
    engine.requestRender();
    updateFoundingPanel();
});

// ── Borehole Hit Testing ──
const _origHitTest2 = hitTestElement;
hitTestElement = function(sheetPos, tolerance) {
    // Only hit-test boreholes when founding strata is active
    if (!foundingStrata.active) return _origHitTest2(sheetPos, tolerance);
    // Check boreholes first
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'borehole') continue;
        const cp = engine.coords.realToSheet(el.x, el.y);
        const hitR = 3 / engine.viewport.zoom + tolerance; // generous hit area
        if (Math.abs(sheetPos.x - cp.x) < hitR && Math.abs(sheetPos.y - cp.y) < hitR) return el;
    }
    return _origHitTest2(sheetPos, tolerance);
};

// ── Borehole 2D Rendering ──
engine.onRender((ctx, eng) => {
    if (!foundingStrata.active) return;
    const zoom = eng.viewport.zoom;
    for (const el of project.getVisibleElements()) {
        if (el.type !== 'borehole') continue;
        const pos = eng.coords.realToScreen(el.x, el.y);
        const selected = isElementSelected(el);
        const color = selected ? '#2B7CD0' : '#8B4513';
        const r = Math.max(4, 8 * zoom);

        // Outer circle
        ctx.beginPath();
        ctx.arc(pos.x, pos.y, r, 0, Math.PI * 2);
        ctx.fillStyle = selected ? 'rgba(43,124,208,0.15)' : 'rgba(139,69,19,0.15)';
        ctx.fill();
        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(1, 1.5 * zoom);
        ctx.stroke();

        // Inner dot
        ctx.beginPath();
        ctx.arc(pos.x, pos.y, Math.max(1.5, 2.5 * zoom), 0, Math.PI * 2);
        ctx.fillStyle = color;
        ctx.fill();

        // Crosshairs
        const chLen = r * 0.6;
        ctx.beginPath();
        ctx.moveTo(pos.x - r, pos.y); ctx.lineTo(pos.x - chLen, pos.y);
        ctx.moveTo(pos.x + chLen, pos.y); ctx.lineTo(pos.x + r, pos.y);
        ctx.moveTo(pos.x, pos.y - r); ctx.lineTo(pos.x, pos.y - chLen);
        ctx.moveTo(pos.x, pos.y + chLen); ctx.lineTo(pos.x, pos.y + r);
        ctx.stroke();

        // Tag label above
        const tagFs = Math.max(1, 2.8 * zoom);
        ctx.font = `bold ${tagFs}px "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = color;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(el.tag, pos.x, pos.y - r - 2);

        // Depth label below
        const depthFs = Math.max(1, 2.2 * zoom);
        ctx.font = `${depthFs}px "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = '#666';
        ctx.textBaseline = 'top';
        ctx.fillText(el.depthToRock.toFixed(2) + 'm', pos.x, pos.y + r + 2);

        // Rock RL label (if available)
        if (el.rockRL !== undefined && levelSystem.groundRL !== 0) {
            const rlFs = Math.max(1, 1.8 * zoom);
            ctx.font = `500 ${rlFs}px "Consolas", monospace`;
            ctx.fillStyle = '#2B7CD0';
            ctx.fillText('Rock RL ' + el.rockRL.toFixed(3), pos.x, pos.y + r + 2 + depthFs + 1);
        }
    }
});

// ══════════════════════════════════════════════════════════
// ── INTERPOLATION ENGINE (Delaunay + IDW) ────────────────
// ══════════════════════════════════════════════════════════

function pointInTri(px, py, ax, ay, bx, by, cx, cy) {
    const d1 = (px-bx)*(ay-by) - (ax-bx)*(py-by);
    const d2 = (px-cx)*(by-cy) - (bx-cx)*(py-cy);
    const d3 = (px-ax)*(cy-ay) - (cx-ax)*(py-ay);
    return !((d1<0||d2<0||d3<0) && (d1>0||d2>0||d3>0));
}

function baryInterp(px, py, ax, ay, az, bx, by, bz, cx, cy, cz) {
    const det = (by-cy)*(ax-cx) + (cx-bx)*(ay-cy);
    if (Math.abs(det) < 1e-12) return null;
    const l1 = ((by-cy)*(px-cx) + (cx-bx)*(py-cy)) / det;
    const l2 = ((cy-ay)*(px-cx) + (ax-cx)*(py-cy)) / det;
    return l1*az + l2*bz + (1-l1-l2)*cz;
}

function inCircumcircle(px, py, a, b, c) {
    const ax=a[0]-px, ay=a[1]-py, bx=b[0]-px, by=b[1]-py, cx=c[0]-px, cy=c[1]-py;
    const det = ax*(by*(cx*cx+cy*cy)-cy*(bx*bx+by*by))
               -ay*(bx*(cx*cx+cy*cy)-cx*(bx*bx+by*by))
               +(ax*ax+ay*ay)*(bx*cy-by*cx);
    const orient = (b[0]-a[0])*(c[1]-a[1])-(b[1]-a[1])*(c[0]-a[0]);
    return orient > 0 ? det > 0 : det < 0;
}

function computeDelaunay(points) {
    if (points.length < 3) return [];
    const xs = points.map(p=>p[0]), ys = points.map(p=>p[1]);
    const pad = 200;
    const mn = [Math.min(...xs)-pad, Math.min(...ys)-pad];
    const d = Math.max(Math.max(...xs)-mn[0], Math.max(...ys)-mn[1]) * 3 + pad;
    const st = [[mn[0]-d, mn[1]-d], [mn[0]+d*3, mn[1]-d], [mn[0], mn[1]+d*3]];
    const all = [...points, ...st];
    const n = points.length;
    let tris = [[n, n+1, n+2]];
    for (let i = 0; i < n; i++) {
        const px = all[i][0], py = all[i][1];
        const bad = [];
        for (let t = 0; t < tris.length; t++) {
            if (inCircumcircle(px, py, all[tris[t][0]], all[tris[t][1]], all[tris[t][2]])) bad.push(t);
        }
        const edges = [];
        for (const t of bad) {
            const tri = tris[t];
            for (let j = 0; j < 3; j++) {
                const e = [tri[j], tri[(j+1)%3]];
                let shared = false;
                for (const t2 of bad) {
                    if (t2===t) continue;
                    const tri2 = tris[t2];
                    for (let k = 0; k < 3; k++) {
                        if ((tri2[k]===e[0]&&tri2[(k+1)%3]===e[1])||(tri2[k]===e[1]&&tri2[(k+1)%3]===e[0])) shared=true;
                    }
                }
                if (!shared) edges.push(e);
            }
        }
        tris = tris.filter((_,idx) => !bad.includes(idx));
        for (const e of edges) tris.push([i, e[0], e[1]]);
    }
    return tris.filter(t => t[0]<n && t[1]<n && t[2]<n);
}

function idwInterp(px, py, coords, depths, pow) {
    let num=0, den=0;
    for (let i=0; i<coords.length; i++) {
        const dx=px-coords[i][0], dy=py-coords[i][1];
        const d = Math.sqrt(dx*dx+dy*dy);
        if (d < 0.01) return depths[i];
        const w = 1/Math.pow(d, pow || 2);
        num += w*depths[i]; den += w;
    }
    return den > 0 ? num/den : 0;
}

function interpAt(px, py) {
    const { _triangles: tris, _coords: coords, _depths: depths } = foundingStrata;
    const method = foundingStrata.settings.method;
    if (method === 'tin' && tris.length > 0) {
        for (const [i,j,k] of tris) {
            if (pointInTri(px, py, coords[i][0],coords[i][1], coords[j][0],coords[j][1], coords[k][0],coords[k][1])) {
                const v = baryInterp(px, py, coords[i][0],coords[i][1],depths[i], coords[j][0],coords[j][1],depths[j], coords[k][0],coords[k][1],depths[k]);
                if (v !== null) return v;
            }
        }
    }
    if (coords.length > 0) return idwInterp(px, py, coords, depths);
    return 0;
}

function isInsideConvexHull(px, py) {
    const { _triangles: tris, _coords: coords } = foundingStrata;
    for (const [i,j,k] of tris) {
        if (pointInTri(px, py, coords[i][0],coords[i][1], coords[j][0],coords[j][1], coords[k][0],coords[k][1])) return true;
    }
    return false;
}

// ── Marching Squares + Chain ──
function marchSquares(grid, rows, cols, level) {
    const segs = [];
    for (let r=0; r<rows-1; r++) {
        for (let c=0; c<cols-1; c++) {
            const tl=grid[r*cols+c], tr=grid[r*cols+c+1], br=grid[(r+1)*cols+c+1], bl=grid[(r+1)*cols+c];
            let code=0;
            if(tl>=level) code|=8; if(tr>=level) code|=4; if(br>=level) code|=2; if(bl>=level) code|=1;
            if(code===0||code===15) continue;
            const lp=(a,b)=>a===b?0.5:(level-a)/(b-a);
            const top=[c+lp(tl,tr),r], right=[c+1,r+lp(tr,br)], bot=[c+lp(bl,br),r+1], left=[c,r+lp(tl,bl)];
            switch(code){
                case 1:case 14:segs.push([left,bot]);break;
                case 2:case 13:segs.push([bot,right]);break;
                case 3:case 12:segs.push([left,right]);break;
                case 4:case 11:segs.push([top,right]);break;
                case 5:segs.push([left,top]);segs.push([bot,right]);break;
                case 6:case 9:segs.push([top,bot]);break;
                case 7:case 8:segs.push([left,top]);break;
                case 10:segs.push([top,right]);segs.push([left,bot]);break;
            }
        }
    }
    return segs;
}

function chainSegments(segments) {
    if (!segments.length) return [];
    const eps = 0.6;
    const near = (a,b) => Math.abs(a[0]-b[0])<eps && Math.abs(a[1]-b[1])<eps;
    const used = new Set();
    const chains = [];
    for (let i=0; i<segments.length; i++) {
        if (used.has(i)) continue;
        used.add(i);
        let chain = [segments[i][0], segments[i][1]];
        let changed = true;
        while (changed) {
            changed = false;
            for (let j=0; j<segments.length; j++) {
                if (used.has(j)) continue;
                const [a,b] = segments[j];
                const head=chain[0], tail=chain[chain.length-1];
                if (near(a,tail)) { chain.push(b); used.add(j); changed=true; }
                else if (near(b,tail)) { chain.push(a); used.add(j); changed=true; }
                else if (near(b,head)) { chain.unshift(a); used.add(j); changed=true; }
                else if (near(a,head)) { chain.unshift(b); used.add(j); changed=true; }
            }
        }
        chains.push(chain);
    }
    return chains;
}

// ── Colour Ramp (warm=shallow, cool=deep) ──
const DEPTH_RAMP = [
    [76, 175, 80],    // green — shallow (good, less blinding)
    [255, 235, 59],   // yellow
    [255, 152, 0],    // orange
    [244, 67, 54],    // red
    [136, 14, 79],    // deep purple — deep (lots of blinding)
];

function depthToRGB(d, mn, mx) {
    const t = mx > mn ? Math.max(0, Math.min(1, (d-mn)/(mx-mn))) : 0.5;
    const idx = t * (DEPTH_RAMP.length-1);
    const lo = Math.floor(idx), hi = Math.min(lo+1, DEPTH_RAMP.length-1);
    const f = idx - lo;
    return [
        Math.round(DEPTH_RAMP[lo][0]*(1-f)+DEPTH_RAMP[hi][0]*f),
        Math.round(DEPTH_RAMP[lo][1]*(1-f)+DEPTH_RAMP[hi][1]*f),
        Math.round(DEPTH_RAMP[lo][2]*(1-f)+DEPTH_RAMP[hi][2]*f),
    ];
}

function depthToCSS(d, mn, mx, a) {
    const [r,g,b] = depthToRGB(d, mn, mx);
    return `rgba(${r},${g},${b},${a||1})`;
}

// ══════════════════════════════════════════════════════════
// ── HEATMAP + CONTOUR RENDERING ──────────────────────────
// ══════════════════════════════════════════════════════════

function recomputeInterpolationGrid() {
    const boreholes = project.elements.filter(el => el.type === 'borehole');
    if (boreholes.length < 2) {
        foundingStrata._grid = null;
        return;
    }

    const coords = boreholes.map(b => [b.x / 1000, b.y / 1000]); // real-world mm → metres
    const depths = boreholes.map(b => b.depthToRock);
    const tris = boreholes.length >= 3 ? computeDelaunay(coords) : [];

    foundingStrata._coords = coords;
    foundingStrata._depths = depths;
    foundingStrata._triangles = tris;

    // Build interpolation grid in sheet coordinates
    const GRID_RES = 6; // pixels — balance of speed vs resolution
    const vp = engine.viewport;
    const cols = Math.ceil(container.clientWidth / GRID_RES);
    const rows = Math.ceil(container.clientHeight / GRID_RES);
    const grid = new Float32Array(cols * rows);
    let mn = Infinity, mx = -Infinity;

    for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
            const screenX = c * GRID_RES + GRID_RES / 2;
            const screenY = r * GRID_RES + GRID_RES / 2;
            const sheetPos = engine.coords.screenToSheet(screenX, screenY);
            const realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);
            const rm = realPos.x / 1000; // mm to metres
            const ry = realPos.y / 1000;
            const v = interpAt(rm, ry);
            grid[r * cols + c] = v;
            if (v < mn) mn = v;
            if (v > mx) mx = v;
        }
    }

    foundingStrata._grid = grid;
    foundingStrata._gridCols = cols;
    foundingStrata._gridRows = rows;
    foundingStrata._minDepth = mn;
    foundingStrata._maxDepth = mx;
}

// Render heatmap + contours as early layer (before elements)
let _lastGridHash = '';
engine.onRender((ctx, eng) => {
    if (!foundingStrata.active) return;
    const boreholes = project.elements.filter(el => el.type === 'borehole');
    if (boreholes.length < 2) return;

    // Recompute grid (debounced by checking hash)
    const hash = boreholes.map(b => b.x + ',' + b.y + ',' + b.depthToRock).join('|')
               + '|' + foundingStrata.settings.method
               + '|' + eng.viewport.zoom + ',' + eng.viewport.panX + ',' + eng.viewport.panY;
    if (hash !== _lastGridHash) {
        _lastGridHash = hash;
        recomputeInterpolationGrid();
    }

    const { _grid: grid, _gridCols: cols, _gridRows: rows, _minDepth: mn, _maxDepth: mx } = foundingStrata;
    if (!grid) return;
    const s = foundingStrata.settings;
    const GRID_RES = 6;

    // ── Heatmap ──
    if (s.showHeatmap) {
        const imgData = ctx.createImageData(cols, rows);
        const alpha = Math.round(s.heatmapOpacity * 255);
        for (let i = 0; i < cols * rows; i++) {
            const [r,g,b] = depthToRGB(grid[i], mn, mx);
            imgData.data[i*4] = r;
            imgData.data[i*4+1] = g;
            imgData.data[i*4+2] = b;
            imgData.data[i*4+3] = alpha;
        }
        const tmpCanvas = document.createElement('canvas');
        tmpCanvas.width = cols; tmpCanvas.height = rows;
        tmpCanvas.getContext('2d').putImageData(imgData, 0, 0);
        ctx.imageSmoothingEnabled = true;
        ctx.drawImage(tmpCanvas, 0, 0, container.clientWidth, container.clientHeight);
    }

    // ── Contour Lines ──
    if (s.showContours) {
        const interval = s.contourInterval;
        const start = Math.ceil(mn / interval) * interval;
        ctx.lineWidth = 0.7;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';

        for (let lev = start; lev <= mx + 0.001; lev += interval) {
            const level = Math.round(lev * 1000) / 1000;
            const segs = marchSquares(grid, rows, cols, level);
            const scaledSegs = segs.map(s => [
                [s[0][0]*GRID_RES, s[0][1]*GRID_RES],
                [s[1][0]*GRID_RES, s[1][1]*GRID_RES]
            ]);
            const chains = chainSegments(scaledSegs);

            ctx.strokeStyle = 'rgba(160,160,160,0.5)';

            for (const chain of chains) {
                if (chain.length < 2) continue;
                ctx.beginPath();
                ctx.moveTo(chain[0][0], chain[0][1]);
                for (let i = 1; i < chain.length; i++) ctx.lineTo(chain[i][0], chain[i][1]);
                ctx.stroke();
            }

            // Label on longest chain
            let longest = chains[0];
            for (const ch of chains) { if (ch.length > (longest?.length||0)) longest = ch; }
            if (longest && longest.length > 4) {
                const mi = Math.floor(longest.length / 2);
                const lx = longest[mi][0], ly = longest[mi][1];
                const txt = level.toFixed(1) + 'm';
                const labelFs = 9;
                ctx.font = `600 ${labelFs}px "Segoe UI", Arial, sans-serif`;
                const tw = ctx.measureText(txt).width + 6;
                ctx.fillStyle = 'rgba(255,255,255,0.85)';
                ctx.beginPath();
                if (ctx.roundRect) ctx.roundRect(lx - tw/2, ly - 7, tw, 14, 2);
                else { ctx.rect(lx - tw/2, ly - 7, tw, 14); }
                ctx.fill();
                ctx.fillStyle = 'rgba(120,120,120,0.8)';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.fillText(txt, lx, ly);
            }
        }
    }

    // ── TIN mesh edges ──
    if (s.showTIN && foundingStrata._triangles.length > 0) {
        ctx.strokeStyle = 'rgba(139,69,19,0.2)';
        ctx.lineWidth = 0.8;
        ctx.setLineDash([4, 3]);
        for (const [i,j,k] of foundingStrata._triangles) {
            const bhs = boreholes;
            const p1 = eng.coords.realToScreen(bhs[i].x, bhs[i].y);
            const p2 = eng.coords.realToScreen(bhs[j].x, bhs[j].y);
            const p3 = eng.coords.realToScreen(bhs[k].x, bhs[k].y);
            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.lineTo(p3.x, p3.y);
            ctx.closePath();
            ctx.stroke();
        }
        ctx.setLineDash([]);
    }
});

// ══════════════════════════════════════════════════════════
// ── BLINDING CALCULATION + FOOTING THEMATIC ──────────────
// ══════════════════════════════════════════════════════════

function calculateBlindingSchedule() {
    const boreholes = project.elements.filter(el => el.type === 'borehole');
    if (boreholes.length < 2) return [];

    const s = foundingStrata.settings;
    const socketM = s.socketDepth / 1000;   // mm → m
    const contM = s.contingency / 1000;      // mm → m

    const allFootings = project.elements.filter(el =>
        el.type === 'footing' || el.type === 'stripFooting'
    );

    const groundRL = levelSystem.groundRL;
    const hasRL = groundRL !== 0;

    const schedule = [];
    for (const ft of allFootings) {
        let ftX, ftY; // centre in real mm
        let ftAreaM2;  // area in m²
        let ftTag;
        let ftDepthBelowFSL; // m — depth to underside of footing below ground surface

        if (ft.type === 'footing') {
            ftX = ft.x;
            ftY = ft.y;
            const pfTypeRef = ft.typeRef || ft.mark || 'PF1';
            const pfTypeData = project.scheduleTypes.padfooting[pfTypeRef] || {};
            const fw = (ft.footingWidth || ft.width || 1000) / 1000; // mm → m
            const fd = (ft.footingDepth || ft.depth || 400) / 1000;
            // Rectangular footings use width × length; square use width × width
            const fl = pfTypeData.rect ? (parseInt(pfTypeData.length) || ft.footingWidth || 1000) / 1000 : fw;
            ftAreaM2 = fw * fl;
            ftTag = ft.mark || 'PF?';
            // Use actual TOF (respects overrides) to derive true depth below ground
            if (hasRL) {
                const tof = getFootingTOF(ft);
                ftDepthBelowFSL = (groundRL - tof) + fd;
            } else {
                ftDepthBelowFSL = ((ft.depthBelowFSL || 200) + (ft.footingDepth || ft.depth || 400)) / 1000;
            }
        } else {
            // Strip footing — use midpoint
            ftX = (ft.x1 + ft.x2) / 2;
            ftY = (ft.y1 + ft.y2) / 2;
            const dx = (ft.x2 - ft.x1) / 1000, dy = (ft.y2 - ft.y1) / 1000; // m
            const len = Math.sqrt(dx*dx + dy*dy);
            const fw = (ft.footingWidth || 300) / 1000;
            const fd = (ft.footingDepth || 500) / 1000;
            ftAreaM2 = len * fw;
            ftTag = ft.tag || 'SF?';
            // Use actual TOF (respects overrides) to derive true depth below ground
            if (hasRL) {
                const tof = getFootingTOF(ft);
                ftDepthBelowFSL = (groundRL - tof) + fd;
            } else {
                ftDepthBelowFSL = ((ft.depthBelowFSL || 200) + (ft.footingDepth || 500)) / 1000;
            }
        }

        // Convert to metres for interpolation
        const xM = ftX / 1000, yM = ftY / 1000;

        // Sample rock depth — take worst case (deepest) at centre + corners for pads
        let rockDepth;
        if (ft.type === 'footing') {
            const fw = (ft.footingWidth || ft.width || 1000) / 1000 / 2;
            const samples = [
                interpAt(xM, yM),
                interpAt(xM - fw, yM - fw),
                interpAt(xM + fw, yM - fw),
                interpAt(xM - fw, yM + fw),
                interpAt(xM + fw, yM + fw),
            ];
            rockDepth = Math.max(...samples); // worst case = deepest rock
        } else {
            // Strip: sample at start, mid, end
            const x1M = ft.x1/1000, y1M = ft.y1/1000, x2M = ft.x2/1000, y2M = ft.y2/1000;
            const samples = [
                interpAt(x1M, y1M),
                interpAt(xM, yM),
                interpAt(x2M, y2M),
            ];
            rockDepth = Math.max(...samples);
        }

        // Blinding = rock depth + socket depth - footing underside depth + contingency
        // If rock is shallower than footing underside → blinding = socket depth only
        const blindingM = Math.max(0, rockDepth + socketM - ftDepthBelowFSL) + contM;
        const volumeM3 = blindingM * ftAreaM2;
        const insideHull = isInsideConvexHull(xM, yM);

        // Calculate RL-based values when ground RL is set
        const tofRL = hasRL ? getFootingTOF(ft) : null;
        const uosRL = hasRL ? tofRL - (ft.type === 'footing' ? (ft.footingDepth || ft.depth || 400) : (ft.footingDepth || 500)) / 1000 : null;
        // Rock RL at footing location (interpolated from borehole RLs if available)
        const rockRLAtFt = hasRL ? (groundRL - rockDepth) : null;

        schedule.push({
            el: ft,
            tag: ftTag,
            rockDepth: rockDepth,
            ftDepthBelowFSL: ftDepthBelowFSL,
            blindingM: blindingM,
            blindingMM: Math.round(blindingM * 1000),
            volumeM3: volumeM3,
            areaM2: ftAreaM2,
            insideHull: insideHull,
            x: ftX, y: ftY,
            tofRL: tofRL,
            uosRL: uosRL,
            rockRL: rockRLAtFt,
        });
    }

    return schedule;
}

// ── Footing blinding colour thematic overlay ──
engine.onRender((ctx, eng) => {
    if (!foundingStrata.active) return;
    const boreholes = project.elements.filter(el => el.type === 'borehole');
    if (boreholes.length < 2) return;

    const schedule = calculateBlindingSchedule();
    if (schedule.length === 0) return;
    const maxBlind = Math.max(...schedule.map(s => s.blindingMM), 1);
    const zoom = eng.viewport.zoom;

    for (const entry of schedule) {
        const ft = entry.el;
        let pos, halfW, halfH;

        if (ft.type === 'footing') {
            pos = eng.coords.realToScreen(ft.x, ft.y);
            const fw = (ft.footingWidth || ft.width || 1000);
            halfW = (fw / 2 / CONFIG.drawingScale) * zoom;
            halfH = halfW;
        } else if (ft.type === 'stripFooting') {
            const mx = (ft.x1 + ft.x2) / 2, my = (ft.y1 + ft.y2) / 2;
            pos = eng.coords.realToScreen(mx, my);
            halfW = 0; halfH = 0;
        }

        // ── Colour fill based on blinding severity ──
        const [cr,cg,cb] = depthToRGB(entry.blindingMM, 0, Math.max(maxBlind, 500));

        if (ft.type === 'footing' && halfW > 0) {
            // Pad footing — square fill
            ctx.fillStyle = `rgba(${cr},${cg},${cb},0.25)`;
            ctx.fillRect(pos.x - halfW, pos.y - halfH, halfW * 2, halfH * 2);
        } else if (ft.type === 'stripFooting' && entry.blindingMM > 0) {
            // Strip footing — fill along the strip length
            const sfTypeRef = ft.typeRef || ft.tag || 'SF1';
            const sfTD = project.scheduleTypes.stripfooting[sfTypeRef] || {};
            const sfW = (parseInt(sfTD.width) || ft.stripWidth || 400) / CONFIG.drawingScale * zoom;
            const p1 = eng.coords.realToScreen(ft.x1, ft.y1);
            const p2 = eng.coords.realToScreen(ft.x2, ft.y2);
            const dx = p2.x - p1.x, dy = p2.y - p1.y;
            const len = Math.sqrt(dx * dx + dy * dy);
            if (len > 0) {
                const nx = -dy / len * (sfW / 2), ny = dx / len * (sfW / 2);
                ctx.fillStyle = `rgba(${cr},${cg},${cb},0.25)`;
                ctx.beginPath();
                ctx.moveTo(p1.x + nx, p1.y + ny);
                ctx.lineTo(p2.x + nx, p2.y + ny);
                ctx.lineTo(p2.x - nx, p2.y - ny);
                ctx.lineTo(p1.x - nx, p1.y - ny);
                ctx.closePath();
                ctx.fill();
            }
        }

        // ── Blinding depth tag (black text, positioned below TOF tag) ──
        if (entry.blindingMM > 0) {
            const blindM = (entry.blindingMM / 1000).toFixed(2);
            const txt = blindM + 'm blinding';
            const labelFs = Math.max(1, 1.6 * zoom);
            ctx.font = `600 ${labelFs}px "Segoe UI", Arial, sans-serif`;
            ctx.fillStyle = '#000000';
            ctx.textAlign = 'left';
            ctx.textBaseline = 'middle';

            if (ft.type === 'footing') {
                // Position below-right of pad — offset down to avoid TOF tag clash
                const tofTagHeight = Math.max(7, 2.5 * zoom) + 6; // fontSize + padY*2 + gap
                ctx.fillText(txt, pos.x + halfW + 3, pos.y + tofTagHeight);
            } else {
                // Position just below the strip footing midpoint
                ctx.textAlign = 'center';
                ctx.fillText(txt, pos.x, pos.y + 4 * zoom);
            }
        }
    }
});

// ══════════════════════════════════════════════════════════
// ── FOUNDING STRATA PANEL UI ─────────────────────────────
// ══════════════════════════════════════════════════════════

function updateFoundingPanel() {
    const content = document.getElementById('fp-content');
    const boreholes = project.elements.filter(el => el.type === 'borehole');
    const s = foundingStrata.settings;

    // Calculate schedule
    const schedule = boreholes.length >= 2 ? calculateBlindingSchedule() : [];
    const totalVol = schedule.reduce((sum, e) => sum + e.volumeM3, 0);

    let html = '';

    // ── Mode hint ──
    html += `<div class="fp-section">
        <div style="font-size:10px;color:#666;line-height:1.5;">
            ${activeTool === 'borehole'
                ? '<b>Click on the plan to place boreholes.</b> Double-click a borehole to edit its depth.'
                : 'Switch to <b>Select</b> to drag boreholes. Double-click to edit depth.'}
        </div>
        <div style="margin-top:6px;">
            <button class="tbtn${activeTool==='borehole'?' active':''}" id="fp-btn-add-bh" style="font-size:10px;padding:3px 8px;">+ Add Borehole</button>
        </div>
    </div>`;

    // ── Total Volume ──
    html += `<div class="fp-section">
        <div class="fp-total-box">
            <div class="fp-total-vol">${totalVol.toFixed(2)} m³</div>
            <div class="fp-total-label">Estimated Blinding Concrete Volume</div>
        </div>
    </div>`;

    // ── Boreholes ──
    html += `<div class="fp-section">
        <div class="fp-section-title">Boreholes (${boreholes.length})</div>`;
    if (boreholes.length === 0) {
        html += '<div style="font-size:10px;color:#999;">No boreholes placed yet. Click on the plan to add.</div>';
    }
    const hasRL = levelSystem.groundRL !== 0;
    for (const bh of boreholes) {
        const rlInfo = (hasRL && bh.rockRL !== undefined)
            ? `<span style="font-size:9px;color:#2B7CD0;margin-left:4px;">Rock RL ${bh.rockRL.toFixed(3)}</span>`
            : '';
        const grlInfo = (hasRL && bh.groundRL)
            ? `<span style="font-size:8px;color:#999;margin-left:2px;">(GL ${bh.groundRL.toFixed(3)})</span>`
            : '';
        html += `<div class="fp-bh-item">
            <span class="fp-bh-tag">${bh.tag}</span>
            <span class="fp-bh-depth">${bh.depthToRock.toFixed(2)}m</span>
            ${rlInfo}${grlInfo}
            <span class="fp-bh-del" data-bh-del="${bh.id}" title="Remove borehole">✕</span>
        </div>`;
    }
    html += '</div>';

    // ── Settings ──
    html += `<div class="fp-section">
        <div class="fp-section-title">Settings</div>
        <div class="fp-row"><label>Socket into rock</label><input type="number" id="fp-socket" value="${s.socketDepth}" min="0" step="50" style="width:70px;"> <span style="font-size:9px;color:#999;">mm</span></div>
        <div class="fp-row"><label>Contingency</label><input type="number" id="fp-contingency" value="${s.contingency}" min="0" step="10" style="width:70px;"> <span style="font-size:9px;color:#999;">mm</span></div>
        <div class="fp-row"><label>Contour interval</label><input type="number" id="fp-contour-int" value="${s.contourInterval}" min="0.05" step="0.05" style="width:70px;"> <span style="font-size:9px;color:#999;">m</span></div>
        <div class="fp-row"><label>Interpolation</label><select id="fp-method"><option value="tin"${s.method==='tin'?' selected':''}>TIN (Delaunay)</option><option value="idw"${s.method==='idw'?' selected':''}>IDW</option></select></div>
        <div class="fp-row"><label><input type="checkbox" id="fp-heatmap" ${s.showHeatmap?'checked':''}> Heatmap</label></div>
        <div class="fp-row"><label><input type="checkbox" id="fp-contours" ${s.showContours?'checked':''}> Contour lines</label></div>
        <div class="fp-row"><label><input type="checkbox" id="fp-tin-mesh" ${s.showTIN?'checked':''}> TIN mesh</label></div>
        <div class="fp-row"><label>Heatmap opacity</label><input type="range" id="fp-opacity" min="0.05" max="0.7" step="0.05" value="${s.heatmapOpacity}" style="flex:1;"></div>
    </div>`;

    // ── Legend ──
    if (boreholes.length >= 2) {
        const mn = foundingStrata._minDepth, mx = foundingStrata._maxDepth;
        const gradStops = [];
        for (let i = 0; i <= 10; i++) {
            const d = mn + (mx - mn) * i / 10;
            const [r,g,b] = depthToRGB(d, mn, mx);
            gradStops.push(`rgb(${r},${g},${b})`);
        }
        html += `<div class="fp-section">
            <div class="fp-section-title">Depth Legend</div>
            <div class="fp-legend-bar" style="background:linear-gradient(to right,${gradStops.join(',')})"></div>
            <div class="fp-legend-labels"><span>${mn.toFixed(1)}m (shallow)</span><span>${mx.toFixed(1)}m (deep)</span></div>
        </div>`;
    }

    // ── Blinding Schedule ──
    if (schedule.length > 0) {
        const showRL = hasRL && schedule[0].tofRL !== null;
        html += `<div class="fp-section">
            <div class="fp-section-title">Blinding Schedule</div>
            <div class="fp-sched-row fp-sched-hdr">
                <span>Mark</span>${showRL ? '<span>TOF RL</span><span>Rock RL</span>' : '<span>Rock (m)</span><span>UOS (m)</span>'}<span>Blind</span><span>Vol</span>
            </div>`;

        for (const entry of schedule) {
            let blindColor = '#4CAF50'; // green
            if (entry.blindingMM > 300) blindColor = '#f44336'; // red
            else if (entry.blindingMM > 150) blindColor = '#FF9800'; // orange
            else if (entry.blindingMM > 50) blindColor = '#FFC107'; // yellow

            const statusIcon = entry.insideHull ? '' : ' ⚠';

            if (showRL) {
                html += `<div class="fp-sched-row">
                    <span style="font-weight:700;">${entry.tag}${statusIcon}</span>
                    <span style="color:var(--accent);font-size:10px;">${entry.tofRL.toFixed(3)}</span>
                    <span style="color:#8B4513;font-size:10px;">${entry.rockRL.toFixed(3)}</span>
                    <span><span class="fp-blind-badge" style="background:${blindColor}20;color:${blindColor};">${entry.blindingMM}mm</span></span>
                    <span style="font-family:var(--font-mono);font-size:9px;">${entry.volumeM3.toFixed(2)}</span>
                </div>`;
            } else {
                html += `<div class="fp-sched-row">
                    <span style="font-weight:700;">${entry.tag}${statusIcon}</span>
                    <span>${entry.rockDepth.toFixed(2)}</span>
                    <span>${entry.ftDepthBelowFSL.toFixed(2)}</span>
                    <span><span class="fp-blind-badge" style="background:${blindColor}20;color:${blindColor};">${entry.blindingMM}mm</span></span>
                    <span style="font-family:var(--font-mono);font-size:9px;">${entry.volumeM3.toFixed(2)}</span>
                </div>`;
            }
        }

        html += `<div style="text-align:right;padding:4px 0;font-weight:700;font-size:11px;">
            Total: ${totalVol.toFixed(2)} m³
        </div>`;
        html += '</div>';
    }

    // ── Disclaimers ──
    html += `<div class="fp-section">
        <div class="fp-note">
            <b>Note:</b> Blinding volumes are estimates based on ${boreholes.length} borehole(s) using
            ${s.method === 'tin' ? 'Delaunay TIN' : 'Inverse Distance Weighting'} interpolation.
            Rock surface is assumed to vary linearly between boreholes.
            Actual conditions may differ — verify during excavation.
            Footings marked ⚠ are outside the borehole convex hull (extrapolated).
        </div>
    </div>`;

    content.innerHTML = html;

    // Wire up panel events
    document.getElementById('fp-btn-add-bh')?.addEventListener('click', () => {
        setActiveTool('borehole');
        updateFoundingPanel();
    });

    document.getElementById('fp-socket')?.addEventListener('change', (e) => {
        s.socketDepth = parseFloat(e.target.value) || 0;
        engine.requestRender(); updateFoundingPanel();
    });
    document.getElementById('fp-contingency')?.addEventListener('change', (e) => {
        s.contingency = parseFloat(e.target.value) || 0;
        engine.requestRender(); updateFoundingPanel();
    });
    document.getElementById('fp-contour-int')?.addEventListener('change', (e) => {
        s.contourInterval = parseFloat(e.target.value) || 0.2;
        _lastGridHash = ''; engine.requestRender();
    });
    document.getElementById('fp-method')?.addEventListener('change', (e) => {
        s.method = e.target.value;
        _lastGridHash = ''; engine.requestRender(); updateFoundingPanel();
    });
    document.getElementById('fp-heatmap')?.addEventListener('change', (e) => {
        s.showHeatmap = e.target.checked;
        engine.requestRender();
    });
    document.getElementById('fp-contours')?.addEventListener('change', (e) => {
        s.showContours = e.target.checked;
        engine.requestRender();
    });
    document.getElementById('fp-tin-mesh')?.addEventListener('change', (e) => {
        s.showTIN = e.target.checked;
        engine.requestRender();
    });
    document.getElementById('fp-opacity')?.addEventListener('input', (e) => {
        s.heatmapOpacity = parseFloat(e.target.value);
        engine.requestRender();
    });

    // Borehole delete buttons
    content.querySelectorAll('[data-bh-del]').forEach(btn => {
        btn.addEventListener('click', () => {
            const bhId = btn.dataset.bhDel;
            const idx = project.elements.findIndex(el => el.id === bhId);
            if (idx !== -1) {
                const removed = project.elements[idx];
                history.execute({
                    description: 'Remove borehole: ' + removed.tag,
                    execute() {
                        const i = project.elements.indexOf(removed);
                        if (i !== -1) project.elements.splice(i, 1);
                    },
                    undo() { project.elements.push(removed); }
                });
                _lastGridHash = '';
                engine.requestRender();
                updateFoundingPanel();
            }
        });
    });
}

// Update panel when elements change (throttled)
let _fpUpdateTimer = null;
const _origRender = engine.requestRender.bind(engine);
const _patchedRender = function() {
    _origRender();
    if (foundingStrata.active) {
        clearTimeout(_fpUpdateTimer);
        _fpUpdateTimer = setTimeout(updateFoundingPanel, 300);
    }
};
engine.requestRender = _patchedRender;

// ── Ensure borehole tool is in setActiveTool ──
const _origSetActive = setActiveTool;
setActiveTool = function(tool) {
    _origSetActive(tool);
    const btn = document.getElementById('btn-founding-strata');
    if (btn) btn.classList.toggle('active', foundingStrata.active);
};

// Add S-GEO layer if it doesn't exist
if (!project.layers['S-GEO']) {
    project.layers['S-GEO'] = { name: 'Geotechnical', color: '#8B4513', visible: true };
}

// Add snap support for boreholes
const _origFindSnap = findSnap;
findSnap = function(sx, sy) {
    const result = _origFindSnap(sx, sy);
    if (result) return result;
    // Also snap to borehole positions
    const snapR = 8;
    for (const el of project.elements) {
        if (el.type !== 'borehole') continue;
        const sp = engine.coords.realToScreen(el.x, el.y);
        const shp = engine.coords.realToSheet(el.x, el.y);
        if (Math.abs(sp.x - sx) < snapR && Math.abs(sp.y - sy) < snapR) {
            return { x: shp.x, y: shp.y, type: 'endpoint' };
        }
    }
    return null;
};

// ══════════════════════════════════════════════════════════
