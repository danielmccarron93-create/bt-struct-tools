// ── 3D ENGINE: TRIMBLE-STYLE STRUCTURAL VIEWER ──────────
// ══════════════════════════════════════════════════════════

let scene3d, camera3d;

// ── Material Palette ──
const MAT_COLORS = {
    steel:    0x8899AA,
    concrete: 0xB0A898,
    timber:   0xC4A672,
    block:    0xA09080,
    stud:     0x999999,
    footing:  0x998877,
    ground:   0xD5D0C8,
    edge:     0x444444,
    selected: 0x2B7CD0,
};

// ── On-sheet 3D view state (frozen snapshot) ──
const isoView = {
    visible: true,
    x: 550, y: 15,
    width: 260, height: 200,
    // Frozen snapshot
    canvas: null,
    snapshotDataURL: null,
    snapshotImage: null,
    showBorder: false,
    // Interaction
    dragging: false, resizing: false,
    dragStart: null, dragOrigPos: null,
    resizeStart: null, resizeOrigSize: null,
    // Selection (shared with fullscreen)
    selectedMesh: null, selectedOrigMat: null,
};


// ══════════════════════════════════════════════════════════
// ── 3D STRUCTURAL VIEWER (Trimble-style) ─────────────────
// ══════════════════════════════════════════════════════════

// ── 3D Material Factory ───────────────────────────────────

function makeMat(color, opts = {}) {
    return new THREE.MeshStandardMaterial({
        color: color,
        metalness: opts.metalness ?? 0.05,
        roughness: opts.roughness ?? 0.85,
        transparent: opts.transparent ?? false,
        opacity: opts.opacity ?? 1.0,
        side: THREE.DoubleSide,
        depthWrite: opts.depthWrite ?? true,
        polygonOffset: true,
        polygonOffsetFactor: 1,
        polygonOffsetUnits: 1,
    });
}

const MAT_STEEL    = () => makeMat(MAT_COLORS.steel, { metalness: 0.15, roughness: 0.65 });
const MAT_CONCRETE = () => makeMat(MAT_COLORS.concrete, { roughness: 0.92, metalness: 0.0 });
const MAT_TIMBER   = () => makeMat(MAT_COLORS.timber, { roughness: 0.9, metalness: 0.0 });
const MAT_BLOCK    = () => makeMat(MAT_COLORS.block, { roughness: 0.95, metalness: 0.0 });
const MAT_STUD     = () => makeMat(MAT_COLORS.stud, { roughness: 0.8 });
const MAT_SLAB     = () => makeMat(MAT_COLORS.concrete, { transparent: true, opacity: 0.82, roughness: 0.92, metalness: 0.0, depthWrite: false });
const MAT_FOOTING  = () => makeMat(MAT_COLORS.footing, { roughness: 0.95, metalness: 0.0 });
const MAT_ROOF     = () => makeMat(0xCC6644, { roughness: 0.85, metalness: 0.05, transparent: true, opacity: 0.7 });
const MAT_ENVWALL  = () => makeMat(0xD5CFC8, { roughness: 0.95, metalness: 0.0, transparent: true, opacity: 0.35 });
const MAT_GROUND   = () => makeMat(MAT_COLORS.ground, { roughness: 1.0, metalness: 0.0 });
const MAT_SELECT   = () => makeMat(MAT_COLORS.selected, { metalness: 0.3, roughness: 0.5 });
const EDGE_MAT     = () => new THREE.LineBasicMaterial({ color: MAT_COLORS.edge });

// ── Geometry Builders ─────────────────────────────────────

/** Build I-beam cross-section with root fillets, extruded to length */
function buildIBeamGeo(sectionData, lengthMM) {
    const { d, bf, tf, tw, r1 } = sectionData;
    const scale = 1 / 1000;
    const hw = d / 2 * scale;
    const hb = bf / 2 * scale;
    const tfS = tf * scale;
    const twS = tw / 2 * scale;
    const L = lengthMM * scale;
    const r = Math.min(r1 || 0, (bf - tw) / 2 - 0.5, tf - 0.5) * scale;

    const sh = new THREE.Shape();
    sh.moveTo(-hb, -hw);
    sh.lineTo(hb, -hw);
    sh.lineTo(hb, -hw + tfS);
    if (r > 0.0005) {
        sh.lineTo(twS + r, -hw + tfS);
        sh.quadraticCurveTo(twS, -hw + tfS, twS, -hw + tfS + r);
    } else {
        sh.lineTo(twS, -hw + tfS);
    }
    if (r > 0.0005) {
        sh.lineTo(twS, hw - tfS - r);
        sh.quadraticCurveTo(twS, hw - tfS, twS + r, hw - tfS);
    } else {
        sh.lineTo(twS, hw - tfS);
    }
    sh.lineTo(hb, hw - tfS);
    sh.lineTo(hb, hw);
    sh.lineTo(-hb, hw);
    sh.lineTo(-hb, hw - tfS);
    if (r > 0.0005) {
        sh.lineTo(-twS - r, hw - tfS);
        sh.quadraticCurveTo(-twS, hw - tfS, -twS, hw - tfS - r);
    } else {
        sh.lineTo(-twS, hw - tfS);
    }
    if (r > 0.0005) {
        sh.lineTo(-twS, -hw + tfS + r);
        sh.quadraticCurveTo(-twS, -hw + tfS, -twS - r, -hw + tfS);
    } else {
        sh.lineTo(-twS, -hw + tfS);
    }
    sh.lineTo(-hb, -hw + tfS);
    sh.lineTo(-hb, -hw);

    const geo = new THREE.ExtrudeGeometry(sh, { depth: L, bevelEnabled: false, steps: 1 });
    geo.rotateY(-Math.PI / 2);
    geo.computeBoundingBox();
    const bb = geo.boundingBox;
    geo.translate(-(bb.max.x + bb.min.x) / 2, -(bb.max.y + bb.min.y) / 2, -(bb.max.z + bb.min.z) / 2);
    geo.computeVertexNormals();
    return geo;
}

/** Build hollow SHS/RHS geometry */
function buildSHSGeo(B, D, t, lengthMM) {
    const scale = 1 / 1000;
    const outerW = B * scale, outerD = (D || B) * scale;
    const innerW = (B - 2 * t) * scale, innerD = ((D || B) - 2 * t) * scale;
    const L = lengthMM * scale;

    const sh = new THREE.Shape();
    sh.moveTo(-outerW / 2, -outerD / 2);
    sh.lineTo(outerW / 2, -outerD / 2);
    sh.lineTo(outerW / 2, outerD / 2);
    sh.lineTo(-outerW / 2, outerD / 2);
    sh.lineTo(-outerW / 2, -outerD / 2);

    const hole = new THREE.Path();
    hole.moveTo(-innerW / 2, -innerD / 2);
    hole.lineTo(innerW / 2, -innerD / 2);
    hole.lineTo(innerW / 2, innerD / 2);
    hole.lineTo(-innerW / 2, innerD / 2);
    hole.lineTo(-innerW / 2, -innerD / 2);
    sh.holes.push(hole);

    const geo = new THREE.ExtrudeGeometry(sh, { depth: L, bevelEnabled: false, steps: 1 });
    geo.rotateY(-Math.PI / 2);
    geo.computeBoundingBox();
    const bb = geo.boundingBox;
    geo.translate(-(bb.max.x + bb.min.x) / 2, -(bb.max.y + bb.min.y) / 2, -(bb.max.z + bb.min.z) / 2);
    geo.computeVertexNormals();
    return geo;
}

/** Build hollow CHS geometry */
function buildCHSGeo(outerDia, t, lengthMM) {
    const scale = 1 / 1000;
    const ro = outerDia / 2 * scale;
    const ri = (outerDia / 2 - t) * scale;
    const L = lengthMM * scale;

    const sh = new THREE.Shape();
    sh.absarc(0, 0, ro, 0, Math.PI * 2, false);
    const hole = new THREE.Path();
    hole.absarc(0, 0, ri, 0, Math.PI * 2, true);
    sh.holes.push(hole);

    const geo = new THREE.ExtrudeGeometry(sh, { depth: L, bevelEnabled: false, steps: 1 });
    geo.rotateY(-Math.PI / 2);
    geo.computeBoundingBox();
    const bb = geo.boundingBox;
    geo.translate(-(bb.max.x + bb.min.x) / 2, -(bb.max.y + bb.min.y) / 2, -(bb.max.z + bb.min.z) / 2);
    geo.computeVertexNormals();
    return geo;
}

/** Build PFC (parallel flange channel) geometry */
function buildPFCGeo(sectionData, lengthMM) {
    const { d, bf, tf, tw } = sectionData;
    const scale = 1 / 1000;
    const D = d * scale, B = bf * scale;
    const TF = tf * scale, TW = tw * scale;
    const L = lengthMM * scale;

    const sh = new THREE.Shape();
    sh.moveTo(0, -D / 2);
    sh.lineTo(B, -D / 2);
    sh.lineTo(B, -D / 2 + TF);
    sh.lineTo(TW, -D / 2 + TF);
    sh.lineTo(TW, D / 2 - TF);
    sh.lineTo(B, D / 2 - TF);
    sh.lineTo(B, D / 2);
    sh.lineTo(0, D / 2);
    sh.lineTo(0, -D / 2);

    const geo = new THREE.ExtrudeGeometry(sh, { depth: L, bevelEnabled: false, steps: 1 });
    geo.rotateY(-Math.PI / 2);
    geo.computeBoundingBox();
    const bb = geo.boundingBox;
    geo.translate(-(bb.max.x + bb.min.x) / 2, -(bb.max.y + bb.min.y) / 2, -(bb.max.z + bb.min.z) / 2);
    geo.computeVertexNormals();
    return geo;
}

/** Build rectangular timber/GLT section */
function buildRectGeo(widthMM, depthMM, lengthMM) {
    const scale = 1 / 1000;
    return new THREE.BoxGeometry(lengthMM * scale, depthMM * scale, widthMM * scale);
}

/** Parse member size string to get section data for 3D */
function parseMemberFor3D(memberSize, category) {
    if (!memberSize) return null;

    const ubKey = memberSize.replace(/\s+/g, '').replace('UB', 'UB').replace('UC', 'UC');
    if (UB_DIMENSIONS[ubKey]) {
        return { type: 'UB', data: UB_DIMENSIONS[ubKey] };
    }

    const cleanSize = memberSize.replace(/\s+/g, '');
    for (const [key, val] of Object.entries(UB_DIMENSIONS)) {
        if (cleanSize.includes(key) || key.includes(cleanSize)) {
            return { type: 'UB', data: val };
        }
    }

    const shsMatch = memberSize.match(/(\d+)(?:x(\d+))?x([\d.]+)\s*SHS/i) ||
                     memberSize.match(/^(\d+)x([\d.]+)$/);
    if (shsMatch || (category && category === 'SHS')) {
        const m = memberSize.match(/(\d+)x([\d.]+)/);
        if (m) return { type: 'SHS', B: parseFloat(m[1]), t: parseFloat(m[2]) };
    }

    const rhsMatch = memberSize.match(/(\d+)x(\d+)x([\d.]+)\s*RHS/i);
    if (rhsMatch || (category && category === 'RHS')) {
        const m = memberSize.match(/(\d+)x(\d+)x([\d.]+)/);
        if (m) return { type: 'RHS', B: parseFloat(m[1]), D: parseFloat(m[2]), t: parseFloat(m[3]) };
    }

    const chsMatch = memberSize.match(/([\d.]+)x([\d.]+)\s*CHS/i);
    if (chsMatch || (category && category === 'CHS')) {
        const m = memberSize.match(/([\d.]+)x([\d.]+)/);
        if (m) return { type: 'CHS', D: parseFloat(m[1]), t: parseFloat(m[2]) };
    }

    if (category === 'PFC' || /PFC/i.test(memberSize)) {
        const pfcSections = {
            '380PFC': { d: 380, bf: 100, tf: 17.5, tw: 10, r1: 14 },
            '300PFC': { d: 300, bf: 90, tf: 16, tw: 8, r1: 14 },
            '250PFC': { d: 250, bf: 90, tf: 15, tw: 8, r1: 12 },
            '200PFC': { d: 200, bf: 75, tf: 12, tw: 6, r1: 12 },
            '150PFC': { d: 150, bf: 75, tf: 9.5, tw: 6, r1: 10 },
        };
        for (const [k, v] of Object.entries(pfcSections)) {
            if (memberSize.includes(k.replace('PFC', ''))) return { type: 'PFC', data: v };
        }
        return { type: 'PFC', data: { d: 200, bf: 75, tf: 12, tw: 6, r1: 12 } };
    }

    if (category === 'TIMBER' || category === 'GLT' || /MGP|GL\d|LVL/i.test(memberSize)) {
        let ply = 1;
        let dimStr = memberSize;
        const plyMatch = memberSize.match(/^(\d+)\//);
        if (plyMatch) { ply = parseInt(plyMatch[1]); dimStr = memberSize.substring(plyMatch[0].length); }
        const dimMatch = dimStr.match(/(\d+)x(\d+)/);
        if (dimMatch) {
            return { type: 'RECT', width: parseFloat(dimMatch[2]) * ply, depth: parseFloat(dimMatch[1]),
                     material: 'timber' };
        }
    }

    const eaMatch = memberSize.match(/(\d+)x(\d+)x(\d+)/);
    if (eaMatch && (category === 'EA' || /EA/i.test(memberSize))) {
        return { type: 'RECT', width: parseFloat(eaMatch[1]), depth: parseFloat(eaMatch[2]), material: 'steel' };
    }

    const fallback = memberSize.match(/(\d+)x(\d+)/);
    if (fallback) {
        return { type: 'RECT', width: parseFloat(fallback[2]) || 50, depth: parseFloat(fallback[1]) || 50 };
    }

    return null;
}

/** Add mesh + edge lines to scene, storing element reference */
function addMeshWithEdges(geo, mat, position, rotation, elementRef) {
    const mesh = new THREE.Mesh(geo, mat);
    mesh.position.copy(position);
    if (rotation) {
        if (rotation.y !== undefined) mesh.rotation.y = rotation.y;
        if (rotation.x !== undefined) mesh.rotation.x = rotation.x;
        if (rotation.z !== undefined) mesh.rotation.z = rotation.z;
    }
    mesh.userData.elementRef = elementRef;
    scene3d.add(mesh);

    const edgeGeo = new THREE.EdgesGeometry(geo, 15);
    const edgeMesh = new THREE.LineSegments(edgeGeo, EDGE_MAT());
    edgeMesh.position.copy(position);
    if (rotation) {
        if (rotation.y !== undefined) edgeMesh.rotation.y = rotation.y;
        if (rotation.x !== undefined) edgeMesh.rotation.x = rotation.x;
        if (rotation.z !== undefined) edgeMesh.rotation.z = rotation.z;
    }
    edgeMesh.userData.isEdge = true;
    scene3d.add(edgeMesh);

    return mesh;
}

// ── 3D Camera Controller (Trimble-style damped spherical) ──

const orbitCtrl = {
    theta: 0.62, phi: 1.18,
    sTheta: 0.62, sPhi: 1.18,
    radius: 30, sRadius: 30,
    target: new THREE.Vector3(),
    sTarget: new THREE.Vector3(),
    damping: 0.13,
    rotSp: 0.007,
    panSp: 1.0,
    zoomSp: 0.0012,
    minR: 0.5,
    maxR: 500,
};

function clamp3(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }

function updateCamera3D(cam, ctrl) {
    const a = ctrl.damping;
    ctrl.sTheta += (ctrl.theta - ctrl.sTheta) * a;
    ctrl.sPhi += (ctrl.phi - ctrl.sPhi) * a;
    ctrl.sRadius += (ctrl.radius - ctrl.sRadius) * a;
    ctrl.sTarget.lerp(ctrl.target, a);

    const r = ctrl.sRadius;
    cam.position.set(
        ctrl.sTarget.x + r * Math.sin(ctrl.sPhi) * Math.cos(ctrl.sTheta),
        ctrl.sTarget.y + r * Math.cos(ctrl.sPhi),
        ctrl.sTarget.z + r * Math.sin(ctrl.sPhi) * Math.sin(ctrl.sTheta)
    );
    cam.lookAt(ctrl.sTarget);
    cam.updateProjectionMatrix();
}

/** Compute bounding box of structural elements only (ignores ground, grids, labels) */
function getStructuralBBox() {
    const bbox = new THREE.Box3();
    scene3d.traverse(obj => {
        if (obj.isMesh && obj.userData.elementRef && !obj.userData.isGround && !obj.userData.isFloor) {
            bbox.expandByObject(obj);
        }
    });
    return bbox;
}

function frame3D(ctrl, cam, snap) {
    if (!scene3d) return;
    // Use structural-only bbox for tight framing
    let bbox = getStructuralBBox();
    if (bbox.isEmpty()) bbox = new THREE.Box3().setFromObject(scene3d); // fallback
    if (bbox.isEmpty()) return;
    const sz = new THREE.Vector3(); bbox.getSize(sz);
    const ctr = new THREE.Vector3(); bbox.getCenter(ctr);
    ctrl.target.copy(ctr);
    const half = Math.max(sz.x, sz.y, sz.z) * 0.5;
    ctrl.radius = clamp3(half / Math.tan((cam.fov || 38) * Math.PI / 360) * 1.4, ctrl.minR, ctrl.maxR);
    if (snap) {
        ctrl.sRadius = ctrl.radius;
        ctrl.sTarget.copy(ctrl.target);
    }
}

/** Frame tightly on a single member */
function frameMember(ctrl, cam, mesh) {
    const bb = new THREE.Box3().setFromObject(mesh);
    if (bb.isEmpty()) return;
    const sz = new THREE.Vector3(); bb.getSize(sz);
    const ctr = new THREE.Vector3(); bb.getCenter(ctr);
    ctrl.target.copy(ctr);
    const half = Math.max(sz.x, sz.y, sz.z) * 0.5;
    ctrl.radius = clamp3(half / Math.tan((cam.fov || 38) * Math.PI / 360) * 2.0, ctrl.minR, ctrl.maxR);
}

function reset3D(ctrl, cam, snap) {
    ctrl.theta = 0.62;
    ctrl.phi = 1.18;
    frame3D(ctrl, cam, false);
    if (snap) {
        ctrl.sTheta = ctrl.theta;
        ctrl.sPhi = ctrl.phi;
        ctrl.sRadius = ctrl.radius;
        ctrl.sTarget.copy(ctrl.target);
    }
}

// ── 3D Text Sprites ──────────────────────────────────────

function createTextSprite(text, color, fontSize) {
    if (typeof THREE === 'undefined') return null;
    const cv = document.createElement('canvas');
    const cx = cv.getContext('2d');
    const fs = fontSize || 28;
    cv.width = 512; cv.height = 64;
    cx.fillStyle = 'rgba(255,255,255,0.88)';
    cx.fillRect(0, 0, 512, 64);
    cx.font = `bold ${fs}px "Segoe UI", Arial, sans-serif`;
    cx.fillStyle = color || '#555555';
    cx.textAlign = 'left';
    cx.textBaseline = 'middle';
    cx.fillText(text, 8, 32);
    const tex = new THREE.CanvasTexture(cv);
    const mat = new THREE.SpriteMaterial({ map: tex, transparent: true, depthWrite: false });
    const sprite = new THREE.Sprite(mat);
    sprite.scale.set(4, 1, 1);
    sprite.userData.isLabel = true;
    return sprite;
}

// ── Init 3D Scene & Renderer ─────────────────────────────

let snapshotRenderer = null; // offscreen renderer for Take Picture

function init3DViewport() {
    if (typeof THREE === 'undefined') {
        console.warn('[3D] Three.js not loaded');
        return;
    }

    orbitCtrl.target = new THREE.Vector3();
    orbitCtrl.sTarget = new THREE.Vector3();

    // Offscreen canvas for snapshot capture
    isoView.canvas = document.createElement('canvas');
    isoView.canvas.width = 1600;
    isoView.canvas.height = 1200;

    // Scene
    scene3d = new THREE.Scene();
    scene3d.background = new THREE.Color(0xFFFFFF); // white for clean snapshot

    // Camera
    camera3d = new THREE.PerspectiveCamera(38, 1600 / 1200, 0.01, 2000);

    // Offscreen renderer for snapshots
    snapshotRenderer = new THREE.WebGLRenderer({
        canvas: isoView.canvas,
        antialias: true,
        alpha: false,
        preserveDrawingBuffer: true
    });
    snapshotRenderer.setSize(1600, 1200);
    snapshotRenderer.setPixelRatio(1); // fixed for consistent snapshots
    snapshotRenderer.shadowMap.enabled = true;
    snapshotRenderer.shadowMap.type = THREE.PCFSoftShadowMap;

    // Lighting
    const ambLight = new THREE.AmbientLight(0xffffff, 0.55);
    scene3d.add(ambLight);

    const dirLight1 = new THREE.DirectionalLight(0xffffff, 0.75);
    dirLight1.position.set(30, 50, 40);
    dirLight1.castShadow = true;
    dirLight1.shadow.mapSize.width = 2048;
    dirLight1.shadow.mapSize.height = 2048;
    dirLight1.shadow.camera.near = 0.1;
    dirLight1.shadow.camera.far = 200;
    dirLight1.shadow.camera.left = -50;
    dirLight1.shadow.camera.right = 50;
    dirLight1.shadow.camera.top = 50;
    dirLight1.shadow.camera.bottom = -50;
    dirLight1.shadow.bias = -0.002;
    scene3d.add(dirLight1);

    const dirLight2 = new THREE.DirectionalLight(0xffffff, 0.30);
    dirLight2.position.set(-20, -15, -25);
    scene3d.add(dirLight2);

    const hemiLight = new THREE.HemisphereLight(0xDDE4EC, 0x8B7D6B, 0.25);
    scene3d.add(hemiLight);

    scene3d.userData.dirLight1 = dirLight1;
}

// ── Rebuild 3D Scene ─────────────────────────────────────

function rebuild3DScene() {
    if (!scene3d || typeof THREE === 'undefined') return;

    const scale3d = 1 / 1000;

    // Clear scene (keep lights and camera)
    const toRemove = [];
    scene3d.traverse(obj => {
        if (obj.isMesh || obj.isLine || obj.isLineSegments || obj.isSprite || obj.isGroup) {
            toRemove.push(obj);
        }
    });
    toRemove.forEach(obj => {
        if (obj.geometry) obj.geometry.dispose();
        if (obj.material) {
            if (obj.material.map) obj.material.map.dispose();
            obj.material.dispose();
        }
        scene3d.remove(obj);
    });

    const da = engine.coords.drawArea;
    const drawW = da.width * project.drawingScale * scale3d;
    const drawD = da.height * project.drawingScale * scale3d;

    // ── Ground Plane (very subtle, for shadows only) ──
    const groundSize = Math.max(drawW, drawD) * 2.5;
    const groundGeo = new THREE.PlaneGeometry(groundSize, groundSize);
    const groundMat = new THREE.MeshStandardMaterial({
        color: 0xFFFFFF, transparent: true, opacity: 0.01,
        roughness: 1.0, side: THREE.DoubleSide, depthWrite: false
    });
    const ground = new THREE.Mesh(groundGeo, groundMat);
    ground.rotation.x = -Math.PI / 2;
    ground.position.y = -0.005;
    ground.receiveShadow = true;
    ground.userData.isGround = true;
    scene3d.add(ground);

    // ── Level Labels (no floor planes — just text references) ──
    for (let li = 0; li < levelSystem.levels.length; li++) {
        const lv = levelSystem.levels[li];
        const elev = lv.elevation * scale3d;

        const labelText = lv.name + ' (RL ' + (lv.elevation / 1000).toFixed(1) + 'm)';
        const sprite = createTextSprite(labelText, '#AAAAAA');
        if (sprite) {
            sprite.position.set(-1.5, elev + 0.15, drawD / 2);
            scene3d.add(sprite);
        }
    }

    // ── Structural Grid Lines (vertical projection) ──
    if (typeof structuralGrids !== 'undefined') {
        const maxElev = levelSystem.levels.reduce((m, l) => Math.max(m, l.elevation + (l.height || 0)), 0) * scale3d;
        const gridLineMat = new THREE.LineBasicMaterial({ color: 0xCCCCCC, transparent: true, opacity: 0.25 });

        for (const g of structuralGrids) {
            const pos = g.position * scale3d;
            let pts;
            if (g.axis === 'V') {
                pts = [new THREE.Vector3(pos, -0.5, -2), new THREE.Vector3(pos, maxElev + 1, -2)];
            } else {
                pts = [new THREE.Vector3(-2, -0.5, pos), new THREE.Vector3(-2, maxElev + 1, pos)];
            }
            const gGeo = new THREE.BufferGeometry().setFromPoints(pts);
            scene3d.add(new THREE.Line(gGeo, gridLineMat));

            const bubbleSprite = createTextSprite(g.label, '#AAAAAA', 22);
            if (bubbleSprite) {
                if (g.axis === 'V') {
                    bubbleSprite.position.set(pos, maxElev + 1.5, -2);
                } else {
                    bubbleSprite.position.set(-2, maxElev + 1.5, pos);
                }
                scene3d.add(bubbleSprite);
            }
        }
    }

    // ── Process Elements (ALL levels for full building view) ──
    const visibleElements = project.elements.filter(el => {
        const layer = project.layers[el.layer];
        return layer && layer.visible;
    });

    for (const el of visibleElements) {
        const lv = levelSystem.levels.find(l => l.id === el.level) || getActiveLevel();
        const elev = lv.elevation * scale3d;

        // ── BEAMS ──
        if (el.type === 'line' && el.layer === 'S-BEAM') {
            const x1 = el.x1 * scale3d, z1 = el.y1 * scale3d;
            const x2 = el.x2 * scale3d, z2 = el.y2 * scale3d;
            const dx = x2 - x1, dz = z2 - z1;
            const len = Math.sqrt(dx * dx + dz * dz);
            const angle = Math.atan2(dz, dx);
            const mx = (x1 + x2) / 2, mz = (z1 + z2) / 2;
            const lenMM = len * 1000;

            const parsed = parseMemberFor3D(el.memberSize, el.memberCategory);
            let geo, mat;

            if (parsed) {
                if (parsed.type === 'UB') {
                    geo = buildIBeamGeo(parsed.data, lenMM);
                    mat = MAT_STEEL();
                } else if (parsed.type === 'SHS') {
                    geo = buildSHSGeo(parsed.B, parsed.B, parsed.t, lenMM);
                    mat = MAT_STEEL();
                } else if (parsed.type === 'RHS') {
                    geo = buildSHSGeo(parsed.B, parsed.D, parsed.t, lenMM);
                    mat = MAT_STEEL();
                } else if (parsed.type === 'CHS') {
                    geo = buildCHSGeo(parsed.D, parsed.t, lenMM);
                    mat = MAT_STEEL();
                } else if (parsed.type === 'PFC') {
                    geo = buildPFCGeo(parsed.data, lenMM);
                    mat = MAT_STEEL();
                } else if (parsed.type === 'RECT') {
                    geo = buildRectGeo(parsed.width, parsed.depth, lenMM);
                    mat = (parsed.material === 'timber') ? MAT_TIMBER() : MAT_STEEL();
                }
            }

            if (!geo) {
                geo = new THREE.BoxGeometry(len, 0.3 * scale3d * 1000, 0.15 * scale3d * 1000);
                mat = MAT_STEEL();
            }

            const beamDepth = geo.boundingBox ? (geo.boundingBox.max.y - geo.boundingBox.min.y) : 0.2;
            const pos = new THREE.Vector3(mx, elev - beamDepth / 2, mz);
            const mesh = addMeshWithEdges(geo, mat, pos, { y: -angle }, el);
            mesh.castShadow = true;
        }

        // ── WALLS ──
        if (el.type === 'wall') {
            const x1 = el.x1 * scale3d, z1 = el.y1 * scale3d;
            const x2 = el.x2 * scale3d, z2 = el.y2 * scale3d;
            const dx = x2 - x1, dz = z2 - z1;
            const len = Math.sqrt(dx * dx + dz * dz);
            const angle = Math.atan2(dz, dx);
            const mx = (x1 + x2) / 2, mz = (z1 + z2) / 2;

            const wallThick = (el.thickness || 90) * scale3d;
            const wallH = (lv.height || 2700) * scale3d;

            let wallMat;
            if (el.wallType === 'concrete') wallMat = MAT_CONCRETE();
            else if (el.wallType === 'block') wallMat = MAT_BLOCK();
            else wallMat = MAT_STUD();

            const wallGeo = new THREE.BoxGeometry(len, wallH, wallThick);
            const pos = new THREE.Vector3(mx, elev + wallH / 2, mz);
            const mesh = addMeshWithEdges(wallGeo, wallMat, pos, { y: -angle }, el);
            mesh.castShadow = true;
            mesh.receiveShadow = true;
        }

        // ── COLUMNS ──
        if (el.type === 'column') {
            const cx = el.x * scale3d, cz = el.y * scale3d;
            const colSize = (el.size || 89) * scale3d;

            const lvIdx = levelSystem.levels.findIndex(l => l.id === el.level);
            const dir = el.extends || 'below';
            let bottomElev = elev, topElev2 = elev;

            if (dir === 'below' && lvIdx > 0) {
                bottomElev = levelSystem.levels[lvIdx - 1].elevation * scale3d;
                topElev2 = elev;
            } else if (dir === 'above') {
                bottomElev = elev;
                const nextLv = lvIdx < levelSystem.levels.length - 1 ? levelSystem.levels[lvIdx + 1] : lv;
                topElev2 = (nextLv.elevation || elev * 1000 + (lv.height || 2700)) * scale3d;
            } else if (dir === 'both') {
                bottomElev = lvIdx > 0 ? levelSystem.levels[lvIdx - 1].elevation * scale3d : elev;
                topElev2 = lvIdx < levelSystem.levels.length - 1 ? levelSystem.levels[lvIdx + 1].elevation * scale3d : elev + (lv.height || 2700) * scale3d;
            } else {
                topElev2 = elev + (lv.height || 2700) * scale3d;
            }

            const colH = Math.max(0.1, topElev2 - bottomElev);

            const parsed = parseMemberFor3D(el.memberSize, el.memberCategory);
            let colGeo;

            if (parsed && parsed.type === 'SHS') {
                colGeo = buildSHSGeo(parsed.B, parsed.B, parsed.t, colH * 1000);
                colGeo.rotateZ(Math.PI / 2);
            } else if (parsed && parsed.type === 'CHS') {
                colGeo = buildCHSGeo(parsed.D, parsed.t, colH * 1000);
                colGeo.rotateZ(Math.PI / 2);
            } else if (parsed && (parsed.type === 'UB')) {
                colGeo = buildIBeamGeo(parsed.data, colH * 1000);
                colGeo.rotateZ(Math.PI / 2);
            } else {
                colGeo = new THREE.BoxGeometry(colSize, colH, colSize);
            }

            const pos = new THREE.Vector3(cx, bottomElev + colH / 2, cz);
            const mesh = addMeshWithEdges(colGeo, MAT_STEEL(), pos, {}, el);
            mesh.castShadow = true;
        }

        // ── SLABS ──
        if (el.type === 'polyline' && el.layer === 'S-SLAB' && (el.closed || (el.points && el.points.length > 3))) {
            try {
                // Shape is created in XY plane; after rotateX(-PI/2), shape-Y maps to -Z.
                // Beams/columns use positive Z for 2D Y, so negate Y here to compensate.
                const pts2d = el.points.map(p => new THREE.Vector2(p.x * scale3d, -p.y * scale3d));
                const shape = new THREE.Shape(pts2d);
                const slabThick = (el.slabThickness || 200) * scale3d;
                // ExtrudeGeometry extrudes in +Z from shape plane.
                // After rotateX(-PI/2), +Z becomes -Y (downward) — correct for slab below beams.
                const slabGeo = new THREE.ExtrudeGeometry(shape, {
                    depth: slabThick, bevelEnabled: false, steps: 1,
                });
                slabGeo.rotateX(-Math.PI / 2);

                const slabMat = MAT_SLAB();
                const pos = new THREE.Vector3(0, elev, 0);
                const mesh = addMeshWithEdges(slabGeo, slabMat, pos, {}, el);
                mesh.castShadow = true;
                mesh.receiveShadow = true;
            } catch (e) { /* skip invalid slab geometry */ }
        }

        // ── FOOTINGS ──
        if (el.type === 'footing') {
            const fx = el.x * scale3d, fz = el.y * scale3d;
            const fw = (el.footingWidth || el.width || 1000) * scale3d;
            const fd = (el.footingDepth || el.depth || 400) * scale3d;

            // Top of footing auto-calculated from slab thickness + sub-slab gap
            const topElev = getFootingTopElevation(el) * scale3d;
            const ftgGeo = new THREE.BoxGeometry(fw, fd, fw);
            const pos = new THREE.Vector3(fx, topElev - fd / 2, fz);
            addMeshWithEdges(ftgGeo, MAT_FOOTING(), pos, {}, el);
        }

        // ── STRIP FOOTINGS ──
        if (el.type === 'stripFooting') {
            const x1 = el.x1 * scale3d, z1 = el.y1 * scale3d;
            const x2 = el.x2 * scale3d, z2 = el.y2 * scale3d;
            const dx = x2 - x1, dz = z2 - z1;
            const len = Math.sqrt(dx * dx + dz * dz);
            const angle = Math.atan2(dz, dx);
            const mx = (x1 + x2) / 2, mz = (z1 + z2) / 2;

            const sfW = (el.footingWidth || 300) * scale3d;
            const sfD = (el.footingDepth || 500) * scale3d;
            // Top of footing auto-calculated from slab thickness + sub-slab gap
            const topElev = getFootingTopElevation(el) * scale3d;
            const sfGeo = new THREE.BoxGeometry(len, sfD, sfW);
            const pos = new THREE.Vector3(mx, topElev - sfD / 2, mz);
            addMeshWithEdges(sfGeo, MAT_FOOTING(), pos, { y: -angle }, el);
        }

        // ── EDGE THICKENINGS ──
        if (el.type === 'edge' || (el.type === 'line' && el.layer === 'S-SLAB')) {
            const x1 = el.x1 * scale3d, z1 = el.y1 * scale3d;
            const x2 = el.x2 * scale3d, z2 = el.y2 * scale3d;
            const dx = x2 - x1, dz = z2 - z1;
            const len = Math.sqrt(dx * dx + dz * dz);
            const angle = Math.atan2(dz, dx);
            const mx = (x1 + x2) / 2, mz = (z1 + z2) / 2;

            const edgeW = 0.3;
            const edgeD = 0.45;
            const edgeGeo = new THREE.BoxGeometry(len, edgeD, edgeW);
            const pos = new THREE.Vector3(mx, elev - edgeD / 2, mz);
            addMeshWithEdges(edgeGeo, MAT_CONCRETE(), pos, { y: -angle }, el);
        }
    }

    // ── ROOF SURFACES ──
    // Priority: skeleton model > legacy ridgeLine model
    {
        const rfLevel = levelSystem.levels.find(l => l.id === 'RF' || l.name.toLowerCase().includes('roof'));
        const eavesElev = rfLevel ? rfLevel.elevation
            : (levelSystem.levels.length > 0 ? levelSystem.levels[levelSystem.levels.length - 1].elevation : 2700);

        let roofModel = null;

        // 1. Skeleton model (preferred)
        const skeleton = typeof findSkeletonElement === 'function'
            ? project.elements.find(el => el.type === 'roofSkeleton')
            : null;
        if (skeleton && skeleton.committed && skeleton.faces &&
            skeleton.faces.some(f => f.pitch != null) &&
            typeof buildRoofModelFromSkeleton === 'function') {
            roofModel = buildRoofModelFromSkeleton(skeleton, eavesElev);
        }

        // 2. Legacy ridgeLine fallback
        if (!roofModel && typeof buildRoofModel === 'function') {
            const envelope = project.elements.find(el => el.type === 'buildingEnvelope');
            const ridge    = project.elements.find(el => el.type === 'ridgeLine');
            if (envelope && ridge) {
                roofModel = buildRoofModel(envelope, ridge, eavesElev);
            }
        }

        if (roofModel) {
            // Helper: add a polygon mesh to the scene
            const _addPolyMesh = (vertices, material) => {
                if (!vertices || vertices.length < 3) return;
                const positions = [];
                for (const v of vertices) {
                    positions.push(v.x * scale3d, (v.z || 0) * scale3d, v.y * scale3d);
                }
                const indices = [];
                // Fan triangulation
                for (let i = 1; i < vertices.length - 1; i++) indices.push(0, i, i + 1);
                const geo = new THREE.BufferGeometry();
                geo.setAttribute('position', new THREE.Float32BufferAttribute(positions, 3));
                geo.setIndex(indices);
                geo.computeVertexNormals();
                const mesh = new THREE.Mesh(geo, material);
                mesh.userData.elementRef = 'roof';
                scene3d.add(mesh);
                const edgeGeo = new THREE.EdgesGeometry(geo);
                const edgeMesh = new THREE.LineSegments(edgeGeo, new THREE.LineBasicMaterial({ color: 0x994433, linewidth: 1 }));
                scene3d.add(edgeMesh);
            };

            // Wall shell
            for (const face of roofModel.wallFaces) {
                if (!face.vertices || face.vertices.length < 3) continue;
                const positions = [];
                for (const v of face.vertices) {
                    positions.push(v.x * scale3d, (v.z || 0) * scale3d, v.y * scale3d);
                }
                const indices = face.vertices.length === 4
                    ? [0, 1, 2, 0, 2, 3]
                    : Array.from({ length: face.vertices.length - 2 }, (_, i) => [0, i + 1, i + 2]).flat();
                const geo = new THREE.BufferGeometry();
                geo.setAttribute('position', new THREE.Float32BufferAttribute(positions, 3));
                geo.setIndex(indices);
                geo.computeVertexNormals();
                const wallMesh = new THREE.Mesh(geo, MAT_ENVWALL());
                wallMesh.userData.elementRef = 'roof';
                scene3d.add(wallMesh);
            }

            // Roof surfaces
            for (const surface of roofModel.surfaces) {
                _addPolyMesh(surface.vertices, MAT_ROOF());
            }

            // Ridge skeleton lines in 3D (skeleton model only)
            if (roofModel.isSkeletonModel && skeleton && roofModel.nodeHeights) {
                const nodeMap = {};
                for (const n of skeleton.nodes) nodeMap[n.id] = n;
                for (const edge of skeleton.edges) {
                    if (edge.source === 'envelope') continue;
                    const nA = nodeMap[edge.a], nB = nodeMap[edge.b];
                    if (!nA || !nB) continue;
                    const hA = roofModel.nodeHeights[edge.a] ?? eavesElev;
                    const hB = roofModel.nodeHeights[edge.b] ?? eavesElev;
                    const pts = [
                        new THREE.Vector3(nA.x * scale3d, hA * scale3d, nA.y * scale3d),
                        new THREE.Vector3(nB.x * scale3d, hB * scale3d, nB.y * scale3d),
                    ];
                    const geo = new THREE.BufferGeometry().setFromPoints(pts);
                    scene3d.add(new THREE.Line(geo, new THREE.LineBasicMaterial({ color: 0xdc2626, linewidth: 2 })));
                }
            }
        }
    }

    // ── Update shadow camera to fit scene ──
    const dl = scene3d.userData.dirLight1;
    if (dl) {
        const bbox = new THREE.Box3().setFromObject(scene3d);
        if (!bbox.isEmpty()) {
            const sz = new THREE.Vector3(); bbox.getSize(sz);
            const ctr = new THREE.Vector3(); bbox.getCenter(ctr);
            const maxDim = Math.max(sz.x, sz.y, sz.z);
            dl.shadow.camera.left = -maxDim;
            dl.shadow.camera.right = maxDim;
            dl.shadow.camera.top = maxDim;
            dl.shadow.camera.bottom = -maxDim;
            dl.shadow.camera.far = maxDim * 4;
            dl.shadow.camera.updateProjectionMatrix();
            dl.position.set(ctr.x + maxDim, ctr.y + maxDim * 1.5, ctr.z + maxDim * 1.2);
            dl.target.position.copy(ctr);
            dl.target.updateMatrixWorld();
        }
    }
}

// ══════════════════════════════════════════════════════════
// ── ON-SHEET ISOMETRIC (frozen snapshot image) ───────────
// ══════════════════════════════════════════════════════════

// Take a snapshot of current 3D view for on-sheet display
function takeSnapshot() {
    if (!snapshotRenderer || !scene3d || !camera3d) return;

    // Use white background for clean sheet appearance
    const oldBg = scene3d.background;
    scene3d.background = new THREE.Color(0xFFFFFF);

    // Render to offscreen canvas
    camera3d.aspect = 1600 / 1200;
    camera3d.updateProjectionMatrix();
    snapshotRenderer.render(scene3d, camera3d);

    // Store as image data URL
    isoView.snapshotDataURL = isoView.canvas.toDataURL('image/png');

    // Create Image object for drawing onto the sheet
    const img = new Image();
    img.onload = () => {
        isoView.snapshotImage = img;
        engine.requestRender();
    };
    img.src = isoView.snapshotDataURL;

    // Restore scene background
    scene3d.background = oldBg;
}

// ── Live Update: re-render snapshot when elements change ──
// Stores the camera state from the last Take Picture (or initial auto-frame)
const snapshotCamera = {
    theta: orbitCtrl.theta,
    phi: orbitCtrl.phi,
    radius: orbitCtrl.radius,
    target: new THREE.Vector3(),
    locked: false, // true once user manually clicks Take Picture; false = auto-frame each update
};

let _liveUpdateTimer = null;
let _lastElementHash = '';

/** Compute a quick hash of element count + types to detect changes */
function elementHash() {
    const els = project.elements;
    if (!els || !els.length) return '0';
    // Hash based on count + last element's coordinates (fast change detection)
    const last = els[els.length - 1];
    return els.length + ':' + (last.x1 || last.x || 0).toFixed(0) + ',' + (last.y1 || last.y || 0).toFixed(0);
}

/** Re-render the on-sheet snapshot from the stored camera angle */
function liveUpdateSnapshot() {
    if (!snapshotRenderer || !scene3d || !camera3d || fullscreen3d) return;

    // Rebuild 3D scene with latest elements
    rebuild3DScene();

    // Check if there are any structural meshes to show
    const bbox = getStructuralBBox();
    if (bbox.isEmpty()) return;

    // Save current camera state (might be in fullscreen orbit)
    const savedTheta = orbitCtrl.theta, savedPhi = orbitCtrl.phi;
    const savedRadius = orbitCtrl.radius;
    const savedTarget = orbitCtrl.target.clone();
    const savedSTheta = orbitCtrl.sTheta, savedSPhi = orbitCtrl.sPhi;
    const savedSRadius = orbitCtrl.sRadius;
    const savedSTarget = orbitCtrl.sTarget.clone();

    if (snapshotCamera.locked) {
        // User has manually taken a photo — use their locked camera angle
        orbitCtrl.theta = snapshotCamera.theta;
        orbitCtrl.phi = snapshotCamera.phi;
        orbitCtrl.radius = snapshotCamera.radius;
        orbitCtrl.sTheta = snapshotCamera.theta;
        orbitCtrl.sPhi = snapshotCamera.phi;
        orbitCtrl.sRadius = snapshotCamera.radius;
        if (snapshotCamera.target) {
            orbitCtrl.target.copy(snapshotCamera.target);
            orbitCtrl.sTarget.copy(snapshotCamera.target);
        }
    } else {
        // No manual photo yet — auto-frame to fit all structural elements
        frame3D(orbitCtrl, camera3d, true);
    }

    // Update camera position and render
    updateCamera3D(camera3d, orbitCtrl);
    takeSnapshot();

    // Make the iso view visible (auto-show on first element)
    isoView.visible = true;

    // Restore camera state
    orbitCtrl.theta = savedTheta; orbitCtrl.phi = savedPhi;
    orbitCtrl.radius = savedRadius; orbitCtrl.target.copy(savedTarget);
    orbitCtrl.sTheta = savedSTheta; orbitCtrl.sPhi = savedSPhi;
    orbitCtrl.sRadius = savedSRadius; orbitCtrl.sTarget.copy(savedSTarget);
}

// Hook into engine render loop — check for element changes with debounce
engine.onRender(() => {
    if (fullscreen3d) return; // don't update while in fullscreen
    const hash = elementHash();
    if (hash !== _lastElementHash) {
        _lastElementHash = hash;
        if (_liveUpdateTimer) clearTimeout(_liveUpdateTimer);
        _liveUpdateTimer = setTimeout(() => {
            liveUpdateSnapshot();
        }, 250); // 250ms debounce after last change
    }
});

// On-sheet compositing: draw frozen snapshot
engine.onRender((ctx, eng) => {
    if (!isoView.visible || !isoView.snapshotImage) return;

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    const tl = coords.sheetToScreen(isoView.x, isoView.y);
    const br = coords.sheetToScreen(isoView.x + isoView.width, isoView.y + isoView.height);
    const sx = tl.x, sy = tl.y;
    const sw = br.x - tl.x, sh = br.y - tl.y;

    // Draw the frozen snapshot image
    ctx.drawImage(isoView.snapshotImage, sx, sy, sw, sh);

    // Dotted border only when hovered
    if (isoView.showBorder) {
        ctx.strokeStyle = '#BBBBBB';
        ctx.lineWidth = Math.max(0.5, 0.35 * zoom);
        ctx.setLineDash([4 * zoom, 3 * zoom]);
        ctx.strokeRect(sx, sy, sw, sh);
        ctx.setLineDash([]);

        // Resize handle (bottom-right corner)
        const rhSize = 8;
        ctx.fillStyle = 'rgba(150,150,150,0.5)';
        ctx.beginPath();
        ctx.moveTo(sx + sw, sy + sh);
        ctx.lineTo(sx + sw - rhSize, sy + sh);
        ctx.lineTo(sx + sw, sy + sh - rhSize);
        ctx.closePath();
        ctx.fill();
    }

    // Title below
    const titleFontSize = Math.max(1, 3.5 * zoom);
    ctx.fillStyle = '#000000';
    ctx.font = `bold ${titleFontSize}px "Segoe UI", Arial, sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText('ISOMETRIC VIEW', sx + sw / 2, sy + sh + 2 * zoom);

    const tw2 = ctx.measureText('ISOMETRIC VIEW').width;
    ctx.strokeStyle = '#000000';
    ctx.lineWidth = Math.max(0.5, 0.2 * zoom);
    ctx.beginPath();
    ctx.moveTo(sx + sw / 2 - tw2 / 2, sy + sh + 2 * zoom + titleFontSize + 1);
    ctx.lineTo(sx + sw / 2 + tw2 / 2, sy + sh + 2 * zoom + titleFontSize + 1);
    ctx.stroke();

    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';
});

// ── On-Sheet Interaction (drag, resize, double-click to enter 3D) ──

function isInIsoView(screenX, screenY) {
    if (!isoView.visible || !isoView.snapshotImage) return false;
    const tl = engine.coords.sheetToScreen(isoView.x, isoView.y);
    const br = engine.coords.sheetToScreen(isoView.x + isoView.width, isoView.y + isoView.height);
    return screenX >= tl.x && screenX <= br.x && screenY >= tl.y && screenY <= br.y;
}

let isoViewHovered = false;

// Hover — show/hide border
container.addEventListener('mousemove', (e) => {
    if (!isoView.visible || !isoView.snapshotImage) return;
    const rect = container.getBoundingClientRect();
    const mx = e.clientX - rect.left, my = e.clientY - rect.top;
    const wasHovered = isoViewHovered;
    isoViewHovered = isInIsoView(mx, my);

    // Show dotted border on hover
    if (isoViewHovered !== wasHovered) {
        isoView.showBorder = isoViewHovered;
        engine.requestRender();
    }

    // Drag iso viewport
    if (isoView.dragging) {
        const sheetPos = engine.coords.screenToSheet(mx, my);
        isoView.x = isoView.dragOrigPos.x + (sheetPos.x - isoView.dragStart.x);
        isoView.y = isoView.dragOrigPos.y + (sheetPos.y - isoView.dragStart.y);
        engine.requestRender();
        return;
    }
    // Resize
    if (isoView.resizing) {
        const sheetPos = engine.coords.screenToSheet(mx, my);
        const newW = Math.max(80, isoView.resizeOrigSize.w + (sheetPos.x - isoView.resizeStart.x));
        const newH = newW * (isoView.resizeOrigSize.h / isoView.resizeOrigSize.w); // maintain aspect
        isoView.width = newW;
        isoView.height = newH;
        engine.requestRender();
        return;
    }

    // Cursor
    if (isoViewHovered) {
        const br2 = engine.coords.sheetToScreen(isoView.x + isoView.width, isoView.y + isoView.height);
        if (Math.abs(mx - br2.x) < 12 && Math.abs(my - br2.y) < 12) {
            container.style.cursor = 'nwse-resize';
        } else {
            container.style.cursor = 'grab';
        }
    }
});

// Mousedown — drag or resize
container.addEventListener('mousedown', (e) => {
    if (!isoView.visible || !isoViewHovered) return;

    const rect = container.getBoundingClientRect();
    const mx = e.clientX - rect.left, my = e.clientY - rect.top;
    const sheetPos = engine.coords.screenToSheet(mx, my);

    // Check resize handle (bottom-right corner)
    const br = engine.coords.sheetToScreen(isoView.x + isoView.width, isoView.y + isoView.height);
    if (Math.abs(mx - br.x) < 14 && Math.abs(my - br.y) < 14) {
        isoView.resizing = true;
        isoView.resizeStart = { x: sheetPos.x, y: sheetPos.y };
        isoView.resizeOrigSize = { w: isoView.width, h: isoView.height };
        e.stopPropagation();
        e.preventDefault();
        return;
    }

    // Drag to reposition
    isoView.dragging = true;
    isoView.dragStart = { x: sheetPos.x, y: sheetPos.y };
    isoView.dragOrigPos = { x: isoView.x, y: isoView.y };
    container.style.cursor = 'grabbing';
    e.stopPropagation();
    e.preventDefault();
});

window.addEventListener('mouseup', () => {
    if (isoView.dragging || isoView.resizing) {
        isoView.dragging = false;
        isoView.resizing = false;
        container.style.cursor = 'default';
    }
});

// Double-click on iso image → enter fullscreen 3D
container.addEventListener('dblclick', (e) => {
    if (!isoView.visible || !isoViewHovered) return;
    e.preventDefault();
    e.stopPropagation();
    enterFullscreen3D();
});

// ── 3D Button → Fullscreen 3D ────────────────────────────

const btn3d = document.getElementById('btn-3d');
btn3d.addEventListener('click', () => {
    if (!isoView.visible) {
        isoView.visible = true;
        btn3d.classList.add('active');
        // Rebuild and take initial snapshot
        rebuild3DScene();
        frame3D(orbitCtrl, camera3d, true);
        updateCamera3D(camera3d, orbitCtrl);
        takeSnapshot();
    }
    enterFullscreen3D();
});

// ══════════════════════════════════════════════════════════
// ── FULLSCREEN 3D VIEWER ─────────────────────────────────
// ══════════════════════════════════════════════════════════

let fullscreen3d = false;
let fullscreen3dDiv = null;
let fsRenderer = null;
let fsAnimId = null;

function enterFullscreen3D() {
    if (fullscreen3d || typeof THREE === 'undefined') return;
    fullscreen3d = true;

    // Rebuild scene with latest elements
    rebuild3DScene();

    // Switch scene bg to dark for the Trimble feel
    scene3d.background = new THREE.Color(0xE8E8E8);

    // Create overlay
    fullscreen3dDiv = document.createElement('div');
    fullscreen3dDiv.id = 'fullscreen-3d-overlay';
    fullscreen3dDiv.style.cssText = `
        position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
        z-index: 10000; background: #E8E8E8;
        display: flex; flex-direction: column;
    `;

    // ── Canvas Container ──
    const canvasContainer = document.createElement('div');
    canvasContainer.style.cssText = `flex: 1; position: relative; overflow: hidden;`;
    fullscreen3dDiv.appendChild(canvasContainer);

    // ── Floating Toolbar (minimal, bottom-centre) ──
    const toolbar = document.createElement('div');
    toolbar.style.cssText = `
        position: absolute; bottom: 24px; left: 50%; transform: translateX(-50%);
        background: rgba(35,35,35,0.92); border-radius: 10px;
        display: flex; align-items: center; gap: 4px; padding: 6px 10px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.35); backdrop-filter: blur(8px);
        font-family: "Segoe UI", Arial, sans-serif; z-index: 2;
    `;
    canvasContainer.appendChild(toolbar);

    // ── Member Info Tooltip (top-left floating) ──
    const infoPanel = document.createElement('div');
    infoPanel.style.cssText = `
        position: absolute; top: 16px; left: 16px;
        background: rgba(35,35,35,0.88); border-radius: 8px;
        padding: 10px 16px; color: #EEE; font-family: "Segoe UI", Arial, sans-serif;
        font-size: 13px; line-height: 1.6; min-width: 180px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.3); backdrop-filter: blur(8px);
        display: none; z-index: 2;
    `;
    canvasContainer.appendChild(infoPanel);

    // ── Status hint (top-centre) ──
    const hint = document.createElement('div');
    hint.style.cssText = `
        position: absolute; top: 12px; left: 50%; transform: translateX(-50%);
        color: rgba(100,100,100,0.7); font-size: 11px; font-family: "Segoe UI", Arial, sans-serif;
        pointer-events: none; z-index: 1;
    `;
    hint.textContent = 'Left-drag: Orbit  ·  Right/Shift-drag: Pan  ·  Scroll: Zoom  ·  Click: Select + orbit centre  ·  Double-click: Frame member';
    canvasContainer.appendChild(hint);

    const makeTBBtn = (text, title, onClick, highlight) => {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.title = title;
        btn.style.cssText = `
            background: ${highlight ? '#2B7CD0' : 'rgba(255,255,255,0.08)'}; border: none;
            color: ${highlight ? '#FFF' : '#CCC'}; padding: 7px 16px;
            border-radius: 6px; cursor: pointer; font-size: 12px; font-family: inherit;
            font-weight: ${highlight ? '600' : '400'};
            transition: background 0.15s, color 0.15s;
        `;
        btn.addEventListener('mouseenter', () => {
            if (!highlight) btn.style.background = 'rgba(255,255,255,0.18)';
        });
        btn.addEventListener('mouseleave', () => {
            if (!highlight) btn.style.background = 'rgba(255,255,255,0.08)';
        });
        btn.addEventListener('click', onClick);
        return btn;
    };

    // Separator
    const sep = () => {
        const s = document.createElement('div');
        s.style.cssText = 'width:1px; height:20px; background:rgba(255,255,255,0.15); margin:0 4px;';
        return s;
    };

    // Wireframe toggle
    let wireframeOn = false;
    const wireBtn = makeTBBtn('Wireframe', 'Toggle wireframe', () => {
        wireframeOn = !wireframeOn;
        wireBtn.style.background = wireframeOn ? '#2B7CD0' : 'rgba(255,255,255,0.08)';
        wireBtn.style.color = wireframeOn ? '#FFF' : '#CCC';
        scene3d.traverse(obj => {
            if (obj.isMesh && obj.material && !obj.userData.isGround && !obj.userData.isFloor) {
                obj.material.wireframe = wireframeOn;
            }
        });
    });
    toolbar.appendChild(wireBtn);

    // X-Ray toggle
    let xrayOn = false;
    const xrayBtn = makeTBBtn('X-Ray', 'Toggle transparency', () => {
        xrayOn = !xrayOn;
        xrayBtn.style.background = xrayOn ? '#2B7CD0' : 'rgba(255,255,255,0.08)';
        xrayBtn.style.color = xrayOn ? '#FFF' : '#CCC';
        scene3d.traverse(obj => {
            if (obj.isMesh && obj.material && !obj.userData.isGround && !obj.userData.isFloor) {
                obj.material.transparent = xrayOn;
                obj.material.opacity = xrayOn ? 0.35 : 1.0;
                obj.material.depthWrite = !xrayOn;
                obj.material.needsUpdate = true;
            }
        });
    });
    toolbar.appendChild(xrayBtn);

    toolbar.appendChild(sep());

    toolbar.appendChild(makeTBBtn('Frame', 'Frame all (F)', () => {
        frame3D(orbitCtrl, camera3d, false);
    }));

    toolbar.appendChild(makeTBBtn('Reset', 'Reset view (R)', () => {
        reset3D(orbitCtrl, camera3d, false);
    }));

    toolbar.appendChild(sep());

    // ★ TAKE PICTURE — the key feature
    toolbar.appendChild(makeTBBtn('📷 Take Picture', 'Capture this view for the drawing sheet', () => {
        // Freeze current camera state
        // Temporarily set white background for clean sheet image
        const oldBg = scene3d.background;
        scene3d.background = new THREE.Color(0xFFFFFF);

        // Render to offscreen snapshot canvas at current camera angle
        camera3d.aspect = 1600 / 1200;
        camera3d.updateProjectionMatrix();
        snapshotRenderer.render(scene3d, camera3d);

        // ★ Save camera state for live updates + lock the angle
        snapshotCamera.theta = orbitCtrl.theta;
        snapshotCamera.phi = orbitCtrl.phi;
        snapshotCamera.radius = orbitCtrl.radius;
        snapshotCamera.target.copy(orbitCtrl.target);
        snapshotCamera.locked = true;

        isoView.snapshotDataURL = isoView.canvas.toDataURL('image/png');
        const img = new Image();
        img.onload = () => {
            isoView.snapshotImage = img;
            isoView.visible = true;
            btn3d.classList.add('active');
            engine.requestRender();
        };
        img.src = isoView.snapshotDataURL;

        scene3d.background = oldBg;

        // Flash feedback
        const flash = document.createElement('div');
        flash.style.cssText = `
            position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
            background: white; z-index: 10001; pointer-events: none;
            transition: opacity 0.4s ease-out; opacity: 0.7;
        `;
        document.body.appendChild(flash);
        requestAnimationFrame(() => {
            flash.style.opacity = '0';
            setTimeout(() => flash.remove(), 500);
        });

        // Brief toast notification
        const toast = document.createElement('div');
        toast.textContent = '✓ View captured to drawing sheet';
        toast.style.cssText = `
            position: absolute; bottom: 80px; left: 50%; transform: translateX(-50%);
            background: rgba(43,124,208,0.95); color: white; padding: 8px 20px;
            border-radius: 6px; font-size: 13px; font-family: inherit;
            z-index: 3; pointer-events: none;
            animation: fadeUp 2s ease-out forwards;
        `;
        canvasContainer.appendChild(toast);
        setTimeout(() => toast.remove(), 2500);
    }, true));

    toolbar.appendChild(sep());

    toolbar.appendChild(makeTBBtn('Close', 'Return to drawing (Esc)', () => {
        exitFullscreen3D();
    }));

    // Add fadeUp animation
    const style = document.createElement('style');
    style.textContent = `@keyframes fadeUp { 0% { opacity:1; transform:translateX(-50%) translateY(0); } 100% { opacity:0; transform:translateX(-50%) translateY(-20px); } }`;
    fullscreen3dDiv.appendChild(style);

    document.body.appendChild(fullscreen3dDiv);

    // ── Create fullscreen renderer ──
    const fsCanvas = document.createElement('canvas');
    canvasContainer.insertBefore(fsCanvas, canvasContainer.firstChild);

    const w = window.innerWidth;
    const h = window.innerHeight;

    fsRenderer = new THREE.WebGLRenderer({
        canvas: fsCanvas,
        antialias: true,
        alpha: false
    });
    fsRenderer.setSize(w, h);
    fsRenderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    fsRenderer.shadowMap.enabled = true;
    fsRenderer.shadowMap.type = THREE.PCFSoftShadowMap;

    camera3d.aspect = w / h;
    camera3d.updateProjectionMatrix();

    // Frame scene on first open
    frame3D(orbitCtrl, camera3d, true);

    // ── Orbit controls (Trimble: left=orbit, right/shift=pan, scroll=zoom) ──
    let fsDragging = false, fsPanning = false;
    let fsLastX = 0, fsLastY = 0;

    fsCanvas.addEventListener('mousedown', (e) => {
        e.preventDefault();
        fsLastX = e.clientX; fsLastY = e.clientY;
        if (e.button === 2 || (e.button === 0 && e.shiftKey)) {
            fsPanning = true;
            fsCanvas.style.cursor = 'move';
        } else if (e.button === 0) {
            fsDragging = true;
            fsCanvas.style.cursor = 'grabbing';
        }
    });

    window._fsMouseMove = (e) => {
        if (fsDragging) {
            const dx = e.clientX - fsLastX, dy = e.clientY - fsLastY;
            orbitCtrl.theta += -dx * orbitCtrl.rotSp;
            orbitCtrl.phi = clamp3(orbitCtrl.phi - dy * orbitCtrl.rotSp, 0.05, Math.PI - 0.05);
            fsLastX = e.clientX; fsLastY = e.clientY;
        }
        if (fsPanning) {
            const dx = e.clientX - fsLastX, dy = e.clientY - fsLastY;
            fsLastX = e.clientX; fsLastY = e.clientY;
            const fwd = new THREE.Vector3();
            camera3d.getWorldDirection(fwd);
            const right = new THREE.Vector3().crossVectors(fwd, camera3d.up).normalize();
            const up = camera3d.up.clone().normalize();
            const panScale = orbitCtrl.sRadius * 0.002;
            orbitCtrl.target.addScaledVector(right, -dx * panScale);
            orbitCtrl.target.addScaledVector(up, dy * panScale);
        }
    };
    window.addEventListener('mousemove', window._fsMouseMove);

    window._fsMouseUp = () => {
        fsDragging = false;
        fsPanning = false;
        if (fsCanvas) fsCanvas.style.cursor = 'grab';
    };
    window.addEventListener('mouseup', window._fsMouseUp);

    fsCanvas.addEventListener('wheel', (e) => {
        e.preventDefault();
        orbitCtrl.radius = clamp3(
            orbitCtrl.radius * Math.exp(e.deltaY * orbitCtrl.zoomSp),
            orbitCtrl.minR, orbitCtrl.maxR
        );
    }, { passive: false });

    fsCanvas.addEventListener('contextmenu', (e) => e.preventDefault());

    // ── Raycaster pick helper ──
    let selectionGlow = null; // glow outline mesh for selected member

    function clearSelection() {
        if (isoView.selectedMesh && isoView.selectedOrigMat) {
            isoView.selectedMesh.material = isoView.selectedOrigMat;
            isoView.selectedMesh = null;
            isoView.selectedOrigMat = null;
        }
        if (selectionGlow && selectionGlow.parent) {
            selectionGlow.parent.remove(selectionGlow);
            if (selectionGlow.geometry) selectionGlow.geometry.dispose();
            if (selectionGlow.material) selectionGlow.material.dispose();
            selectionGlow = null;
        }
        if (infoPanel) infoPanel.style.display = 'none';
    }

    function pickMemberAt(clientX, clientY) {
        const rect = fsCanvas.getBoundingClientRect();
        const ndcX = ((clientX - rect.left) / rect.width) * 2 - 1;
        const ndcY = -((clientY - rect.top) / rect.height) * 2 + 1;

        const raycaster = new THREE.Raycaster();
        raycaster.setFromCamera(new THREE.Vector2(ndcX, ndcY), camera3d);

        const meshes = [];
        scene3d.traverse(obj => {
            if (obj.isMesh && !obj.userData.isGround && !obj.userData.isFloor && obj.userData.elementRef) {
                meshes.push(obj);
            }
        });
        const hits = raycaster.intersectObjects(meshes, false);
        return hits.length > 0 ? hits[0].object : null;
    }

    function selectMember(picked) {
        clearSelection();
        if (!picked) return;

        // Store original material (don't replace — keep visual identity)
        isoView.selectedOrigMat = picked.material;
        isoView.selectedMesh = picked;

        // Glow outline: slightly larger wireframe clone in selection colour
        const glowMat = new THREE.MeshBasicMaterial({
            color: 0x2B7CD0, wireframe: true, transparent: true, opacity: 0.6,
            depthTest: true, depthWrite: false,
        });
        selectionGlow = new THREE.Mesh(picked.geometry.clone(), glowMat);
        selectionGlow.position.copy(picked.position);
        selectionGlow.rotation.copy(picked.rotation);
        selectionGlow.scale.copy(picked.scale).multiplyScalar(1.03);
        selectionGlow.userData.isGlow = true;
        scene3d.add(selectionGlow);

        // Move orbit centre to this member (smooth via damping)
        const bb = new THREE.Box3().setFromObject(picked);
        const ctr = new THREE.Vector3(); bb.getCenter(ctr);
        orbitCtrl.target.copy(ctr);

        // Show member info
        const el = picked.userData.elementRef;
        if (el && infoPanel) {
            let html = '';
            if (el.memberSize) html += `<div style="font-size:16px;font-weight:600;color:#7CB4E8;margin-bottom:4px;">${el.memberSize}</div>`;
            if (el.layer) html += `<div><span style="color:#888;">Layer:</span> ${el.layer}</div>`;
            if (el.type) html += `<div><span style="color:#888;">Type:</span> ${el.type}</div>`;
            if (el.memberCategory) html += `<div><span style="color:#888;">Category:</span> ${el.memberCategory}</div>`;
            if (el.wallType) html += `<div><span style="color:#888;">Wall type:</span> ${el.wallType}</div>`;
            if (el.thickness) html += `<div><span style="color:#888;">Thickness:</span> ${el.thickness}mm</div>`;
            if (el.x1 !== undefined && el.x2 !== undefined) {
                const span = Math.round(Math.sqrt(Math.pow(el.x2 - el.x1, 2) + Math.pow(el.y2 - el.y1, 2)));
                html += `<div><span style="color:#888;">Span:</span> ${span}mm</div>`;
            }
            const lvl = levelSystem.levels.find(l => l.id === el.level);
            if (lvl) html += `<div><span style="color:#888;">Level:</span> ${lvl.name}</div>`;
            infoPanel.innerHTML = html;
            infoPanel.style.display = 'block';
        }
    }

    // ── Single-click: select member + set orbit centre ──
    let fsClickStart = { x: 0, y: 0, t: 0 };
    fsCanvas.addEventListener('mousedown', (e2) => {
        fsClickStart = { x: e2.clientX, y: e2.clientY, t: Date.now() };
    }, true);

    fsCanvas.addEventListener('mouseup', (e2) => {
        // Only treat as click if mouse didn't move much (not a drag) and was quick
        const dx = e2.clientX - fsClickStart.x;
        const dy = e2.clientY - fsClickStart.y;
        const dt = Date.now() - fsClickStart.t;
        if (Math.sqrt(dx*dx + dy*dy) > 5 || dt > 300) return; // was a drag, ignore

        const picked = pickMemberAt(e2.clientX, e2.clientY);
        if (picked) {
            selectMember(picked);
        } else {
            clearSelection();
        }
    }, true);

    // ── Double-click: frame tightly on member ──
    fsCanvas.addEventListener('dblclick', (e) => {
        const picked = pickMemberAt(e.clientX, e.clientY);
        if (picked) {
            selectMember(picked);
            frameMember(orbitCtrl, camera3d, picked);
        } else {
            // Double-click empty → frame all
            frame3D(orbitCtrl, camera3d, false);
        }
    });

    // ── Keyboard shortcuts ──
    window._fsKeydown = (e) => {
        if (e.key === 'Escape') exitFullscreen3D();
        if (e.key === 'f' || e.key === 'F') frame3D(orbitCtrl, camera3d, false);
        if (e.key === 'r' || e.key === 'R') reset3D(orbitCtrl, camera3d, false);
    };
    window.addEventListener('keydown', window._fsKeydown);

    // ── Resize handler ──
    window._fsResize = () => {
        if (!fsRenderer || !canvasContainer) return;
        const rw = window.innerWidth;
        const rh = window.innerHeight;
        fsRenderer.setSize(rw, rh);
        camera3d.aspect = rw / rh;
        camera3d.updateProjectionMatrix();
    };
    window.addEventListener('resize', window._fsResize);

    // ── Animation loop ──
    fsCanvas.style.cursor = 'grab';
    function fsAnimate() {
        fsAnimId = requestAnimationFrame(fsAnimate);
        updateCamera3D(camera3d, orbitCtrl);
        try { fsRenderer.render(scene3d, camera3d); } catch (e) { /* silent */ }
    }
    fsAnimate();
}

function exitFullscreen3D() {
    if (!fullscreen3d) return;
    fullscreen3d = false;

    if (fsAnimId) cancelAnimationFrame(fsAnimId);
    fsAnimId = null;

    if (window._fsMouseMove) window.removeEventListener('mousemove', window._fsMouseMove);
    if (window._fsMouseUp) window.removeEventListener('mouseup', window._fsMouseUp);
    if (window._fsResize) window.removeEventListener('resize', window._fsResize);
    if (window._fsKeydown) window.removeEventListener('keydown', window._fsKeydown);
    window._fsMouseMove = null;
    window._fsMouseUp = null;
    window._fsResize = null;
    window._fsKeydown = null;

    if (fsRenderer) { fsRenderer.dispose(); fsRenderer = null; }

    // Reset wireframe/xray
    scene3d.traverse(obj => {
        if (obj.isMesh && obj.material && !obj.userData.isGround && !obj.userData.isFloor && !obj.userData.isGlow) {
            obj.material.wireframe = false;
            obj.material.transparent = false;
            obj.material.opacity = 1.0;
            obj.material.depthWrite = true;
            obj.material.needsUpdate = true;
        }
    });

    // Deselect and remove glow
    if (isoView.selectedMesh && isoView.selectedOrigMat) {
        isoView.selectedMesh.material = isoView.selectedOrigMat;
        isoView.selectedMesh = null;
        isoView.selectedOrigMat = null;
    }
    // Remove any selection glow meshes
    const glows = [];
    scene3d.traverse(obj => { if (obj.userData && obj.userData.isGlow) glows.push(obj); });
    glows.forEach(g => {
        if (g.geometry) g.geometry.dispose();
        if (g.material) g.material.dispose();
        scene3d.remove(g);
    });

    // Restore scene bg to white for snapshots
    scene3d.background = new THREE.Color(0xFFFFFF);

    // Restore camera aspect for snapshot renderer
    camera3d.aspect = 1600 / 1200;
    camera3d.updateProjectionMatrix();

    if (fullscreen3dDiv && fullscreen3dDiv.parentNode) {
        fullscreen3dDiv.parentNode.removeChild(fullscreen3dDiv);
    }
    fullscreen3dDiv = null;

    engine.requestRender();
}

// Fullscreen 3D button (separate from toggle)
const btnFs3d = document.getElementById('btn-fullscreen-3d');
if (btnFs3d) {
    btnFs3d.addEventListener('click', () => {
        if (fullscreen3d) exitFullscreen3D();
        else enterFullscreen3D();
    });
}

// ── Level Switch → Rebuild 3D ────────────────────────────

const origSwitchLevel = switchToLevel;
switchToLevel = function(index) {
    origSwitchLevel(index);
    // Don't auto-rebuild snapshot on level switch — user controls via Take Picture
};

// ── Initialise 3D ────────────────────────────────────────

try {
    init3DViewport();
    rebuild3DScene();
    // Auto-frame and take initial snapshot
    if (scene3d && camera3d && orbitCtrl.target) {
        frame3D(orbitCtrl, camera3d, true);
        updateCamera3D(camera3d, orbitCtrl);
        takeSnapshot();
        // ★ Save initial camera state for live updates
        snapshotCamera.theta = orbitCtrl.theta;
        snapshotCamera.phi = orbitCtrl.phi;
        snapshotCamera.radius = orbitCtrl.radius;
        snapshotCamera.target.copy(orbitCtrl.target);
    }
} catch (e3d) {
    console.error('[3D] Init error (non-fatal):', e3d.message);
}

// ══════════════════════════════════════════════════════════
// ── 3D PHASE 5: SECTION CUT ─────────────────────────────
// ══════════════════════════════════════════════════════════

// ── Section Tool State ───────────────────────────────────

const sectionState = {
    placing: false,
    startPoint: null,
    currentEnd: null,
};

document.getElementById('btn-section').addEventListener('click', () => setActiveTool('section'));

// Extend setActiveTool
// X key = section tool
window.addEventListener('keydown', (e) => {
    if (document.activeElement !== document.body) return;
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'x') setActiveTool('section');
});

function getSectionPos(e) {
    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left, sy = e.clientY - rect.top;
    const snap = findSnap(sx, sy);
    let pos = snap ? { x: snap.x, y: snap.y } : engine.coords.screenToSheet(sx, sy);
    if (sectionState.placing && sectionState.startPoint) {
        pos = applyOrtho(pos.x, pos.y, sectionState.startPoint.x, sectionState.startPoint.y);
    }
    return pos;
}

container.addEventListener('mousemove', (e) => {
    if (typeof detailsSystem !== 'undefined') return; // V18: new system handles this
    if (activeTool === 'section') {
        sectionState.currentEnd = getSectionPos(e);
        engine.requestRender();
    }
});

container.addEventListener('mousedown', (e) => {
    if (typeof detailsSystem !== 'undefined') return; // V18: new system handles this
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;
    if (activeTool !== 'section') return;

    const pos = getSectionPos(e);

    if (!sectionState.placing) {
        sectionState.placing = true;
        sectionState.startPoint = pos;
    } else {
        // Second click — generate section
        const start = sectionState.startPoint;
        const end = pos;
        sectionState.placing = false;
        sectionState.startPoint = null;

        const realStart = engine.coords.sheetToReal(start.x, start.y);
        const realEnd = engine.coords.sheetToReal(end.x, end.y);

        generateSection(realStart, realEnd);
        setActiveTool('select');
    }
});

container.addEventListener('contextmenu', (e) => {
    if (typeof detailsSystem !== 'undefined') return; // V18: new system handles this
    if (activeTool === 'section' && sectionState.placing) {
        e.preventDefault();
        sectionState.placing = false;
        sectionState.startPoint = null;
        engine.requestRender();
    }
});

// Draw section line preview on 2D canvas (V18: disabled when details system loaded)
engine.onRender((ctx, eng) => {
    if (typeof detailsSystem !== 'undefined') return; // V18: new system handles preview
    if (activeTool !== 'section' || !sectionState.placing || !sectionState.startPoint || !sectionState.currentEnd) return;
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    const sp = coords.sheetToScreen(sectionState.startPoint.x, sectionState.startPoint.y);
    const ep = coords.sheetToScreen(sectionState.currentEnd.x, sectionState.currentEnd.y);

    // Thick dashed red line
    ctx.strokeStyle = '#CC0000';
    ctx.lineWidth = 2.5;
    ctx.setLineDash([8, 4]);
    ctx.beginPath(); ctx.moveTo(sp.x, sp.y); ctx.lineTo(ep.x, ep.y); ctx.stroke();
    ctx.setLineDash([]);

    // Arrow indicators at each end (perpendicular to section line)
    const dx = ep.x - sp.x, dy = ep.y - sp.y;
    const len = Math.sqrt(dx*dx + dy*dy);
    if (len > 10) {
        const nx = -dy/len, ny = dx/len; // perpendicular
        const arrowLen = 12;
        ctx.fillStyle = '#CC0000';
        // Arrow at start
        ctx.beginPath();
        ctx.moveTo(sp.x + nx*arrowLen, sp.y + ny*arrowLen);
        ctx.lineTo(sp.x + nx*arrowLen - dy/len*6, sp.y + ny*arrowLen + dx/len*6);
        ctx.lineTo(sp.x + nx*arrowLen + dy/len*6, sp.y + ny*arrowLen - dx/len*6);
        ctx.fill();
        // Labels
        ctx.font = 'bold 12px "Segoe UI", Arial, sans-serif';
        ctx.fillStyle = '#CC0000';
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillText('A', sp.x + nx*20, sp.y + ny*20);
        ctx.fillText('A', ep.x + nx*20, ep.y + ny*20);
    }

    // Start dot
    ctx.fillStyle = '#CC0000';
    ctx.beginPath(); ctx.arc(sp.x, sp.y, 4, 0, Math.PI*2); ctx.fill();
    ctx.textAlign = 'left'; ctx.textBaseline = 'top';
});

// ── Section View Generation ──────────────────────────────

const sectionOverlay = document.getElementById('section-overlay');
const sectionCanvas = document.getElementById('section-canvas');
const sectionCtx = sectionCanvas.getContext('2d');

function generateSection(realStart, realEnd) {
    // Section line direction
    const sdx = realEnd.x - realStart.x;
    const sdy = realEnd.y - realStart.y;
    const sectionLen = Math.sqrt(sdx*sdx + sdy*sdy);
    if (sectionLen < 100) return; // too short

    // Unit vector along section line
    const ux = sdx / sectionLen;
    const uy = sdy / sectionLen;

    // For each element, project onto the section line to get its position along the section.
    // Elements within a tolerance of the section line are "cut" and shown.
    const cutTolerance = 500 * CONFIG.drawingScale * 0.01; // ~5m real-world catch zone

    const sectionElements = [];

    for (const el of project.elements) {
        const layer = project.layers[el.layer];
        if (!layer || !layer.visible) continue;
        const lv = levelSystem.levels.find(l => l.id === el.level);
        if (!lv) continue;

        if (el.type === 'column') {
            // Project column position onto section line
            const dx = el.x - realStart.x, dy = el.y - realStart.y;
            const along = dx * ux + dy * uy; // distance along section line
            const perp = Math.abs(dx * (-uy) + dy * ux); // perpendicular distance

            if (perp < cutTolerance && along > -cutTolerance && along < sectionLen + cutTolerance) {
                const dir = el.extends || 'below';
                const lvIdx = levelSystem.levels.findIndex(l => l.id === el.level);
                let botElev, topElev;
                if (dir === 'below' && lvIdx > 0) {
                    botElev = levelSystem.levels[lvIdx-1].elevation;
                    topElev = lv.elevation;
                } else if (dir === 'above' && lvIdx < levelSystem.levels.length - 1) {
                    botElev = lv.elevation;
                    topElev = levelSystem.levels[lvIdx+1].elevation;
                } else if (dir === 'both') {
                    botElev = lvIdx > 0 ? levelSystem.levels[lvIdx-1].elevation : lv.elevation;
                    topElev = lvIdx < levelSystem.levels.length-1 ? levelSystem.levels[lvIdx+1].elevation : lv.elevation + lv.height;
                } else {
                    botElev = lv.elevation;
                    topElev = lv.elevation + lv.height;
                }
                sectionElements.push({
                    type: 'column', along, tag: el.tag, size: el.size || 89,
                    botElev, topElev
                });
            }
        }

        if (el.type === 'line' && (el.layer === 'S-BEAM' || el.layer === 'S-WALL')) {
            // Check if line crosses the section
            const mx = (el.x1 + el.x2) / 2, my = (el.y1 + el.y2) / 2;
            const dx = mx - realStart.x, dy = my - realStart.y;
            const along = dx * ux + dy * uy;
            const perp = Math.abs(dx * (-uy) + dy * ux);

            if (perp < cutTolerance && along > -cutTolerance && along < sectionLen + cutTolerance) {
                sectionElements.push({
                    type: el.layer === 'S-WALL' ? 'wall' : 'beam',
                    along, elev: lv.elevation,
                    height: el.layer === 'S-WALL' ? lv.height * 0.8 : 200,
                    width: el.layer === 'S-WALL' ? 200 : 300,
                });
            }
        }

        if (el.type === 'polyline' && el.layer === 'S-SLAB' && (el.closed || el.points.length > 3)) {
            // Check if slab crosses section
            for (const pt of el.points) {
                const dx = pt.x - realStart.x, dy = pt.y - realStart.y;
                const along = dx * ux + dy * uy;
                const perp = Math.abs(dx * (-uy) + dy * ux);
                if (perp < cutTolerance && along > 0 && along < sectionLen) {
                    sectionElements.push({
                        type: 'slab', along: 0, width: sectionLen,
                        elev: lv.elevation, thickness: 200,
                    });
                    break;
                }
            }
        }
    }

    renderSection(sectionElements, sectionLen);
}

function renderSection(elements, sectionLen) {
    const wrap = document.getElementById('section-canvas-wrap');
    const w = wrap.clientWidth || 900;
    const h = wrap.clientHeight || 500;
    const dpr = window.devicePixelRatio || 1;

    sectionCanvas.width = w * dpr;
    sectionCanvas.height = h * dpr;
    sectionCanvas.style.width = w + 'px';
    sectionCanvas.style.height = h + 'px';
    sectionCtx.setTransform(dpr, 0, 0, dpr, 0, 0);

    const ctx = sectionCtx;

    // Background
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, w, h);

    // Compute bounds
    const topLevel = levelSystem.levels[levelSystem.levels.length - 1];
    const maxElev = topLevel.elevation + (topLevel.height || 2700);
    const padding = 60;

    // Scale to fit
    const scaleX = (w - padding * 2) / sectionLen;
    const scaleY = (h - padding * 2) / maxElev;
    const scale = Math.min(scaleX, scaleY);

    const ox = padding; // origin x
    const oy = h - padding; // origin y (bottom of view)

    function toScreen(along, elev) {
        return { x: ox + along * scale, y: oy - elev * scale };
    }

    // ── Draw level lines ──
    for (const lv of levelSystem.levels) {
        const y = toScreen(0, lv.elevation).y;
        ctx.strokeStyle = '#CCCCCC';
        ctx.lineWidth = 1;
        ctx.setLineDash([6, 4]);
        ctx.beginPath();
        ctx.moveTo(ox - 30, y);
        ctx.lineTo(ox + sectionLen * scale + 10, y);
        ctx.stroke();
        ctx.setLineDash([]);

        // Level label
        ctx.font = '11px "Segoe UI", Arial, sans-serif';
        ctx.fillStyle = '#888888';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(lv.name, ox - 35, y);

        // RL
        ctx.font = '9px "Consolas", monospace';
        ctx.fillStyle = '#AAAAAA';
        ctx.fillText('RL ' + (lv.elevation/1000).toFixed(1) + 'm', ox - 35, y + 12);
    }

    // ── Draw ground hatch ──
    const groundY = toScreen(0, 0).y;
    ctx.fillStyle = '#F0EDE8';
    ctx.fillRect(ox - 30, groundY, sectionLen * scale + 40, h - groundY);
    // Ground line
    ctx.strokeStyle = '#333333';
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(ox - 30, groundY);
    ctx.lineTo(ox + sectionLen * scale + 10, groundY);
    ctx.stroke();

    // ── Draw section elements ──
    for (const el of elements) {
        if (el.type === 'column') {
            const p1 = toScreen(el.along, el.botElev);
            const p2 = toScreen(el.along, el.topElev);
            const colW = el.size * scale * 0.5;

            // Column rectangle
            ctx.fillStyle = 'rgba(43,102,170,0.2)';
            ctx.fillRect(p1.x - colW/2, p2.y, colW, p1.y - p2.y);
            ctx.strokeStyle = '#2266AA';
            ctx.lineWidth = 1.5;
            ctx.strokeRect(p1.x - colW/2, p2.y, colW, p1.y - p2.y);

            // X cross
            ctx.beginPath();
            ctx.moveTo(p1.x - colW/2, p2.y); ctx.lineTo(p1.x + colW/2, p1.y);
            ctx.moveTo(p1.x + colW/2, p2.y); ctx.lineTo(p1.x - colW/2, p1.y);
            ctx.stroke();

            // Tag
            if (el.tag) {
                ctx.font = 'bold 10px "Segoe UI", Arial, sans-serif';
                ctx.fillStyle = '#2266AA';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                ctx.fillText(el.tag, p1.x, p2.y - 4);
            }
        }

        if (el.type === 'beam') {
            const p = toScreen(el.along, el.elev);
            const bw = Math.max(8, el.width * scale);
            const bh = Math.max(4, el.height * scale);
            ctx.fillStyle = 'rgba(0,0,0,0.15)';
            ctx.fillRect(p.x - bw/2, p.y - bh, bw, bh);
            ctx.strokeStyle = '#333333';
            ctx.lineWidth = 1;
            ctx.strokeRect(p.x - bw/2, p.y - bh, bw, bh);
        }

        if (el.type === 'wall') {
            const p = toScreen(el.along, el.elev);
            const ww = Math.max(6, el.width * scale);
            const wh = Math.max(10, el.height * scale);
            ctx.fillStyle = 'rgba(100,100,100,0.15)';
            ctx.fillRect(p.x - ww/2, p.y - wh, ww, wh);
            ctx.strokeStyle = '#666666';
            ctx.lineWidth = 1;
            ctx.strokeRect(p.x - ww/2, p.y - wh, ww, wh);
            // Hatch
            ctx.strokeStyle = '#CCCCCC';
            ctx.lineWidth = 0.5;
            for (let hy = p.y; hy > p.y - wh; hy -= 4) {
                ctx.beginPath(); ctx.moveTo(p.x - ww/2, hy); ctx.lineTo(p.x + ww/2, hy); ctx.stroke();
            }
        }

        if (el.type === 'slab') {
            const p1 = toScreen(0, el.elev);
            const p2 = toScreen(el.width, el.elev);
            const th = Math.max(3, el.thickness * scale);
            ctx.fillStyle = 'rgba(180,180,180,0.4)';
            ctx.fillRect(p1.x, p1.y - th, p2.x - p1.x, th);
            ctx.strokeStyle = '#999999';
            ctx.lineWidth = 1;
            ctx.strokeRect(p1.x, p1.y - th, p2.x - p1.x, th);
        }
    }

    // ── Title ──
    ctx.font = 'bold 14px "Segoe UI", Arial, sans-serif';
    ctx.fillStyle = '#333333';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText('SECTION A-A', w / 2, 10);

    // Scale bar
    ctx.font = '10px "Consolas", monospace';
    ctx.fillStyle = '#888888';
    ctx.textAlign = 'left';
    ctx.textBaseline = 'bottom';
    ctx.fillText('Section length: ' + (sectionLen/1000).toFixed(1) + 'm', ox, h - 10);

    ctx.textAlign = 'left'; ctx.textBaseline = 'top';

    // Show overlay
    sectionOverlay.classList.remove('hidden');
    document.getElementById('section-title').textContent = 'Section A-A';
}

// ── Section Overlay Controls ─────────────────────────────

document.getElementById('section-close').addEventListener('click', () => {
    sectionOverlay.classList.add('hidden');
});

sectionOverlay.addEventListener('click', (e) => {
    if (e.target === sectionOverlay) sectionOverlay.classList.add('hidden');
});

document.getElementById('section-export').addEventListener('click', () => {
    const link = document.createElement('a');
    link.download = 'section-A-A.png';
    link.href = sectionCanvas.toDataURL('image/png');
    link.click();
});

window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !sectionOverlay.classList.contains('hidden')) {
        sectionOverlay.classList.add('hidden');
    }
});

// ══════════════════════════════════════════════════════════
// ── 3D PHASE 6: MEMBER LIBRARY + SMART TAGS ──────────────
// ══════════════════════════════════════════════════════════

// ── Australian Section Library ───────────────────────────

const SECTION_LIBRARY = {
    UB: [
        '150UB14', '150UB18', '180UB16', '180UB18', '180UB22',
        '200UB18', '200UB22', '200UB25', '200UB30',
        '250UB25', '250UB31', '250UB37',
        '310UB32', '310UB40', '310UB46',
        '360UB45', '360UB51', '360UB57',
        '410UB54', '410UB60',
        '460UB67', '460UB74', '460UB82',
        '530UB82', '530UB92',
        '610UB101', '610UB113', '610UB125',
    ],
    UC: [
        '100UC15', '150UC23', '150UC30', '150UC37',
        '200UC46', '200UC52', '200UC60',
        '250UC73', '250UC89',
        '310UC97', '310UC118', '310UC137', '310UC158',
    ],
    PFC: [
        '75PFC', '100PFC', '125PFC', '150PFC',
        '200PFC', '250PFC', '300PFC', '380PFC',
    ],
    SHS: [
        '50x50x3SHS', '50x50x5SHS',
        '65x65x3SHS', '65x65x5SHS',
        '75x75x3SHS', '75x75x5SHS', '75x75x6SHS',
        '89x89x3.5SHS', '89x89x5SHS', '89x89x6SHS',
        '100x100x4SHS', '100x100x5SHS', '100x100x6SHS',
        '125x125x5SHS', '125x125x6SHS',
        '150x150x5SHS', '150x150x6SHS', '150x150x8SHS',
        '200x200x6SHS', '200x200x8SHS', '200x200x10SHS',
        '250x250x8SHS', '250x250x10SHS',
    ],
    RHS: [
        '75x50x3RHS', '75x50x5RHS',
        '100x50x3RHS', '100x50x5RHS',
        '125x75x4RHS', '125x75x5RHS',
        '150x50x4RHS', '150x50x5RHS',
        '150x100x5RHS', '150x100x6RHS',
        '200x100x5RHS', '200x100x6RHS',
        '250x150x6RHS', '250x150x8RHS',
        '300x200x8RHS', '300x200x10RHS',
    ],
    CHS: [
        '48.3x3.2CHS', '60.3x3.6CHS',
        '76.1x3.6CHS', '88.9x4.0CHS',
        '101.6x4.0CHS', '114.3x4.5CHS',
        '139.7x4.5CHS', '165.1x5.0CHS',
        '168.3x5.0CHS', '219.1x6.4CHS',
        '273.1x6.4CHS', '323.9x8.0CHS',
    ],
    EA: [
        '50x50x3EA', '50x50x5EA', '50x50x6EA',
        '65x65x5EA', '65x65x6EA',
        '75x75x5EA', '75x75x6EA', '75x75x8EA',
        '90x90x6EA', '90x90x8EA',
        '100x100x6EA', '100x100x8EA', '100x100x10EA',
        '125x125x8EA', '125x125x10EA',
        '150x150x10EA', '150x150x12EA',
    ],
    TIMBER: [
        '70x35 MGP10', '90x35 MGP10', '90x45 MGP10',
        '140x35 MGP10', '140x45 MGP10',
        '190x35 MGP10', '190x45 MGP10',
        '240x35 MGP10', '240x45 MGP10',
        '290x45 MGP10',
        '70x35 MGP12', '90x45 MGP12',
        '140x45 MGP12', '190x45 MGP12', '240x45 MGP12',
        '2/90x45 MGP10', '2/140x45 MGP10', '2/190x45 MGP10', '2/240x45 MGP10',
        '3/90x45 MGP10', '3/140x45 MGP10', '3/190x45 MGP10',
    ],
    GLT: [
        '130x63 GL13', '180x63 GL13', '240x63 GL13',
        '300x63 GL13', '360x63 GL13',
        '130x63 GL17', '180x63 GL17', '240x63 GL17',
        '300x63 GL17', '360x63 GL17', '450x63 GL17',
        '180x63 GL18', '240x63 GL18', '300x63 GL18',
        '360x63 GL18', '450x63 GL18',
        '200x65 LVL', '240x65 LVL', '300x65 LVL',
        '360x65 LVL', '400x65 LVL',
        '2/200x65 LVL', '2/240x65 LVL', '2/300x65 LVL',
    ],
    CUSTOM: ['Custom'],
};

// ── UB Section Dimensions Database (Australian Standard) ──

const UB_DIMENSIONS = {
  "610UB125":{ d:612, bf:229, tf:19.6, tw:11.9, r1:14.0 },
  "610UB113":{ d:607, bf:228, tf:17.3, tw:11.2, r1:14.0 },
  "610UB101":{ d:602, bf:228, tf:14.8, tw:10.6, r1:14.0 },
  "530UB92.4":{ d:533, bf:209, tf:15.6, tw:10.2, r1:14.0 },
  "530UB82":{ d:528, bf:209, tf:13.2, tw:9.6, r1:14.0 },
  "460UB82.1":{ d:460, bf:191, tf:16.0, tw:9.9, r1:11.4 },
  "460UB74.6":{ d:457, bf:190, tf:14.5, tw:9.1, r1:11.4 },
  "460UB67.1":{ d:454, bf:190, tf:12.7, tw:8.5, r1:11.4 },
  "410UB59.7":{ d:406, bf:178, tf:12.8, tw:7.8, r1:11.4 },
  "410UB53.7":{ d:403, bf:178, tf:10.9, tw:7.6, r1:11.4 },
  "360UB56.7":{ d:359, bf:172, tf:13.0, tw:8.0, r1:11.4 },
  "360UB50.7":{ d:356, bf:171, tf:11.5, tw:7.3, r1:11.4 },
  "360UB44.7":{ d:352, bf:171, tf:9.7, tw:6.9, r1:11.4 },
  "310UB46.2":{ d:307, bf:166, tf:11.8, tw:6.7, r1:11.4 },
  "310UB40.4":{ d:304, bf:165, tf:10.2, tw:6.1, r1:11.4 },
  "310UB32":{ d:298, bf:149, tf:8.0, tw:5.5, r1:13.0 },
  "250UB37.3":{ d:256, bf:146, tf:10.9, tw:6.4, r1:8.9 },
  "250UB31.4":{ d:252, bf:146, tf:8.6, tw:6.1, r1:8.9 },
  "250UB25.7":{ d:248, bf:124, tf:8.0, tw:5.0, r1:12.0 },
  "200UB29.8":{ d:207, bf:134, tf:9.6, tw:6.3, r1:8.9 },
  "200UB25.4":{ d:203, bf:133, tf:7.8, tw:5.8, r1:8.9 },
  "200UB22.3":{ d:202, bf:133, tf:7.0, tw:5.0, r1:8.9 },
  "200UB18.2":{ d:198, bf:99, tf:7.0, tw:4.5, r1:11.0 },
  "180UB22.2":{ d:179, bf:90, tf:10.0, tw:6.0, r1:8.9 },
  "180UB18.1":{ d:175, bf:90, tf:8.0, tw:5.0, r1:8.9 },
  "180UB16.1":{ d:173, bf:90, tf:7.0, tw:4.5, r1:8.9 },
  "150UB18":{ d:155, bf:75, tf:9.5, tw:6.0, r1:8.0 },
  "150UB14":{ d:150, bf:75, tf:7.0, tw:5.0, r1:8.0 },
};

// ── Auto-numbering for smart tags ────────────────────────

let _beamNum = 1;
let _colNum = _colNextNum; // continue from existing counter

function getSmartTag(elementType, sectionSize) {
    if (elementType === 'column') {
        return 'SC' + (_colNum++) + '-' + sectionSize;
    }
    // Beam/general prefix based on layer
    const prefix = 'B';
    return prefix + (_beamNum++) + '-' + sectionSize;
}

// ── Member Size Selector UI ──────────────────────────────

const memberCatSelect = document.getElementById('member-category');
const memberSizeSelect = document.getElementById('member-size');

function populateSizeDropdown() {
    const cat = memberCatSelect.value;
    const sizes = SECTION_LIBRARY[cat] || ['Custom'];
    memberSizeSelect.innerHTML = '';
    for (const s of sizes) {
        const opt = document.createElement('option');
        opt.value = s;
        opt.textContent = s;
        memberSizeSelect.appendChild(opt);
    }
}

memberCatSelect.addEventListener('change', populateSizeDropdown);
populateSizeDropdown(); // initial population
// Default to 89x5.0 SHS on first load
if (memberSizeSelect.querySelector('option[value="89x89x5SHS"]')) {
    memberSizeSelect.value = '89x89x5SHS';
}

function getSelectedMemberSize() {
    if (memberCatSelect.value === 'CUSTOM') {
        return prompt('Enter custom section size:', '200x10 PLATE') || 'Custom';
    }
    return memberSizeSelect.value;
}

// ── Patch Line Tool: store member data + smart tag ───────

// We need to intercept line creation to add member data.
// Patch the line tool's mousedown handler by wrapping the history execute.
// The line tool pushes elements via history.execute, which already stamps level.
// We'll add a post-creation hook.

const origHistExec2 = history.execute;
history.execute = function(cmd) {
    origHistExec2.call(this, cmd);

    // After execution, find newly added elements and enrich with member/schedule data
    for (const el of project.elements) {
        if (el.type === 'line' && (el.layer === 'S-BEAM' || el.layer === 'S-WALL' || el.layer === 'S-FTNG')) {
            // If element has a schedule typeRef, use that as the tag (schedule is source of truth)
            if (el.typeRef && !el.tag) {
                el.tag = el.typeRef;
            }
            // Legacy path: if no typeRef, use old smart tag system
            if (!el.memberSize && !el.typeRef) {
                el.memberSize = getSelectedMemberSize();
                el.memberCategory = memberCatSelect.value;
                if (!el.tag) {
                    el.tag = getSmartTag('beam', el.memberSize);
                }
            }
        }
    }
};

// ── Patch Column Tool: store section size ────────────────

// Columns already get tags. We need to enrich with section data.
// The column mousedown creates the element — we patch similarly.
// Add a mutation observer approach: after each render, check for columns without memberSize.
engine.onRender(() => {
    for (const el of project.elements) {
        if (el.type === 'column' && !el.memberSize && !el.typeRef) {
            // Legacy path: auto-assign from toolbar selector if no schedule ref
            if (memberCatSelect.value === 'SHS' || memberCatSelect.value === 'CHS' || memberCatSelect.value === 'UC') {
                el.memberSize = memberSizeSelect.value;
                el.memberCategory = memberCatSelect.value;
            } else {
                el.memberSize = '89x89x5SHS';
                el.memberCategory = 'SHS';
            }
            // Update tag to include section
            if (el.tag && !el.tag.includes('-')) {
                el.tag = el.tag + '-' + el.memberSize;
            }
        }
    }
});

// ── Smart Tag Rendering Enhancement ──────────────────────

// The existing text rendering for tags and beam labels already works.
// Beam tags are now auto-generated with section sizes.
// We'll add rendering for beam tags along the line midpoint.

function drawBeamTags(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;
    if (zoom < 0.4) return;

    for (const el of project.getVisibleElements()) {
        if (el.type !== 'line' || !el.tag) continue;
        if (el.layer !== 'S-BEAM' && el.layer !== 'S-WALL' && el.layer !== 'S-FTNG') continue;

        const p1 = coords.realToScreen(el.x1, el.y1);
        const p2 = coords.realToScreen(el.x2, el.y2);
        const mx = (p1.x + p2.x) / 2;
        const my = (p1.y + p2.y) / 2;

        // Calculate angle of the beam line
        let angle = Math.atan2(p2.y - p1.y, p2.x - p1.x);

        // Keep text readable: if angle would make text upside-down, flip it
        if (angle > Math.PI / 2) angle -= Math.PI;
        if (angle < -Math.PI / 2) angle += Math.PI;

        const fontSize = Math.max(1, 2.5 * zoom);
        const offset = fontSize * 0.6 + 2; // offset above the line

        // Apply custom tag offset if user has dragged it
        const tagDx = (el._tagOffsetX || 0) / CONFIG.drawingScale * zoom;
        const tagDy = (el._tagOffsetY || 0) / CONFIG.drawingScale * zoom;

        ctx.save();
        ctx.translate(mx + tagDx, my + tagDy);
        ctx.rotate(angle);
        ctx.font = `${fontSize}px "Architects Daughter", cursive`;
        ctx.fillStyle = isElementSelected(el) ? '#2B7CD0' : '#333333';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        // Show member size from schedule if assigned, otherwise type ref
        let displayTag = el.tag;
        if (el.layer === 'S-BEAM') {
            const typeRef = el.typeRef || el.tag;
            const schedData = project.scheduleTypes.beam[typeRef];
            if (schedData && schedData.size) {
                displayTag = schedData.size;
            } else {
                displayTag = typeRef;
            }
        }
        ctx.fillText(displayTag, 0, -offset);
        ctx.restore();
    }
    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';
}

engine.onRender(drawBeamTags);

// ── Column Tag Update (show full section in tag) ─────────

// The existing column rendering already shows el.tag.
// Now tags include the section size (e.g., "SC1-89x89x5SHS").
// No rendering changes needed — it's already rendered via the column draw.

// ── Save/Load: member data persists automatically ────────
// Elements already save all properties including memberSize, memberCategory, tag.

// ══════════════════════════════════════════════════════════
