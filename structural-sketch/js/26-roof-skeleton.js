// ============================================================
// 26-roof-skeleton.js — Roof Skeleton Drawing Tool
// Draw ridge + hip lines → DCEL face detection → per-face
// pitch entry → transparent 3D mass model
// ============================================================

// ── Drawing State (sheet-mm, temporary) ─────────────────────

const roofSkeletonState = {
    nodes: [],           // [{id, x, y}] in sheet-mm
    edges: [],           // [{id, a, b}]
    chainStartId: null,  // node id the next edge starts from
    hoverPoint: null,    // {x, y} in sheet-mm (current mouse pos)
    snapTarget: null,    // {x, y, type} — current snap highlight
};

let _rsnCounter = 0;
function _rsnId() { return 'rn_' + (++_rsnCounter); }
function _rseId() { return 're_' + (++_rsnCounter); }

// ── Helpers ──────────────────────────────────────────────────

function findSkeletonElement() {
    return project.elements.find(el => el.type === 'roofSkeleton');
}

function startDrawSkeleton() {
    // Switch to RF (roof) level
    const rfIndex = levelSystem.levels.findIndex(
        l => l.id === 'RF' || l.name.toLowerCase().includes('roof')
    );
    if (rfIndex !== -1) switchToLevel(rfIndex);
    // Clear any previous drawing state
    roofSkeletonState.nodes = [];
    roofSkeletonState.edges = [];
    roofSkeletonState.chainStartId = null;
    roofSkeletonState.hoverPoint = null;
    roofSkeletonState.snapTarget = null;
    setActiveTool('roofSkeleton');
}

// Distance between two points (sheet-mm)
function _dist2d(ax, ay, bx, by) {
    return Math.sqrt((ax - bx) ** 2 + (ay - by) ** 2);
}

// ── Custom Snap for Skeleton Tool ────────────────────────────

function _findSkeletonSnap(screenX, screenY) {
    const sheetPos = engine.coords.screenToSheet(screenX, screenY);
    const radiusMM = (snapState.snapRadius * 1.8) / engine.viewport.zoom; // boosted

    let best = null;
    let bestDist = radiusMM;

    // 1. Existing drawing-state nodes (highest priority — close the chain)
    for (const n of roofSkeletonState.nodes) {
        const d = _dist2d(sheetPos.x, sheetPos.y, n.x, n.y);
        if (d < bestDist) {
            bestDist = d;
            best = { x: n.x, y: n.y, type: 'skeletonNode', nodeId: n.id };
        }
    }

    // 2. Envelope corners (second priority) — find across all levels
    const envelope = _findEnvelopeAnyLevel();
    if (envelope && envelope.points) {
        for (const pt of envelope.points) {
            const s = engine.coords.realToSheet(pt.x, pt.y);
            const d = _dist2d(sheetPos.x, sheetPos.y, s.x, s.y);
            if (d < bestDist) {
                bestDist = d;
                best = { x: s.x, y: s.y, type: 'envelopeCorner', realX: pt.x, realY: pt.y };
            }
        }
    }

    // 3. Envelope edge midpoints (tertiary)
    if (envelope && envelope.points && envelope.points.length >= 2) {
        const pts = envelope.points;
        for (let i = 0; i < pts.length; i++) {
            const j = (i + 1) % pts.length;
            const mx = (pts[i].x + pts[j].x) / 2;
            const my = (pts[i].y + pts[j].y) / 2;
            const s = engine.coords.realToSheet(mx, my);
            const d = _dist2d(sheetPos.x, sheetPos.y, s.x, s.y);
            if (d < bestDist * 0.8) { // slightly smaller radius for midpoints
                bestDist = d;
                best = { x: s.x, y: s.y, type: 'envelopeMidpoint', realX: mx, realY: my };
            }
        }
    }

    return best;
}

// ── Mouse: Click (left) ──────────────────────────────────────

container.addEventListener('mousedown', function _skeletonMousedown(e) {
    if (e.button !== 0) return;
    if (activeTool !== 'roofSkeleton') return;
    if (engine._spaceDown || engine._isPanning) return;
    e.preventDefault();

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    // Determine snapped position in sheet-mm
    const skelSnap = _findSkeletonSnap(sx, sy);
    let sheetPos;
    if (skelSnap) {
        sheetPos = { x: skelSnap.x, y: skelSnap.y };
    } else {
        const stdSnap = findSnap(sx, sy);
        sheetPos = stdSnap
            ? { x: stdSnap.x, y: stdSnap.y }
            : engine.coords.screenToSheet(sx, sy);
    }

    // Apply shift constraint (45°/90°) when chain is active
    if (e.shiftKey && roofSkeletonState.chainStartId !== null) {
        const startNode = roofSkeletonState.nodes.find(
            n => n.id === roofSkeletonState.chainStartId
        );
        if (startNode) {
            sheetPos = _applyAngleConstraint(sheetPos, startNode);
        }
    }

    // Check if snapping to an existing skeleton node (close/connect chain)
    if (skelSnap && skelSnap.type === 'skeletonNode' && skelSnap.nodeId) {
        // Connect to existing node — don't create a duplicate
        if (roofSkeletonState.chainStartId !== null &&
            roofSkeletonState.chainStartId !== skelSnap.nodeId) {
            // Add edge from chainStart to this existing node
            roofSkeletonState.edges.push({
                id: _rseId(),
                a: roofSkeletonState.chainStartId,
                b: skelSnap.nodeId,
            });
        }
        roofSkeletonState.chainStartId = skelSnap.nodeId;
        engine.requestRender();
        return;
    }

    // Create new node at the click position
    const newNode = { id: _rsnId(), x: sheetPos.x, y: sheetPos.y };
    roofSkeletonState.nodes.push(newNode);

    if (roofSkeletonState.chainStartId !== null) {
        // Add edge from chain start to new node
        roofSkeletonState.edges.push({
            id: _rseId(),
            a: roofSkeletonState.chainStartId,
            b: newNode.id,
        });
    }

    roofSkeletonState.chainStartId = newNode.id;
    engine.requestRender();
});

// ── Mouse: Move ───────────────────────────────────────────────

container.addEventListener('mousemove', function _skeletonMousemove(e) {
    if (activeTool !== 'roofSkeleton') return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;

    const skelSnap = _findSkeletonSnap(sx, sy);
    let sheetPos;
    if (skelSnap) {
        sheetPos = { x: skelSnap.x, y: skelSnap.y };
    } else {
        const stdSnap = findSnap(sx, sy);
        sheetPos = stdSnap
            ? { x: stdSnap.x, y: stdSnap.y }
            : engine.coords.screenToSheet(sx, sy);
    }

    if (e.shiftKey && roofSkeletonState.chainStartId !== null) {
        const startNode = roofSkeletonState.nodes.find(
            n => n.id === roofSkeletonState.chainStartId
        );
        if (startNode) {
            sheetPos = _applyAngleConstraint(sheetPos, startNode);
        }
    }

    roofSkeletonState.hoverPoint = sheetPos;
    roofSkeletonState.snapTarget = skelSnap;
    engine.requestRender();
});

// ── Right-click: End current chain ───────────────────────────

container.addEventListener('contextmenu', function _skeletonContextmenu(e) {
    if (activeTool !== 'roofSkeleton') return;
    e.preventDefault();
    // End current chain — next click starts a fresh one
    roofSkeletonState.chainStartId = null;
    engine.requestRender();
});

// ── Keyboard: Enter to commit, Escape to cancel ──────────────

document.addEventListener('keydown', function _skeletonKeydown(e) {
    if (activeTool !== 'roofSkeleton') return;

    if (e.key === 'Enter') {
        e.preventDefault();
        if (roofSkeletonState.edges.length >= 1) {
            _commitSkeletonTool();
        }
    }

    if (e.key === 'Escape') {
        e.preventDefault();
        // Cancel the current in-progress segment, keep committed edges
        roofSkeletonState.chainStartId = null;
        roofSkeletonState.hoverPoint = null;
        roofSkeletonState.snapTarget = null;
        engine.requestRender();
    }
});

// ── Angle Constraint (Shift) ─────────────────────────────────

function _applyAngleConstraint(sheetPos, fromNode) {
    const dx = sheetPos.x - fromNode.x;
    const dy = sheetPos.y - fromNode.y;
    const len = Math.sqrt(dx * dx + dy * dy);
    if (len < 0.01) return sheetPos;
    // Snap to nearest 45° increment
    const angle = Math.atan2(dy, dx);
    const snapped = Math.round(angle / (Math.PI / 4)) * (Math.PI / 4);
    return {
        x: fromNode.x + len * Math.cos(snapped),
        y: fromNode.y + len * Math.sin(snapped),
    };
}

// ── Commit: Convert to Real-World MM, Merge Envelope, Detect Faces ──

function _getRFLevelId() {
    if (typeof levelSystem === 'undefined') return 'RF';
    const rfLevel = levelSystem.levels.find(
        l => l.id === 'RF' || l.name.toLowerCase().includes('roof')
    );
    return rfLevel ? rfLevel.id : 'RF';
}

function _switchToRFLevel() {
    if (typeof levelSystem === 'undefined') return;
    const rfIndex = levelSystem.levels.findIndex(
        l => l.id === 'RF' || l.name.toLowerCase().includes('roof')
    );
    if (rfIndex !== -1 && levelSystem.activeLevelIndex !== rfIndex) {
        switchToLevel(rfIndex);
    }
}

function _commitSkeletonTool() {
    if (roofSkeletonState.edges.length === 0) return;

    const currentLevel = _getRFLevelId(); // always commit to RF level

    // Convert drawing-state nodes from sheet-mm to real-world mm
    const realNodes = roofSkeletonState.nodes.map(n => {
        const r = engine.coords.sheetToReal(n.x, n.y);
        return { id: n.id, x: r.x, y: r.y, source: 'drawn' };
    });
    const drawnEdges = roofSkeletonState.edges.map(e => ({
        id: e.id, a: e.a, b: e.b, source: 'drawn',
    }));

    // Merge with envelope to form a closed planar graph
    const { mergedNodes, mergedEdges } = _mergeWithEnvelope(realNodes, drawnEdges);

    // Detect faces using DCEL halfedge traversal
    const faces = _detectFaces(mergedNodes, mergedEdges);

    // Preserve pitches from any previous skeleton
    const existing = findSkeletonElement();
    if (existing && existing.faces) {
        for (const newFace of faces) {
            // Match by approximate centroid proximity
            const oldMatch = existing.faces.find(of => {
                const dx = Math.abs(of.centroid.x - newFace.centroid.x);
                const dy = Math.abs(of.centroid.y - newFace.centroid.y);
                return dx < 500 && dy < 500; // within 500mm
            });
            if (oldMatch && oldMatch.pitch != null) {
                newFace.pitch = oldMatch.pitch;
            }
        }
    }

    const skeleton = {
        id: generateId(),
        type: 'roofSkeleton',
        layer: 'S-RIDGE',
        level: currentLevel,
        nodes: mergedNodes,
        edges: mergedEdges,
        faces: faces,
        committed: true,
    };

    history.execute({
        description: 'Draw roof skeleton',
        execute() {
            if (existing) {
                const idx = project.elements.indexOf(existing);
                if (idx !== -1) project.elements.splice(idx, 1);
            }
            project.elements.push(skeleton);
        },
        undo() {
            const i = project.elements.indexOf(skeleton);
            if (i !== -1) project.elements.splice(i, 1);
            if (existing) project.elements.push(existing);
        },
    });

    // Clear drawing state
    roofSkeletonState.nodes = [];
    roofSkeletonState.edges = [];
    roofSkeletonState.chainStartId = null;
    roofSkeletonState.hoverPoint = null;
    roofSkeletonState.snapTarget = null;

    if (typeof markComplianceDirty === 'function') markComplianceDirty();
    engine.requestRender();
    if (typeof updateBracingSummaryPanel === 'function') updateBracingSummaryPanel();
    setActiveTool('select');
}

// ── Merge Envelope into Node/Edge Graph ──────────────────────

function _mergeWithEnvelope(drawnNodes, drawnEdges) {
    const MERGE_TOL = 20; // mm — snap tolerance for matching to envelope corners

    const envelope = _findEnvelopeAnyLevel(); // envelope lives on GF, skeleton on RF
    const allNodes = [...drawnNodes];
    const allEdges = [...drawnEdges];

    if (!envelope || !envelope.points || envelope.points.length < 3) {
        return { mergedNodes: allNodes, mergedEdges: allEdges };
    }

    const envPts = envelope.points;
    const envNodeIds = [];

    // For each envelope corner, either find a matching drawn node or create a new one
    for (let i = 0; i < envPts.length; i++) {
        const pt = envPts[i];
        // Find closest drawn node within tolerance
        let matchId = null;
        let matchDist = MERGE_TOL;
        for (const dn of drawnNodes) {
            const d = _dist2d(dn.x, dn.y, pt.x, pt.y);
            if (d < matchDist) { matchDist = d; matchId = dn.id; }
        }
        if (matchId) {
            // Merge: snap the drawn node to the envelope corner
            const node = allNodes.find(n => n.id === matchId);
            if (node) { node.x = pt.x; node.y = pt.y; node.source = 'envelope'; }
            envNodeIds.push(matchId);
        } else {
            // Add new envelope corner node
            const envNode = {
                id: 'env_' + i + '_' + (envelope.id || '0'),
                x: pt.x, y: pt.y, source: 'envelope',
            };
            allNodes.push(envNode);
            envNodeIds.push(envNode.id);
        }
    }

    // Add envelope boundary edges
    for (let i = 0; i < envPts.length; i++) {
        const j = (i + 1) % envPts.length;
        const idA = envNodeIds[i];
        const idB = envNodeIds[j];
        // Avoid duplicating an edge that was already drawn
        const alreadyExists = allEdges.some(
            e => (e.a === idA && e.b === idB) || (e.a === idB && e.b === idA)
        );
        if (!alreadyExists) {
            allEdges.push({ id: 'env_e_' + i, a: idA, b: idB, source: 'envelope' });
        }
    }

    return { mergedNodes: allNodes, mergedEdges: allEdges };
}

// ── DCEL Face Detection ───────────────────────────────────────

function _detectFaces(nodes, edges) {
    if (!nodes || !edges || nodes.length < 3 || edges.length < 3) return [];

    const nodeMap = {};
    for (const n of nodes) nodeMap[n.id] = n;

    // Build directed halfedges
    const halfedges = [];
    const heIndex = {}; // id → index
    const twinMap = {}; // hid → twin hid

    for (const e of edges) {
        if (!nodeMap[e.a] || !nodeMap[e.b]) continue;
        const abId = e.id + '_ab';
        const baId = e.id + '_ba';
        halfedges.push({ id: abId, from: e.a, to: e.b });
        halfedges.push({ id: baId, from: e.b, to: e.a });
        twinMap[abId] = baId;
        twinMap[baId] = abId;
    }

    for (let i = 0; i < halfedges.length; i++) heIndex[halfedges[i].id] = i;

    // For each node, sort outgoing halfedges by angle
    const outgoing = {}; // nodeId → [halfedge index]
    for (const n of nodes) outgoing[n.id] = [];
    for (let i = 0; i < halfedges.length; i++) {
        const h = halfedges[i];
        if (outgoing[h.from] !== undefined) outgoing[h.from].push(i);
    }

    for (const nodeId of Object.keys(outgoing)) {
        const n = nodeMap[nodeId];
        if (!n) continue;
        outgoing[nodeId].sort((ai, bi) => {
            const ha = halfedges[ai], hb = halfedges[bi];
            const na = nodeMap[ha.to], nb = nodeMap[hb.to];
            const angA = Math.atan2(na.y - n.y, na.x - n.x);
            const angB = Math.atan2(nb.y - n.y, nb.x - n.x);
            return angA - angB;
        });
    }

    // Build nextHe map using twin-rotation (face to the left of each halfedge)
    const nextHe = {}; // hid → next hid in face
    for (const h of halfedges) {
        const twinId = twinMap[h.id];
        const twinIdx = heIndex[twinId];
        const outList = outgoing[h.to] || [];
        const posInList = outList.indexOf(twinIdx);
        if (posInList === -1 || outList.length === 0) continue;
        // Predecessor in CCW sort = clockwise rotation
        const prevPos = (posInList - 1 + outList.length) % outList.length;
        nextHe[h.id] = halfedges[outList[prevPos]].id;
    }

    // Traverse faces
    const visited = new Set();
    const rawFaces = [];

    for (const h of halfedges) {
        if (visited.has(h.id) || nextHe[h.id] === undefined) continue;

        const faceHalfedgeIds = [];
        let cur = h.id;
        let guard = 0;
        while (!visited.has(cur) && guard < halfedges.length) {
            visited.add(cur);
            faceHalfedgeIds.push(cur);
            cur = nextHe[cur];
            guard++;
        }

        if (faceHalfedgeIds.length < 3) continue;

        const faceNodeIds = faceHalfedgeIds.map(hid => halfedges[heIndex[hid]].from);

        // Compute signed area (shoelace) in y-down coordinate system
        let area = 0;
        for (let i = 0; i < faceNodeIds.length; i++) {
            const j = (i + 1) % faceNodeIds.length;
            const ni = nodeMap[faceNodeIds[i]];
            const nj = nodeMap[faceNodeIds[j]];
            if (!ni || !nj) continue;
            area += (ni.x * nj.y) - (nj.x * ni.y);
        }
        area /= 2;

        rawFaces.push({ nodeIds: faceNodeIds, area });
    }

    // Interior faces have positive area (CW in y-down = interior)
    // Outer face has negative area (CCW in y-down = exterior)
    // Also guard: area must be > small minimum to exclude degenerate sliver faces
    const minArea = 500 * 500; // 0.25 m² minimum face area
    const interiorFaces = rawFaces.filter(f => f.area > minArea);

    // Sort faces by centroid compass direction (for labelling)
    const centreX = nodes.reduce((s, n) => s + n.x, 0) / nodes.length;
    const centreY = nodes.reduce((s, n) => s + n.y, 0) / nodes.length;

    return interiorFaces.map((f, i) => {
        const cx = f.nodeIds.reduce((s, id) => s + (nodeMap[id]?.x || 0), 0) / f.nodeIds.length;
        const cy = f.nodeIds.reduce((s, id) => s + (nodeMap[id]?.y || 0), 0) / f.nodeIds.length;
        const bearing = Math.atan2(cy - centreY, cx - centreX) * 180 / Math.PI;
        const label = _compassLabel(bearing);
        return {
            id: 'face_' + i,
            nodeIds: f.nodeIds,
            area: f.area,
            centroid: { x: cx, y: cy },
            label: label + ' slope',
            pitch: null,
        };
    });
}

function _compassLabel(deg) {
    // deg: atan2 angle where +x=east, +y=south (y-down)
    const norm = ((deg % 360) + 360) % 360;
    if (norm >= 337.5 || norm < 22.5) return 'East';
    if (norm < 67.5) return 'SE';
    if (norm < 112.5) return 'South';
    if (norm < 157.5) return 'SW';
    if (norm < 202.5) return 'West';
    if (norm < 247.5) return 'NW';
    if (norm < 292.5) return 'North';
    return 'NE';
}

// ── Level Helpers ─────────────────────────────────────────────

/** Returns true when the active level is the RF (roof) level */
function _isOnRFLevel() {
    if (typeof levelSystem === 'undefined') return false;
    const active = levelSystem.levels[levelSystem.activeLevelIndex];
    if (!active) return false;
    return active.id === 'RF' || active.name.toLowerCase().includes('roof');
}

/**
 * Find the building envelope regardless of which level it lives on.
 * (Envelope is typically on GF; skeleton lives on RF — cross-level lookup needed.)
 */
function _findEnvelopeAnyLevel() {
    return project.elements.find(el => el.type === 'buildingEnvelope');
}

// ── Canvas Rendering ─────────────────────────────────────────

// Face fill colour palette (semi-transparent)
const FACE_COLORS = [
    'rgba(59,130,246,0.18)',   // blue
    'rgba(16,185,129,0.18)',   // green
    'rgba(245,158,11,0.18)',   // amber
    'rgba(168,85,247,0.18)',   // purple
    'rgba(239,68,68,0.18)',    // red
    'rgba(236,72,153,0.18)',   // pink
];
const FACE_COLORS_HOVER = [
    'rgba(59,130,246,0.42)',
    'rgba(16,185,129,0.42)',
    'rgba(245,158,11,0.42)',
    'rgba(168,85,247,0.42)',
    'rgba(239,68,68,0.42)',
    'rgba(236,72,153,0.42)',
];

let _hoveredFaceId = null;

function drawRoofSkeleton(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;
    const skeleton = findSkeletonElement();
    const onRF = _isOnRFLevel();

    // ── Ghost envelope underlay — always visible on RF level ──
    // Shows the building footprint as a dotted reference when viewing roof plan
    if (onRF) {
        const envelope = _findEnvelopeAnyLevel();
        if (envelope && envelope.points && envelope.points.length >= 3) {
            ctx.save();
            ctx.setLineDash([6, 6]);
            ctx.strokeStyle = 'rgba(99,102,241,0.45)';
            ctx.lineWidth = 1.2;
            ctx.beginPath();
            for (let i = 0; i < envelope.points.length; i++) {
                const s = coords.realToSheet(envelope.points[i].x, envelope.points[i].y);
                const sc = coords.sheetToScreen(s.x, s.y);
                i === 0 ? ctx.moveTo(sc.x, sc.y) : ctx.lineTo(sc.x, sc.y);
            }
            ctx.closePath();
            ctx.stroke();
            // Highlight corners as snap targets only when actively drawing skeleton
            if (activeTool === 'roofSkeleton') {
                for (const pt of envelope.points) {
                    const s = coords.realToSheet(pt.x, pt.y);
                    const sc = coords.sheetToScreen(s.x, s.y);
                    ctx.strokeStyle = 'rgba(99,102,241,0.8)';
                    ctx.lineWidth = 1.2;
                    ctx.setLineDash([]);
                    ctx.beginPath();
                    ctx.arc(sc.x, sc.y, 5, 0, Math.PI * 2);
                    ctx.stroke();
                }
            }
            ctx.restore();
        }
    }

    // ── Draw committed skeleton (RF level only) ───────────────
    if (skeleton && skeleton.committed && onRF) {
        const nodeMap = {};
        for (const n of skeleton.nodes) nodeMap[n.id] = n;

        // Face fills
        if (skeleton.faces && skeleton.faces.length > 0) {
            ctx.save();
            for (let fi = 0; fi < skeleton.faces.length; fi++) {
                const face = skeleton.faces[fi];
                if (!face.nodeIds || face.nodeIds.length < 3) continue;

                const isHover = _hoveredFaceId === face.id;
                ctx.fillStyle = isHover
                    ? FACE_COLORS_HOVER[fi % FACE_COLORS_HOVER.length]
                    : FACE_COLORS[fi % FACE_COLORS.length];

                ctx.beginPath();
                for (let i = 0; i < face.nodeIds.length; i++) {
                    const n = nodeMap[face.nodeIds[i]];
                    if (!n) continue;
                    const s = coords.realToSheet(n.x, n.y);
                    const sc = coords.sheetToScreen(s.x, s.y);
                    i === 0 ? ctx.moveTo(sc.x, sc.y) : ctx.lineTo(sc.x, sc.y);
                }
                ctx.closePath();
                ctx.fill();

                // Face label at centroid
                const cs = coords.realToSheet(face.centroid.x, face.centroid.y);
                const csc = coords.sheetToScreen(cs.x, cs.y);
                ctx.font = `bold ${Math.max(9, 10 * zoom)}px sans-serif`;
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                if (face.pitch != null) {
                    ctx.fillStyle = '#1d4ed8';
                    ctx.fillText(face.pitch + '°', csc.x, csc.y);
                } else {
                    ctx.fillStyle = '#9ca3af';
                    ctx.fillText('dbl-click', csc.x, csc.y - 6);
                    ctx.font = `${Math.max(8, 9 * zoom)}px sans-serif`;
                    ctx.fillText('to set pitch', csc.x, csc.y + 6);
                }
            }
            ctx.restore();
        }

        // Draw drawn edges (ridge + hip lines) — solid red
        ctx.save();
        ctx.strokeStyle = '#dc2626';
        ctx.lineWidth = Math.max(1, 1.5 * zoom);
        ctx.setLineDash([]);
        for (const edge of skeleton.edges) {
            if (edge.source === 'envelope') continue; // don't redraw envelope outline
            const nA = nodeMap[edge.a], nB = nodeMap[edge.b];
            if (!nA || !nB) continue;
            const sA = coords.realToSheet(nA.x, nA.y);
            const sB = coords.realToSheet(nB.x, nB.y);
            const scA = coords.sheetToScreen(sA.x, sA.y);
            const scB = coords.sheetToScreen(sB.x, sB.y);
            ctx.beginPath();
            ctx.moveTo(scA.x, scA.y);
            ctx.lineTo(scB.x, scB.y);
            ctx.stroke();
        }
        // Draw drawn nodes (small dots at ridge/hip intersections)
        ctx.fillStyle = '#dc2626';
        for (const n of skeleton.nodes) {
            if (n.source === 'envelope') continue;
            const s = coords.realToSheet(n.x, n.y);
            const sc = coords.sheetToScreen(s.x, s.y);
            ctx.beginPath();
            ctx.arc(sc.x, sc.y, Math.max(2.5, 3 * zoom), 0, Math.PI * 2);
            ctx.fill();
        }
        ctx.restore();
    }

    // ── Draw in-progress drawing state ───────────────────────
    if (activeTool === 'roofSkeleton') {
        const state = roofSkeletonState;

        // Draw committed drawing-state edges
        ctx.save();
        ctx.strokeStyle = '#dc2626';
        ctx.lineWidth = Math.max(1, 1.5 * zoom);
        for (const edge of state.edges) {
            const nA = state.nodes.find(n => n.id === edge.a);
            const nB = state.nodes.find(n => n.id === edge.b);
            if (!nA || !nB) continue;
            const scA = coords.sheetToScreen(nA.x, nA.y);
            const scB = coords.sheetToScreen(nB.x, nB.y);
            ctx.beginPath();
            ctx.moveTo(scA.x, scA.y);
            ctx.lineTo(scB.x, scB.y);
            ctx.stroke();
        }

        // Draw nodes
        ctx.fillStyle = '#dc2626';
        for (const n of state.nodes) {
            const sc = coords.sheetToScreen(n.x, n.y);
            ctx.beginPath();
            ctx.arc(sc.x, sc.y, 3, 0, Math.PI * 2);
            ctx.fill();
        }

        // Draw in-progress segment (dashed) from chainStart to hover
        if (state.chainStartId !== null && state.hoverPoint) {
            const startNode = state.nodes.find(n => n.id === state.chainStartId);
            if (startNode) {
                ctx.setLineDash([5, 5]);
                ctx.strokeStyle = '#dc2626';
                ctx.lineWidth = 1;
                const scStart = coords.sheetToScreen(startNode.x, startNode.y);
                const scEnd = coords.sheetToScreen(state.hoverPoint.x, state.hoverPoint.y);
                ctx.beginPath();
                ctx.moveTo(scStart.x, scStart.y);
                ctx.lineTo(scEnd.x, scEnd.y);
                ctx.stroke();
                // Segment length label
                const sheetLen = _dist2d(state.hoverPoint.x, state.hoverPoint.y, startNode.x, startNode.y);
                const realLen = sheetLen * CONFIG.drawingScale;
                if (realLen > 100) {
                    const mx = (scStart.x + scEnd.x) / 2;
                    const my = (scStart.y + scEnd.y) / 2;
                    ctx.setLineDash([]);
                    ctx.font = '10px sans-serif';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'bottom';
                    ctx.fillStyle = '#dc2626';
                    ctx.fillText((realLen / 1000).toFixed(2) + 'm', mx, my - 4);
                }
            }
        }
        ctx.restore();

        // Draw snap highlight
        if (state.snapTarget) {
            const sc = coords.sheetToScreen(state.snapTarget.x, state.snapTarget.y);
            ctx.save();
            if (state.snapTarget.type === 'envelopeCorner' || state.snapTarget.type === 'envelopeMidpoint') {
                ctx.strokeStyle = '#6366f1';
                ctx.lineWidth = 1.5;
            } else {
                ctx.strokeStyle = '#dc2626';
                ctx.lineWidth = 1.5;
            }
            ctx.beginPath();
            ctx.arc(sc.x, sc.y, 6, 0, Math.PI * 2);
            ctx.stroke();
            ctx.restore();
        }

        // Status instruction
        ctx.save();
        ctx.font = '11px sans-serif';
        ctx.fillStyle = '#71717a';
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';
        const instrText = state.chainStartId
            ? 'Click to add point  |  Shift: 45°/90°  |  Right-click: end chain  |  Enter: commit'
            : 'Click to start a ridge or hip line  |  Right-click: end chain  |  Enter: commit';
        ctx.fillText(instrText, 10, 30);
        ctx.restore();
    }
}

// ── Double-click to Set Face Pitch ────────────────────────────

container.addEventListener('dblclick', function _skeletonDblclick(e) {
    const skeleton = findSkeletonElement();
    if (!skeleton || !skeleton.committed || !skeleton.faces) return;
    if (!_isOnRFLevel()) return;
    // Only intercept if not in an active drawing mode
    if (activeTool !== 'select' && activeTool !== 'roofSkeleton') return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;
    const sheetPos = engine.coords.screenToSheet(sx, sy);
    const realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);

    const nodeMap = {};
    for (const n of skeleton.nodes) nodeMap[n.id] = n;

    // Find face containing click point
    let clickedFace = null;
    for (const face of skeleton.faces) {
        if (_pointInFace(realPos, face, nodeMap)) {
            clickedFace = face;
            break;
        }
    }
    if (!clickedFace) return;
    e.preventDefault();
    e.stopPropagation();

    // Show floating input at the face centroid (in screen coords)
    _showPitchInput(clickedFace, skeleton);
});

function _pointInFace(pt, face, nodeMap) {
    // Ray casting point-in-polygon
    const verts = face.nodeIds.map(id => nodeMap[id]).filter(Boolean);
    if (verts.length < 3) return false;
    let inside = false;
    for (let i = 0, j = verts.length - 1; i < verts.length; j = i++) {
        const xi = verts[i].x, yi = verts[i].y;
        const xj = verts[j].x, yj = verts[j].y;
        if (((yi > pt.y) !== (yj > pt.y)) &&
            (pt.x < (xj - xi) * (pt.y - yi) / (yj - yi) + xi)) {
            inside = !inside;
        }
    }
    return inside;
}

function _showPitchInput(face, skeleton) {
    // Remove any existing input
    const oldInput = document.getElementById('face-pitch-input-overlay');
    if (oldInput) oldInput.remove();

    const cs = engine.coords.realToSheet(face.centroid.x, face.centroid.y);
    const csc = engine.coords.sheetToScreen(cs.x, cs.y);
    const rect = container.getBoundingClientRect();

    const wrapper = document.createElement('div');
    wrapper.id = 'face-pitch-input-overlay';
    wrapper.style.cssText = `
        position: fixed;
        left: ${rect.left + csc.x - 55}px;
        top: ${rect.top + csc.y - 32}px;
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 6px;
        padding: 6px 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.12);
        z-index: 9999;
        display: flex;
        flex-direction: column;
        gap: 4px;
        min-width: 110px;
    `;

    const label = document.createElement('div');
    label.textContent = face.label;
    label.style.cssText = 'font-size:10px;color:#71717a;font-weight:600;text-transform:uppercase;letter-spacing:0.5px;';
    wrapper.appendChild(label);

    const row = document.createElement('div');
    row.style.cssText = 'display:flex;align-items:center;gap:4px;';

    const input = document.createElement('input');
    input.type = 'number';
    input.min = '0'; input.max = '60'; input.step = '0.5';
    input.value = face.pitch != null ? face.pitch : '';
    input.placeholder = '22.5';
    input.style.cssText = 'width:60px;font-size:12px;padding:2px 4px;border:1px solid #d1d5db;border-radius:4px;';

    const deg = document.createElement('span');
    deg.textContent = '°';
    deg.style.cssText = 'font-size:12px;color:#374151;';

    row.appendChild(input);
    row.appendChild(deg);
    wrapper.appendChild(row);

    const applyAll = document.createElement('button');
    applyAll.textContent = 'Apply to all faces';
    applyAll.style.cssText = 'font-size:9px;padding:2px 6px;background:#f3f4f6;border:1px solid #e5e7eb;border-radius:3px;cursor:pointer;margin-top:2px;';
    wrapper.appendChild(applyAll);

    document.body.appendChild(wrapper);
    input.focus();
    input.select();

    const commit = () => {
        const val = parseFloat(input.value);
        if (!isNaN(val) && val >= 0 && val <= 60) {
            face.pitch = val;
            if (typeof markComplianceDirty === 'function') markComplianceDirty();
            engine.requestRender();
            if (typeof updateBracingSummaryPanel === 'function') updateBracingSummaryPanel();
        }
        wrapper.remove();
    };

    applyAll.addEventListener('click', () => {
        const val = parseFloat(input.value);
        if (!isNaN(val) && val >= 0 && val <= 60) {
            for (const f of skeleton.faces) f.pitch = val;
            if (typeof markComplianceDirty === 'function') markComplianceDirty();
            engine.requestRender();
            if (typeof updateBracingSummaryPanel === 'function') updateBracingSummaryPanel();
        }
        wrapper.remove();
    });

    input.addEventListener('keydown', ev => {
        if (ev.key === 'Enter') { ev.preventDefault(); commit(); }
        if (ev.key === 'Escape') { wrapper.remove(); }
    });
    input.addEventListener('blur', () => { setTimeout(() => wrapper.remove(), 150); });
}

// ── Hover Detection for Face Highlight ───────────────────────

container.addEventListener('mousemove', function _skeletonHoverFaces(e) {
    const skeleton = findSkeletonElement();
    if (!skeleton || !skeleton.committed || !skeleton.faces) return;
    if (!_isOnRFLevel()) return;
    if (activeTool !== 'select' && activeTool !== 'roofSkeleton') return;

    const rect = container.getBoundingClientRect();
    const sx = e.clientX - rect.left;
    const sy = e.clientY - rect.top;
    const sheetPos = engine.coords.screenToSheet(sx, sy);
    const realPos = engine.coords.sheetToReal(sheetPos.x, sheetPos.y);

    const nodeMap = {};
    for (const n of skeleton.nodes) nodeMap[n.id] = n;

    let newHovered = null;
    for (const face of skeleton.faces) {
        if (_pointInFace(realPos, face, nodeMap)) {
            newHovered = face.id;
            break;
        }
    }
    if (newHovered !== _hoveredFaceId) {
        _hoveredFaceId = newHovered;
        engine.requestRender();
    }
});

// ── 3D Model from Skeleton ────────────────────────────────────

function buildRoofModelFromSkeleton(skeleton, eavesHeightMM) {
    if (!skeleton || !skeleton.faces || skeleton.faces.length === 0) return null;

    const nodeMap = {};
    for (const n of skeleton.nodes) nodeMap[n.id] = n;

    // Compute 3D height for every node
    const nodeHeights = {};

    // Envelope-source nodes sit at eaves height
    for (const n of skeleton.nodes) {
        if (n.source === 'envelope') nodeHeights[n.id] = eavesHeightMM;
    }

    // Interior (drawn) nodes: height = perpendicular distance to nearest
    // eaves edge × tan(adjacent face pitch)
    for (const n of skeleton.nodes) {
        if (n.source !== 'drawn') continue;

        const adjacentFaces = skeleton.faces.filter(f => f.nodeIds.includes(n.id));
        const contributions = [];

        for (const face of adjacentFaces) {
            if (face.pitch == null || face.pitch <= 0) continue;
            const pitchRad = face.pitch * Math.PI / 180;

            // Find eaves edges of this face (edges where both endpoints are envelope-source)
            for (let i = 0; i < face.nodeIds.length; i++) {
                const j = (i + 1) % face.nodeIds.length;
                const nA = nodeMap[face.nodeIds[i]];
                const nB = nodeMap[face.nodeIds[j]];
                if (!nA || !nB) continue;
                if (nA.source === 'envelope' && nB.source === 'envelope') {
                    // Perpendicular distance from n to this eaves edge
                    const d = _perpDistToSegment(n, nA, nB);
                    contributions.push(d * Math.tan(pitchRad));
                }
            }
        }

        nodeHeights[n.id] = contributions.length > 0
            ? eavesHeightMM + Math.min(...contributions)
            : eavesHeightMM;
    }

    // Build 3D surfaces from faces
    const surfaces = [];
    for (const face of skeleton.faces) {
        if (!face.nodeIds || face.nodeIds.length < 3) continue;
        const vertices = face.nodeIds.map(nid => {
            const n = nodeMap[nid];
            return { x: n.x, y: n.y, z: nodeHeights[nid] ?? eavesHeightMM };
        });
        surfaces.push({ type: 'slope', label: face.label, pitch: face.pitch, vertices });
    }

    // Build wall faces from envelope edges
    const wallFaces = [];
    for (const edge of skeleton.edges) {
        if (edge.source !== 'envelope') continue;
        const nA = nodeMap[edge.a], nB = nodeMap[edge.b];
        if (!nA || !nB) continue;
        wallFaces.push({
            type: 'wall',
            vertices: [
                { x: nA.x, y: nA.y, z: 0 },
                { x: nB.x, y: nB.y, z: 0 },
                { x: nB.x, y: nB.y, z: eavesHeightMM },
                { x: nA.x, y: nA.y, z: eavesHeightMM },
            ]
        });
    }

    return {
        surfaces,
        wallFaces,
        eavesHeightMM,
        nodeHeights,
        isSkeletonModel: true,
    };
}

function _perpDistToSegment(P, A, B) {
    const dx = B.x - A.x, dy = B.y - A.y;
    const lenSq = dx * dx + dy * dy;
    if (lenSq < 0.01) return Math.sqrt((P.x - A.x) ** 2 + (P.y - A.y) ** 2);
    let t = ((P.x - A.x) * dx + (P.y - A.y) * dy) / lenSq;
    t = Math.max(0, Math.min(1, t));
    const cx = A.x + t * dx, cy = A.y + t * dy;
    return Math.sqrt((P.x - cx) ** 2 + (P.y - cy) ** 2);
}

// ── Auto-generate Hip Roof from Envelope ─────────────────────

function generateHipRoofFromEnvelope() {
    const envelope = findEnvelopeElement();
    if (!envelope || !envelope.points || envelope.points.length < 3) {
        alert('Please draw a building envelope first.');
        return;
    }

    const pts = envelope.points;
    let xmin = Infinity, xmax = -Infinity, ymin = Infinity, ymax = -Infinity;
    for (const p of pts) {
        xmin = Math.min(xmin, p.x); xmax = Math.max(xmax, p.x);
        ymin = Math.min(ymin, p.y); ymax = Math.max(ymax, p.y);
    }
    const W = xmax - xmin, H = ymax - ymin;
    const shortSpan = Math.min(W, H);
    const xmid = (xmin + xmax) / 2, ymid = (ymin + ymax) / 2;

    // Determine long axis and place ridge accordingly
    let r1, r2;
    if (W >= H) {
        // Long axis = X, short span = H, hip offset = H/2
        r1 = { id: _rsnId(), x: xmin + shortSpan / 2, y: ymid, source: 'drawn' };
        r2 = { id: _rsnId(), x: xmax - shortSpan / 2, y: ymid, source: 'drawn' };
    } else {
        // Long axis = Y
        r1 = { id: _rsnId(), x: xmid, y: ymin + shortSpan / 2, source: 'drawn' };
        r2 = { id: _rsnId(), x: xmid, y: ymax - shortSpan / 2, source: 'drawn' };
    }

    const drawnNodes = [r1, r2];
    const drawnEdges = [
        { id: _rseId(), a: r1.id, b: r2.id, source: 'drawn' }, // ridge
    ];

    // Find the 4 corners of the bounding rectangle in the envelope
    // Match each corner to the nearest envelope point
    const corners = [
        { x: xmin, y: ymin }, { x: xmax, y: ymin },
        { x: xmax, y: ymax }, { x: xmin, y: ymax },
    ];

    const matchedCorners = corners.map(c => {
        let best = pts[0], bestDist = Infinity;
        for (const p of pts) {
            const d = _dist2d(p.x, p.y, c.x, c.y);
            if (d < bestDist) { bestDist = d; best = p; }
        }
        return best;
    });

    // Determine which ridge point each corner connects to (nearest ridge end)
    for (const corner of matchedCorners) {
        const d1 = _dist2d(corner.x, corner.y, r1.x, r1.y);
        const d2 = _dist2d(corner.x, corner.y, r2.x, r2.y);
        // Create a temporary node at the corner position — will be merged with envelope
        const tempId = _rsnId();
        drawnNodes.push({ id: tempId, x: corner.x, y: corner.y, source: 'drawn' });
        drawnEdges.push({
            id: _rseId(),
            a: tempId,
            b: d1 <= d2 ? r1.id : r2.id,
            source: 'drawn',
        });
    }

    // Merge with envelope and detect faces
    const { mergedNodes, mergedEdges } = _mergeWithEnvelope(drawnNodes, drawnEdges);
    const faces = _detectFaces(mergedNodes, mergedEdges);

    // Default pitch 22.5°
    for (const f of faces) f.pitch = 22.5;

    _switchToRFLevel(); // always place skeleton on RF level
    const currentLevel = _getRFLevelId();
    const existing = findSkeletonElement();

    const skeleton = {
        id: generateId(),
        type: 'roofSkeleton',
        layer: 'S-RIDGE',
        level: currentLevel,
        nodes: mergedNodes,
        edges: mergedEdges,
        faces: faces,
        committed: true,
    };

    history.execute({
        description: 'Generate hip roof skeleton',
        execute() {
            if (existing) {
                const idx = project.elements.indexOf(existing);
                if (idx !== -1) project.elements.splice(idx, 1);
            }
            project.elements.push(skeleton);
        },
        undo() {
            const i = project.elements.indexOf(skeleton);
            if (i !== -1) project.elements.splice(i, 1);
            if (existing) project.elements.push(existing);
        },
    });

    if (typeof markComplianceDirty === 'function') markComplianceDirty();
    engine.requestRender();
    if (typeof updateBracingSummaryPanel === 'function') updateBracingSummaryPanel();
}

// ── Register Renderer ─────────────────────────────────────────

engine.onRender(drawRoofSkeleton);

// ── Global Exports ───────────────────────────────────────────

window.findSkeletonElement         = findSkeletonElement;
window.startDrawSkeleton           = startDrawSkeleton;
window.buildRoofModelFromSkeleton  = buildRoofModelFromSkeleton;
window.generateHipRoofFromEnvelope = generateHipRoofFromEnvelope;
