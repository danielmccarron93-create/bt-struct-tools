// ── PER-LEVEL PDF + MOVE PDF ─────────────────────────────
// ══════════════════════════════════════════════════════════

// Save current PDF position to level data when it changes
function savePdfToLevel() {
    const lv = typeof getActiveLevel === 'function' ? getActiveLevel() : null;
    if (!lv || !pdfState.loaded) return;
    const cached = pdfState.pageCanvases[pdfState.currentPage];
    pdfState.levelPdfs[lv.id] = {
        pdfDoc: pdfState.pdfDoc,
        pageNum: pdfState.currentPage,
        pageCanvas: cached,
        sheetX: pdfState.sheetX,
        sheetY: pdfState.sheetY,
        sheetW: pdfState.sheetWidth,
        sheetH: pdfState.sheetHeight,
        nativeW: pdfState.nativeWidth,
        nativeH: pdfState.nativeHeight,
    };
}

// Update the PDF level indicator badge
function updatePdfLevelIndicator() {
    const indicator = document.getElementById('pdf-level-indicator');
    const removeBtn = document.getElementById('btn-remove-pdf');
    if (!indicator) return;
    const lv = typeof getActiveLevel === 'function' ? getActiveLevel() : null;
    if (lv && pdfState.loaded) {
        indicator.textContent = lv.name;
        indicator.style.display = '';
        if (removeBtn) removeBtn.style.display = '';
    } else {
        indicator.style.display = 'none';
        if (removeBtn) removeBtn.style.display = 'none';
    }

    // Also show dots on level tabs for levels with PDFs
    const levelTabsDiv = document.getElementById('level-tabs');
    if (levelTabsDiv) {
        const tabs = levelTabsDiv.querySelectorAll('.level-tab');
        tabs.forEach(tab => {
            const lvId = tab.dataset.levelId;
            const dot = tab.querySelector('.pdf-dot');
            const hasPdf = pdfState.levelPdfs[lvId] && pdfState.levelPdfs[lvId].pageCanvas;
            if (hasPdf && !dot) {
                const d = document.createElement('span');
                d.className = 'pdf-dot';
                d.style.cssText = 'display:inline-block;width:5px;height:5px;border-radius:50%;background:#3b82f6;margin-left:4px;vertical-align:middle;';
                d.title = 'PDF loaded';
                tab.appendChild(d);
            } else if (!hasPdf && dot) {
                dot.remove();
            }
        });
    }
}

// Restore PDF from level data when switching levels
function restorePdfFromLevel(levelId) {
    const lpdf = pdfState.levelPdfs[levelId];
    if (lpdf && lpdf.pageCanvas) {
        pdfState.loaded = true;
        pdfState.pdfDoc = lpdf.pdfDoc;
        pdfState.currentPage = lpdf.pageNum;
        pdfState.pageCanvases[lpdf.pageNum] = lpdf.pageCanvas;
        pdfState.sheetX = lpdf.sheetX;
        pdfState.sheetY = lpdf.sheetY;
        pdfState.sheetWidth = lpdf.sheetW;
        pdfState.sheetHeight = lpdf.sheetH;
        pdfState.nativeWidth = lpdf.nativeW;
        pdfState.nativeHeight = lpdf.nativeH;
        pdfState.visible = true;
        showPdfControls(true);
        document.getElementById('btn-move-pdf').style.display = '';
    } else {
        // This level has no PDF — hide controls but don't clear other levels
        pdfState.loaded = false;
        document.getElementById('btn-pdf-toggle').style.display = 'none';
        document.getElementById('pdf-page-select').style.display = 'none';
        document.getElementById('pdf-controls').style.display = 'none';
        document.getElementById('pdf-calib-group').style.display = 'none';
        document.getElementById('btn-move-pdf').style.display = 'none';
    }
    updatePdfLevelIndicator();
}

// Remove PDF from current level
document.getElementById('btn-remove-pdf').addEventListener('click', () => {
    const lv = typeof getActiveLevel === 'function' ? getActiveLevel() : null;
    if (!lv) return;
    if (!confirm('Remove PDF from ' + lv.name + '?')) return;
    delete pdfState.levelPdfs[lv.id];
    pdfState.loaded = false;
    pdfState.pdfDoc = null;
    pdfState.pageCanvases = {};
    showPdfControls(false);
    document.getElementById('btn-move-pdf').style.display = 'none';
    updatePdfLevelIndicator();
    engine.requestRender();
    console.log('[PDF] Removed PDF from ' + lv.name);
});

// Hook into level switching — save current PDF pos, restore next level's
const origSwitchLevelPdf = switchToLevel;
switchToLevel = function(index) {
    // Save current level's PDF position
    savePdfToLevel();
    // Switch level
    origSwitchLevelPdf(index);
    // Restore new level's PDF
    const newLv = levelSystem.levels[index];
    if (newLv) restorePdfFromLevel(newLv.id);
    updatePdfLevelIndicator();
    engine.requestRender();
};

// ── Move PDF Tool ────────────────────────────────────────

const movePdfBtn = document.getElementById('btn-move-pdf');

movePdfBtn.addEventListener('click', () => {
    setActiveTool('movepdf');
});

// Add movepdf to setActiveTool (simple inline patch)
container.addEventListener('mousedown', (e) => {
    if (activeTool !== 'movepdf') return;
    if (e.button !== 0 || engine._spaceDown) return;
    if (!pdfState.loaded) return;

    pdfState.draggingPdf = true;
    const sheetPos = engine.getSheetPos(e);
    pdfState.pdfDragStart = { x: sheetPos.x, y: sheetPos.y };
    pdfState.pdfDragOrigPos = { x: pdfState.sheetX, y: pdfState.sheetY };
    container.style.cursor = 'move';
    e.stopPropagation();
});

window.addEventListener('mousemove', (e) => {
    if (!pdfState.draggingPdf) return;
    const sheetPos = engine.getSheetPos(e);
    const dx = sheetPos.x - pdfState.pdfDragStart.x;
    const dy = sheetPos.y - pdfState.pdfDragStart.y;

    pdfState.sheetX = pdfState.pdfDragOrigPos.x + dx;
    pdfState.sheetY = pdfState.pdfDragOrigPos.y + dy;

    // Snap to grid if snap is enabled
    if (snapState.enabled && CONFIG.gridVisible) {
        const minor = CONFIG.GRID_MINOR_MM;
        pdfState.sheetX = Math.round(pdfState.sheetX / minor) * minor;
        pdfState.sheetY = Math.round(pdfState.sheetY / minor) * minor;
    }

    engine.requestRender();
});

window.addEventListener('mouseup', (e) => {
    if (!pdfState.draggingPdf) return;
    pdfState.draggingPdf = false;
    container.style.cursor = activeTool === 'movepdf' ? 'move' : '';
    savePdfToLevel(); // save new position
    engine.requestRender();
});

// Enter key confirms PDF move position (exits move tool)
window.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && activeTool === 'movepdf') {
        e.preventDefault();
        setActiveTool('select');
        savePdfToLevel();
    }
});

// When page selector changes, update per-level data
document.getElementById('pdf-page-select').addEventListener('change', () => {
    setTimeout(savePdfToLevel, 200);
});

// ══════════════════════════════════════════════════════════
// ── AUTO MEMBER SCHEDULE ON SHEET ────────────────────────
// ══════════════════════════════════════════════════════════

// Schedule position and size state (in sheet-mm) — draggable + resizable
const schedLayout = {
    // Default position: right side, above title block (set on first render)
    x: null, y: null,       // top-left corner in sheet-mm (null = auto-position)
    width: 108,             // total width in sheet-mm (20% bigger: was 90)
    markColRatio: 0.28,     // ratio of width for MARK column
    rowH: 5,               // row height mm (20% bigger: was 4.2)
    headerH: 7.2,          // header height mm
    titleH: 7.2,           // title height mm
    // Drag state
    dragging: false,
    resizing: false,
    dragStartSheet: null,
    dragOrigPos: null,
    resizeStartSheet: null,
    resizeOrigW: null,
};

function getSchedEntries() {
    const activeId = typeof getActiveLevel === 'function' ? getActiveLevel().id : null;
    const members = {};
    const tagOrder = { SC: 1, SB: 2, B: 3, S: 4, BT: 5, J: 6, R: 7, WH: 8, XBR: 9, L: 10, PP: 11, BL: 12, ET: 13, SF: 14, BW: 15 };

    for (const el of project.elements) {
        if (activeId && el.level !== activeId) continue;
        if (!el.tag) continue;
        let mark, member;
        if (el.type === 'column') {
            const cRef = el.typeRef || (el.tag.includes('-') ? el.tag.split('-')[0] : el.tag);
            mark = cRef;
            const cSched = project.scheduleTypes.column[cRef];
            member = (cSched && cSched.size) ? cSched.size : (el.memberSize || (el.size ? el.size + 'x' + el.size + ' SHS' : ''));
        } else if (el.type === 'line' && el.tag && (el.layer === 'S-BEAM' || el.layer === 'S-WALL' || el.layer === 'S-FTNG')) {
            const bRef = el.typeRef || (el.tag.includes('-') ? el.tag.split('-')[0] : el.tag);
            mark = bRef;
            const bSched = project.scheduleTypes.beam[bRef];
            member = (bSched && bSched.size) ? bSched.size : (el.memberSize || '');
        } else if (el.type === 'stripFooting' && el.tag) {
            mark = el.tag;
            member = `${el.footingWidth}W x ${el.footingDepth}D STRIP FTG`;
        } else continue;
        if (!mark || !member) continue;
        if (!members[mark]) members[mark] = member;
    }
    const entries = Object.entries(members);
    entries.sort((a, b) => {
        const pA = a[0].replace(/[0-9]/g, ''), pB = b[0].replace(/[0-9]/g, '');
        const nA = parseInt(a[0].replace(/[^0-9]/g, '')) || 0, nB = parseInt(b[0].replace(/[^0-9]/g, '')) || 0;
        const oA = tagOrder[pA] || 99, oB = tagOrder[pB] || 99;
        return oA !== oB ? oA - oB : nA - nB;
    });
    return entries;
}

function getSchedBounds(entries) {
    const c = CONFIG;
    const sl = schedLayout;
    const totalH = sl.titleH + sl.headerH + entries.length * sl.rowH + 2;

    // Auto-position on first render: below isometric view + title text, left-aligned
    if (sl.x === null) {
        sl.x = isoView.x;
        sl.y = isoView.y + isoView.height + 15; // +15mm clears the "ISOMETRIC VIEW" title + underline
    }

    return { left: sl.x, top: sl.y, width: sl.width, height: totalH };
}

function drawAutoMemberSchedule(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;
    if (zoom < 0.2) return;

    const entries = getSchedEntries();
    if (entries.length === 0) return;

    const sl = schedLayout;
    const bounds = getSchedBounds(entries);
    const { left: sL, top: sT, width: sW, height: sH } = bounds;
    const colMarkW = sW * sl.markColRatio;
    const colMemberW = sW - colMarkW;

    // Background
    const tlS = coords.sheetToScreen(sL, sT);
    const brS = coords.sheetToScreen(sL + sW, sT + sH);
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(tlS.x, tlS.y, brS.x - tlS.x, brS.y - tlS.y);

    // Border
    ctx.strokeStyle = '#000000';
    ctx.lineWidth = Math.max(0.5, 0.35 * zoom);
    ctx.strokeRect(tlS.x, tlS.y, brS.x - tlS.x, brS.y - tlS.y);

    // Title
    const titleP = coords.sheetToScreen(sL + sW / 2, sT + sl.titleH * 0.55);
    const titleFs = Math.max(1, 3.6 * zoom);
    ctx.font = `bold ${titleFs}px "Segoe UI", Arial, sans-serif`;
    ctx.fillStyle = '#000000';
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText('MEMBER SCHEDULE', titleP.x, titleP.y);

    // Title line
    const tlY = sT + sl.titleH;
    const t1 = coords.sheetToScreen(sL, tlY), t2 = coords.sheetToScreen(sL + sW, tlY);
    ctx.beginPath(); ctx.moveTo(t1.x, t1.y); ctx.lineTo(t2.x, t2.y); ctx.stroke();

    // Header
    const hY = sT + sl.titleH + sl.headerH * 0.5;
    const hFs = Math.max(1, 2.6 * zoom);
    ctx.font = `bold ${hFs}px "Segoe UI", Arial, sans-serif`;
    const hMP = coords.sheetToScreen(sL + colMarkW * 0.5, hY);
    const hMeP = coords.sheetToScreen(sL + colMarkW + colMemberW * 0.5, hY);
    ctx.fillText('MARK', hMP.x, hMP.y);
    ctx.fillText('MEMBER', hMeP.x, hMeP.y);

    // Header line
    const hlY = sT + sl.titleH + sl.headerH;
    const h1 = coords.sheetToScreen(sL, hlY), h2 = coords.sheetToScreen(sL + sW, hlY);
    ctx.lineWidth = Math.max(0.5, 0.3 * zoom);
    ctx.beginPath(); ctx.moveTo(h1.x, h1.y); ctx.lineTo(h2.x, h2.y); ctx.stroke();

    // Column divider
    const dX = sL + colMarkW;
    const dT = coords.sheetToScreen(dX, sT + sl.titleH), dB = coords.sheetToScreen(dX, sT + sH);
    ctx.beginPath(); ctx.moveTo(dT.x, dT.y); ctx.lineTo(dB.x, dB.y); ctx.stroke();

    // Data rows
    const dFs = Math.max(1, 2.2 * zoom);
    ctx.lineWidth = Math.max(0.3, 0.15 * zoom);
    for (let i = 0; i < entries.length; i++) {
        const [mark, member] = entries[i];
        const rT = sT + sl.titleH + sl.headerH + i * sl.rowH;
        const rM = rT + sl.rowH * 0.5;

        // Row separator
        const r1 = coords.sheetToScreen(sL, rT + sl.rowH), r2 = coords.sheetToScreen(sL + sW, rT + sl.rowH);
        ctx.strokeStyle = '#CCCCCC';
        ctx.beginPath(); ctx.moveTo(r1.x, r1.y); ctx.lineTo(r2.x, r2.y); ctx.stroke();

        // Mark
        const mP = coords.sheetToScreen(sL + colMarkW * 0.5, rM);
        ctx.font = `bold ${dFs}px "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = '#000000'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillText(mark, mP.x, mP.y);

        // Member
        const meP = coords.sheetToScreen(sL + colMarkW + 2, rM);
        ctx.font = `${dFs}px "Segoe UI", Arial, sans-serif`;
        ctx.textAlign = 'left';
        ctx.fillText(member, meP.x, meP.y);
    }

    // Resize handle (small triangle at bottom-right corner)
    const brCorner = coords.sheetToScreen(sL + sW, sT + sH);
    const rhSize = 6;
    ctx.fillStyle = '#AAAAAA';
    ctx.beginPath();
    ctx.moveTo(brCorner.x, brCorner.y);
    ctx.lineTo(brCorner.x - rhSize, brCorner.y);
    ctx.lineTo(brCorner.x, brCorner.y - rhSize);
    ctx.closePath();
    ctx.fill();

    ctx.strokeStyle = '#000000'; ctx.textAlign = 'left'; ctx.textBaseline = 'top';
}

engine.onRender(drawAutoMemberSchedule);

// ── AUTO STRIP FOOTING SCHEDULE ON SHEET ──
function getStripFtgEntries() {
    const activeId = typeof getActiveLevel === 'function' ? getActiveLevel().id : null;
    const entries = {};
    for (const el of project.elements) {
        if (el.type !== 'stripFooting') continue;
        if (activeId && el.level !== activeId) continue;
        if (!el.tag) continue;
        if (!entries[el.tag]) {
            entries[el.tag] = {
                mark: el.tag,
                size: `${el.footingWidth}W x ${el.footingDepth}D`,
                reinforcement: el.reinforcement || ''
            };
        }
    }
    return Object.values(entries).sort((a, b) => {
        const nA = parseInt(a.mark.replace(/[^0-9]/g, '')) || 0;
        const nB = parseInt(b.mark.replace(/[^0-9]/g, '')) || 0;
        return nA - nB;
    });
}

function drawStripFtgSchedule(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;
    if (zoom < 0.2) return;

    const sfEntries = getStripFtgEntries();
    if (sfEntries.length === 0) return;

    // Position below the member schedule
    const memberEntries = getSchedEntries();
    const memberBounds = getSchedBounds(memberEntries);
    const sl = schedLayout;

    const titleH = 7.2, headerH = 7.2, rowH = 5;
    const totalH = titleH + headerH + sfEntries.length * rowH + 2;
    const sL = sl.x;
    const sT = memberBounds.top + memberBounds.height + 5; // 5mm gap below member schedule
    const sW = sl.width;

    // 3-column layout: MARK | SIZE | REINFORCEMENT
    const col1W = sW * 0.2;
    const col2W = sW * 0.35;
    const col3W = sW - col1W - col2W;

    // Background
    const tlS = coords.sheetToScreen(sL, sT);
    const brS = coords.sheetToScreen(sL + sW, sT + totalH);
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(tlS.x, tlS.y, brS.x - tlS.x, brS.y - tlS.y);

    // Border
    ctx.strokeStyle = '#000000';
    ctx.lineWidth = Math.max(0.5, 0.35 * zoom);
    ctx.strokeRect(tlS.x, tlS.y, brS.x - tlS.x, brS.y - tlS.y);

    // Title
    const titleP = coords.sheetToScreen(sL + sW / 2, sT + titleH * 0.55);
    const titleFs = Math.max(1, 3.6 * zoom);
    ctx.font = `bold ${titleFs}px "Segoe UI", Arial, sans-serif`;
    ctx.fillStyle = '#000000';
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText('STRIP FOOTING SCHEDULE', titleP.x, titleP.y);

    // Title line
    const tlY = sT + titleH;
    const t1 = coords.sheetToScreen(sL, tlY), t2 = coords.sheetToScreen(sL + sW, tlY);
    ctx.beginPath(); ctx.moveTo(t1.x, t1.y); ctx.lineTo(t2.x, t2.y); ctx.stroke();

    // Header
    const hY = sT + titleH + headerH * 0.5;
    const hFs = Math.max(1, 2.6 * zoom);
    ctx.font = `bold ${hFs}px "Segoe UI", Arial, sans-serif`;
    const h1P = coords.sheetToScreen(sL + col1W * 0.5, hY);
    const h2P = coords.sheetToScreen(sL + col1W + col2W * 0.5, hY);
    const h3P = coords.sheetToScreen(sL + col1W + col2W + col3W * 0.5, hY);
    ctx.fillText('MARK', h1P.x, h1P.y);
    ctx.fillText('SIZE', h2P.x, h2P.y);
    ctx.fillText('REINF.', h3P.x, h3P.y);

    // Header line
    const hlY = sT + titleH + headerH;
    const hl1 = coords.sheetToScreen(sL, hlY), hl2 = coords.sheetToScreen(sL + sW, hlY);
    ctx.lineWidth = Math.max(0.5, 0.3 * zoom);
    ctx.beginPath(); ctx.moveTo(hl1.x, hl1.y); ctx.lineTo(hl2.x, hl2.y); ctx.stroke();

    // Column dividers
    const d1x = sL + col1W;
    const d2x = sL + col1W + col2W;
    const dT1 = coords.sheetToScreen(d1x, sT + titleH), dB1 = coords.sheetToScreen(d1x, sT + totalH);
    const dT2 = coords.sheetToScreen(d2x, sT + titleH), dB2 = coords.sheetToScreen(d2x, sT + totalH);
    ctx.beginPath(); ctx.moveTo(dT1.x, dT1.y); ctx.lineTo(dB1.x, dB1.y); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(dT2.x, dT2.y); ctx.lineTo(dB2.x, dB2.y); ctx.stroke();

    // Data rows
    const dFs = Math.max(1, 2.2 * zoom);
    ctx.lineWidth = Math.max(0.3, 0.15 * zoom);
    for (let i = 0; i < sfEntries.length; i++) {
        const entry = sfEntries[i];
        const rT = sT + titleH + headerH + i * rowH;
        const rM = rT + rowH * 0.5;

        // Row separator
        const r1 = coords.sheetToScreen(sL, rT + rowH), r2 = coords.sheetToScreen(sL + sW, rT + rowH);
        ctx.strokeStyle = '#CCCCCC';
        ctx.beginPath(); ctx.moveTo(r1.x, r1.y); ctx.lineTo(r2.x, r2.y); ctx.stroke();

        // Mark
        const mP = coords.sheetToScreen(sL + col1W * 0.5, rM);
        ctx.font = `bold ${dFs}px "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = '#000000'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillText(entry.mark, mP.x, mP.y);

        // Size
        const sP = coords.sheetToScreen(sL + col1W + col2W * 0.5, rM);
        ctx.font = `${dFs}px "Segoe UI", Arial, sans-serif`;
        ctx.fillText(entry.size, sP.x, sP.y);

        // Reinforcement
        const rP = coords.sheetToScreen(sL + col1W + col2W + 2, rM);
        ctx.textAlign = 'left';
        ctx.fillText(entry.reinforcement, rP.x, rP.y);
    }

    ctx.strokeStyle = '#000000'; ctx.textAlign = 'left'; ctx.textBaseline = 'top';
}

engine.onRender(drawStripFtgSchedule);

// ── AUTO PAD FOOTING SCHEDULE ON SHEET ──
function getPadFtgEntries() {
    const activeId = typeof getActiveLevel === 'function' ? getActiveLevel().id : null;
    const entries = {};
    for (const el of project.elements) {
        if (el.type !== 'footing') continue;
        if (activeId && el.level !== activeId) continue;
        if (!el.mark) continue;
        if (!entries[el.mark]) {
            const fw = el.footingWidth || el.width || 1000;
            const fd = el.footingDepth || el.depth || 400;
            entries[el.mark] = {
                mark: el.mark,
                size: `${fw}x${fw}x${fd}D`,
                reinforcement: el.reinforcement || ''
            };
        }
    }
    return Object.values(entries).sort((a, b) => {
        const nA = parseInt(a.mark.replace(/[^0-9]/g, '')) || 0;
        const nB = parseInt(b.mark.replace(/[^0-9]/g, '')) || 0;
        return nA - nB;
    });
}

function drawPadFtgSchedule(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;
    if (zoom < 0.2) return;

    const pfEntries = getPadFtgEntries();
    if (pfEntries.length === 0) return;

    // Position below the strip footing schedule (or member schedule if no strip footings)
    const memberEntries = getSchedEntries();
    const memberBounds = getSchedBounds(memberEntries);
    const sl = schedLayout;

    const sfEntries = getStripFtgEntries();
    let topY;
    if (sfEntries.length > 0) {
        // Below strip footing schedule
        const sfTitleH = 7.2, sfHeaderH = 7.2, sfRowH = 5;
        const sfTotalH = sfTitleH + sfHeaderH + sfEntries.length * sfRowH + 2;
        const sfTop = memberBounds.top + memberBounds.height + 5;
        topY = sfTop + sfTotalH + 5; // 5mm gap below strip footing schedule
    } else {
        // Below member schedule
        topY = memberBounds.top + memberBounds.height + 5;
    }

    const titleH = 7.2, headerH = 7.2, rowH = 5;
    const totalH = titleH + headerH + pfEntries.length * rowH + 2;
    const sL = sl.x;
    const sT = topY;
    const sW = sl.width;

    // 3-column layout: MARK | SIZE | REINFORCEMENT
    const col1W = sW * 0.2;
    const col2W = sW * 0.35;
    const col3W = sW - col1W - col2W;

    // Background
    const tlS = coords.sheetToScreen(sL, sT);
    const brS = coords.sheetToScreen(sL + sW, sT + totalH);
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(tlS.x, tlS.y, brS.x - tlS.x, brS.y - tlS.y);

    // Border
    ctx.strokeStyle = '#000000';
    ctx.lineWidth = Math.max(0.5, 0.35 * zoom);
    ctx.strokeRect(tlS.x, tlS.y, brS.x - tlS.x, brS.y - tlS.y);

    // Title
    const titleP = coords.sheetToScreen(sL + sW / 2, sT + titleH * 0.55);
    const titleFs = Math.max(1, 3.6 * zoom);
    ctx.font = `bold ${titleFs}px "Segoe UI", Arial, sans-serif`;
    ctx.fillStyle = '#000000';
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText('PAD FOOTING SCHEDULE', titleP.x, titleP.y);

    // Title line
    const tlY = sT + titleH;
    const t1 = coords.sheetToScreen(sL, tlY), t2 = coords.sheetToScreen(sL + sW, tlY);
    ctx.beginPath(); ctx.moveTo(t1.x, t1.y); ctx.lineTo(t2.x, t2.y); ctx.stroke();

    // Header
    const hY = sT + titleH + headerH * 0.5;
    const hFs = Math.max(1, 2.6 * zoom);
    ctx.font = `bold ${hFs}px "Segoe UI", Arial, sans-serif`;
    const h1P = coords.sheetToScreen(sL + col1W * 0.5, hY);
    const h2P = coords.sheetToScreen(sL + col1W + col2W * 0.5, hY);
    const h3P = coords.sheetToScreen(sL + col1W + col2W + col3W * 0.5, hY);
    ctx.fillText('MARK', h1P.x, h1P.y);
    ctx.fillText('SIZE', h2P.x, h2P.y);
    ctx.fillText('REINF.', h3P.x, h3P.y);

    // Header line
    const hlY = sT + titleH + headerH;
    const hl1 = coords.sheetToScreen(sL, hlY), hl2 = coords.sheetToScreen(sL + sW, hlY);
    ctx.lineWidth = Math.max(0.5, 0.3 * zoom);
    ctx.beginPath(); ctx.moveTo(hl1.x, hl1.y); ctx.lineTo(hl2.x, hl2.y); ctx.stroke();

    // Column dividers
    const d1x = sL + col1W;
    const d2x = sL + col1W + col2W;
    const dT1 = coords.sheetToScreen(d1x, sT + titleH), dB1 = coords.sheetToScreen(d1x, sT + totalH);
    const dT2 = coords.sheetToScreen(d2x, sT + titleH), dB2 = coords.sheetToScreen(d2x, sT + totalH);
    ctx.beginPath(); ctx.moveTo(dT1.x, dT1.y); ctx.lineTo(dB1.x, dB1.y); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(dT2.x, dT2.y); ctx.lineTo(dB2.x, dB2.y); ctx.stroke();

    // Data rows
    const dFs = Math.max(1, 2.2 * zoom);
    ctx.lineWidth = Math.max(0.3, 0.15 * zoom);
    for (let i = 0; i < pfEntries.length; i++) {
        const entry = pfEntries[i];
        const rT = sT + titleH + headerH + i * rowH;
        const rM = rT + rowH * 0.5;

        // Row separator
        const r1 = coords.sheetToScreen(sL, rT + rowH), r2 = coords.sheetToScreen(sL + sW, rT + rowH);
        ctx.strokeStyle = '#CCCCCC';
        ctx.beginPath(); ctx.moveTo(r1.x, r1.y); ctx.lineTo(r2.x, r2.y); ctx.stroke();

        // Mark
        const mP = coords.sheetToScreen(sL + col1W * 0.5, rM);
        ctx.font = `bold ${dFs}px "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = '#000000'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillText(entry.mark, mP.x, mP.y);

        // Size
        const sP = coords.sheetToScreen(sL + col1W + col2W * 0.5, rM);
        ctx.font = `${dFs}px "Segoe UI", Arial, sans-serif`;
        ctx.fillText(entry.size, sP.x, sP.y);

        // Reinforcement
        const rP = coords.sheetToScreen(sL + col1W + col2W + 2, rM);
        ctx.textAlign = 'left';
        ctx.fillText(entry.reinforcement, rP.x, rP.y);
    }

    ctx.strokeStyle = '#000000'; ctx.textAlign = 'left'; ctx.textBaseline = 'top';
}

engine.onRender(drawPadFtgSchedule);

// ── Schedule Drag + Resize ───────────────────────────────

container.addEventListener('mousedown', (e) => {
    if (activeTool !== 'select' || e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;

    const entries = getSchedEntries();
    if (entries.length === 0) return;
    const bounds = getSchedBounds(entries);

    const sheetPos = engine.getSheetPos(e);
    const { left: sL, top: sT, width: sW, height: sH } = bounds;

    // Check resize handle (bottom-right corner, 3mm hit zone)
    if (Math.abs(sheetPos.x - (sL + sW)) < 3 && Math.abs(sheetPos.y - (sT + sH)) < 3) {
        schedLayout.resizing = true;
        schedLayout.resizeStartSheet = { x: sheetPos.x, y: sheetPos.y };
        schedLayout.resizeOrigW = schedLayout.width;
        container.style.cursor = 'nwse-resize';
        e.stopPropagation();
        return;
    }

    // Check if clicking inside the schedule title bar area (top 7mm)
    if (sheetPos.x >= sL && sheetPos.x <= sL + sW &&
        sheetPos.y >= sT && sheetPos.y <= sT + schedLayout.titleH) {
        schedLayout.dragging = true;
        schedLayout.dragStartSheet = { x: sheetPos.x, y: sheetPos.y };
        schedLayout.dragOrigPos = { x: schedLayout.x, y: schedLayout.y };
        container.style.cursor = 'move';
        e.stopPropagation();
        return;
    }
}, true);

window.addEventListener('mousemove', (e) => {
    if (schedLayout.dragging) {
        const sheetPos = engine.getSheetPos(e);
        const dx = sheetPos.x - schedLayout.dragStartSheet.x;
        const dy = sheetPos.y - schedLayout.dragStartSheet.y;
        schedLayout.x = schedLayout.dragOrigPos.x + dx;
        schedLayout.y = schedLayout.dragOrigPos.y + dy;
        engine.requestRender();
    }
    if (schedLayout.resizing) {
        const sheetPos = engine.getSheetPos(e);
        const dx = sheetPos.x - schedLayout.resizeStartSheet.x;
        schedLayout.width = Math.max(60, schedLayout.resizeOrigW + dx);
        engine.requestRender();
    }
});


window.addEventListener('mouseup', (e) => {
    if (schedLayout.dragging) {
        schedLayout.dragging = false;
        container.style.cursor = '';
        engine.requestRender();
    }
    if (schedLayout.resizing) {
        schedLayout.resizing = false;
        container.style.cursor = '';
        engine.requestRender();
    }
});

window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && fullscreen3d) exitFullscreen3D();
});

// ══════════════════════════════════════════════════════════
