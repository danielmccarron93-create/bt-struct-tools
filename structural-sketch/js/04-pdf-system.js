// ── PHASE 2: PDF UNDERLAY SYSTEM ─────────────────────────
// ══════════════════════════════════════════════════════════

// Initialise pdf.js worker
if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

// PDF state
const pdfState = {
    loaded: false,
    visible: true,
    opacity: 0.30,
    pdfDoc: null,          // pdf.js document object
    totalPages: 0,
    currentPage: 1,
    pageCanvases: {},      // cache of rendered page canvases by page number
    // Positioning on sheet (in sheet-mm)
    sheetX: 0,
    sheetY: 0,
    sheetWidth: 0,
    sheetHeight: 0,
    nativeWidth: 0,
    nativeHeight: 0,
    renderScale: 2.5,      // render at 2.5x for crisp quality
    // Calibration
    calibrating: false,
    calibPoints: [],       // [{x, y}] in sheet-mm, up to 2
    // Per-level PDF underlays: { levelId: { pdfDoc, pageNum, pageCanvas, sheetX, sheetY, sheetW, sheetH } }
    levelPdfs: {},
    // PDF drag state
    draggingPdf: false,
    pdfDragStart: null,
    pdfDragOrigPos: null,
};

// ── PDF: Scale Auto-Detection ────────────────────────────

/**
 * Scan PDF text content for common scale notations.
 * Searches all pages (or first 3) for patterns like:
 *   "1:100", "1 : 200", "SCALE 1:50", "@ 1:100"
 * Returns the detected scale integer (e.g. 100) or null.
 */
async function detectPdfScale(pdfDoc) {
    const pagesToScan = Math.min(pdfDoc.numPages, 3);
    const scalePattern = /(?:scale\s*[:=]?\s*)?1\s*:\s*(\d{1,4})/gi;
    const validScales = [10, 20, 25, 50, 75, 100, 150, 200, 250, 500, 1000];

    for (let i = 1; i <= pagesToScan; i++) {
        try {
            const page = await pdfDoc.getPage(i);
            const textContent = await page.getTextContent();
            const fullText = textContent.items.map(item => item.str).join(' ');

            let match;
            while ((match = scalePattern.exec(fullText)) !== null) {
                const scaleVal = parseInt(match[1], 10);
                // Only accept standard engineering scales
                if (validScales.includes(scaleVal)) {
                    console.log('[PDF Scale] Detected 1:' + scaleVal +
                        ' on page ' + i + ' from text: "' + match[0] + '"');
                    return scaleVal;
                }
            }
        } catch (err) {
            console.warn('[PDF Scale] Could not read text from page ' + i, err);
        }
    }
    return null;
}

/**
 * Show the scale confirmation modal.
 * Returns a Promise that resolves with the selected scale (int)
 * or null if user clicks Skip.
 */
function showScaleConfirmModal(detectedScale) {
    return new Promise((resolve) => {
        const modal = document.getElementById('pdf-scale-modal');
        const select = document.getElementById('pdf-scale-select-modal');
        const msg = document.getElementById('pdf-scale-detect-msg');
        const applyBtn = document.getElementById('pdf-scale-apply');
        const cancelBtn = document.getElementById('pdf-scale-cancel');

        // Pre-fill with detected scale or default 1:100
        if (detectedScale) {
            select.value = String(detectedScale);
            msg.innerHTML = 'Detected <strong>1:' + detectedScale +
                '</strong> from the PDF title block. Confirm or change below.';
            msg.style.color = '#2E8B57';
        } else {
            select.value = '100';
            msg.textContent = 'No scale detected in PDF. Select the drawing scale below.';
            msg.style.color = '#555';
        }

        modal.classList.remove('hidden');

        function cleanup() {
            modal.classList.add('hidden');
            applyBtn.removeEventListener('click', onApply);
            cancelBtn.removeEventListener('click', onCancel);
        }

        function onApply() {
            const val = parseInt(select.value, 10);
            cleanup();
            resolve(val);
        }

        function onCancel() {
            cleanup();
            resolve(null);
        }

        applyBtn.addEventListener('click', onApply);
        cancelBtn.addEventListener('click', onCancel);
    });
}

// ── PDF: Load & Render ───────────────────────────────────

async function loadPDF(file) {
    if (typeof pdfjsLib === 'undefined') {
        alert('PDF.js library failed to load. Check your internet connection and reload.');
        return;
    }
    try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

        pdfState.pdfDoc = pdf;
        pdfState.totalPages = pdf.numPages;
        pdfState.currentPage = 1;
        pdfState.pageCanvases = {};

        // Populate page selector
        const pageSelect = document.getElementById('pdf-page-select');
        pageSelect.innerHTML = '';
        for (let i = 1; i <= pdf.numPages; i++) {
            const opt = document.createElement('option');
            opt.value = i;
            opt.textContent = 'Page ' + i;
            pageSelect.appendChild(opt);
        }
        pageSelect.value = 1;

        await renderPDFPage(1);

        pdfState.loaded = true;
        pdfState.visible = true;

        // Store per-level PDF data
        const activeLevel = typeof getActiveLevel === 'function' ? getActiveLevel() : null;
        if (activeLevel) {
            pdfState.levelPdfs[activeLevel.id] = {
                pdfDoc: pdf,
                pageNum: pdfState.currentPage,
                pageCanvas: pdfState.pageCanvases[pdfState.currentPage],
                sheetX: pdfState.sheetX,
                sheetY: pdfState.sheetY,
                sheetW: pdfState.sheetWidth,
                sheetH: pdfState.sheetHeight,
                nativeW: pdfState.nativeWidth,
                nativeH: pdfState.nativeHeight,
            };
        }

        // Show PDF controls
        showPdfControls(true);
        document.getElementById('btn-move-pdf').style.display = '';
        if (typeof updatePdfLevelIndicator === 'function') updatePdfLevelIndicator();
        engine.requestRender();

        console.log('[PDF] Loaded ' + pdf.numPages + ' page(s) for ' +
            (activeLevel ? activeLevel.name : 'current level'));

        // Auto-detect scale from PDF text and show confirmation modal
        const detectedScale = await detectPdfScale(pdf);
        const confirmedScale = await showScaleConfirmModal(detectedScale);

        if (confirmedScale && confirmedScale > 0) {
            CONFIG.drawingScale = confirmedScale;
            project.drawingScale = confirmedScale;
            project.projectInfo.scale = '1:' + confirmedScale;
            const scaleSelEl = document.getElementById('scale-select');
            if (scaleSelEl) scaleSelEl.value = String(confirmedScale);
            updateStatusBar();
            engine.requestRender();
            console.log('[PDF Scale] Set to 1:' + confirmedScale);
        }
    } catch (err) {
        alert('Error loading PDF: ' + err.message);
        console.error('[PDF]', err);
    }
}

async function renderPDFPage(pageNum) {
    // Check cache first
    if (pdfState.pageCanvases[pageNum]) {
        positionPDFFromCache(pageNum);
        return;
    }

    const page = await pdfState.pdfDoc.getPage(pageNum);
    const scale = pdfState.renderScale;
    const viewport = page.getViewport({ scale });

    // Get physical page dimensions at scale=1 (PDF points, 72pt = 1 inch = 25.4mm)
    const rawViewport = page.getViewport({ scale: 1 });
    const physicalWidthMM = rawViewport.width * 25.4 / 72;
    const physicalHeightMM = rawViewport.height * 25.4 / 72;

    const offscreen = document.createElement('canvas');
    offscreen.width = viewport.width;
    offscreen.height = viewport.height;
    const offCtx = offscreen.getContext('2d');

    await page.render({
        canvasContext: offCtx,
        viewport: viewport
    }).promise;

    // Cache it (include physical dimensions)
    pdfState.pageCanvases[pageNum] = {
        canvas: offscreen,
        nativeWidth: viewport.width,
        nativeHeight: viewport.height,
        physicalWidthMM: physicalWidthMM,
        physicalHeightMM: physicalHeightMM
    };

    console.log('[PDF] Page ' + pageNum + ' physical size: ' +
        physicalWidthMM.toFixed(1) + 'mm × ' + physicalHeightMM.toFixed(1) + 'mm');

    positionPDFFromCache(pageNum);
}

function positionPDFFromCache(pageNum) {
    const cached = pdfState.pageCanvases[pageNum];
    const c = CONFIG;

    pdfState.nativeWidth = cached.nativeWidth;
    pdfState.nativeHeight = cached.nativeHeight;

    const drawAreaW = c.SHEET_WIDTH_MM - c.MARGIN_LEFT - c.MARGIN_RIGHT;
    const drawAreaH = c.SHEET_HEIGHT_MM - c.MARGIN_TOP - c.MARGIN_BOTTOM - c.TITLE_BLOCK_HEIGHT_MM;

    // ── Physical scale positioning ────────────────────────
    // Display the PDF at its actual physical page size on the sheet.
    // An A1 PDF (841×594mm) fills the A1 sheet exactly.
    // An A3 PDF (420×297mm) fills approximately half the sheet.
    // This ensures PDF content at the same scale as the drawing
    // lines up 1:1 with the drawing coordinate system.
    if (cached.physicalWidthMM && cached.physicalHeightMM) {
        pdfState.sheetWidth = cached.physicalWidthMM;
        pdfState.sheetHeight = cached.physicalHeightMM;

        // If the PDF is larger than the sheet, scale it down to fit
        if (pdfState.sheetWidth > c.SHEET_WIDTH_MM || pdfState.sheetHeight > c.SHEET_HEIGHT_MM) {
            const scaleDown = Math.min(
                c.SHEET_WIDTH_MM / pdfState.sheetWidth,
                c.SHEET_HEIGHT_MM / pdfState.sheetHeight
            );
            pdfState.sheetWidth *= scaleDown;
            pdfState.sheetHeight *= scaleDown;
        }
    } else {
        // Fallback for cached pages without physical dims (legacy)
        const aspect = cached.nativeHeight / cached.nativeWidth;
        pdfState.sheetWidth = drawAreaW * 0.85;
        pdfState.sheetHeight = pdfState.sheetWidth * aspect;
        if (pdfState.sheetHeight > drawAreaH * 0.90) {
            pdfState.sheetHeight = drawAreaH * 0.90;
            pdfState.sheetWidth = pdfState.sheetHeight / aspect;
        }
    }

    // Position PDF on the sheet.
    // If the PDF is close to sheet size (A1), align to sheet origin (0,0)
    // so the PDF's title block, margins, etc. overlay the app's sheet exactly.
    // Otherwise centre in the drawing area.
    const isFullSheet = (cached.physicalWidthMM &&
        Math.abs(cached.physicalWidthMM - c.SHEET_WIDTH_MM) < 20 &&
        Math.abs(cached.physicalHeightMM - c.SHEET_HEIGHT_MM) < 20);

    if (isFullSheet) {
        // Align to sheet origin — PDF represents the full sheet
        pdfState.sheetX = 0;
        pdfState.sheetY = 0;
        console.log('[PDF] Full-sheet PDF detected, aligned to sheet origin');
    } else {
        // Centre in drawing area
        pdfState.sheetX = c.MARGIN_LEFT + (drawAreaW - pdfState.sheetWidth) / 2;
        pdfState.sheetY = c.MARGIN_TOP + (drawAreaH - pdfState.sheetHeight) / 2;
    }

    pdfState.currentPage = pageNum;
}

function showPdfControls(show) {
    const display = show ? '' : 'none';
    document.getElementById('btn-pdf-toggle').style.display = display;
    document.getElementById('btn-pdf-calibrate').style.display = display;
    document.getElementById('pdf-page-select').style.display =
        (show && pdfState.totalPages > 1) ? '' : 'none';
    document.getElementById('pdf-controls').style.display = display;
    document.getElementById('pdf-calib-group').style.display = display;
    const removeBtn = document.getElementById('btn-remove-pdf');
    if (removeBtn) removeBtn.style.display = display;
    if (show) {
        document.getElementById('btn-pdf-toggle').classList.add('active');
    }
    if (typeof updatePdfLevelIndicator === 'function') updatePdfLevelIndicator();
}

// ── PDF: Draw Underlay ───────────────────────────────────

function drawPdfUnderlay(ctx, eng) {
    if (!pdfState.loaded || !pdfState.visible) return;

    const cached = pdfState.pageCanvases[pdfState.currentPage];
    if (!cached) return;

    const coords = eng.coords;

    ctx.save();
    ctx.globalAlpha = pdfState.opacity;

    // Clip to drawing area
    const da = coords.drawArea;
    const clipTL = coords.sheetToScreen(da.left, da.top);
    const clipBR = coords.sheetToScreen(da.right, da.bottom);
    ctx.beginPath();
    ctx.rect(clipTL.x, clipTL.y, clipBR.x - clipTL.x, clipBR.y - clipTL.y);
    ctx.clip();

    // Draw the PDF image
    const tl = coords.sheetToScreen(pdfState.sheetX, pdfState.sheetY);
    const w = pdfState.sheetWidth * eng.viewport.zoom;
    const h = pdfState.sheetHeight * eng.viewport.zoom;
    ctx.drawImage(cached.canvas, tl.x, tl.y, w, h);

    ctx.restore();

    // Draw calibration markers (not affected by opacity)
    if (pdfState.calibrating && pdfState.calibPoints.length > 0) {
        drawCalibrationMarkers(ctx, eng);
    }
}

function drawCalibrationMarkers(ctx, eng) {
    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    for (let i = 0; i < pdfState.calibPoints.length; i++) {
        const pt = pdfState.calibPoints[i];
        const sp = coords.sheetToScreen(pt.x, pt.y);

        // Crosshair
        const r = 8;
        ctx.strokeStyle = '#FF3300';
        ctx.lineWidth = 2;
        ctx.beginPath();
        ctx.moveTo(sp.x - r, sp.y); ctx.lineTo(sp.x + r, sp.y);
        ctx.moveTo(sp.x, sp.y - r); ctx.lineTo(sp.x, sp.y + r);
        ctx.stroke();

        // Circle
        ctx.beginPath();
        ctx.arc(sp.x, sp.y, r, 0, Math.PI * 2);
        ctx.stroke();

        // Label
        ctx.font = 'bold 11px "Segoe UI", Arial, sans-serif';
        ctx.fillStyle = '#FF3300';
        ctx.textAlign = 'left';
        ctx.textBaseline = 'bottom';
        ctx.fillText('P' + (i + 1), sp.x + r + 3, sp.y - 3);
    }

    // Line between points
    if (pdfState.calibPoints.length === 2) {
        const p1 = coords.sheetToScreen(pdfState.calibPoints[0].x, pdfState.calibPoints[0].y);
        const p2 = coords.sheetToScreen(pdfState.calibPoints[1].x, pdfState.calibPoints[1].y);

        ctx.strokeStyle = '#FF3300';
        ctx.lineWidth = 1.5;
        ctx.setLineDash([6, 4]);
        ctx.beginPath();
        ctx.moveTo(p1.x, p1.y);
        ctx.lineTo(p2.x, p2.y);
        ctx.stroke();
        ctx.setLineDash([]);
    }

    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';
}

// Register the PDF render callback — runs BEFORE the main render callbacks
// We insert it at position 0 so it draws under everything else
engine._renderCallbacks.unshift(drawPdfUnderlay);

// ── PDF: Toolbar Event Handlers ──────────────────────────

// Import PDF button → trigger file picker
document.getElementById('btn-import-pdf').addEventListener('click', () => {
    document.getElementById('pdf-file-input').click();
});

// File input change
document.getElementById('pdf-file-input').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) loadPDF(file);
    e.target.value = ''; // allow re-importing same file
});

// Toggle visibility
document.getElementById('btn-pdf-toggle').addEventListener('click', () => {
    pdfState.visible = !pdfState.visible;
    const btn = document.getElementById('btn-pdf-toggle');
    if (pdfState.visible) {
        btn.classList.add('active');
        btn.querySelector('svg + span, svg ~ *') ||
            (btn.lastChild.textContent = ' Visible');
    } else {
        btn.classList.remove('active');
    }
    // Update button text
    const textNode = Array.from(btn.childNodes).find(n => n.nodeType === 3 && n.textContent.trim());
    if (textNode) textNode.textContent = pdfState.visible ? '\n                    Visible\n                ' : '\n                    Hidden\n                ';
    engine.requestRender();
});

// Page selector
document.getElementById('pdf-page-select').addEventListener('change', async (e) => {
    const pageNum = parseInt(e.target.value, 10);
    if (pageNum >= 1 && pageNum <= pdfState.totalPages) {
        await renderPDFPage(pageNum);
        engine.requestRender();
    }
});

// Opacity slider
const opacitySlider = document.getElementById('pdf-opacity');
const opacityVal = document.getElementById('pdf-opacity-val');
opacitySlider.addEventListener('input', () => {
    pdfState.opacity = parseInt(opacitySlider.value, 10) / 100;
    opacityVal.textContent = opacitySlider.value + '%';
    engine.requestRender();
});

// ── PDF: Scale Calibration ───────────────────────────────

const calibBanner = document.getElementById('calib-banner');
const calibModal = document.getElementById('calib-modal');

// Scale preset dropdown
const scalePreset = document.getElementById('scale-preset');
scalePreset.addEventListener('change', () => {
    const val = scalePreset.value;
    if (val === 'calibrate') {
        // Start calibration mode
        pdfState.calibrating = true;
        pdfState.calibPoints = [];
        calibBanner.classList.remove('hidden');
        calibBanner.innerHTML = '<strong>Set Scale:</strong> Click <strong>Point 1</strong> on a known dimension';
        engine.container.style.cursor = 'crosshair';
        engine.requestRender();
        scalePreset.value = ''; // reset dropdown
    } else {
        // Quick-set scale
        const scale = parseInt(val, 10);
        if (scale > 0) {
            CONFIG.drawingScale = scale;
            project.drawingScale = scale;
            project.projectInfo.scale = '1:' + scale;
            document.getElementById('scale-select').value = scale;
            engine.requestRender();
            updateStatusBar();
            scalePreset.value = ''; // reset dropdown
            console.log('[Scale] Set to 1:' + scale);
        }
    }
});

// Canvas click handler for calibration
engine.container.addEventListener('mousedown', (e) => {
    if (!pdfState.calibrating) return;
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;

    const sheetPos = engine.getSheetPos(e);

    if (pdfState.calibPoints.length === 0) {
        // First point
        pdfState.calibPoints.push({ x: sheetPos.x, y: sheetPos.y });
        calibBanner.innerHTML = '<strong>Set Scale:</strong> Click <strong>Point 2</strong> — end of known dimension';
        engine.requestRender();
    } else if (pdfState.calibPoints.length === 1) {
        // Second point — show distance modal
        pdfState.calibPoints.push({ x: sheetPos.x, y: sheetPos.y });
        engine.requestRender();

        calibBanner.classList.add('hidden');
        engine.container.style.cursor = '';
        calibModal.classList.remove('hidden');

        // Focus the input
        const distInput = document.getElementById('calib-distance');
        distInput.value = '';
        distInput.focus();
    }
});

// Calibration modal: Apply
document.getElementById('calib-apply').addEventListener('click', () => {
    applyCalibration();
});

// Calibration modal: Cancel
document.getElementById('calib-cancel').addEventListener('click', () => {
    cancelCalibration();
});

// Enter key in distance input
document.getElementById('calib-distance').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') applyCalibration();
    if (e.key === 'Escape') cancelCalibration();
});

function applyCalibration() {
    const distInput = document.getElementById('calib-distance');
    const realDistMM = parseFloat(distInput.value);

    if (!realDistMM || realDistMM <= 0) {
        distInput.style.borderColor = 'var(--danger)';
        return;
    }

    // Compute the sheet-mm distance between the two points
    const p1 = pdfState.calibPoints[0];
    const p2 = pdfState.calibPoints[1];
    const sheetDist = Math.sqrt(
        Math.pow(p2.x - p1.x, 2) + Math.pow(p2.y - p1.y, 2)
    );

    if (sheetDist < 0.1) {
        alert('Points are too close together. Try again with points further apart.');
        cancelCalibration();
        return;
    }

    // New scale = real distance / sheet distance
    const newScale = realDistMM / sheetDist;

    // Round to nearest standard scale if close
    const standardScales = [10, 20, 25, 50, 75, 100, 150, 200, 250, 500];
    let bestScale = newScale;
    for (const ss of standardScales) {
        if (Math.abs(newScale - ss) / ss < 0.08) {
            bestScale = ss;
            break;
        }
    }

    CONFIG.drawingScale = bestScale;
    project.drawingScale = bestScale;
    project.projectInfo.scale = '1:' + Math.round(bestScale);

    // Update the scale dropdown
    const scaleSelect = document.getElementById('scale-select');
    const rounded = Math.round(bestScale);
    // Check if this scale exists in the dropdown
    let found = false;
    for (const opt of scaleSelect.options) {
        if (parseInt(opt.value) === rounded) {
            scaleSelect.value = rounded;
            found = true;
            break;
        }
    }
    if (!found) {
        // Add a custom option
        const opt = document.createElement('option');
        opt.value = rounded;
        opt.textContent = '1:' + rounded;
        scaleSelect.appendChild(opt);
        scaleSelect.value = rounded;
    }

    // End calibration
    pdfState.calibrating = false;
    pdfState.calibPoints = [];
    calibModal.classList.add('hidden');

    engine.requestRender();
    updateStatusBar();

    console.log('[Calibrate] Sheet dist: ' + sheetDist.toFixed(2) +
        'mm, Real dist: ' + realDistMM + 'mm → Scale 1:' + Math.round(bestScale));
}

function cancelCalibration() {
    pdfState.calibrating = false;
    pdfState.calibPoints = [];
    calibBanner.classList.add('hidden');
    calibModal.classList.add('hidden');
    engine.container.style.cursor = '';
    engine.requestRender();
}

// Escape key cancels calibration
window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && pdfState.calibrating) {
        cancelCalibration();
    }
});

// ── PDF: Calibrate Button Popover ────────────────────────

const calibPopover = document.getElementById('pdf-calib-popover');

// Open calibration popover
document.getElementById('btn-pdf-calibrate').addEventListener('click', () => {
    calibPopover.classList.remove('hidden');
});

// Close calibration popover
document.getElementById('pdf-calib-popover-close').addEventListener('click', () => {
    calibPopover.classList.add('hidden');
});

// Quick-set scale buttons in popover
document.querySelectorAll('.pdf-calib-scale-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        const scale = parseInt(btn.dataset.scale, 10);
        if (scale > 0) {
            CONFIG.drawingScale = scale;
            project.drawingScale = scale;
            project.projectInfo.scale = '1:' + scale;
            const scaleSelEl = document.getElementById('scale-select');
            if (scaleSelEl) scaleSelEl.value = String(scale);
            updateStatusBar();
            engine.requestRender();
            calibPopover.classList.add('hidden');
            console.log('[Calibrate] Quick-set scale to 1:' + scale);
        }
    });

    // Hover highlight
    btn.addEventListener('mouseenter', () => {
        btn.style.background = '#d0e2f4';
        btn.style.borderColor = '#7baed4';
    });
    btn.addEventListener('mouseleave', () => {
        btn.style.background = '#f5f5f5';
        btn.style.borderColor = '#ccc';
    });
});

// Two-point calibrate button in popover → triggers existing calibration mode
document.getElementById('pdf-calib-2pt-btn').addEventListener('click', () => {
    calibPopover.classList.add('hidden');
    // Start calibration mode (same as existing calibrate flow)
    pdfState.calibrating = true;
    pdfState.calibPoints = [];
    calibBanner.classList.remove('hidden');
    calibBanner.innerHTML = '<strong>Calibrate:</strong> Click <strong>Point 1</strong> on a known dimension';
    engine.container.style.cursor = 'crosshair';
    engine.requestRender();
});

// ══════════════════════════════════════════════════════════
