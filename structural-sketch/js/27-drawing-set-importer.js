/* ═══════════════════════════════════════════════════════════════════════
   27-drawing-set-importer.js  —  Drawing Set Importer
   Parses a multi-page architectural PDF, identifies floor-plan pages,
   lets the user verify assignments, then loads each page into
   StructuralSketch as a separate level with cross-aligned grids.
   ═══════════════════════════════════════════════════════════════════════ */

/* ── Constants ───────────────────────────────────────────────────────── */

const DSI_VERSION = '1.0.0';

/** Drawing-number prefixes that indicate a floor-plan sheet (GAP = General
 *  Arrangement Plan).  RCP, elevations, sections, schedules, details are
 *  excluded.  Matching is case-insensitive. */
const DSI_PLAN_DRAWING_PREFIXES = ['A2', 'A-2', 'A-GA', 'AGA', 'A-FP'];
const DSI_PLAN_DRAWING_CODES = [
    'A20', 'A21',               // floor plans (A20xx, A21xx common)
];

/** Sheet-type keywords — used to classify drawing title text. */
const DSI_FLOOR_PLAN_KEYWORDS = [
    'floor plan', 'ground floor', 'level', 'basement',
    'mezzanine', 'podium', 'transfer', 'plant room',
    'roof plan', 'roof terrace',
    'general arrangement', 'gap',
];

const DSI_EXCLUDE_KEYWORDS = [
    'elevation', 'section', 'schedule', 'detail',
    'reflected ceiling', 'rcp', 'reflected',
    'electrical', 'hydraulic', 'mechanical',
    'door', 'window', 'finish', 'legend',
    'specification', 'cover sheet', 'index',
];

/** Map drawing titles to canonical level assignments. Ordered so the
 *  first match wins — more specific before more general. */
const DSI_LEVEL_PATTERNS = [
    { regex: /\bbasement\s*2\b/i,          level: 'B2',  name: 'Basement 2' },
    { regex: /\bbasement\s*1?\b/i,         level: 'B1',  name: 'Basement 1' },
    { regex: /\blower\s*ground\b/i,        level: 'LG',  name: 'Lower Ground' },
    { regex: /\bground\s*floor\b/i,        level: 'GF',  name: 'Ground Floor' },
    { regex: /\bground\b/i,               level: 'GF',  name: 'Ground Floor' },
    { regex: /\bmezzanine\b/i,            level: 'MZ',  name: 'Mezzanine' },
    { regex: /\bpodium\b/i,              level: 'PD',  name: 'Podium' },
    { regex: /\btransfer\b/i,            level: 'TF',  name: 'Transfer Level' },
    { regex: /\blevel\s*(\d+)/i,          level: null,  name: null }, // dynamic
    { regex: /\bfloor\s*(\d+)/i,          level: null,  name: null }, // dynamic
    { regex: /\broof\s*terrace\b/i,       level: 'RT',  name: 'Roof Terrace' },
    { regex: /\broof\s*plan\b/i,          level: 'RF',  name: 'Roof' },
    { regex: /\broof\b/i,                level: 'RF',  name: 'Roof' },
    { regex: /\bsite\s*plan\b/i,          level: 'SP',  name: 'Site Plan' },
];

/* ── State ───────────────────────────────────────────────────────────── */

const dsiState = {
    pdfDoc: null,           // pdf.js document (multi-page)
    pdfArrayBuffer: null,   // raw ArrayBuffer for re-loading pages
    pages: [],              // [{pageNum, drawingNo, drawingTitle, assignedLevel, assignedName, isFloorPlan}]
    confirmed: false,       // user has confirmed the table
    importing: false,       // import in progress
    referenceGrids: [],     // grids detected on the reference (lowest) level
};

/* ═══════════════════════════════════════════════════════════════════════
   STAGE 1 — Parse
   ═══════════════════════════════════════════════════════════════════════ */

/**
 * Opens the import dialog and starts the parse pipeline.
 * Called from the "Import Drawing Set" toolbar button.
 */
async function dsiStartImport(file) {
    if (!file || !file.name.toLowerCase().endsWith('.pdf')) {
        alert('Please select a PDF file.');
        return;
    }
    if (typeof pdfjsLib === 'undefined') {
        alert('PDF.js library not available. Check your internet connection.');
        return;
    }

    dsiShowModal('loading');

    try {
        const arrayBuffer = await file.arrayBuffer();
        dsiState.pdfArrayBuffer = arrayBuffer.slice(0); // keep a copy
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        dsiState.pdfDoc = pdf;

        const pages = [];
        for (let i = 1; i <= pdf.numPages; i++) {
            dsiUpdateLoadingProgress(i, pdf.numPages);
            const info = await dsiExtractPageInfo(pdf, i);
            pages.push(info);
        }

        dsiState.pages = pages;
        dsiState.confirmed = false;

        console.log('[DSI] Parsed ' + pdf.numPages + ' pages');
        console.table(pages.map(p => ({
            page: p.pageNum,
            drawingNo: p.drawingNo,
            title: p.drawingTitle,
            level: p.assignedLevel,
            plan: p.isFloorPlan
        })));

        // Move to Stage 2 — show verification table
        dsiShowVerifyTable();

    } catch (err) {
        dsiHideModal();
        alert('Error parsing PDF: ' + err.message);
        console.error('[DSI]', err);
    }
}

/**
 * Extracts drawing number, drawing title, and level assignment from a
 * single PDF page by reading text in the title-block region.
 */
async function dsiExtractPageInfo(pdf, pageNum) {
    const page = await pdf.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1 });
    const w = viewport.width;
    const h = viewport.height;

    // Strategy: try bottom-right title-block first, then bottom ribbon,
    // then fall back to full page text.
    let drawingNo = '';
    let drawingTitle = '';

    // ── Pass 1: structured text items with position ─────────────
    const textContent = await page.getTextContent();
    const items = textContent.items.filter(it => it.str && it.str.trim());

    // Title block region: bottom-right quadrant (x > 55% width, y > 70% height)
    // Note: PDF y=0 is at bottom, but text items use transformed coords
    // where y increases downward from the top of the viewport.
    const tbItems = items.filter(it => {
        const x = it.transform[4];
        const y = it.transform[5];
        // In PDF coordinate space, y=0 is at bottom of page
        // Title block is at bottom-right
        return x > w * 0.55 && y < h * 0.30;
    });

    // Also get bottom ribbon items (full width, bottom 10%)
    const ribbonItems = items.filter(it => {
        const y = it.transform[5];
        return y < h * 0.10;
    });

    // Merge title block + ribbon, dedup
    const candidateItems = [...tbItems, ...ribbonItems];
    const candidateTexts = candidateItems.map(it => it.str.trim());

    // ── Pass 2: find "Drawing No." / "Drawing Name" field labels ────
    // Look for the field labels and take the text that appears just
    // before or after them in the text stream.
    const allTexts = items.map(it => it.str.trim());
    const allFullText = allTexts.join('\n');

    // Strategy A: Find by field labels
    drawingNo = dsiFindFieldValue(candidateItems, items, 'Drawing No');
    drawingTitle = dsiFindFieldValue(candidateItems, items, 'Drawing Name');

    // Strategy B: If field-label approach failed, try pattern matching
    if (!drawingNo) {
        drawingNo = dsiFindDrawingNumber(candidateTexts) || dsiFindDrawingNumber(allTexts);
    }
    if (!drawingTitle) {
        drawingTitle = dsiFindDrawingTitle(candidateTexts, drawingNo);
    }

    // Strategy C: Full page fallback
    if (!drawingTitle) {
        drawingTitle = dsiFindDrawingTitle(allTexts, drawingNo);
    }

    // ── Pass 3: classify ────────────────────────────────────────
    const classification = dsiClassifyPage(drawingNo, drawingTitle);

    return {
        pageNum,
        drawingNo: drawingNo || '(unknown)',
        drawingTitle: drawingTitle || '(unknown)',
        assignedLevel: classification.level,
        assignedName: classification.name,
        isFloorPlan: classification.isFloorPlan,
    };
}

/**
 * Find a field value by locating the field label (e.g. "Drawing No.") in
 * the text items, then returning the text item that appears spatially
 * just above it (typical title block layout: value sits above its label).
 */
function dsiFindFieldValue(regionItems, allItems, fieldLabel) {
    const labelLower = fieldLabel.toLowerCase().replace(/[.\s]/g, '');

    // Find items matching the field label
    for (const item of allItems) {
        const text = item.str.trim().toLowerCase().replace(/[.\s]/g, '');
        if (text.includes(labelLower)) {
            const labelX = item.transform[4];
            const labelY = item.transform[5];
            const labelSize = Math.abs(item.transform[0]) || 8;

            // Look for the nearest text item ABOVE this label
            // (higher y in PDF coords = above)
            let best = null;
            let bestDist = Infinity;

            for (const other of allItems) {
                if (other === item) continue;
                const ox = other.transform[4];
                const oy = other.transform[5];
                const otherText = other.str.trim();

                // Must be above (higher y), within horizontal range, non-empty, not another label
                const dy = oy - labelY;
                const dx = Math.abs(ox - labelX);

                if (dy > 0 && dy < labelSize * 6 && dx < 200 && otherText.length > 0) {
                    const isLabel = /^(drawing|project|revision|scale|date|status|checked|drawn)/i.test(otherText);
                    if (!isLabel && dy < bestDist) {
                        bestDist = dy;
                        best = otherText;
                    }
                }
            }
            if (best) return best;
        }
    }
    return '';
}

/**
 * Find drawing number by pattern matching — looks for typical
 * architectural drawing number formats.
 */
function dsiFindDrawingNumber(texts) {
    // Common formats: A-GA-001, A2000, A-2000, AR-001, A0103
    const patterns = [
        /\b[A-Z]{1,3}[-.]?[A-Z]{0,2}[-.]?\d{2,4}\b/,          // A-GA-001, A2000
        /\b[A-Z]{1,2}\d{4}\b/,                                   // A2000
        /\b[A-Z]{1,2}[-]\d{2,4}\b/,                             // A-001
    ];

    for (const text of texts) {
        for (const pat of patterns) {
            const m = text.match(pat);
            if (m) return m[0];
        }
    }
    return '';
}

/**
 * Find drawing title — looks for text that describes the drawing content.
 * Excludes the drawing number, project info lines, and copyright text.
 */
function dsiFindDrawingTitle(texts, drawingNo) {
    const excludePatterns = [
        /^[A-Z]{1,3}[-.]?\d{2,4}$/,                            // drawing numbers
        /copyright|©|ABN|Pty\s*Ltd|www\.|\.com/i,               // legal
        /project\s*(name|number|address)/i,                      // project fields
        /drawing\s*(name|no|number|scale)/i,                     // field labels
        /revision|status|date|checked|drawn|scale/i,             // other labels
        /^\d{5}$/,                                               // project number
        /^\d+\s*[-–]\s*\d+\s/,                                  // addresses
        /^[A-Z][a-z]+$/,                                         // single city names (Brisbane, Sydney, etc.)
        /^\+?\d[\d\s]+$/,                                        // phone numbers
        /^@\s/,                                                   // scale annotations
    ];

    // Also exclude texts that match the drawing number
    const drawingNoClean = (drawingNo || '').trim();

    for (const text of texts) {
        const t = text.trim();
        if (t.length < 3 || t.length > 80) continue;
        if (drawingNoClean && t === drawingNoClean) continue;

        let excluded = false;
        for (const pat of excludePatterns) {
            if (pat.test(t)) { excluded = true; break; }
        }
        if (excluded) continue;

        // Check if this looks like a drawing title — should contain
        // at least one floor-plan or sheet-type keyword, or be a
        // substantial descriptive phrase
        const tLower = t.toLowerCase();
        const hasKeyword = DSI_FLOOR_PLAN_KEYWORDS.some(k => tLower.includes(k)) ||
                          DSI_EXCLUDE_KEYWORDS.some(k => tLower.includes(k)) ||
                          /plan|elevation|section|detail|sheet/i.test(t);

        if (hasKeyword) return t;
    }

    // Fallback — return the first substantial non-excluded text
    for (const text of texts) {
        const t = text.trim();
        if (t.length >= 5 && t.length <= 60) {
            let excluded = false;
            for (const pat of excludePatterns) {
                if (pat.test(t)) { excluded = true; break; }
            }
            if (!excluded && drawingNoClean !== t) return t;
        }
    }
    return '';
}

/**
 * Classify a page by its drawing number and title.
 * Returns { isFloorPlan, level, name }.
 */
function dsiClassifyPage(drawingNo, drawingTitle) {
    const titleLower = (drawingTitle || '').toLowerCase();
    const numUpper = (drawingNo || '').toUpperCase();

    // ── Check exclusions first ──────────────────────────────────
    for (const kw of DSI_EXCLUDE_KEYWORDS) {
        if (titleLower.includes(kw)) {
            return { isFloorPlan: false, level: '', name: '' };
        }
    }

    // ── Check drawing number prefix ─────────────────────────────
    // A2xxx range typically = floor plans, A23xx = RCP (excluded above)
    let likelyPlan = false;
    if (numUpper) {
        // A20xx, A21xx = floor plans; A23xx = RCP; A0xxx = site
        if (/^A-?2[01]\d{2}$/.test(numUpper)) likelyPlan = true;
        if (/^A-?GA/i.test(numUpper)) likelyPlan = true;
        if (/^A-?FP/i.test(numUpper)) likelyPlan = true;
    }

    // ── Check title keywords ────────────────────────────────────
    const hasPlanKeyword = DSI_FLOOR_PLAN_KEYWORDS.some(k => titleLower.includes(k));

    if (!likelyPlan && !hasPlanKeyword) {
        return { isFloorPlan: false, level: '', name: '' };
    }

    // ── Assign level ────────────────────────────────────────────
    for (const pat of DSI_LEVEL_PATTERNS) {
        const m = titleLower.match(pat.regex);
        if (m) {
            let levelId = pat.level;
            let levelName = pat.name;

            // Dynamic level (Level N / Floor N)
            if (!levelId && m[1]) {
                const n = parseInt(m[1]);
                levelId = 'L' + n;
                levelName = 'Level ' + n;
            }

            return { isFloorPlan: true, level: levelId || '', name: levelName || drawingTitle };
        }
    }

    // Fallback — it looks like a plan but we can't determine the level
    return { isFloorPlan: true, level: '', name: drawingTitle };
}


/* ═══════════════════════════════════════════════════════════════════════
   STAGE 2 — Verify
   ═══════════════════════════════════════════════════════════════════════ */

/**
 * Build and display the verification table modal.
 */
function dsiShowVerifyTable() {
    const pages = dsiState.pages;
    const floorPlans = pages.filter(p => p.isFloorPlan);
    const skipped = pages.filter(p => !p.isFloorPlan);

    let html = `
        <div class="dsi-verify-header">
            <h2>Drawing Set Import</h2>
            <p class="dsi-subtitle">
                Found <strong>${pages.length}</strong> pages —
                <strong>${floorPlans.length}</strong> identified as floor plans,
                ${skipped.length} skipped.
                Review and correct assignments below.
            </p>
        </div>
        <div class="dsi-table-wrap">
            <table class="dsi-table">
                <thead>
                    <tr>
                        <th style="width:40px">
                            <input type="checkbox" id="dsi-select-all" title="Select/deselect all" checked>
                        </th>
                        <th style="width:55px">Page</th>
                        <th style="width:100px">Drawing No.</th>
                        <th>Drawing Title</th>
                        <th style="width:180px">Assigned Level</th>
                    </tr>
                </thead>
                <tbody>`;

    for (const pg of pages) {
        const checked = pg.isFloorPlan ? 'checked' : '';
        const rowClass = pg.isFloorPlan ? '' : 'dsi-row-skipped';
        html += `
                    <tr class="${rowClass}" data-page="${pg.pageNum}">
                        <td><input type="checkbox" class="dsi-page-check" data-page="${pg.pageNum}" ${checked}></td>
                        <td class="dsi-cell-center">${pg.pageNum}</td>
                        <td>${dsiEscHtml(pg.drawingNo)}</td>
                        <td>${dsiEscHtml(pg.drawingTitle)}</td>
                        <td>
                            <select class="dsi-level-select" data-page="${pg.pageNum}">
                                <option value="">(skip)</option>
                                <option value="B2" ${pg.assignedLevel === 'B2' ? 'selected' : ''}>Basement 2</option>
                                <option value="B1" ${pg.assignedLevel === 'B1' ? 'selected' : ''}>Basement 1</option>
                                <option value="LG" ${pg.assignedLevel === 'LG' ? 'selected' : ''}>Lower Ground</option>
                                <option value="GF" ${pg.assignedLevel === 'GF' ? 'selected' : ''}>Ground Floor</option>
                                <option value="MZ" ${pg.assignedLevel === 'MZ' ? 'selected' : ''}>Mezzanine</option>
                                <option value="PD" ${pg.assignedLevel === 'PD' ? 'selected' : ''}>Podium</option>
                                <option value="L1" ${pg.assignedLevel === 'L1' ? 'selected' : ''}>Level 1</option>
                                <option value="L2" ${pg.assignedLevel === 'L2' ? 'selected' : ''}>Level 2</option>
                                <option value="L3" ${pg.assignedLevel === 'L3' ? 'selected' : ''}>Level 3</option>
                                <option value="L4" ${pg.assignedLevel === 'L4' ? 'selected' : ''}>Level 4</option>
                                <option value="L5" ${pg.assignedLevel === 'L5' ? 'selected' : ''}>Level 5</option>
                                <option value="L6" ${pg.assignedLevel === 'L6' ? 'selected' : ''}>Level 6</option>
                                <option value="L7" ${pg.assignedLevel === 'L7' ? 'selected' : ''}>Level 7</option>
                                <option value="L8" ${pg.assignedLevel === 'L8' ? 'selected' : ''}>Level 8</option>
                                <option value="L9" ${pg.assignedLevel === 'L9' ? 'selected' : ''}>Level 9</option>
                                <option value="L10" ${pg.assignedLevel === 'L10' ? 'selected' : ''}>Level 10</option>
                                <option value="TF" ${pg.assignedLevel === 'TF' ? 'selected' : ''}>Transfer Level</option>
                                <option value="RT" ${pg.assignedLevel === 'RT' ? 'selected' : ''}>Roof Terrace</option>
                                <option value="RF" ${pg.assignedLevel === 'RF' ? 'selected' : ''}>Roof</option>
                                <option value="SP" ${pg.assignedLevel === 'SP' ? 'selected' : ''}>Site Plan</option>
                            </select>
                        </td>
                    </tr>`;
    }

    html += `
                </tbody>
            </table>
        </div>
        <div class="dsi-actions">
            <button class="dsi-btn dsi-btn-secondary" id="dsi-btn-cancel">Cancel</button>
            <button class="dsi-btn dsi-btn-primary" id="dsi-btn-proceed">
                Import ${floorPlans.length} Floor Plan${floorPlans.length !== 1 ? 's' : ''}
            </button>
        </div>`;

    dsiShowModal('verify', html);

    // Wire up interactions
    document.getElementById('dsi-select-all').addEventListener('change', (e) => {
        const checks = document.querySelectorAll('.dsi-page-check');
        checks.forEach(c => { c.checked = e.target.checked; });
        dsiUpdateProceedButton();
    });

    document.querySelectorAll('.dsi-page-check').forEach(c => {
        c.addEventListener('change', () => dsiUpdateProceedButton());
    });

    document.querySelectorAll('.dsi-level-select').forEach(sel => {
        sel.addEventListener('change', (e) => {
            const pageNum = parseInt(e.target.dataset.page);
            const pg = dsiState.pages.find(p => p.pageNum === pageNum);
            if (pg) {
                const val = e.target.value;
                pg.assignedLevel = val;
                pg.isFloorPlan = val !== '';
                // Update the checkbox to match
                const check = document.querySelector(`.dsi-page-check[data-page="${pageNum}"]`);
                if (check) check.checked = val !== '';
            }
            dsiUpdateProceedButton();
        });
    });

    document.getElementById('dsi-btn-cancel').addEventListener('click', () => {
        dsiHideModal();
        dsiResetState();
    });

    document.getElementById('dsi-btn-proceed').addEventListener('click', () => {
        dsiProceedToImport();
    });
}

function dsiUpdateProceedButton() {
    const checks = document.querySelectorAll('.dsi-page-check:checked');
    const count = checks.length;
    const btn = document.getElementById('dsi-btn-proceed');
    if (btn) {
        btn.textContent = `Import ${count} Floor Plan${count !== 1 ? 's' : ''}`;
        btn.disabled = count === 0;
    }
}


/* ═══════════════════════════════════════════════════════════════════════
   STAGE 3 — Automate (Import + Grid Detection + Alignment)
   ═══════════════════════════════════════════════════════════════════════ */

/**
 * Gather the checked pages, create/match levels, load PDFs, detect grids.
 */
async function dsiProceedToImport() {
    // Collect selected pages with valid level assignments
    const checks = document.querySelectorAll('.dsi-page-check:checked');
    const selectedPages = [];
    checks.forEach(c => {
        const pageNum = parseInt(c.dataset.page);
        const pg = dsiState.pages.find(p => p.pageNum === pageNum);
        const levelSelect = document.querySelector(`.dsi-level-select[data-page="${pageNum}"]`);
        const levelId = levelSelect ? levelSelect.value : (pg ? pg.assignedLevel : '');
        if (pg && levelId) {
            // Get the display name from the select option
            const opt = levelSelect ? levelSelect.options[levelSelect.selectedIndex] : null;
            selectedPages.push({
                pageNum: pg.pageNum,
                levelId: levelId,
                levelName: opt ? opt.textContent : pg.assignedName,
                drawingNo: pg.drawingNo,
                drawingTitle: pg.drawingTitle,
            });
        }
    });

    if (selectedPages.length === 0) {
        alert('No pages selected with valid level assignments.');
        return;
    }

    // Sort by natural level order (basement → ground → level 1, 2... → roof)
    selectedPages.sort((a, b) => dsiLevelSortOrder(a.levelId) - dsiLevelSortOrder(b.levelId));

    console.log('[DSI] Importing ' + selectedPages.length + ' pages:', selectedPages);

    dsiShowModal('importing');
    dsiState.importing = true;

    try {
        // ── Step 1: Set up levels ───────────────────────────────
        dsiUpdateImportProgress('Setting up levels...', 0, selectedPages.length);

        // Clear existing levels and build fresh ones from the drawing set
        const newLevels = [];
        let elevation = 0;
        const defaultHeight = 2700; // mm

        for (let i = 0; i < selectedPages.length; i++) {
            const sp = selectedPages[i];
            const isRoof = sp.levelId === 'RF' || sp.levelId === 'RT';
            newLevels.push({
                id: sp.levelId,
                name: sp.levelName,
                height: isRoof ? 0 : defaultHeight,
                elevation: elevation,
            });
            if (!isRoof) elevation += defaultHeight;
        }

        // Replace level system (preserve groundRL)
        const savedGroundRL = levelSystem.groundRL;
        levelSystem.levels = newLevels;
        levelSystem.groundRL = savedGroundRL;
        levelSystem.activeLevelIndex = 0;

        if (typeof recalcElevations === 'function') recalcElevations();
        if (typeof buildLevelTabs === 'function') buildLevelTabs();

        // ── Step 2: Load PDF for each level ─────────────────────
        // We need to load the same PDF doc but show different pages per level.
        // Re-use the already-loaded pdf.js document.
        const pdf = dsiState.pdfDoc;

        // Detect scale from the PDF (use the first floor plan page)
        const detectedScale = await detectPdfScale(pdf);
        let confirmedScale = detectedScale || CONFIG.drawingScale || 100;

        for (let i = 0; i < selectedPages.length; i++) {
            const sp = selectedPages[i];
            dsiUpdateImportProgress(
                `Loading ${sp.levelName} (page ${sp.pageNum})...`,
                i, selectedPages.length
            );

            // Switch to this level
            const lvIndex = levelSystem.levels.findIndex(l => l.id === sp.levelId);
            if (lvIndex === -1) continue;

            // Direct level switch (avoid the hook that saves/restores PDFs
            // since we're building fresh)
            levelSystem.activeLevelIndex = lvIndex;

            // Set up the PDF state for this page
            pdfState.pdfDoc = pdf;
            pdfState.totalPages = pdf.numPages;
            pdfState.currentPage = sp.pageNum;

            // Render this specific page
            await renderPDFPage(sp.pageNum);

            pdfState.loaded = true;
            pdfState.visible = true;

            // Store per-level PDF data
            const cached = pdfState.pageCanvases[sp.pageNum];
            pdfState.levelPdfs[sp.levelId] = {
                pdfDoc: pdf,
                pageNum: sp.pageNum,
                pageCanvas: cached,
                sheetX: pdfState.sheetX,
                sheetY: pdfState.sheetY,
                sheetW: pdfState.sheetWidth,
                sheetH: pdfState.sheetHeight,
                nativeW: pdfState.nativeWidth,
                nativeH: pdfState.nativeHeight,
            };

            console.log('[DSI] Loaded page ' + sp.pageNum + ' for ' + sp.levelName);
        }

        // ── Step 3: Set confirmed scale ─────────────────────────
        if (confirmedScale && confirmedScale > 0) {
            CONFIG.drawingScale = confirmedScale;
            project.drawingScale = confirmedScale;
            project.projectInfo.scale = '1:' + confirmedScale;
            const scaleSelEl = document.getElementById('scale-select');
            if (scaleSelEl) scaleSelEl.value = String(confirmedScale);
        }

        // ── Step 4: Detect grids on the reference level ─────────
        dsiUpdateImportProgress('Detecting grid lines...', selectedPages.length - 1, selectedPages.length);

        // Use the lowest floor plan as reference (first non-site-plan)
        const refPage = selectedPages.find(sp => sp.levelId !== 'SP') || selectedPages[0];
        const refLvIndex = levelSystem.levels.findIndex(l => l.id === refPage.levelId);

        if (refLvIndex !== -1) {
            levelSystem.activeLevelIndex = refLvIndex;

            // Restore PDF for the reference level
            const refPdf = pdfState.levelPdfs[refPage.levelId];
            if (refPdf && refPdf.pageCanvas) {
                pdfState.loaded = true;
                pdfState.pdfDoc = refPdf.pdfDoc;
                pdfState.currentPage = refPdf.pageNum;
                pdfState.pageCanvases[refPdf.pageNum] = refPdf.pageCanvas;
                pdfState.sheetX = refPdf.sheetX;
                pdfState.sheetY = refPdf.sheetY;
                pdfState.sheetWidth = refPdf.sheetW;
                pdfState.sheetHeight = refPdf.sheetH;
                pdfState.nativeWidth = refPdf.nativeW;
                pdfState.nativeHeight = refPdf.nativeH;
            }

            // Run text-based grid detection on the reference page
            const grids = await dsiDetectGridsFromText(pdf, refPage.pageNum);

            if (grids.length > 0) {
                // Clear existing grids and add detected ones
                // All auto-detected grids are type:'angled' (finite extent)
                structuralGrids.length = 0;
                for (const g of grids) {
                    structuralGrids.push({
                        id: generateId(),
                        type: 'angled',
                        x1: g.x1,
                        y1: g.y1,
                        x2: g.x2,
                        y2: g.y2,
                        angle: g.angle,
                        label: g.label,
                        zone: g.zone || 'main',
                    });
                }

                // Update grid label state based on detected labels
                // Find the last letter and last number across all grids
                const letterGrids = grids.filter(g => /^[A-Z]$/.test(g.label));
                const numberGrids = grids.filter(g => /^\d+$/.test(g.label));

                if (letterGrids.length > 0) {
                    const lastLetter = letterGrids
                        .map(g => g.label)
                        .sort()
                        .pop();
                    gridLabelState.V.scheme = 'alpha';
                    gridLabelState.V.nextAlpha = lastLetter.charCodeAt(0) - 64 + 1;
                }
                if (numberGrids.length > 0) {
                    const lastNum = numberGrids
                        .map(g => parseInt(g.label))
                        .sort((a, b) => a - b)
                        .pop();
                    gridLabelState.H.scheme = 'num';
                    gridLabelState.H.nextNum = lastNum + 1;
                }

                // Log summary per zone
                const zoneCounts = {};
                for (const g of grids) {
                    const z = g.zone || 'main';
                    zoneCounts[z] = (zoneCounts[z] || 0) + 1;
                }
                console.log('[DSI] Grid summary: ' + grids.length + ' total — ' +
                    Object.entries(zoneCounts).map(([z, c]) => z + ':' + c).join(', '));

                dsiState.referenceGrids = grids;
                console.log('[DSI] Detected ' + grids.length + ' grid lines from reference page');
            } else {
                console.log('[DSI] No grids detected from text — user can add manually');
            }
        }

        // ── Step 5: Switch to first level and finish ────────────
        switchToLevel(0);
        if (typeof updatePdfLevelIndicator === 'function') updatePdfLevelIndicator();
        showPdfControls(true);
        if (typeof updateStatusBar === 'function') updateStatusBar();
        engine.requestRender();

        // Show completion
        dsiShowModal('complete', `
            <div class="dsi-complete">
                <div class="dsi-complete-icon">✓</div>
                <h2>Import Complete</h2>
                <p>Successfully imported <strong>${selectedPages.length}</strong> floor plan${selectedPages.length !== 1 ? 's' : ''}
                   with <strong>${structuralGrids.length}</strong> grid lines detected.</p>
                <div class="dsi-complete-summary">
                    ${selectedPages.map(sp => `<div class="dsi-complete-row">
                        <span class="dsi-complete-level">${sp.levelName}</span>
                        <span class="dsi-complete-dwg">${sp.drawingNo}</span>
                    </div>`).join('')}
                </div>
                <button class="dsi-btn dsi-btn-primary" id="dsi-btn-done" style="margin-top:16px;">
                    Start Working
                </button>
                <p class="dsi-hint" style="margin-top:10px;">
                    Tip: Use Page Up / Page Down to switch between levels.
                    ${structuralGrids.length === 0 ? 'No grids were auto-detected — use the Grid tool to add them manually.' : ''}
                </p>
            </div>
        `);

        document.getElementById('dsi-btn-done').addEventListener('click', () => {
            dsiHideModal();
            dsiResetState();
            // Show the scale confirmation if needed
            if (confirmedScale) {
                showScaleConfirmModal(confirmedScale);
            }
        });

    } catch (err) {
        dsiHideModal();
        alert('Import failed: ' + err.message);
        console.error('[DSI Import]', err);
    } finally {
        dsiState.importing = false;
    }
}

/**
 * Detect structural grids by finding grid-bubble text labels in the PDF.
 * Supports both orthogonal (H/V) and angled grids with finite extent.
 *
 * Algorithm:
 *  1. Collect single-char (A-Z) and small-number (1-99) text at the
 *     dominant grid-bubble font size.
 *  2. Group by label text; compute pairwise angles between instances of
 *     the same label to discover grid directions.
 *  3. Cluster the angles to identify grid "zones" — each zone is a
 *     family of parallel grid lines at a common angle.
 *  4. For each zone, emit either type:'ortho' (if angle ≈ 0° or 90°)
 *     or type:'angled' (with finite start/end endpoints).
 */
async function dsiDetectGridsFromText(pdf, pageNum) {
    const page = await pdf.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1 });
    const textContent = await page.getTextContent();
    const items = textContent.items;

    const w = viewport.width;
    const h = viewport.height;

    // ── 1. Collect grid-label candidates ────────────────────────
    const candidates = [];
    const fontSizes = [];

    for (const item of items) {
        const text = item.str.trim();
        if (!text) continue;

        const isGridLabel = /^[A-Z]$/.test(text) || /^[1-9]\d?$/.test(text);
        if (!isGridLabel) continue;

        const x = item.transform[4];
        const y = item.transform[5];
        const fontSize = Math.abs(item.transform[0]) || Math.abs(item.transform[3]) || 8;

        candidates.push({ text, x, y, fontSize });
        fontSizes.push(Math.round(fontSize * 10) / 10);
    }

    if (candidates.length === 0) return [];

    // ── 2. Find dominant grid-label font size ───────────────────
    const sizeCount = {};
    for (const s of fontSizes) {
        sizeCount[s] = (sizeCount[s] || 0) + 1;
    }
    const sizesWithCounts = Object.entries(sizeCount)
        .map(([s, c]) => ({ size: parseFloat(s), count: c }))
        .sort((a, b) => b.size - a.size);

    let gridFontSize = null;
    for (const sc of sizesWithCounts) {
        if (sc.count >= 4 && sc.size >= 6) {
            gridFontSize = sc.size;
            break;
        }
    }
    if (!gridFontSize) return [];

    const tolerance = gridFontSize * 0.15;
    const filtered = candidates.filter(c =>
        Math.abs(c.fontSize - gridFontSize) <= tolerance
    );

    // ── 3. Group by label text ──────────────────────────────────
    const byText = {};
    for (const c of filtered) {
        if (!byText[c.text]) byText[c.text] = [];
        byText[c.text].push(c);
    }

    // ── 4. Compute pairwise angles for each label ───────────────
    // For each label that appears 2+ times, compute angles between
    // all pairs of its instances. These angles tell us the direction
    // of the grid line that connects two bubbles of the same label.
    const angleSamples = [];  // { label, angle (0-180°), pair: [{x,y},{x,y}] }

    for (const [label, pts] of Object.entries(byText)) {
        if (pts.length < 2) continue;
        for (let i = 0; i < pts.length; i++) {
            for (let j = i + 1; j < pts.length; j++) {
                const dx = pts[j].x - pts[i].x;
                const dy = pts[j].y - pts[i].y;
                const dist = Math.sqrt(dx * dx + dy * dy);
                // Ignore pairs that are very close (same bubble, duplicate text)
                if (dist < w * 0.05) continue;
                // Angle in degrees 0-180 (direction, not signed)
                let angle = Math.atan2(-dy, dx) * 180 / Math.PI; // -dy because PDF y-up
                if (angle < 0) angle += 180;
                if (angle >= 180) angle -= 180;
                angleSamples.push({
                    label,
                    angle,
                    p1: { x: pts[i].x, y: pts[i].y },
                    p2: { x: pts[j].x, y: pts[j].y },
                    dist
                });
            }
        }
    }

    if (angleSamples.length === 0) return [];

    // ── 5. Cluster angles ───────────────────────────────────────
    // Group angle samples into clusters where all angles are within
    // ±ANGLE_TOL of the cluster mean. This identifies distinct grid
    // directions (e.g. 0° horizontal, 90° vertical, 67° angled wing).
    const ANGLE_TOL = 8; // degrees tolerance for same direction
    const angleClusters = dsiClusterAngles(angleSamples, ANGLE_TOL);

    console.log('[DSI] Angle clusters:', angleClusters.map(c =>
        c.meanAngle.toFixed(1) + '° (' + c.samples.length + ' samples)'
    ).join(', '));

    // ── 5b. Deduplicate: a label can legitimately appear in both an
    //    ortho and angled cluster (e.g. grid "A" in both main body
    //    and wing). But WITHIN the same type (both ortho or both
    //    angled), keep a label only in the cluster with more samples
    //    to prevent spurious cross-zone matches. ──────────────────
    for (const cluster of angleClusters) {
        const labelCounts = {};
        for (const s of cluster.samples) {
            labelCounts[s.label] = (labelCounts[s.label] || 0) + 1;
        }
        cluster._labelCounts = labelCounts;
        cluster._isOrtho = dsiIsOrthoAngle(cluster.meanAngle);
    }
    // For each label, find the best cluster per type (ortho / angled)
    const labelBestOrtho = {};   // label → { clusterIdx, count }
    const labelBestAngled = {};
    for (let ci = 0; ci < angleClusters.length; ci++) {
        const map = angleClusters[ci]._isOrtho ? labelBestOrtho : labelBestAngled;
        for (const [label, count] of Object.entries(angleClusters[ci]._labelCounts)) {
            if (!map[label] || count > map[label].count) {
                map[label] = { clusterIdx: ci, count };
            }
        }
    }
    // Remove samples from non-best clusters of the SAME type
    for (let ci = 0; ci < angleClusters.length; ci++) {
        const map = angleClusters[ci]._isOrtho ? labelBestOrtho : labelBestAngled;
        angleClusters[ci].samples = angleClusters[ci].samples.filter(s => {
            return map[s.label] && map[s.label].clusterIdx === ci;
        });
    }
    // Remove empty clusters
    const prunedClusters = angleClusters.filter(c => c.samples.length > 0);

    console.log('[DSI] After dedup:', prunedClusters.map(c => {
        const labels = [...new Set(c.samples.map(s => s.label))].sort();
        return c.meanAngle.toFixed(1) + '° → ' + labels.join(',');
    }).join(' | '));

    // ── 6. For each angle cluster, create a zone and build grid lines ──
    //    ALL auto-detected grids are type:'angled' (finite extent) so
    //    they match the architect's grid line extent, not the full page.
    const grids = [];
    let angledZoneIdx = 0;

    // Clear existing zones and ensure 'Main' exists for ortho grids
    gridZones.length = 0;
    ensureMainZone();

    for (const cluster of prunedClusters) {
        const meanAngle = cluster.meanAngle;
        const isOrtho = dsiIsOrthoAngle(meanAngle);

        // Create zone for this cluster
        let zoneId;
        if (isOrtho) {
            zoneId = 'main';
        } else {
            angledZoneIdx++;
            const zoneName = 'Zone ' + angledZoneIdx + ' (' + meanAngle.toFixed(0) + '°)';
            const zone = createGridZone(zoneName, meanAngle);
            zoneId = zone.id;
        }

        // Gather unique labels in this cluster
        const clusterLabels = {};
        for (const s of cluster.samples) {
            if (!clusterLabels[s.label]) clusterLabels[s.label] = [];
            clusterLabels[s.label].push(s);
        }

        // For each label, determine endpoints from the best pair
        const labelTexts = Object.keys(clusterLabels).sort((a, b) => {
            const aNum = parseInt(a);
            const bNum = parseInt(b);
            if (!isNaN(aNum) && !isNaN(bNum)) return aNum - bNum;
            return a.localeCompare(b);
        });

        // Find sequential subset
        const sequential = dsiFindSequentialLabels(labelTexts);
        if (sequential.length < 2) continue;

        // Build grids — ALL as type:'angled' with finite endpoints
        const clusterGrids = [];
        for (const labelText of sequential) {
            const samples = clusterLabels[labelText];
            if (!samples || samples.length === 0) continue;

            // Pick the pair with the longest distance
            let best = samples[0];
            for (const s of samples) {
                if (s.dist > best.dist) best = s;
            }
            const g = dsiPdfPairToAngledGrid(
                best.p1.x, best.p1.y,
                best.p2.x, best.p2.y,
                w, h, labelText, zoneId
            );
            if (g) clusterGrids.push(g);
        }

        // ── 6b. Orphan label recovery ──────────────────────────
        // Some labels may be sequential but missing from this cluster
        // because they only had 1 instance in the zone (no pair).
        // Use the cluster's established direction and typical length
        // to project a grid line through the orphan's known position.
        if (clusterGrids.length >= 3) {
            // Compute the cluster's direction vector (unit) in PDF coords
            const refAngleRad = meanAngle * Math.PI / 180;
            // direction in PDF.js space: dx = cos(a), dy = -sin(a) (y-up)
            const udx = Math.cos(refAngleRad);
            const udy = -Math.sin(refAngleRad);

            // Compute typical grid line half-length from existing grids
            const lengths = clusterGrids.map(g => {
                const dx = g.x2 - g.x1, dy = g.y2 - g.y1;
                return Math.sqrt(dx * dx + dy * dy);
            });
            const avgLen = lengths.reduce((s, l) => s + l, 0) / lengths.length;
            // Use half-length for projection from centre point
            const halfLen = avgLen / 2;

            // Build full sequential range to find orphans.
            // Only consider labels of the SAME TYPE as the cluster's
            // existing labels (letters or numbers) to avoid cross-
            // contamination (e.g. recovering letter grids into a
            // number-based cluster).
            const clusterIsNumeric = sequential.every(l => /^\d+$/.test(l));
            const sameTypeLabels = Object.keys(byText)
                .filter(l => clusterIsNumeric ? /^\d+$/.test(l) : /^[A-Z]$/.test(l))
                .sort((a, b) => {
                    const aN = parseInt(a), bN = parseInt(b);
                    if (!isNaN(aN) && !isNaN(bN)) return aN - bN;
                    return a.localeCompare(b);
                });
            const expandedSeq = dsiFindSequentialLabels(sameTypeLabels);

            const existingLabels = new Set(clusterGrids.map(g => g.label));

            for (const labelText of expandedSeq) {
                if (existingLabels.has(labelText)) continue;

                // Find instances of this label
                const pts = byText[labelText] || [];
                if (pts.length === 0) continue;

                // Pick the instance that lies closest to the cluster's
                // spatial extent — i.e., the instance most likely in the
                // same zone. Use perpendicular distance to the line
                // through any existing grid's midpoint at the cluster angle.
                let bestPt = null, bestOrphanDist = Infinity;
                const refGrid = clusterGrids[0];
                // Reference line midpoint in drawing-mm
                const rmx = (refGrid.x1 + refGrid.x2) / 2;
                const rmy = (refGrid.y1 + refGrid.y2) / 2;

                for (const pt of pts) {
                    // Convert pt to drawing-mm for comparison
                    const nX = pt.x / w;
                    const nY = 1 - (pt.y / h);
                    const sX = pdfState.sheetX + nX * pdfState.sheetWidth;
                    const sY = pdfState.sheetY + nY * pdfState.sheetHeight;
                    const da = engine.coords.drawArea;
                    const dmx = (sX - da.left) * CONFIG.drawingScale;
                    const dmy = (sY - da.top) * CONFIG.drawingScale;

                    // Perpendicular distance from this point to the
                    // cluster direction through the ref midpoint
                    // (measures how far off-axis this instance is)
                    const dx = dmx - rmx, dy = dmy - rmy;
                    // Cross product with unit direction gives perp dist
                    const angRad = meanAngle * Math.PI / 180;
                    const ux = Math.cos(angRad), uy = -Math.sin(angRad);
                    const perpDist = Math.abs(dx * uy - dy * ux);

                    if (perpDist < bestOrphanDist) {
                        bestOrphanDist = perpDist;
                        bestPt = pt;
                    }
                }

                if (!bestPt) continue;

                // Only recover if the orphan instance is reasonably
                // close to the cluster in BOTH dimensions:
                // 1) Perpendicular distance (across grids) < 3× avg spacing
                // 2) Along-axis distance: orphan must project within or
                //    near the cluster's existing spatial extent

                // Compute average spacing between consecutive grids
                const spacings = [];
                for (let gi = 1; gi < clusterGrids.length; gi++) {
                    const g0 = clusterGrids[gi - 1];
                    const g1 = clusterGrids[gi];
                    const mx0 = (g0.x1 + g0.x2) / 2, my0 = (g0.y1 + g0.y2) / 2;
                    const mx1 = (g1.x1 + g1.x2) / 2, my1 = (g1.y1 + g1.y2) / 2;
                    spacings.push(Math.sqrt((mx1-mx0)**2 + (my1-my0)**2));
                }
                const avgSpacing = spacings.length > 0
                    ? spacings.reduce((s,v) => s+v, 0) / spacings.length
                    : avgLen;
                // Reject if perpendicular distance exceeds 3× avg spacing
                if (bestOrphanDist > avgSpacing * 3) continue;

                // Along-axis check: project orphan onto the line
                // perpendicular to the cluster direction (i.e. the
                // direction grids are spread along). Compute how far
                // this projection is from the cluster's existing range.
                const da2 = engine.coords.drawArea;
                const nXo = bestPt.x / w;
                const nYo = 1 - (bestPt.y / h);
                const sXo = pdfState.sheetX + nXo * pdfState.sheetWidth;
                const sYo = pdfState.sheetY + nYo * pdfState.sheetHeight;
                const dmxO = (sXo - da2.left) * CONFIG.drawingScale;
                const dmyO = (sYo - da2.top) * CONFIG.drawingScale;

                // Compute dot product with perpendicular direction
                // (the direction grids are spread along)
                const angRad2 = meanAngle * Math.PI / 180;
                // Perpendicular to grid direction = spread direction
                const perpDx = Math.sin(angRad2);
                const perpDy = Math.cos(angRad2);

                // Project all existing grid midpoints onto spread axis
                const projections = clusterGrids.map(g => {
                    const mx = (g.x1 + g.x2) / 2;
                    const my = (g.y1 + g.y2) / 2;
                    return mx * perpDx + my * perpDy;
                });
                const orphanProj = dmxO * perpDx + dmyO * perpDy;
                const minProj = Math.min(...projections);
                const maxProj = Math.max(...projections);
                const extent = maxProj - minProj || avgSpacing;
                // Reject if orphan projects more than 1.5× extent beyond
                // the cluster's existing range
                const margin = extent * 0.5 + avgSpacing * 2;
                if (orphanProj < minProj - margin || orphanProj > maxProj + margin) continue;

                // Project a grid line through this instance at the
                // cluster angle, with the typical half-length
                // Convert the half-length from drawing-mm to PDF coords
                const halfLenPdf = (halfLen / CONFIG.drawingScale) /
                    (pdfState.sheetWidth / w);  // approximate scale

                const cx = bestPt.x;
                const cy = bestPt.y;
                const px1 = cx - udx * halfLenPdf;
                const py1 = cy - udy * halfLenPdf;
                const px2 = cx + udx * halfLenPdf;
                const py2 = cy + udy * halfLenPdf;

                const g = dsiPdfPairToAngledGrid(
                    px1, py1, px2, py2, w, h, labelText, zoneId
                );
                if (g) {
                    clusterGrids.push(g);
                    console.log('[DSI] Recovered orphan grid ' + labelText +
                        ' in cluster ' + meanAngle.toFixed(1) + '°');
                }
            }
        }

        grids.push(...clusterGrids);
    }

    return grids;
}

/**
 * Convert a pair of PDF-coordinate points into a type:'angled' grid
 * object in sheet-mm coordinates.
 */
function dsiPdfPairToAngledGrid(px1, py1, px2, py2, pageW, pageH, label, zoneId) {
    // PDF coords → normalised (0..1) → sheet-mm
    const normX1 = px1 / pageW;
    const normY1 = 1 - (py1 / pageH); // PDF y-up → sheet y-down
    const normX2 = px2 / pageW;
    const normY2 = 1 - (py2 / pageH);

    const sx1 = pdfState.sheetX + normX1 * pdfState.sheetWidth;
    const sy1 = pdfState.sheetY + normY1 * pdfState.sheetHeight;
    const sx2 = pdfState.sheetX + normX2 * pdfState.sheetWidth;
    const sy2 = pdfState.sheetY + normY2 * pdfState.sheetHeight;

    // Convert from sheet-mm (on-screen) to drawing-mm (real-world)
    // sheet-mm = drawArea.left + drawingMm / drawingScale
    // drawingMm = (sheetMm - drawArea.left) * drawingScale
    // But the grid data model stores in drawing-mm already (matching ortho position units)
    const da = engine.coords.drawArea;
    const x1mm = (sx1 - da.left) * CONFIG.drawingScale;
    const y1mm = (sy1 - da.top) * CONFIG.drawingScale;
    const x2mm = (sx2 - da.left) * CONFIG.drawingScale;
    const y2mm = (sy2 - da.top) * CONFIG.drawingScale;

    // Compute angle (degrees from horizontal, measured in sheet space)
    const dx = x2mm - x1mm;
    const dy = y2mm - y1mm;
    const angle = Math.atan2(dy, dx) * 180 / Math.PI;

    return {
        type: 'angled',
        x1: Math.round(x1mm),
        y1: Math.round(y1mm),
        x2: Math.round(x2mm),
        y2: Math.round(y2mm),
        angle: Math.round(angle * 10) / 10,
        label: label,
        zone: zoneId || undefined,
    };
}

/**
 * Cluster angle samples into groups where all angles in a group are
 * within ±tol degrees of the group mean.
 * Uses simple greedy clustering — sort by angle, merge adjacent.
 */
function dsiClusterAngles(samples, tol) {
    if (samples.length === 0) return [];

    // Sort by angle
    const sorted = [...samples].sort((a, b) => a.angle - b.angle);

    const clusters = [];
    let current = { samples: [sorted[0]], sum: sorted[0].angle };

    for (let i = 1; i < sorted.length; i++) {
        const mean = current.sum / current.samples.length;
        if (Math.abs(sorted[i].angle - mean) <= tol) {
            current.samples.push(sorted[i]);
            current.sum += sorted[i].angle;
        } else {
            current.meanAngle = current.sum / current.samples.length;
            clusters.push(current);
            current = { samples: [sorted[i]], sum: sorted[i].angle };
        }
    }
    current.meanAngle = current.sum / current.samples.length;
    clusters.push(current);

    // Also check wrap-around: angles near 0° and near 180° may be same
    // direction (horizontal lines). If first and last clusters are within
    // tol of each other (accounting for 180° wrap), merge them.
    if (clusters.length >= 2) {
        const first = clusters[0];
        const last = clusters[clusters.length - 1];
        const gap = (first.meanAngle + 180) - last.meanAngle;
        if (gap <= tol * 2) {
            // Merge last into first
            first.samples.push(...last.samples);
            first.sum += last.sum;
            first.meanAngle = first.sum / first.samples.length;
            clusters.pop();
        }
    }

    return clusters;
}

/**
 * Check if an angle (0-180°) is approximately orthogonal:
 * ~0° or ~180° = horizontal, ~90° = vertical.
 */
function dsiIsOrthoAngle(angle) {
    const ORTHO_TOL = 10; // degrees
    return (angle < ORTHO_TOL) ||
           (angle > 180 - ORTHO_TOL) ||
           (Math.abs(angle - 90) < ORTHO_TOL);
}

/**
 * Cluster a set of grid-label candidates into structural grid lines.
 * LEGACY fallback — used when angle-based detection produces no results.
 * Determines axis (H or V) by checking whether labels are spread
 * primarily along X or Y.
 */
function dsiClusterGridLabels(labels, pageW, pageH) {
    if (labels.length < 2) return [];

    // Group by unique label text
    const byText = {};
    for (const l of labels) {
        if (!byText[l.text]) byText[l.text] = [];
        byText[l.text].push(l);
    }

    const uniqueLabels = Object.keys(byText).sort((a, b) => {
        const aNum = parseInt(a);
        const bNum = parseInt(b);
        if (!isNaN(aNum) && !isNaN(bNum)) return aNum - bNum;
        return a.localeCompare(b);
    });

    if (uniqueLabels.length < 2) return [];

    const avgPositions = uniqueLabels.map(text => {
        const items = byText[text];
        const avgX = items.reduce((s, i) => s + i.x, 0) / items.length;
        const avgY = items.reduce((s, i) => s + i.y, 0) / items.length;
        return { text, avgX, avgY, count: items.length };
    });

    const xs = avgPositions.map(p => p.avgX);
    const ys = avgPositions.map(p => p.avgY);
    const xSpread = Math.max(...xs) - Math.min(...xs);
    const ySpread = Math.max(...ys) - Math.min(...ys);

    const isVerticalGrids = xSpread > ySpread;
    const axis = isVerticalGrids ? 'V' : 'H';

    const sequential = dsiFindSequentialLabels(uniqueLabels);
    if (sequential.length < 2) return [];

    // Convert PDF → sheet-mm → drawing-mm
    const da = engine.coords.drawArea;
    const grids = [];
    for (const labelText of sequential) {
        const pos = avgPositions.find(p => p.text === labelText);
        if (!pos) continue;

        let realPos;
        if (isVerticalGrids) {
            const normX = pos.avgX / pageW;
            const sheetX = pdfState.sheetX + normX * pdfState.sheetWidth;
            realPos = (sheetX - da.left) * CONFIG.drawingScale;
        } else {
            const normY = 1 - (pos.avgY / pageH);
            const sheetY = pdfState.sheetY + normY * pdfState.sheetHeight;
            realPos = (sheetY - da.top) * CONFIG.drawingScale;
        }

        grids.push({
            type: 'ortho',
            axis: axis,
            position: Math.round(realPos),
            label: labelText,
        });
    }

    return grids;
}

/**
 * Find the longest sequential run in an array of labels.
 * For letters: A,B,C,D,E (allows gaps of up to 2)
 * For numbers: 1,2,3,4 (allows gaps of up to 2)
 */
function dsiFindSequentialLabels(labels) {
    if (labels.length === 0) return [];

    // Check if these are letters or numbers
    const allNumbers = labels.every(l => /^\d+$/.test(l));

    if (allNumbers) {
        const nums = labels.map(l => parseInt(l)).sort((a, b) => a - b);
        return dsiFindLongestRun(nums).map(String);
    } else {
        // Letters
        const letterLabels = labels.filter(l => /^[A-Z]$/.test(l));
        const codes = letterLabels.map(l => l.charCodeAt(0)).sort((a, b) => a - b);
        const run = dsiFindLongestRun(codes);
        return run.map(c => String.fromCharCode(c));
    }
}

function dsiFindLongestRun(sorted) {
    if (sorted.length <= 1) return sorted;

    let bestRun = [sorted[0]];
    let currentRun = [sorted[0]];

    for (let i = 1; i < sorted.length; i++) {
        const gap = sorted[i] - sorted[i - 1];
        if (gap >= 1 && gap <= 2) {
            currentRun.push(sorted[i]);
        } else {
            if (currentRun.length > bestRun.length) bestRun = currentRun;
            currentRun = [sorted[i]];
        }
    }
    if (currentRun.length > bestRun.length) bestRun = currentRun;

    return bestRun;
}


/* ═══════════════════════════════════════════════════════════════════════
   UI — Modal System
   ═══════════════════════════════════════════════════════════════════════ */

function dsiShowModal(mode, content) {
    let modal = document.getElementById('dsi-modal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'dsi-modal';
        document.getElementById('app').appendChild(modal);
    }
    modal.className = 'dsi-modal';
    modal.style.display = 'flex';

    if (mode === 'loading') {
        modal.innerHTML = `
            <div class="dsi-modal-content dsi-loading">
                <div class="dsi-spinner"></div>
                <h3>Parsing Drawing Set</h3>
                <p id="dsi-loading-msg">Reading PDF pages...</p>
                <div class="dsi-progress-bar">
                    <div class="dsi-progress-fill" id="dsi-progress-fill" style="width:0%"></div>
                </div>
            </div>`;
    } else if (mode === 'importing') {
        modal.innerHTML = `
            <div class="dsi-modal-content dsi-loading">
                <div class="dsi-spinner"></div>
                <h3>Importing Drawing Set</h3>
                <p id="dsi-import-msg">Setting up levels...</p>
                <div class="dsi-progress-bar">
                    <div class="dsi-progress-fill" id="dsi-import-progress" style="width:0%"></div>
                </div>
            </div>`;
    } else if (mode === 'verify') {
        modal.innerHTML = `<div class="dsi-modal-content dsi-verify">${content}</div>`;
    } else if (mode === 'complete') {
        modal.innerHTML = `<div class="dsi-modal-content">${content}</div>`;
    }
}

function dsiHideModal() {
    const modal = document.getElementById('dsi-modal');
    if (modal) modal.style.display = 'none';
}

function dsiUpdateLoadingProgress(current, total) {
    const msg = document.getElementById('dsi-loading-msg');
    const fill = document.getElementById('dsi-progress-fill');
    if (msg) msg.textContent = `Reading page ${current} of ${total}...`;
    if (fill) fill.style.width = Math.round((current / total) * 100) + '%';
}

function dsiUpdateImportProgress(message, current, total) {
    const msg = document.getElementById('dsi-import-msg');
    const fill = document.getElementById('dsi-import-progress');
    if (msg) msg.textContent = message;
    if (fill) fill.style.width = Math.round(((current + 1) / total) * 100) + '%';
}

function dsiResetState() {
    dsiState.pdfDoc = null;
    dsiState.pdfArrayBuffer = null;
    dsiState.pages = [];
    dsiState.confirmed = false;
    dsiState.importing = false;
    dsiState.referenceGrids = [];
}


/* ── Helpers ─────────────────────────────────────────────────────────── */

function dsiEscHtml(s) {
    const d = document.createElement('div');
    d.textContent = s;
    return d.innerHTML;
}

/** Return a sort key for level IDs so they appear in natural building order. */
function dsiLevelSortOrder(levelId) {
    const order = {
        'B2': 10, 'B1': 20, 'LG': 30, 'GF': 40, 'MZ': 45, 'PD': 43,
        'L1': 50, 'L2': 60, 'L3': 70, 'L4': 80, 'L5': 90,
        'L6': 100, 'L7': 110, 'L8': 120, 'L9': 130, 'L10': 140,
        'TF': 145, 'RT': 150, 'RF': 160, 'SP': 5,
    };
    return order[levelId] || 200;
}


/* ═══════════════════════════════════════════════════════════════════════
   INIT — Wire up UI
   ═══════════════════════════════════════════════════════════════════════ */

(function dsiInit() {
    // File input for drawing set import
    const fileInput = document.getElementById('dsi-file-input');
    const importBtn = document.getElementById('btn-import-drawing-set');

    if (importBtn && fileInput) {
        importBtn.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) dsiStartImport(file);
            fileInput.value = ''; // reset so same file can be re-selected
        });
    }

    console.log('[DSI] Drawing Set Importer v' + DSI_VERSION + ' loaded');
})();
