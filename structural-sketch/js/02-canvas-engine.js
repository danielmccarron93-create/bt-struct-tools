// ── canvas-engine.js ───────────────────────────────────────

class CanvasEngine {
    constructor(container, config) {
        this.container = container;
        this.config = config;

        // Create canvases
        this.mainCanvas = document.createElement('canvas');
        this.mainCanvas.id = 'main-canvas';
        this.mainCanvas.style.cssText = 'position:absolute;top:0;left:0;';
        container.appendChild(this.mainCanvas);
        this.ctx = this.mainCanvas.getContext('2d');

        // PDF underlay canvas (behind main)
        this.pdfCanvas = document.createElement('canvas');
        this.pdfCanvas.id = 'pdf-canvas';
        this.pdfCanvas.style.cssText = 'position:absolute;top:0;left:0;';
        container.insertBefore(this.pdfCanvas, this.mainCanvas);
        this.pdfCtx = this.pdfCanvas.getContext('2d');

        // Viewport state
        this.viewport = {
            panX: 0,
            panY: 0,
            zoom: 1,
        };

        // Canvas pixel dimensions (CSS px, not device px)
        this.width = 0;
        this.height = 0;
        this.dpr = window.devicePixelRatio || 1;

        // Coordinate system
        this.coords = new CoordinateSystem(config, this.viewport);

        // Pan state
        this._isPanning = false;
        this._panStart = { x: 0, y: 0 };
        this._panStartViewport = { x: 0, y: 0 };
        this._spaceDown = false;

        // Render callbacks
        this._renderCallbacks = [];

        // Bind events
        this._bindEvents();
        this.resize();
        this.fitToView();
    }

    /** Register a callback to be called during render (after sheet is drawn) */
    onRender(fn) {
        this._renderCallbacks.push(fn);
    }

    /** Resize canvases to fill container */
    resize() {
        const rect = this.container.getBoundingClientRect();
        this.width = rect.width;
        this.height = rect.height;
        this.dpr = window.devicePixelRatio || 1;

        for (const canvas of [this.mainCanvas, this.pdfCanvas]) {
            canvas.width = Math.round(this.width * this.dpr);
            canvas.height = Math.round(this.height * this.dpr);
            canvas.style.width = this.width + 'px';
            canvas.style.height = this.height + 'px';
        }

        this.requestRender();
    }

    /** Fit the A1 sheet into view with padding */
    fitToView(paddingPx = 60) {
        const c = this.config;
        const zoomX = (this.width - paddingPx * 2) / c.SHEET_WIDTH_MM;
        const zoomY = (this.height - paddingPx * 2) / c.SHEET_HEIGHT_MM;
        this.viewport.zoom = Math.min(zoomX, zoomY);
        this.viewport.panX = (this.width - c.SHEET_WIDTH_MM * this.viewport.zoom) / 2;
        this.viewport.panY = (this.height - c.SHEET_HEIGHT_MM * this.viewport.zoom) / 2;
        this.requestRender();
    }

    _renderRequested = false;

    requestRender() {
        if (this._renderRequested) return;
        this._renderRequested = true;
        requestAnimationFrame(() => {
            this._renderRequested = false;
            this.render();
        });
    }

    render() {
        const ctx = this.ctx;
        const dpr = this.dpr;
        const w = this.width;
        const h = this.height;

        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        ctx.clearRect(0, 0, w, h);

        // Background (workspace area outside the sheet)
        ctx.fillStyle = this.config.CANVAS_BG;
        ctx.fillRect(0, 0, w, h);

        // Draw sheet
        this._drawSheet(ctx);
        this._drawBackgroundGrid(ctx);
        this._drawDrawingFrame(ctx);

        // Let registered renderers draw their content
        for (const fn of this._renderCallbacks) {
            fn(ctx, this);
        }
    }

    _drawSheet(ctx) {
        const c = this.config;
        const v = this.viewport;

        const x = v.panX;
        const y = v.panY;
        const sw = c.SHEET_WIDTH_MM * v.zoom;
        const sh = c.SHEET_HEIGHT_MM * v.zoom;

        // Sheet shadow
        ctx.save();
        ctx.shadowColor = 'rgba(0,0,0,0.15)';
        ctx.shadowBlur = 20;
        ctx.shadowOffsetX = 4;
        ctx.shadowOffsetY = 4;
        ctx.fillStyle = c.SHEET_BG;
        ctx.fillRect(x, y, sw, sh);
        ctx.restore();

        // Sheet border
        ctx.strokeStyle = c.SHEET_BORDER;
        ctx.lineWidth = 1;
        ctx.strokeRect(x, y, sw, sh);
    }

    _drawBackgroundGrid(ctx) {
        const c = this.config;
        if (!c.gridVisible) return;

        const v = this.viewport;
        const zoom = v.zoom;

        // Only draw grid if zoomed in enough to see it
        const minorScreenSize = c.GRID_MINOR_MM * zoom;
        if (minorScreenSize < 4) return; // too dense, skip

        const da = this.coords.drawArea;

        // Minor grid
        ctx.strokeStyle = c.GRID_MINOR_COLOR;
        ctx.lineWidth = 0.5;
        ctx.beginPath();

        const step = c.GRID_MINOR_MM;
        for (let sx = da.left; sx <= da.right; sx += step) {
            const p = this.coords.sheetToScreen(sx, da.top);
            const p2 = this.coords.sheetToScreen(sx, da.bottom);
            ctx.moveTo(Math.round(p.x) + 0.5, p.y);
            ctx.lineTo(Math.round(p2.x) + 0.5, p2.y);
        }
        for (let sy = da.top; sy <= da.bottom; sy += step) {
            const p = this.coords.sheetToScreen(da.left, sy);
            const p2 = this.coords.sheetToScreen(da.right, sy);
            ctx.moveTo(p.x, Math.round(p.y) + 0.5);
            ctx.lineTo(p2.x, Math.round(p2.y) + 0.5);
        }
        ctx.stroke();

        // Major grid
        const majorScreenSize = c.GRID_MAJOR_MM * zoom;
        if (majorScreenSize < 20) return;

        ctx.strokeStyle = c.GRID_MAJOR_COLOR;
        ctx.lineWidth = 0.5;
        ctx.beginPath();

        const majorStep = c.GRID_MAJOR_MM;
        for (let sx = da.left; sx <= da.right; sx += majorStep) {
            const p = this.coords.sheetToScreen(sx, da.top);
            const p2 = this.coords.sheetToScreen(sx, da.bottom);
            ctx.moveTo(Math.round(p.x) + 0.5, p.y);
            ctx.lineTo(Math.round(p2.x) + 0.5, p2.y);
        }
        for (let sy = da.top; sy <= da.bottom; sy += majorStep) {
            const p = this.coords.sheetToScreen(da.left, sy);
            const p2 = this.coords.sheetToScreen(da.right, sy);
            ctx.moveTo(p.x, Math.round(p.y) + 0.5);
            ctx.lineTo(p2.x, Math.round(p2.y) + 0.5);
        }
        ctx.stroke();
    }

    _drawDrawingFrame(ctx) {
        const c = this.config;
        const v = this.viewport;
        const zoom = v.zoom;

        // Outer margin frame
        const left = c.MARGIN_LEFT;
        const top = c.MARGIN_TOP;
        const right = c.SHEET_WIDTH_MM - c.MARGIN_RIGHT;
        const bottom = c.SHEET_HEIGHT_MM - c.MARGIN_BOTTOM;

        const tl = this.coords.sheetToScreen(left, top);
        const br = this.coords.sheetToScreen(right, bottom);

        ctx.strokeStyle = '#000000';
        ctx.lineWidth = Math.max(1, 0.7 * zoom);
        ctx.strokeRect(tl.x, tl.y, br.x - tl.x, br.y - tl.y);

        // Title block divider line
        const tbTop = bottom - c.TITLE_BLOCK_HEIGHT_MM;
        const tbLine = this.coords.sheetToScreen(left, tbTop);
        const tbLineR = this.coords.sheetToScreen(right, tbTop);

        ctx.beginPath();
        ctx.moveTo(tbLine.x, tbLine.y);
        ctx.lineTo(tbLineR.x, tbLineR.y);
        ctx.stroke();

        // Title block placeholder text (Phase 8 will do the full title block)
        if (zoom > 0.4) {
            this._drawTitleBlockPlaceholder(ctx, left, tbTop, right, bottom);
        }
    }

    /**
     * Full 5-zone title block matching Revit structural drawing conventions.
     * Zones (left → right): Rev Table | Company | Project Details | Drawing Title | Sign-off/Metadata
     */
    _drawTitleBlockPlaceholder(ctx, left, top, right, bottom) {
        const zoom = this.viewport.zoom;
        const c = this.config;
        const tb = this._titleBlockData || {};

        const w = right - left;
        const h = bottom - top;

        // Zone widths (mm) — matching the existing app layout
        const revW = 120;
        const compW = 110;
        const projW = 175;
        const titleW = 160;
        const signW = w - revW - compW - projW - titleW;

        const midRow = h * 0.5;
        const lineColor = '#000000';
        const labelColor = '#888888';
        const textColor = '#000000';
        const annoColor = '#666666';
        const font = '"Segoe UI", Arial, sans-serif';

        // Helper: draw a line in sheet coords
        const tbLine = (x1, y1, x2, y2) => {
            const p1 = this.coords.sheetToScreen(left + x1, top + y1);
            const p2 = this.coords.sheetToScreen(left + x2, top + y2);
            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.stroke();
        };

        // Helper: draw text at sheet-mm position (relative to title block origin)
        const tbText = (text, x, y, fontStr, color, align, baseline) => {
            const p = this.coords.sheetToScreen(left + x, top + y);
            ctx.font = fontStr;
            ctx.fillStyle = color;
            ctx.textAlign = align || 'left';
            ctx.textBaseline = baseline || 'top';
            ctx.fillText(text || '', p.x, p.y);
        };

        // Scaling factor for text sizes — fonts are in sheet-mm, need screen conversion
        const fs = (mm) => Math.max(1, mm * zoom);

        ctx.strokeStyle = lineColor;
        ctx.lineWidth = Math.max(0.5, 0.35 * zoom);

        // ── Zone divider lines ──
        tbLine(revW, 0, revW, h);
        tbLine(revW + compW, 0, revW + compW, h);
        tbLine(revW + compW + projW, 0, revW + compW + projW, h);
        tbLine(revW + compW + projW + titleW, 0, revW + compW + projW + titleW, h);

        if (zoom < 0.35) return; // Too small to render text

        // ── ZONE 1: Revision Table (far left) ──
        const rx = 0;
        const revColW = [18, 30, revW - 18 - 30 - 22 - 18, 22, 18];
        const revHeaders = ['Rev.', 'Date', 'Revision Details', 'Drn', 'App.'];

        // Header labels
        let rcx = rx;
        for (let ri = 0; ri < revHeaders.length; ri++) {
            tbText(revHeaders[ri], rcx + 1.5, 1.5, `${fs(3)}px ${font}`, labelColor);
            rcx += revColW[ri];
        }
        // Header separator
        tbLine(rx, 5, rx + revW, 5);

        // Revision data row
        const revRow = [
            tb.revision || 'A',
            tb.date || '',
            tb.revDesc || '',
            tb.drawnBy || '',
            ''
        ];
        rcx = rx;
        for (let ci = 0; ci < revRow.length; ci++) {
            tbText(revRow[ci], rcx + 1.5, 6, `${fs(3.3)}px ${font}`, textColor);
            rcx += revColW[ci];
        }

        // ── ZONE 2: Company / Logo ──
        const cx = revW;
        tbText(tb.company || '', cx + compW / 2, h * 0.32,
            `bold ${fs(7)}px ${font}`, textColor, 'center', 'middle');
        tbText(tb.companySubtitle || 'STRUCTURAL ENGINEERS', cx + compW / 2, h * 0.58,
            `${fs(3.8)}px ${font}`, annoColor, 'center', 'middle');
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';

        // ── ZONE 3: Project Details ──
        const px = cx + compW;
        tbText('Project Details', px + 3, 1.5, `${fs(3)}px ${font}`, labelColor);
        tbText(tb.projectName || 'PROJECT NAME', px + 3, 8,
            `bold ${fs(6.5)}px ${font}`, textColor);
        tbText(tb.address || 'SITE ADDRESS', px + 3, 16,
            `${fs(4.5)}px ${font}`, annoColor);

        // ── ZONE 4: Drawing Title ──
        const tx = px + projW;
        tbText('Drawing Title:', tx + 3, 1.5, `${fs(3)}px ${font}`, labelColor);
        tbText(tb.drawingTitle || 'STRUCTURAL PLAN', tx + 3, 8,
            `bold ${fs(5.5)}px ${font}`, textColor);

        // Status stamp
        if (tb.status) {
            tbText(tb.status, tx + titleW / 2, h * 0.72,
                `bold ${fs(5)}px ${font}`, '#CC0000', 'center', 'middle');
            ctx.textAlign = 'left';
            ctx.textBaseline = 'top';
        }

        // ── ZONE 5: Sign-off grid + metadata (far right) ──
        const sx = tx + titleW;
        const cellH = h / 5;

        // Sign-off rows
        const signRows = [
            { label: 'Drawn:', val: tb.drawnBy || '' },
            { label: 'Designed:', val: tb.designedBy || '' },
            { label: 'Checked:', val: tb.checkedBy || '' },
            { label: 'Approved:', val: tb.approvedBy || '' },
        ];

        let sry = 0;
        const labelW = 28;
        const nameW = 22;
        const signedW = 18;

        for (const row of signRows) {
            tbText(row.label, sx + 2, sry + 1, `${fs(3)}px ${font}`, labelColor);
            tbText(row.val, sx + labelW + 1, sry + 1, `${fs(3.8)}px ${font}`, textColor);
            tbText('Signed:', sx + labelW + nameW + 1, sry + 1, `${fs(3)}px ${font}`, labelColor);
            tbText('Date:', sx + labelW + nameW + signedW + 1, sry + 1, `${fs(3)}px ${font}`, labelColor);
            tbText(tb.date || '', sx + labelW + nameW + signedW + 1, sry + 4.5, `${fs(3.8)}px ${font}`, textColor);
            // Row separator
            tbLine(sx, sry + cellH, sx + signW, sry + cellH);
            sry += cellH;
        }

        // Last row: metadata fields
        const metaY = sry;
        const metaCellW = signW / 5;
        const metaFields = [
            { label: 'Scale:', val: tb.scale || '1:100' },
            { label: 'Project No:', val: tb.projectNumber || '' },
            { label: 'Drawing No:', val: tb.drawingNumber || 'S-001' },
            { label: 'Sheet Size:', val: 'A1' },
            { label: 'Rev:', val: tb.revision || 'A' },
        ];

        let mxx = sx;
        for (const mf of metaFields) {
            tbText(mf.label, mxx + 1, metaY + 0.5, `${fs(2.8)}px ${font}`, labelColor);
            tbText(mf.val, mxx + 1, metaY + 3.5, `bold ${fs(4.2)}px ${font}`, textColor);
            // Vertical separator
            tbLine(mxx, metaY, mxx, h);
            mxx += metaCellW;
        }
    }

    _bindEvents() {
        const el = this.container;

        // Zoom with scroll wheel
        el.addEventListener('wheel', (e) => {
            e.preventDefault();
            const zoomFactor = e.deltaY < 0 ? 1.04 : 1 / 1.04;
            const rect = el.getBoundingClientRect();
            const mx = e.clientX - rect.left;
            const my = e.clientY - rect.top;

            // Zoom toward cursor
            const oldZoom = this.viewport.zoom;
            const newZoom = Math.max(0.05, Math.min(20, oldZoom * zoomFactor));

            this.viewport.panX = mx - (mx - this.viewport.panX) * (newZoom / oldZoom);
            this.viewport.panY = my - (my - this.viewport.panY) * (newZoom / oldZoom);
            this.viewport.zoom = newZoom;

            this.requestRender();
        }, { passive: false });

        // Pan with middle mouse or Space+left
        el.addEventListener('mousedown', (e) => {
            if (e.button === 1 || (e.button === 0 && this._spaceDown)) {
                e.preventDefault();
                this._isPanning = true;
                this._panStart.x = e.clientX;
                this._panStart.y = e.clientY;
                this._panStartViewport.x = this.viewport.panX;
                this._panStartViewport.y = this.viewport.panY;
                el.style.cursor = 'grabbing';
            }
        });

        window.addEventListener('mousemove', (e) => {
            if (!this._isPanning) return;
            this.viewport.panX = this._panStartViewport.x + (e.clientX - this._panStart.x);
            this.viewport.panY = this._panStartViewport.y + (e.clientY - this._panStart.y);
            this.requestRender();
        });

        window.addEventListener('mouseup', (e) => {
            if (this._isPanning) {
                this._isPanning = false;
                el.style.cursor = '';
            }
        });

        // Space key for pan mode
        window.addEventListener('keydown', (e) => {
            if (e.code === 'Space' && !e.repeat && document.activeElement === document.body) {
                e.preventDefault();
                this._spaceDown = true;
                el.style.cursor = 'grab';
            }
        });

        window.addEventListener('keyup', (e) => {
            if (e.code === 'Space') {
                this._spaceDown = false;
                if (!this._isPanning) el.style.cursor = '';
            }
        });

        // Resize
        const ro = new ResizeObserver(() => this.resize());
        ro.observe(this.container);
    }

    /** Get mouse position in sheet-mm from a mouse event */
    getSheetPos(e) {
        const rect = this.container.getBoundingClientRect();
        const px = e.clientX - rect.left;
        const py = e.clientY - rect.top;
        return this.coords.screenToSheet(px, py);
    }

    /** Get mouse position in real-world mm from a mouse event */
    getRealPos(e) {
        const rect = this.container.getBoundingClientRect();
        const px = e.clientX - rect.left;
        const py = e.clientY - rect.top;
        return this.coords.screenToReal(px, py);
    }
}

