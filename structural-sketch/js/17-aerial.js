// ── AERIAL MAP OVERLAY SYSTEM ────────────────────────────
// ══════════════════════════════════════════════════════════

// ── Aerial State ─────────────────────────────────────────

const aerialState = {
    active: false,          // overlay visible
    loading: false,         // tiles currently loading
    loaded: false,          // tiles have been fetched at least once

    // Geographic centre
    latitude: -27.4698,     // default: Brisbane CBD
    longitude: 153.0251,
    mapZoom: 20,            // tile zoom level (auto-calculated)

    // Display
    opacity: 0.35,
    rotation: 0,            // degrees CW from north (Phase 4)
    locked: false,

    // Position on sheet (sheet-mm) — top-left corner of the aerial image
    sheetX: 0,
    sheetY: 0,
    sheetWidth: 0,          // width in sheet-mm
    sheetHeight: 0,         // height in sheet-mm

    // Offscreen canvas with stitched tiles
    tileCanvas: null,
    tileCanvasWidth: 0,     // pixels
    tileCanvasHeight: 0,    // pixels

    // Tile grid info
    tilesX: 0,
    tilesY: 0,
    tileSize: 256,
    centerTileX: 0,
    centerTileY: 0,

    // Real-world coverage (metres)
    coverageWidthM: 0,
    coverageHeightM: 0,

    // Drag state
    dragging: false,
    dragStartScreen: null,
    dragOrigSheetPos: null,

    // Move tool active
    moveToolActive: false,

    // North arrow
    showNorthArrow: true,

    // Address for display
    address: '',

    // Overview mode: initial wide view before snapping to scale
    overviewMode: false,
    matchedZoom: 20,            // the zoom level that matches drawing scale
    overviewZoomFactor: 0.25,   // 0.25 = 1:400 when drawing is 1:100, 1.0 = matched
    matchedSheetWidth: 0,       // sheet-mm at exact drawing scale
    matchedSheetHeight: 0,      // sheet-mm at exact drawing scale
};

// ── Tile Math Utilities ──────────────────────────────────

/** Convert lat/lng to tile coordinates at given zoom */
function latLngToTile(lat, lng, zoom) {
    const n = Math.pow(2, zoom);
    const x = Math.floor((lng + 180) / 360 * n);
    const latRad = lat * Math.PI / 180;
    const y = Math.floor((1 - Math.log(Math.tan(latRad) + 1 / Math.cos(latRad)) / Math.PI) / 2 * n);
    return { x, y };
}

/** Convert lat/lng to fractional tile coordinates (for sub-tile positioning) */
function latLngToTileFrac(lat, lng, zoom) {
    const n = Math.pow(2, zoom);
    const x = (lng + 180) / 360 * n;
    const latRad = lat * Math.PI / 180;
    const y = (1 - Math.log(Math.tan(latRad) + 1 / Math.cos(latRad)) / Math.PI) / 2 * n;
    return { x, y };
}

/** Convert tile coordinates to lat/lng (top-left corner of tile) */
function tileToLatLng(tileX, tileY, zoom) {
    const n = Math.pow(2, zoom);
    const lng = tileX / n * 360 - 180;
    const latRad = Math.atan(Math.sinh(Math.PI * (1 - 2 * tileY / n)));
    const lat = latRad * 180 / Math.PI;
    return { lat, lng };
}

/** Ground resolution in metres per pixel at a given latitude and zoom */
function groundResolution(lat, zoom) {
    return 156543.03392 * Math.cos(lat * Math.PI / 180) / Math.pow(2, zoom);
}

/**
 * Calculate optimal tile zoom level for the current drawing scale.
 * We want the map ground resolution to match the drawing so that
 * 1 pixel on the tile ≈ 1 pixel on screen when viewport.zoom = 1.
 */
function calcOptimalTileZoom(lat, drawingScale) {
    // At viewport.zoom ~1, 1mm sheet ≈ 1px screen
    // 1mm sheet = drawingScale mm real = drawingScale/1000 m real
    // So we want ground resolution ≈ drawingScale/1000 m/px
    // But we want crisp tiles, so aim slightly higher res
    const targetRes = drawingScale / 1000;
    let bestZ = 1;
    for (let z = 1; z <= 21; z++) {
        const res = groundResolution(lat, z);
        if (res <= targetRes) {
            bestZ = z;
            break;
        }
        bestZ = z;
    }
    return Math.min(bestZ, 20); // ESRI max zoom is typically 20
}

/** Get ESRI World Imagery tile URL */
function getEsriTileUrl(z, y, x) {
    return 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/' + z + '/' + y + '/' + x;
}

// ── Geocoding (Nominatim / OpenStreetMap) ────────────────

async function geocodeAddress(address) {
    const url = 'https://nominatim.openstreetmap.org/search?format=json&q=' +
        encodeURIComponent(address) + '&limit=1&countrycodes=au';
    try {
        const response = await fetch(url, {
            headers: { 'User-Agent': 'StructuralSketch3D/1.0' }
        });
        const data = await response.json();
        if (data.length > 0) {
            return {
                lat: parseFloat(data[0].lat),
                lng: parseFloat(data[0].lon),
                displayName: data[0].display_name
            };
        }
        return null;
    } catch (err) {
        console.error('[Aerial] Geocoding failed:', err);
        return null;
    }
}

// ── Tile Fetching & Stitching ────────────────────────────

/**
 * Fetch satellite tiles centered on lat/lng and stitch into
 * an offscreen canvas. Returns the canvas.
 *
 * gridRadius: number of tiles in each direction from centre.
 * A radius of 5 gives an 11×11 grid = 121 tiles covering
 * ~370m × ~370m at zoom 20.
 */
async function fetchAndStitchTiles(lat, lng, zoom, gridRadius) {
    gridRadius = gridRadius || 5;

    const centerTile = latLngToTile(lat, lng, zoom);
    const tileSize = aerialState.tileSize;
    const gridDiam = gridRadius * 2 + 1;

    // Create offscreen canvas
    const canvasW = gridDiam * tileSize;
    const canvasH = gridDiam * tileSize;
    const offscreen = document.createElement('canvas');
    offscreen.width = canvasW;
    offscreen.height = canvasH;
    const ctx = offscreen.getContext('2d');

    // Grey placeholder
    ctx.fillStyle = '#E0E0E0';
    ctx.fillRect(0, 0, canvasW, canvasH);

    // Fetch all tiles in parallel
    const startTileX = centerTile.x - gridRadius;
    const startTileY = centerTile.y - gridRadius;
    const promises = [];

    for (let dy = 0; dy < gridDiam; dy++) {
        for (let dx = 0; dx < gridDiam; dx++) {
            const tx = startTileX + dx;
            const ty = startTileY + dy;
            const url = getEsriTileUrl(zoom, ty, tx);

            const p = new Promise((resolve) => {
                const img = new Image();
                img.crossOrigin = 'anonymous';
                img.onload = () => {
                    ctx.drawImage(img, dx * tileSize, dy * tileSize, tileSize, tileSize);
                    resolve(true);
                };
                img.onerror = () => {
                    // Draw a placeholder for failed tiles
                    ctx.fillStyle = '#D0D0D0';
                    ctx.fillRect(dx * tileSize, dy * tileSize, tileSize, tileSize);
                    ctx.strokeStyle = '#BBB';
                    ctx.strokeRect(dx * tileSize, dy * tileSize, tileSize, tileSize);
                    resolve(false);
                };
                img.src = url;
            });
            promises.push(p);
        }
    }

    await Promise.all(promises);

    // Store tile grid info
    aerialState.centerTileX = centerTile.x;
    aerialState.centerTileY = centerTile.y;
    aerialState.tilesX = gridDiam;
    aerialState.tilesY = gridDiam;
    aerialState.tileCanvas = offscreen;
    aerialState.tileCanvasWidth = canvasW;
    aerialState.tileCanvasHeight = canvasH;

    // Calculate real-world coverage
    const res = groundResolution(lat, zoom);
    aerialState.coverageWidthM = canvasW * res;
    aerialState.coverageHeightM = canvasH * res;

    // Calculate sheet-mm size at drawing scale (the "matched" size)
    const scale = CONFIG.drawingScale;
    aerialState.matchedSheetWidth = (aerialState.coverageWidthM * 1000) / scale;
    aerialState.matchedSheetHeight = (aerialState.coverageHeightM * 1000) / scale;
    // Display size — may differ in overview mode
    aerialState.sheetWidth = aerialState.matchedSheetWidth;
    aerialState.sheetHeight = aerialState.matchedSheetHeight;

    // Position: centre the aerial in the drawing area
    const c = CONFIG;
    const daLeft = c.MARGIN_LEFT;
    const daTop = c.MARGIN_TOP;
    const daW = c.SHEET_WIDTH_MM - c.MARGIN_LEFT - c.MARGIN_RIGHT;
    const daH = c.SHEET_HEIGHT_MM - c.MARGIN_TOP - c.MARGIN_BOTTOM - c.TITLE_BLOCK_HEIGHT_MM;

    aerialState.sheetX = daLeft + (daW - aerialState.sheetWidth) / 2;
    aerialState.sheetY = daTop + (daH - aerialState.sheetHeight) / 2;

    return offscreen;
}

// ── Load Aerial for a Location ───────────────────────────

/**
 * Load aerial imagery for a location.
 * mode = 'overview': loads at wide zoom (matchedZoom - 2) for site context
 * mode = 'matched':  loads at exact drawing scale match
 */
async function loadAerial(lat, lng, address, mode) {
    mode = mode || 'overview';

    aerialState.loading = true;
    aerialState.latitude = lat;
    aerialState.longitude = lng;
    aerialState.address = address || '';

    // Calculate optimal tile zoom for the current drawing scale
    aerialState.matchedZoom = calcOptimalTileZoom(lat, CONFIG.drawingScale);

    if (mode === 'overview') {
        // Wide view: 2 zoom levels lower = ~4× the area
        aerialState.mapZoom = Math.max(1, aerialState.matchedZoom - 2);
        aerialState.overviewMode = true;
    } else {
        aerialState.mapZoom = aerialState.matchedZoom;
        aerialState.overviewMode = false;
    }

    // Show loading state
    const goBtn = document.getElementById('btn-aerial-go');
    const origText = goBtn.textContent;
    goBtn.textContent = '...';
    goBtn.disabled = true;

    // Grid radius: larger for overview to cover more area
    const gridRadius = (mode === 'overview') ? 8 : 6;

    try {
        await fetchAndStitchTiles(lat, lng, aerialState.mapZoom, gridRadius);

        aerialState.loaded = true;
        aerialState.active = true;
        aerialState.loading = false;

        if (mode === 'overview') {
            // Apply overview zoom factor: display at 1/4 size (1:400 for 1:100 drawing)
            aerialState.overviewZoomFactor = 0.25;
            aerialState.sheetWidth = aerialState.matchedSheetWidth * aerialState.overviewZoomFactor;
            aerialState.sheetHeight = aerialState.matchedSheetHeight * aerialState.overviewZoomFactor;

            // Re-centre in drawing area at the overview size
            const c2 = CONFIG;
            const daW2 = c2.SHEET_WIDTH_MM - c2.MARGIN_LEFT - c2.MARGIN_RIGHT;
            const daH2 = c2.SHEET_HEIGHT_MM - c2.MARGIN_TOP - c2.MARGIN_BOTTOM - c2.TITLE_BLOCK_HEIGHT_MM;
            aerialState.sheetX = c2.MARGIN_LEFT + (daW2 - aerialState.sheetWidth) / 2;
            aerialState.sheetY = c2.MARGIN_TOP + (daH2 - aerialState.sheetHeight) / 2;
        }

        // Show controls
        document.getElementById('btn-aerial-toggle').classList.add('active');
        document.getElementById('btn-aerial-move').style.display = '';
        document.getElementById('aerial-controls').style.display = '';

        if (mode === 'overview') {
            // In overview: show Set Scale prominently, auto-activate Move
            document.getElementById('btn-aerial-set-scale').style.display = '';
            document.getElementById('btn-aerial-lock').style.display = 'none';
            document.getElementById('btn-aerial-set-north').style.display = 'none';

            // Auto-activate move mode so user can immediately drag
            aerialState.moveToolActive = true;
            const moveBtn = document.getElementById('btn-aerial-move');
            moveBtn.classList.add('active');
            moveBtn.textContent = 'Moving...';
            engine.container.style.cursor = 'move';

            // Show instruction banner
            const banner = document.getElementById('calib-banner');
            banner.classList.remove('hidden');
            banner.innerHTML = '<strong>Overview (1:' + (CONFIG.drawingScale * 4) + '):</strong> Scroll to zoom in, drag to position your building, then click <strong>Set Scale</strong>';
        } else {
            // Matched mode: show full controls
            document.getElementById('btn-aerial-set-scale').style.display = 'none';
            document.getElementById('btn-aerial-lock').style.display = '';
            document.getElementById('btn-aerial-set-north').style.display = '';
        }

        engine.requestRender();

        console.log('[Aerial] Loaded in ' + mode + ' mode at zoom ' + aerialState.mapZoom +
            ' (' + aerialState.tilesX + '×' + aerialState.tilesY + ' tiles)' +
            ' covering ' + aerialState.coverageWidthM.toFixed(0) + 'm × ' +
            aerialState.coverageHeightM.toFixed(0) + 'm' +
            ' → ' + aerialState.sheetWidth.toFixed(1) + '×' +
            aerialState.sheetHeight.toFixed(1) + 'mm on sheet');
    } catch (err) {
        console.error('[Aerial] Failed to load tiles:', err);
        alert('Failed to load aerial imagery. Check your internet connection.');
        aerialState.loading = false;
    }

    goBtn.textContent = origText;
    goBtn.disabled = false;
}

/**
 * Set Scale: snap from overview to matched scale.
 * Calculates the lat/lng at the centre of the current aerial position
 * on the sheet, then re-fetches tiles at the matched zoom.
 */
async function setAerialToMatchedScale() {
    if (!aerialState.loaded) return;

    // Find the lat/lng at the CENTRE OF THE DRAWING AREA on the sheet.
    // This is the point the user has positioned their building on.
    const c = CONFIG;
    const daCentreX = c.MARGIN_LEFT + (c.SHEET_WIDTH_MM - c.MARGIN_LEFT - c.MARGIN_RIGHT) / 2;
    const daCentreY = c.MARGIN_TOP + (c.SHEET_HEIGHT_MM - c.MARGIN_TOP - c.MARGIN_BOTTOM - c.TITLE_BLOCK_HEIGHT_MM) / 2;

    // What fraction of the aerial tile canvas is at the drawing area centre?
    // The aerial spans from sheetX to sheetX+sheetWidth on the sheet.
    const fracX = (daCentreX - aerialState.sheetX) / aerialState.sheetWidth;
    const fracY = (daCentreY - aerialState.sheetY) / aerialState.sheetHeight;

    // Clamp fractions (in case drawing centre is outside the aerial)
    const clampedFracX = Math.max(0, Math.min(1, fracX));
    const clampedFracY = Math.max(0, Math.min(1, fracY));

    // The tile grid covers a range of lat/lng. Find lat/lng at this fraction.
    const gridDiam = aerialState.tilesX;
    const startTileX = aerialState.centerTileX - Math.floor(gridDiam / 2);
    const startTileY = aerialState.centerTileY - Math.floor(gridDiam / 2);

    // Fractional tile position at the drawing area centre
    const tileFracX = startTileX + clampedFracX * gridDiam;
    const tileFracY = startTileY + clampedFracY * gridDiam;

    // Convert tile fraction to lat/lng
    const zoom = aerialState.mapZoom;
    const n = Math.pow(2, zoom);
    const centreLng = tileFracX / n * 360 - 180;
    const centreLatRad = Math.atan(Math.sinh(Math.PI * (1 - 2 * tileFracY / n)));
    const centreLat = centreLatRad * 180 / Math.PI;

    console.log('[Aerial] Set Scale: centre at lat=' + centreLat.toFixed(5) +
        ' lng=' + centreLng.toFixed(5) +
        ' (frac: ' + clampedFracX.toFixed(3) + ', ' + clampedFracY.toFixed(3) + ')');

    // Hide the overview banner
    document.getElementById('calib-banner').classList.add('hidden');

    // Deactivate move mode
    aerialState.moveToolActive = false;
    const moveBtn = document.getElementById('btn-aerial-move');
    moveBtn.classList.remove('active');
    moveBtn.textContent = 'Move';
    engine.container.style.cursor = '';

    // Re-load at matched scale, centred on the calculated lat/lng
    await loadAerial(centreLat, centreLng, aerialState.address, 'matched');
}

// ── Render Callback: Draw Aerial Underlay ────────────────

function drawAerialUnderlay(ctx, eng) {
    if (!aerialState.active || !aerialState.loaded || !aerialState.tileCanvas) return;

    const coords = eng.coords;
    const zoom = eng.viewport.zoom;

    ctx.save();
    ctx.globalAlpha = aerialState.opacity;

    // Clip to drawing area
    const da = coords.drawArea;
    const clipTL = coords.sheetToScreen(da.left, da.top);
    const clipBR = coords.sheetToScreen(da.right, da.bottom);
    ctx.beginPath();
    ctx.rect(clipTL.x, clipTL.y, clipBR.x - clipTL.x, clipBR.y - clipTL.y);
    ctx.clip();

    // Draw the stitched tile canvas at correct position and size
    const tl = coords.sheetToScreen(aerialState.sheetX, aerialState.sheetY);
    const w = aerialState.sheetWidth * zoom;
    const h = aerialState.sheetHeight * zoom;

    // Apply rotation if needed (Phase 4)
    if (aerialState.rotation !== 0) {
        const cx = tl.x + w / 2;
        const cy = tl.y + h / 2;
        ctx.translate(cx, cy);
        ctx.rotate(aerialState.rotation * Math.PI / 180);
        ctx.translate(-cx, -cy);
    }

    ctx.drawImage(aerialState.tileCanvas, tl.x, tl.y, w, h);

    ctx.restore();

    // North arrow (drawn at full opacity, not clipped)
    if (aerialState.showNorthArrow && aerialState.active) {
        drawNorthArrow(ctx, eng);
    }
}

// ── North Arrow ──────────────────────────────────────────

function drawNorthArrow(ctx, eng) {
    const coords = eng.coords;
    const da = coords.drawArea;

    // Position: top-right of drawing area, slightly inset
    const arrowSheetX = da.right - 15;
    const arrowSheetY = da.top + 15;
    const sp = coords.sheetToScreen(arrowSheetX, arrowSheetY);
    const zoom = eng.viewport.zoom;

    const size = Math.max(8, 18 * zoom);

    ctx.save();
    ctx.translate(sp.x, sp.y);

    // Rotate by aerial rotation (north rotates with the aerial)
    if (aerialState.rotation !== 0) {
        ctx.rotate(-aerialState.rotation * Math.PI / 180);
    }

    // Draw arrow body
    ctx.beginPath();
    ctx.moveTo(0, -size);           // tip (north)
    ctx.lineTo(-size * 0.35, size * 0.5);  // bottom-left
    ctx.lineTo(0, size * 0.25);     // notch
    ctx.lineTo(size * 0.35, size * 0.5);   // bottom-right
    ctx.closePath();

    // Left half filled dark, right half lighter
    ctx.fillStyle = 'rgba(30, 30, 30, 0.85)';
    ctx.fill();
    ctx.strokeStyle = 'rgba(0, 0, 0, 0.9)';
    ctx.lineWidth = Math.max(0.5, 0.8 * zoom);
    ctx.stroke();

    // "N" label
    const fontSize = Math.max(6, 8 * zoom);
    ctx.font = 'bold ' + fontSize + 'px "Segoe UI", Arial, sans-serif';
    ctx.fillStyle = 'rgba(30, 30, 30, 0.9)';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText('N', 0, -size - fontSize - 2);

    // White background circle behind arrow for contrast
    // Draw it behind (we'll redo this more cleanly)

    ctx.restore();
}

// Register aerial render callback — draw BEFORE PDF underlay and elements
// Insert at position 0 (behind everything)
engine._renderCallbacks.unshift(drawAerialUnderlay);

// ── Aerial: Toolbar Event Handlers ───────────────────────

// Toggle aerial visibility
document.getElementById('btn-aerial-toggle').addEventListener('click', () => {
    if (!aerialState.loaded) {
        // If not loaded yet, trigger a search or use default Brisbane
        const addrInput = document.getElementById('aerial-address');
        if (addrInput.value.trim()) {
            triggerAerialSearch();
        } else {
            // Load default Brisbane CBD
            loadAerial(-27.4698, 153.0251, 'Brisbane CBD');
        }
        return;
    }

    aerialState.active = !aerialState.active;
    const btn = document.getElementById('btn-aerial-toggle');
    if (aerialState.active) {
        btn.classList.add('active');
        document.getElementById('aerial-controls').style.display = '';
        document.getElementById('btn-aerial-move').style.display = '';
        if (aerialState.overviewMode) {
            document.getElementById('btn-aerial-set-scale').style.display = '';
            document.getElementById('btn-aerial-lock').style.display = 'none';
            document.getElementById('btn-aerial-set-north').style.display = 'none';
        } else {
            document.getElementById('btn-aerial-set-scale').style.display = 'none';
            document.getElementById('btn-aerial-lock').style.display = '';
            document.getElementById('btn-aerial-set-north').style.display = '';
        }
    } else {
        btn.classList.remove('active');
        document.getElementById('aerial-controls').style.display = 'none';
    }
    engine.requestRender();
});

// Address search
async function triggerAerialSearch() {
    const addrInput = document.getElementById('aerial-address');
    const address = addrInput.value.trim();
    if (!address) {
        addrInput.style.borderColor = '#D32F2F';
        setTimeout(() => { addrInput.style.borderColor = '#ccc'; }, 1500);
        return;
    }

    // Append Australia if not already specified
    let searchAddr = address;
    if (!/australia/i.test(searchAddr) && !/\bau\b/i.test(searchAddr)) {
        searchAddr += ', Australia';
    }

    const result = await geocodeAddress(searchAddr);
    if (result) {
        addrInput.style.borderColor = '#2E8B57';
        setTimeout(() => { addrInput.style.borderColor = '#ccc'; }, 1500);
        console.log('[Aerial] Found: ' + result.displayName +
            ' (' + result.lat.toFixed(5) + ', ' + result.lng.toFixed(5) + ')');
        await loadAerial(result.lat, result.lng, address);
    } else {
        addrInput.style.borderColor = '#D32F2F';
        setTimeout(() => { addrInput.style.borderColor = '#ccc'; }, 2000);
        alert('Address not found. Try a more specific address.');
    }
}

document.getElementById('btn-aerial-go').addEventListener('click', triggerAerialSearch);

// Set Scale button — snap from overview to matched scale
document.getElementById('btn-aerial-set-scale').addEventListener('click', async () => {
    await setAerialToMatchedScale();
});

document.getElementById('aerial-address').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
        e.preventDefault();
        triggerAerialSearch();
    }
});

// Populate address from project info if available
if (project.projectInfo.address && project.projectInfo.address !== 'SITE ADDRESS') {
    document.getElementById('aerial-address').value = project.projectInfo.address;
}

// Opacity slider
const aerialOpacitySlider = document.getElementById('aerial-opacity');
const aerialOpacityVal = document.getElementById('aerial-opacity-val');
aerialOpacitySlider.addEventListener('input', () => {
    aerialState.opacity = parseInt(aerialOpacitySlider.value, 10) / 100;
    aerialOpacityVal.textContent = aerialOpacitySlider.value + '%';
    engine.requestRender();
});

// ── Aerial: Overview Scroll-Zoom ─────────────────────────
// In overview mode, scroll wheel zooms the aerial view toward the cursor
// (like Google Maps zoom). This lets the user zoom in to find their building
// before hitting Set Scale.

container.addEventListener('wheel', (e) => {
    if (!aerialState.overviewMode || !aerialState.active || !aerialState.loaded) return;
    // Don't intercept if Shift is held (that's the rotation handler)
    if (e.shiftKey) return;

    e.preventDefault();
    e.stopPropagation();

    const zoomFactor = e.deltaY < 0 ? 1.10 : 1 / 1.10;
    const oldZF = aerialState.overviewZoomFactor;
    // Allow zooming from ~0.1 (very wide) up to 1.2 (past matched, for comfort)
    const newZF = Math.max(0.08, Math.min(1.2, oldZF * zoomFactor));

    // Zoom toward cursor: keep the point under the cursor fixed
    const rect = container.getBoundingClientRect();
    const cursorScreenX = e.clientX - rect.left;
    const cursorScreenY = e.clientY - rect.top;
    const cursorSheetPos = engine.coords.screenToSheet(cursorScreenX, cursorScreenY);

    // Fraction of the aerial under the cursor
    const oldW = aerialState.sheetWidth;
    const oldH = aerialState.sheetHeight;
    const fracX = (cursorSheetPos.x - aerialState.sheetX) / oldW;
    const fracY = (cursorSheetPos.y - aerialState.sheetY) / oldH;

    // New display size
    const newW = aerialState.matchedSheetWidth * newZF;
    const newH = aerialState.matchedSheetHeight * newZF;

    // Adjust position so cursor point stays fixed
    aerialState.sheetX = cursorSheetPos.x - fracX * newW;
    aerialState.sheetY = cursorSheetPos.y - fracY * newH;
    aerialState.sheetWidth = newW;
    aerialState.sheetHeight = newH;
    aerialState.overviewZoomFactor = newZF;

    // Update banner with current effective scale
    const effectiveScale = Math.round(CONFIG.drawingScale / newZF);
    const banner = document.getElementById('calib-banner');
    if (banner && !banner.classList.contains('hidden')) {
        banner.innerHTML = '<strong>Overview (1:' + effectiveScale + '):</strong> Scroll to zoom in, drag to position your building, then click <strong>Set Scale</strong>';
    }

    engine.requestRender();
}, { passive: false, capture: true });

// ── Aerial: Move / Drag Tool ─────────────────────────────

document.getElementById('btn-aerial-move').addEventListener('click', () => {
    if (aerialState.locked) {
        alert('Aerial is locked. Click Unlock first.');
        return;
    }
    aerialState.moveToolActive = !aerialState.moveToolActive;
    const btn = document.getElementById('btn-aerial-move');
    if (aerialState.moveToolActive) {
        btn.classList.add('active');
        btn.textContent = 'Moving...';
        engine.container.style.cursor = 'move';
    } else {
        btn.classList.remove('active');
        btn.textContent = 'Move';
        engine.container.style.cursor = '';
    }
});

// Lock / Unlock
document.getElementById('btn-aerial-lock').addEventListener('click', () => {
    aerialState.locked = !aerialState.locked;
    const btn = document.getElementById('btn-aerial-lock');
    if (aerialState.locked) {
        btn.textContent = 'Unlock';
        btn.classList.add('active');
        // Deactivate move tool if active
        if (aerialState.moveToolActive) {
            aerialState.moveToolActive = false;
            const moveBtn = document.getElementById('btn-aerial-move');
            moveBtn.classList.remove('active');
            moveBtn.textContent = 'Move';
            engine.container.style.cursor = '';
        }
    } else {
        btn.textContent = 'Lock';
        btn.classList.remove('active');
    }
});

// Mouse handlers for aerial drag
container.addEventListener('mousedown', (e) => {
    if (!aerialState.moveToolActive || !aerialState.active) return;
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;

    e.preventDefault();
    e.stopPropagation();

    aerialState.dragging = true;
    aerialState.dragStartScreen = { x: e.clientX, y: e.clientY };
    aerialState.dragOrigSheetPos = { x: aerialState.sheetX, y: aerialState.sheetY };
    engine.container.style.cursor = 'grabbing';
});

window.addEventListener('mousemove', (e) => {
    if (!aerialState.dragging) return;

    const dx = e.clientX - aerialState.dragStartScreen.x;
    const dy = e.clientY - aerialState.dragStartScreen.y;
    const zoom = engine.viewport.zoom;

    // Convert screen pixel delta to sheet-mm delta
    aerialState.sheetX = aerialState.dragOrigSheetPos.x + dx / zoom;
    aerialState.sheetY = aerialState.dragOrigSheetPos.y + dy / zoom;

    engine.requestRender();
});

window.addEventListener('mouseup', (e) => {
    if (aerialState.dragging) {
        aerialState.dragging = false;
        if (aerialState.moveToolActive) {
            engine.container.style.cursor = 'move';
        }
    }
});

// Arrow key nudging (1mm increments, 5mm with Shift)
window.addEventListener('keydown', (e) => {
    if (!aerialState.moveToolActive || !aerialState.active) return;
    if (aerialState.locked) return;

    const step = e.shiftKey ? 5 : 1; // mm
    let handled = false;

    if (e.key === 'ArrowLeft')  { aerialState.sheetX -= step; handled = true; }
    if (e.key === 'ArrowRight') { aerialState.sheetX += step; handled = true; }
    if (e.key === 'ArrowUp')    { aerialState.sheetY -= step; handled = true; }
    if (e.key === 'ArrowDown')  { aerialState.sheetY += step; handled = true; }

    if (handled) {
        e.preventDefault();
        engine.requestRender();
    }

    // Enter or Escape exits move mode and locks position
    if ((e.key === 'Enter' || e.key === 'Escape') && aerialState.moveToolActive) {
        aerialState.moveToolActive = false;
        const btn = document.getElementById('btn-aerial-move');
        btn.classList.remove('active');
        btn.textContent = 'Move';
        engine.container.style.cursor = '';
    }
});

// ── Aerial: Rotation Control ─────────────────────────────

const aerialRotationInput = document.getElementById('aerial-rotation');
aerialRotationInput.addEventListener('input', () => {
    let val = parseFloat(aerialRotationInput.value);
    if (isNaN(val)) return;
    // Clamp to -180..180
    val = Math.max(-180, Math.min(180, val));
    aerialState.rotation = val;
    engine.requestRender();
});

// Shift+scroll on canvas to rotate aerial (1° per tick)
container.addEventListener('wheel', (e) => {
    if (!e.shiftKey || !aerialState.active || !aerialState.loaded) return;
    if (aerialState.locked) return;
    e.preventDefault();
    e.stopPropagation();
    const delta = e.deltaY > 0 ? 1 : -1;
    aerialState.rotation = Math.max(-180, Math.min(180, aerialState.rotation + delta));
    aerialRotationInput.value = Math.round(aerialState.rotation);
    engine.requestRender();
}, { passive: false, capture: true });

// ── Aerial: Set North from Architect's Arrow (Two-Click) ─

const northCalibState = {
    active: false,
    points: [],      // [{x, y}] in sheet-mm, up to 2
    prevRotation: 0, // for undo
};

// Set North button
document.getElementById('btn-aerial-set-north').addEventListener('click', () => {
    if (!aerialState.active || !aerialState.loaded) {
        alert('Load aerial imagery first.');
        return;
    }

    // Deactivate move tool if active
    if (aerialState.moveToolActive) {
        aerialState.moveToolActive = false;
        const moveBtn = document.getElementById('btn-aerial-move');
        moveBtn.classList.remove('active');
        moveBtn.textContent = 'Move';
    }

    northCalibState.active = true;
    northCalibState.points = [];
    northCalibState.prevRotation = aerialState.rotation;

    const btn = document.getElementById('btn-aerial-set-north');
    btn.classList.add('active');
    btn.textContent = 'Pick...';

    // Show instruction banner (reuse calib-banner)
    const banner = document.getElementById('calib-banner');
    banner.classList.remove('hidden');
    banner.innerHTML = '<strong>Set North:</strong> Click the <strong>base</strong> of the architect\'s north arrow';

    engine.container.style.cursor = 'crosshair';
    engine.requestRender();
});

// Canvas click handler for north calibration
container.addEventListener('mousedown', (e) => {
    if (!northCalibState.active) return;
    if (e.button !== 0) return;
    if (engine._spaceDown || engine._isPanning) return;

    e.preventDefault();
    e.stopPropagation();

    const sheetPos = engine.getSheetPos(e);
    const banner = document.getElementById('calib-banner');

    if (northCalibState.points.length === 0) {
        // First point — base of north arrow
        northCalibState.points.push({ x: sheetPos.x, y: sheetPos.y });
        banner.innerHTML = '<strong>Set North:</strong> Now click the <strong>tip</strong> (north end) of the arrow';
        engine.requestRender();
    } else if (northCalibState.points.length === 1) {
        // Second point — tip of north arrow
        northCalibState.points.push({ x: sheetPos.x, y: sheetPos.y });

        const p1 = northCalibState.points[0]; // base
        const p2 = northCalibState.points[1]; // tip (north)

        // Calculate angle from vertical (positive Y is down on screen/sheet)
        // atan2(dx, -dy) gives angle from "up" direction
        const dx = p2.x - p1.x;
        const dy = p2.y - p1.y;
        const dist = Math.sqrt(dx * dx + dy * dy);

        if (dist < 1) {
            alert('Points are too close together. Try again.');
            cancelNorthCalib();
            return;
        }

        // Angle in degrees: 0° = straight up, positive = clockwise
        let angleDeg = Math.atan2(dx, -dy) * 180 / Math.PI;

        // Round to nearest 0.5°
        angleDeg = Math.round(angleDeg * 2) / 2;

        // Apply rotation
        aerialState.rotation = angleDeg;
        aerialRotationInput.value = Math.round(angleDeg);

        // End calibration mode
        northCalibState.active = false;
        northCalibState.points = [];
        engine.container.style.cursor = '';

        const northBtn = document.getElementById('btn-aerial-set-north');
        northBtn.classList.remove('active');
        northBtn.textContent = 'Set North';

        // Show confirmation banner with undo
        banner.innerHTML = '<strong>North set to ' + angleDeg.toFixed(1) +
            '° from vertical</strong> &nbsp; ' +
            '<a href="#" id="north-undo-link" style="color:#2B7CD0;text-decoration:underline;pointer-events:auto;">Undo</a>';
        banner.classList.remove('hidden');

        // Wire up undo link
        const undoLink = document.getElementById('north-undo-link');
        if (undoLink) {
            undoLink.addEventListener('click', (ev) => {
                ev.preventDefault();
                aerialState.rotation = northCalibState.prevRotation;
                aerialRotationInput.value = Math.round(northCalibState.prevRotation);
                banner.classList.add('hidden');
                engine.requestRender();
                console.log('[North] Undone, reverted to ' + northCalibState.prevRotation + '°');
            });
        }

        // Auto-hide banner after 6 seconds
        setTimeout(() => {
            if (banner.innerHTML.includes('North set to')) {
                banner.classList.add('hidden');
            }
        }, 6000);

        engine.requestRender();
        console.log('[North] Set to ' + angleDeg.toFixed(1) + '° from vertical');
    }
}, true); // capture phase so it fires before other handlers

// Draw north calibration points on the canvas
engine.onRender((ctx, eng) => {
    if (!northCalibState.active || northCalibState.points.length === 0) return;

    const zoom = eng.viewport.zoom;
    ctx.save();

    for (let i = 0; i < northCalibState.points.length; i++) {
        const pt = northCalibState.points[i];
        const sp = eng.coords.sheetToScreen(pt.x, pt.y);
        const r = Math.max(4, 6 * zoom);

        // Crosshair marker
        ctx.strokeStyle = '#FF4444';
        ctx.lineWidth = Math.max(1.5, 2 * zoom);

        ctx.beginPath();
        ctx.moveTo(sp.x - r, sp.y);
        ctx.lineTo(sp.x + r, sp.y);
        ctx.moveTo(sp.x, sp.y - r);
        ctx.lineTo(sp.x, sp.y + r);
        ctx.stroke();

        // Circle
        ctx.beginPath();
        ctx.arc(sp.x, sp.y, r, 0, Math.PI * 2);
        ctx.stroke();

        // Label
        const fontSize = Math.max(8, 10 * zoom);
        ctx.font = 'bold ' + fontSize + 'px "Segoe UI", sans-serif';
        ctx.fillStyle = '#FF4444';
        ctx.textAlign = 'left';
        ctx.textBaseline = 'bottom';
        ctx.fillText(i === 0 ? 'Base' : 'North', sp.x + r + 3, sp.y - 2);
    }

    // Draw line between points if both exist
    if (northCalibState.points.length === 2) {
        const sp1 = eng.coords.sheetToScreen(northCalibState.points[0].x, northCalibState.points[0].y);
        const sp2 = eng.coords.sheetToScreen(northCalibState.points[1].x, northCalibState.points[1].y);

        ctx.strokeStyle = '#FF4444';
        ctx.lineWidth = Math.max(1, 1.5 * zoom);
        ctx.setLineDash([4, 3]);
        ctx.beginPath();
        ctx.moveTo(sp1.x, sp1.y);
        ctx.lineTo(sp2.x, sp2.y);
        ctx.stroke();
        ctx.setLineDash([]);

        // Arrow head at tip
        const angle = Math.atan2(sp2.y - sp1.y, sp2.x - sp1.x);
        const headLen = Math.max(8, 12 * zoom);
        ctx.beginPath();
        ctx.moveTo(sp2.x, sp2.y);
        ctx.lineTo(sp2.x - headLen * Math.cos(angle - 0.4), sp2.y - headLen * Math.sin(angle - 0.4));
        ctx.moveTo(sp2.x, sp2.y);
        ctx.lineTo(sp2.x - headLen * Math.cos(angle + 0.4), sp2.y - headLen * Math.sin(angle + 0.4));
        ctx.stroke();
    }

    ctx.restore();
});

function cancelNorthCalib() {
    northCalibState.active = false;
    northCalibState.points = [];

    const banner = document.getElementById('calib-banner');
    banner.classList.add('hidden');
    engine.container.style.cursor = '';

    const btn = document.getElementById('btn-aerial-set-north');
    btn.classList.remove('active');
    btn.textContent = 'Set North';

    engine.requestRender();
}

// Escape cancels north calibration
window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && northCalibState.active) {
        cancelNorthCalib();
    }
});

// ── Aerial: Scale-Responsive Resizing ────────────────────
// When drawingScale changes, recalculate aerial sheet dimensions
// so the aerial always matches the drawing scale.

function recalcAerialForScale() {
    if (!aerialState.loaded || !aerialState.tileCanvas) return;

    const oldSheetW = aerialState.sheetWidth;
    const oldSheetH = aerialState.sheetHeight;
    const scale = CONFIG.drawingScale;

    // Recalculate sheet-mm size from real-world coverage
    const newSheetW = (aerialState.coverageWidthM * 1000) / scale;
    const newSheetH = (aerialState.coverageHeightM * 1000) / scale;

    // Adjust position so the aerial centre stays in the same place
    const oldCentreX = aerialState.sheetX + oldSheetW / 2;
    const oldCentreY = aerialState.sheetY + oldSheetH / 2;
    aerialState.sheetX = oldCentreX - newSheetW / 2;
    aerialState.sheetY = oldCentreY - newSheetH / 2;

    aerialState.sheetWidth = newSheetW;
    aerialState.sheetHeight = newSheetH;

    console.log('[Aerial] Scale changed to 1:' + scale +
        ' → aerial now ' + newSheetW.toFixed(1) + '×' +
        newSheetH.toFixed(1) + 'mm on sheet');

    engine.requestRender();
}

// Patch the three places where drawing scale gets set:

// 1. Scale selector dropdown (#scale-select)
const _scaleSelect = document.getElementById('scale-select');
if (_scaleSelect) {
    _scaleSelect.addEventListener('change', recalcAerialForScale);
}

// 2. PDF underlay scale preset (#scale-preset)
const _scalePreset = document.getElementById('scale-preset');
if (_scalePreset) {
    _scalePreset.addEventListener('change', () => {
        // Delay slightly so the main handler sets CONFIG.drawingScale first
        setTimeout(recalcAerialForScale, 50);
    });
}

// 3. Observe CONFIG.drawingScale for any other changes (calibration etc.)
// Use a polling approach as a safety net — check every render
let _lastKnownScale = CONFIG.drawingScale;
const _origAerialRender = drawAerialUnderlay;
engine._renderCallbacks[engine._renderCallbacks.indexOf(drawAerialUnderlay)] =
    function(ctx, eng) {
        if (CONFIG.drawingScale !== _lastKnownScale && aerialState.loaded) {
            _lastKnownScale = CONFIG.drawingScale;
            // Recalculate inline (don't re-trigger render to avoid loop)
            const scale = CONFIG.drawingScale;
            const oldW = aerialState.sheetWidth;
            const oldH = aerialState.sheetHeight;
            const newW = (aerialState.coverageWidthM * 1000) / scale;
            const newH = (aerialState.coverageHeightM * 1000) / scale;
            const cx = aerialState.sheetX + oldW / 2;
            const cy = aerialState.sheetY + oldH / 2;
            aerialState.sheetX = cx - newW / 2;
            aerialState.sheetY = cy - newH / 2;
            aerialState.sheetWidth = newW;
            aerialState.sheetHeight = newH;
            console.log('[Aerial] Auto-resize for 1:' + scale);
        }
        _origAerialRender(ctx, eng);
    };

// ── Aerial: Save/Load Integration ────────────────────────

// Patch toJSON to include aerial state
const _origToJSON = project.toJSON.bind(project);
project.toJSON = function() {
    const data = _origToJSON();
    data.groundRL = levelSystem.groundRL;
    data.showTOFTags = levelSystem.showTOFTags;
    if (aerialState.loaded) {
        data.aerialOverlay = {
            latitude: aerialState.latitude,
            longitude: aerialState.longitude,
            address: aerialState.address,
            sheetX: aerialState.sheetX,
            sheetY: aerialState.sheetY,
            opacity: aerialState.opacity,
            rotation: aerialState.rotation,
            locked: aerialState.locked,
            active: aerialState.active,
        };
    }
    return data;
};

// Patch fromJSON to restore aerial state
const _origFromJSON = ProjectData.fromJSON;
ProjectData.fromJSON = function(data) {
    const proj = _origFromJSON(data);
    if (data.aerialOverlay) {
        const ao = data.aerialOverlay;
        aerialState.opacity = ao.opacity || 0.35;
        aerialState.rotation = ao.rotation || 0;
        aerialState.locked = ao.locked || false;
        aerialState.address = ao.address || '';

        // Update UI
        aerialOpacitySlider.value = Math.round(aerialState.opacity * 100);
        aerialOpacityVal.textContent = Math.round(aerialState.opacity * 100) + '%';
        aerialRotationInput.value = Math.round(aerialState.rotation);

        if (ao.address) {
            document.getElementById('aerial-address').value = ao.address;
        }

        // Re-fetch tiles at saved position
        if (ao.latitude && ao.longitude) {
            loadAerial(ao.latitude, ao.longitude, ao.address).then(() => {
                // Restore exact position after reload
                if (ao.sheetX !== undefined) aerialState.sheetX = ao.sheetX;
                if (ao.sheetY !== undefined) aerialState.sheetY = ao.sheetY;
                aerialState.active = ao.active !== false;
                aerialState.locked = ao.locked || false;

                // Update lock button state
                const lockBtn = document.getElementById('btn-aerial-lock');
                if (aerialState.locked) {
                    lockBtn.textContent = 'Unlock';
                    lockBtn.classList.add('active');
                }

                // Update toggle
                const toggleBtn = document.getElementById('btn-aerial-toggle');
                if (aerialState.active) {
                    toggleBtn.classList.add('active');
                } else {
                    toggleBtn.classList.remove('active');
                }

                engine.requestRender();
            });
        }
    }
    return proj;
};

console.log('[Aerial] Aerial map overlay system ready');

// ── Initial render ─────────────────────────────────────────

updateStatusBar();
console.log('[StructuralSketch V5] All phases ready — Founding Strata module loaded');
