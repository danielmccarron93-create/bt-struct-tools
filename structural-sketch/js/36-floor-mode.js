// ═══════════════════════════════════════════════════════════════════
// 36-floor-mode.js — Floor Design Mode shell (Phase A)
//
// Floor Mode is a canvas context transformation — not a tab or panel.
// Entering swaps the ribbon to floor-workflow tools, pins a mode pill,
// frames the sheet in an amber border, and scopes activity to the
// currently active level.
//
// Phase A scope (this module):
//   • State object + enter/exit API
//   • Floor Design ribbon button wiring
//   • Keyboard: F to toggle, Esc to exit
//   • Body class .floor-mode-active drives ribbon swap via CSS
//   • Mode pill (fixed top-right amber badge)
//   • Render callback: amber inset border + subtle sheet tint
//   • Guard: requires an active level (prompts if none)
//   • Listeners: .on('enter' | 'exit', fn) for future phases
//
// Deferred to later phases:
//   • Layer dimming for non-floor elements (Phase B — needs layer
//     filter hook into the main render path, not implemented yet)
//   • Left workbench rail (Phase B)
//   • Summary chip (Phase C)
//   • Live utilization checks (Phase D)
//
// Namespace: window.floorMode
// ═══════════════════════════════════════════════════════════════════

(function () {
    'use strict';

    // ── State ────────────────────────────────────────────────────

    var state = {
        active: false,
        levelId: null,
        _listeners: { enter: [], exit: [] },
    };

    // ── Core: enter/exit ─────────────────────────────────────────

    function enter() {
        if (state.active) return;

        // Guard: require an active level
        var lv = (typeof getActiveLevel === 'function') ? getActiveLevel() : null;
        if (!lv) {
            // Non-blocking notification instead of alert (which blocks JS thread)
            if (typeof engine !== 'undefined' && engine._statusMsg) {
                engine._statusMsg = 'Select a level before entering Floor Design Mode.';
                engine.requestRender();
            }
            console.warn('[floor-mode] Cannot enter: no active level');
            return;
        }

        state.active = true;
        state.levelId = lv.id;

        // Body class drives CSS ribbon swap, pill visibility, cursor tint
        document.body.classList.add('floor-mode-active');

        // Update pill text
        updatePill(lv);

        // Mark the Floor Design ribbon button as active
        var btn = document.getElementById('btn-floor-mode');
        if (btn) btn.classList.add('active');

        // Request a canvas redraw so the amber border paints
        if (typeof engine !== 'undefined' && engine.requestRender) {
            engine.requestRender();
        }

        // Fire listeners
        _fire('enter', { levelId: state.levelId, level: lv });

        console.log('[floor-mode] ENTERED · level=' + state.levelId);
    }

    function exit() {
        if (!state.active) return;

        var prevLevel = state.levelId;
        state.active = false;
        state.levelId = null;

        document.body.classList.remove('floor-mode-active');

        var btn = document.getElementById('btn-floor-mode');
        if (btn) btn.classList.remove('active');

        if (typeof engine !== 'undefined' && engine.requestRender) {
            engine.requestRender();
        }

        _fire('exit', { levelId: prevLevel });

        console.log('[floor-mode] EXITED');
    }

    function toggle() {
        if (state.active) exit(); else enter();
    }

    // ── Level changes while active ──────────────────────────────
    // If the user switches level while in floor mode, update the pill
    // and re-request render. We poll the level index on a short
    // interval; the level system doesn't expose an event bus.

    var _lastSeenLevelIndex = null;
    setInterval(function () {
        if (!state.active) return;
        if (typeof levelSystem === 'undefined' || !levelSystem.levels) return;
        var idx = levelSystem.activeLevelIndex;
        if (idx === _lastSeenLevelIndex) return;
        _lastSeenLevelIndex = idx;
        var lv = levelSystem.levels[idx];
        if (!lv) return;
        state.levelId = lv.id;
        updatePill(lv);
        if (engine && engine.requestRender) engine.requestRender();
    }, 250);

    // ── Mode pill ────────────────────────────────────────────────

    function updatePill(lv) {
        var pill = document.getElementById('floor-mode-pill');
        if (!pill) return;
        var label = lv ? (lv.name || lv.id) : '—';
        pill.innerHTML =
            '<span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:#fff;margin-right:7px;box-shadow:0 0 6px rgba(255,255,255,0.9);"></span>' +
            '<strong style="letter-spacing:0.5px;">FLOOR MODE</strong>' +
            '<span style="opacity:0.85;margin:0 8px;">·</span>' +
            '<span>' + escapeHtml(label) + '</span>' +
            '<span style="opacity:0.7;margin-left:10px;font-size:10px;">press F to exit</span>';
    }

    function escapeHtml(s) {
        if (s == null) return '';
        return String(s).replace(/[&<>"]/g, function (c) {
            return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
        });
    }

    // ── Render callback: amber border + subtle sheet tint ───────
    // Phase A visual treatment. Full non-floor layer dimming comes
    // in Phase B. For now, a clear frame around the sheet signals
    // "you're in a mode" without needing render-path changes.

    function _floorModePaint(ctx, eng) {
        if (!state.active) return;

        var v = eng.viewport;
        var c = eng.config;
        var x = v.panX;
        var y = v.panY;
        var sw = c.SHEET_WIDTH_MM * v.zoom;
        var sh = c.SHEET_HEIGHT_MM * v.zoom;

        ctx.save();

        // Subtle warm tint over the sheet (very low alpha)
        ctx.fillStyle = 'rgba(245, 158, 11, 0.035)'; // amber 500 @ 3.5%
        ctx.fillRect(x, y, sw, sh);

        // Amber inset border — 3px line, inset 2px from sheet edge
        var inset = 2;
        ctx.strokeStyle = 'rgba(217, 119, 6, 0.85)'; // amber 600
        ctx.lineWidth = 3;
        ctx.setLineDash([]);
        ctx.strokeRect(x + inset, y + inset, sw - inset * 2, sh - inset * 2);

        // Corner markers to make the frame read as "mode"
        var cornerLen = Math.min(40, sw * 0.08);
        ctx.strokeStyle = 'rgba(217, 119, 6, 1)';
        ctx.lineWidth = 5;
        ctx.lineCap = 'square';
        [
            [x + inset, y + inset, 1, 1],
            [x + sw - inset, y + inset, -1, 1],
            [x + inset, y + sh - inset, 1, -1],
            [x + sw - inset, y + sh - inset, -1, -1],
        ].forEach(function (c4) {
            var cx = c4[0], cy = c4[1], dx = c4[2], dy = c4[3];
            ctx.beginPath();
            ctx.moveTo(cx + cornerLen * dx, cy);
            ctx.lineTo(cx, cy);
            ctx.lineTo(cx, cy + cornerLen * dy);
            ctx.stroke();
        });

        ctx.restore();
    }

    // ── Listener API (for future phases) ─────────────────────────

    function on(event, fn) {
        if (!state._listeners[event]) return;
        state._listeners[event].push(fn);
    }

    function _fire(event, payload) {
        var list = state._listeners[event];
        if (!list) return;
        for (var i = 0; i < list.length; i++) {
            try { list[i](payload); } catch (e) { console.error('[floor-mode listener]', e); }
        }
    }

    // ── Wiring ───────────────────────────────────────────────────

    function _wire() {
        // Floor Design button in main ribbon
        var btn = document.getElementById('btn-floor-mode');
        if (btn) {
            btn.addEventListener('click', function () { toggle(); });
        }

        // Exit button in contextual ribbon
        var exitBtn = document.getElementById('btn-floor-mode-exit');
        if (exitBtn) {
            exitBtn.addEventListener('click', function () { exit(); });
        }

        // Keyboard: F toggles, Esc exits (only if mode active and no other tool catches it)
        window.addEventListener('keydown', function (e) {
            // Ignore if typing in a form
            var t = e.target;
            if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable)) return;
            if (e.ctrlKey || e.metaKey || e.altKey) return;

            if (e.key === 'f' || e.key === 'F') {
                e.preventDefault();
                toggle();
                return;
            }

            // Esc exits, but only if mode is active and no drawing tool is owning Esc
            // (drawing tools like Floor Load polygon use Esc to cancel — leave them to it)
            if (e.key === 'Escape' && state.active) {
                var activeDrawingTool = (typeof activeTool !== 'undefined' && activeTool && activeTool !== 'pointer' && activeTool !== 'select');
                if (!activeDrawingTool) {
                    exit();
                }
            }
        });

        // Hook paint callback into the canvas render loop
        if (typeof engine !== 'undefined' && Array.isArray(engine._renderCallbacks)) {
            engine._renderCallbacks.push(_floorModePaint);
        }

        // Wire contextual ribbon's "floor tools" to the existing buttons
        // by proxying clicks. Phase A: each contextual button simply
        // triggers the same action as its main-ribbon counterpart.
        _wireContextualButtons();

        // Ensure the pill is populated if mode somehow already active
        if (state.active) {
            var lv = getActiveLevel && getActiveLevel();
            if (lv) updatePill(lv);
        }
    }

    /**
     * Contextual ribbon buttons proxy to existing main-ribbon buttons
     * so we don't duplicate tool wiring logic. Each cbtn-* has a
     * data-target attribute pointing at the real button id to click.
     */
    function _wireContextualButtons() {
        var ctxBtns = document.querySelectorAll('#floor-mode-toolbar [data-proxy]');
        for (var i = 0; i < ctxBtns.length; i++) {
            (function (btn) {
                var targetId = btn.getAttribute('data-proxy');
                btn.addEventListener('click', function (e) {
                    var target = document.getElementById(targetId);
                    if (target) target.click();
                });
            })(ctxBtns[i]);
        }
    }

    // ── Public API ───────────────────────────────────────────────

    window.floorMode = {
        version: '1.0-phaseA',
        isActive: function () { return state.active; },
        getLevelId: function () { return state.levelId; },
        enter: enter,
        exit: exit,
        toggle: toggle,
        on: on,
    };

    // ── Init ─────────────────────────────────────────────────────

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', _wire);
    } else {
        _wire();
    }

    console.log('[floor-mode] 1.0-phaseA loaded — press F to enter');
})();
