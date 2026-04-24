// ═══════════════════════════════════════════════════════════════════
// 38-joist-zone-3d.js — 3D modelling for joist zone elements
//
// Adds rectangular timber members to the 3D scene for each joist
// line computed by 37-joist-zone.js. Called from rebuild3DScene()
// in 11-3d-engine.js.
//
// Joist positioning:
//   SS joists:         top = level.elevation - fflOffset - D/2
//   Continuous joists: top = level.elevation - fflOffset - D/2
//                      (sit on top of beam; beam top adjusts)
//
// Uses buildRectGeo(widthMM, depthMM, lengthMM) from 11-3d-engine.js
// ═══════════════════════════════════════════════════════════════════

(function () {
    'use strict';

    var TIMBER_COLOR = 0xC4956A; // warm timber brown
    var scale3d = 1 / 1000; // mm → metres for Three.js

    /**
     * Build 3D geometry for a joistZone element.
     * Called from the rebuild3DScene element loop.
     *
     * @param {object} el      — joistZone element
     * @param {object} lv      — level object { elevation, ... }
     * @param {object} scene   — THREE.Scene
     * @param {number} scale   — scale factor (1/1000)
     */
    window.buildJoistZone3D = function (el, lv, scene, scale) {
        if (!el || !el.computed || !el.computed.joistLines) return;
        if (typeof THREE === 'undefined') return;
        if (typeof buildRectGeo !== 'function') return;

        var gov = el.computed.governingResult;
        if (!gov || !gov.section) return;

        var D_mm = gov.section.D_mm || 240;
        var B_mm = gov.section.B_mm || 45;
        var fflOffset = el.fflOffset_mm || 19;
        var elev = (lv.elevation || 0) * scale;

        // Material
        var mat = new THREE.MeshLambertMaterial({
            color: TIMBER_COLOR,
            transparent: true,
            opacity: 0.85,
        });

        var joistLines = el.computed.joistLines;
        for (var i = 0; i < joistLines.length; i++) {
            var jl = joistLines[i];
            for (var s = 0; s < jl.spans.length; s++) {
                var span = jl.spans[s];
                var spanLen = span.span_mm;
                if (spanLen < 50) continue;

                // Use per-span section if available, else governing
                var sect = (span.result && span.result.section) ? span.result.section : gov.section;
                var sD = sect.D_mm || D_mm;
                var sB = sect.B_mm || B_mm;

                // Build geometry
                var geo = buildRectGeo(sB, sD, spanLen);

                // Position
                var mx = ((span.x1 + span.x2) / 2) * scale;
                var mz = ((span.y1 + span.y2) / 2) * scale;
                // Y: top of joist at FFL - offset
                var yPos = elev - (fflOffset + sD / 2) * scale;

                var pos = new THREE.Vector3(mx, yPos, mz);

                // Rotation: match joist line angle
                var dx = span.x2 - span.x1;
                var dy = span.y2 - span.y1;
                var angle = Math.atan2(dy, dx);

                var mesh = new THREE.Mesh(geo, mat);
                mesh.position.copy(pos);
                mesh.rotation.y = -angle;

                // Edge wireframe for visibility
                var edges = new THREE.EdgesGeometry(geo);
                var edgeMat = new THREE.LineBasicMaterial({ color: 0x8B6914, opacity: 0.5, transparent: true });
                var edgeMesh = new THREE.LineSegments(edges, edgeMat);
                edgeMesh.position.copy(pos);
                edgeMesh.rotation.y = -angle;

                scene.add(mesh);
                scene.add(edgeMesh);
            }
        }
    };

    console.log('[joist-zone-3d] 1.0 loaded');
})();
