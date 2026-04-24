# BT Inspection Report

A Progressive Web Application for producing structural site inspection reports at Bligh Tanner. Mobile-first, offline-capable, designed to replace the current hybrid paper + Word + Bluebeam workflow.

See [`../site-inspection-app-north-star.md`](../site-inspection-app-north-star.md) for the full design specification. This README covers only what's needed to run and develop the app.

---

## Current status

**v2.0.0 — Phase 10 (V2 features landed)** ✅

Production-shape app covering the full north-star v1 scope **plus** the v2 features:

- Mobile-first drawing viewer with pan / zoom / rotate, distance calibration, and **two-point grid calibration** — new pins auto-label with `"3/B"` style references.
- 4-tier severity (Observation / Defect / **Hold point** / Closed) with red banner on the PDF cover.
- **Libre Franklin** embedded in generated PDFs — matches UI typography and the Bligh Tanner house style.
- Photo capture with **annotation** (circle / arrow / freehand / text) — vector overlay stored separately and composited at report time.
- **AI comment expansion** via Claude Sonnet 4.6 — short-form note → polished, AS-referenced comment. Bring your own Anthropic API key (pasted in Settings, stored on-device only).
- **Voice dictation** on the comment field — native keyboard mic on iOS, Web Speech API button on Chrome / Edge.
- Comment library of ~120 AS-referenced entries including hold-point templates.
- Generated PDF now has a **close-out page**: QR code that opens a pre-filled `mailto:` to the builder with a tracking token; an "Items awaiting rectification" list.
- **Export bundle** — one-click ZIP in the north-star §9.5 layout (`Report.pdf + Report_source.json + Photos/`). Drop into OneDrive / SharePoint manually; real MSAL-based sync is the v2.x follow-up.
- RPEQ signature image upload in Settings, rendered above the typed name on the PDF sign-off.
- Offline-first IndexedDB persistence, service worker caching all shell + library + font assets.

Still v2.x+ (intentionally deferred):

- Live OneDrive / SharePoint sync (needs Azure AD app registration).
- Self-serve builder upload portal (the QR close-out is currently `mailto:`-based).
- BIM / Revit / Aconex export.

---

## Running the app

The app is static HTML/CSS/JS — no build step. There are two ways to run it:

### Option 1 — one-click launcher (recommended)

- **macOS:** double-click `start-app.command`. A Terminal window opens and your browser loads `http://localhost:8080` automatically. Close the Terminal to stop the server.
- **Windows:** double-click `start-app.bat`. A command window opens and your browser loads `http://localhost:8080`. Close the window to stop the server.

Both scripts use Python (which ships with macOS and is usually installed on dev Windows machines). If Python isn't present they fall back to `npx serve`.

> ⚠️ **Don't just double-click `index.html`.** The app needs to be served over `http://` — V2 features (font embedding, QR generation, photo annotation, service worker) all rely on `fetch()` which browsers block on `file://` URLs.

> **First time on macOS** — if double-clicking `start-app.command` opens it in TextEdit instead of running it, right-click → **Open**, confirm once, and from then on a double-click just works.

### Option 2 — serve manually

```bash
# From the BT Inspection Report/ folder:
python3 -m http.server 8080    # macOS / Linux
python -m http.server 8080     # Windows (or the `py` launcher)
# …then open http://localhost:8080
```

### Option 3 — install to iPhone / iPad / Mac as a PWA

1. Serve the app over HTTPS (ngrok, Azure Static Web Apps, SharePoint, anywhere that gives it a public URL).
2. Visit the URL in Safari (or Chrome) on the target device.
3. **iPhone / iPad:** Share icon → **Add to Home Screen**.
4. **Mac (Safari):** Share icon → **Add to Dock**.
5. **Chrome (any OS):** the install icon in the address bar.

Once installed it launches full-screen and works offline — you no longer need the launcher.

---

## File structure

```
BT Inspection Report/
├── index.html              # App shell — single entry point
├── manifest.json           # PWA manifest (name, icons, theme)
├── service-worker.js       # Offline caching
├── assets/
│   ├── logo.svg            # Bligh Tanner logo (replace with licensed file if desired)
│   ├── logo-white.svg      # Reversed logo for dark backgrounds
│   ├── icon-maskable.svg   # PWA icon (iOS / Android install)
│   └── favicon.svg         # Browser tab icon
├── css/
│   └── styles.css          # Design system + component styles
└── js/
    ├── app.js              # Entry point — bootstraps router, DB, service worker
    ├── router.js           # Hash-based SPA routing
    ├── db.js               # IndexedDB schema via Dexie
    ├── state.js            # Small shared app state
    ├── views/              # One file per screen
    │   ├── home.js
    │   ├── projects.js
    │   ├── inspections.js
    │   └── settings.js
    └── components/         # Reusable UI bits
        └── header.js
```

---

## Tech stack

- **Vanilla JavaScript** with ES modules — no build pipeline.
- **Dexie.js** (CDN) — IndexedDB wrapper for offline storage.
- **PDF.js** (CDN) — renders uploaded structural drawings.
- **jsPDF** (CDN) — generates output report PDFs.
- **JSZip** (CDN) — builds the export bundle.
- **qrcode-generator** (CDN) — QR code on the close-out page.
- **Libre Franklin** (OFL, bundled under `assets/fonts/`) — embedded in the PDF; matches the Bligh Tanner house typeface.

Service worker pre-caches every dependency so the full app — fonts included — runs offline after first load.

---

## Design system

Brand tokens are defined as CSS custom properties in `css/styles.css`:

| Token | Value | Use |
|---|---|---|
| `--bt-blue` | `#00aaec` | Primary brand — headings, links, accents |
| `--bt-warm-grey` | `#4a4442` | Body text, secondary UI |
| `--bt-white` | `#ffffff` | Backgrounds |
| `--bt-sand` | `#f5f4f2` | Muted / panel backgrounds |
| `--bt-red` | `#d44426` | Hold points, destructive actions |
| `--bt-yellow` | `#ffc727` | Extent-of-inspection highlights |
| `--font-body` | `'Libre Franklin', …` | All UI text |

Typography rules follow the Bligh Tanner brand guide: ragged-left, strong hierarchy, Franklin-Gothic-equivalent weights.

---

## Replacing the logo

The supplied `assets/logo.svg` is a faithful text reproduction of the Bligh Tanner logotype using Libre Franklin Black. To use the licensed asset instead:

1. Drop `BT_Logo_RGB.svg` (or `.png` / `.jpg`) into `assets/`.
2. Rename to `logo.svg` (or `logo.png`), overwriting the placeholder.
3. Hard-refresh the browser.

The logo is referenced only from `index.html` and `css/styles.css` — no JS changes needed.

---

## Browser support

- **Primary targets:** iOS Safari 15+, iPad Safari 15+ (field use).
- **Secondary:** Chrome / Edge on Windows (office use, report review).
- Not tested on Android / Firefox — should work, not a focus.

---

## Known limitations

- Cloud sync is still manual — use **Export bundle** to drop the ZIP into OneDrive / SharePoint until the MSAL integration lands.
- Close-out is `mailto:`-based — builders reply with photos and the engineer pastes them against items inside the app.
- AI comment expansion depends on an Anthropic API key pasted in Settings; nothing is sent anywhere if the field is blank.
- Title-block extraction is best-effort — drawings sometimes need a manual description edit after upload.

See [`../site-inspection-app-north-star.md`](../site-inspection-app-north-star.md) for the full specification.
