/**
 * Persistent app header with Bligh Tanner logo.
 */

export function mountHeader(hostEl) {
  // Typographic logo — matches the boot screen and PDF output. Previously we
  // shipped an SVG here but the svg had a 280x180 viewBox and two stacked
  // text lines, which browsers rendered at a width that didn't respect our
  // max-width constraint at narrow viewports (the "BLIGH TANNE_I_" bug).
  hostEl.innerHTML = `
    <div class="app-header__inner">
      <a class="app-header__logo" href="#/home" aria-label="Bligh Tanner — Home">
        <span class="app-header__logo-blue">BLIGH</span><span class="app-header__logo-grey">TANNER</span>
      </a>
      <span class="app-header__title">Inspection Report</span>
      <div class="app-header__spacer"></div>
      <button class="app-header__action" id="header-sync" aria-label="Sync status" title="Sync status">
        <svg viewBox="0 0 24 24"><path d="M21 12a9 9 0 0 1-15 6.7L3 21"/><path d="M3 12a9 9 0 0 1 15-6.7L21 3"/><path d="M21 3v5h-5"/><path d="M3 21v-5h5"/></svg>
      </button>
    </div>
  `;
}
