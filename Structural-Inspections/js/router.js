/**
 * Hash-based SPA router.
 *
 * Routes are defined by lazy-loading the view module and calling its
 * default-exported `render(outletElement)` function.
 */

const routes = {
  home:           () => import('./views/home.js'),
  projects:       () => import('./views/projects.js'),
  project:        () => import('./views/project-detail.js'),
  drawing:        () => import('./views/drawing-viewer.js'),
  inspections:    () => import('./views/inspections.js'),
  inspection:     () => import('./views/inspection-detail.js'),
  rectifications: () => import('./views/rectifications.js'),
  // Markup reuses drawing-viewer.js but exposes `renderMarkup` as `render`
  // so the router's module shape stays uniform.
  markup:         () => import('./views/drawing-viewer.js').then((m) => ({ render: m.renderMarkup })),
  settings:       () => import('./views/settings.js')
};

/** Which nav item (if any) should be highlighted for this route. */
const NAV_ROUTE_MAP = {
  project:    'projects',
  drawing:    'projects',
  inspection: 'inspections',
  markup:     'inspections'
};

const DEFAULT_ROUTE = 'home';

let outlet = null;
let navEl = null;

export function initRouter(outletEl, navElement) {
  outlet = outletEl;
  navEl = navElement;
  window.addEventListener('hashchange', handleRoute);
  // On first load, redirect bare `#` to the default
  if (!location.hash || location.hash === '#') {
    history.replaceState(null, '', `#/${DEFAULT_ROUTE}`);
  }
  handleRoute();
}

async function handleRoute() {
  const { route, params } = parseHash();
  const loader = routes[route] || routes[DEFAULT_ROUTE];

  updateNavActive(route);

  try {
    const mod = await loader();
    if (typeof mod.render !== 'function') {
      throw new Error(`View "${route}" does not export a render() function`);
    }
    outlet.innerHTML = '';
    const viewRoot = document.createElement('section');
    viewRoot.className = 'view';
    viewRoot.setAttribute('data-route', route);
    outlet.appendChild(viewRoot);
    await mod.render(viewRoot, params);
    // Move focus to main heading for a11y
    const heading = viewRoot.querySelector('h1, h2');
    if (heading) heading.setAttribute('tabindex', '-1');
  } catch (err) {
    console.error('Route load failed:', err);
    outlet.innerHTML = `
      <div class="card">
        <h2>Something went wrong</h2>
        <p class="muted">We couldn't load that screen. Try returning <a href="#/home">home</a>.</p>
      </div>`;
  }
}

function parseHash() {
  // "#/projects/123" → { route: "projects", params: ["123"] }
  const raw = (location.hash || '').replace(/^#\/?/, '');
  const parts = raw.split('/').filter(Boolean);
  const route = parts.shift() || DEFAULT_ROUTE;
  return { route, params: parts };
}

function updateNavActive(route) {
  if (!navEl) return;
  const highlighted = NAV_ROUTE_MAP[route] || route;
  navEl.querySelectorAll('[data-route]').forEach((item) => {
    if (item.dataset.route === highlighted) {
      item.setAttribute('aria-current', 'page');
    } else {
      item.removeAttribute('aria-current');
    }
  });
}

export function go(route, ...params) {
  const tail = params.length ? '/' + params.join('/') : '';
  const next = `#/${route}${tail}`;
  if (location.hash === next) {
    // Already there — force a re-render (hashchange won't fire)
    handleRoute();
  } else {
    location.hash = next;
  }
}

/** Force a re-render of the current route. Useful after data mutations. */
export function refresh() {
  handleRoute();
}
