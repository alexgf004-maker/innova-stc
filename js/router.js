/**
 * router.js
 * Enrutador SPA liviano.
 * - Carga fragmentos HTML en #content-area
 * - Genera sidebar y bottom nav según rol
 * - Protege rutas por rol
 * - Para agregar un módulo: registrar en ROUTES y NAV_ITEMS
 */

import { initAdmin }   from './modules/admin.js';
import { initHome }    from './modules/home.js';
import { initKardex }  from './modules/kardex.js';
import { initCambios } from './modules/cambios.js';

// ─────────────────────────────────────────
// REGISTRO DE RUTAS
// Para agregar un módulo: agrega una entrada aquí.
// ─────────────────────────────────────────
const ROUTES = {
  '/':                 { view: 'views/home.html',           init: initHome,    roles: ['admin', 'coordinadora', 'campo'] },
  '/kardex':           { view: 'views/kardex.html',         init: initKardex,  roles: ['admin', 'coordinadora', 'campo'] },
  '/otc':              { view: 'views/kardex.html',         init: initKardex,  roles: ['campo'] },
  '/cm':               { view: 'views/cambios.html',        init: initCambios, roles: ['campo', 'admin', 'coordinadora'] },
  '/admin/usuarios':   { view: 'views/admin-usuarios.html', init: initAdmin,   roles: ['admin'] },
  '/perfil':           { view: 'views/perfil.html',         init: null,        roles: ['admin', 'coordinadora', 'campo'] },
};

// ─────────────────────────────────────────
// ITEMS DE NAVEGACIÓN
// Para agregar un módulo: agrega una entrada aquí.
// ─────────────────────────────────────────
const NAV_ITEMS = [
  {
    path:  '/',
    label: 'Inicio',
    roles: ['admin', 'coordinadora', 'campo'],
    icon:  `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>
            </svg>`,
  },
  {
    path:  '/kardex',
    label: 'Kardex',
    roles: ['admin', 'coordinadora'],
    icon:  `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z"/>
              <polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/>
            </svg>`,
  },
  {
    path:  '/otc',
    label: 'OTC',
    roles: ['campo'],
    area:  'OTC',
    icon:  `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>
            </svg>`,
  },
  {
    path:  '/cm',
    label: 'CM',
    roles: ['campo'],
    area:  'CAMBIOS',
    icon:  `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <rect x="2" y="4" width="20" height="16" rx="2"/><path d="M12 14a2 2 0 100-4 2 2 0 000 4z"/><path d="M6 14l3.5-3.5M18 14l-3.5-3.5"/><line x1="12" y1="8" x2="12" y2="6"/>
            </svg>`,
  },
  {
    path:  '/cm',
    label: 'Cambios',
    roles: ['admin', 'coordinadora'],
    icon:  `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <rect x="2" y="4" width="20" height="16" rx="2"/><path d="M12 14a2 2 0 100-4 2 2 0 000 4z"/><path d="M6 14l3.5-3.5M18 14l-3.5-3.5"/><line x1="12" y1="8" x2="12" y2="6"/>
            </svg>`,
  },
  {
    path:  '/admin/usuarios',
    label: 'Usuarios',
    roles: ['admin'],
    icon:  `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/>
              <path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/>
            </svg>`,
  },
  // Fase 3: { path: '/otc',     label: 'OTC',     roles: [...], icon: `...` },
];

// ─────────────────────────────────────────
// INICIALIZACIÓN
// ─────────────────────────────────────────
export function initRouter(session) {
  buildNav(session);
  navigate(getCurrentPath(), session, true);

  // Interceptar clicks en links internos
  document.addEventListener('click', (e) => {
    const link = e.target.closest('[data-route]');
    if (link) {
      e.preventDefault();
      const path = link.dataset.route;
      navigate(path, session);
    }
  });

  // Manejar botón atrás del navegador
  window.addEventListener('popstate', () => {
    navigate(getCurrentPath(), session, true);
  });
}

// ─────────────────────────────────────────
// NAVEGACIÓN
// ─────────────────────────────────────────
export function navigate(path, session, replace = false) {
  const _session = session || window.__session;
  const route    = ROUTES[path] || ROUTES['/'];

  // Verificar rol
  if (!route.roles.includes(_session.role)) {
    loadView('/', _session);
    return;
  }

  if (!replace) {
    history.pushState({ path }, '', '#' + path);
  }

  updateActiveNav(path);
  loadView(path, _session);
}

async function loadView(path, session) {
  const route       = ROUTES[path] || ROUTES['/'];
  const contentArea = document.getElementById('content-area');

  contentArea.innerHTML = `
    <div class="flex items-center justify-center h-32">
      <div class="flex items-center gap-3 text-gray-400 text-sm">
        <svg class="animate-spin w-5 h-5" viewBox="0 0 24 24" fill="none">
          <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/>
          <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
        </svg>
        Cargando...
      </div>
    </div>`;

  try {
    const res  = await fetch(route.view);
    const html = await res.text();
    contentArea.innerHTML = html;

    // Inicializar módulo si tiene función init
    if (typeof route.init === 'function') {
      await route.init(session);
    }
  } catch (err) {
    contentArea.innerHTML = `
      <div class="flex flex-col items-center justify-center h-32 gap-2 text-gray-400">
        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/>
          <line x1="12" y1="16" x2="12.01" y2="16"/>
        </svg>
        <p class="text-sm">Error al cargar la vista.</p>
      </div>`;
    console.error('Error cargando vista:', route.view, err);
  }
}

// ─────────────────────────────────────────
// NAV
// ─────────────────────────────────────────
function buildNav(session) {
  const sidebarNav = document.getElementById('sidebar-nav');
  const bottomNav  = document.getElementById('bottom-nav');

  const visible = NAV_ITEMS.filter(item => item.roles.includes(session.role));

  // Sidebar (desktop) — todos los items visibles
  sidebarNav.innerHTML = visible.map(item => `
    <a data-route="${item.path}"
       class="nav-link sidebar-link flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors cursor-pointer"
       data-path="${item.path}">
      <span class="shrink-0">${item.icon}</span>
      <span>${item.label}</span>
    </a>
  `).join('');

  // Bottom nav (móvil) — máximo 4 items + logout
  const bottomItems = visible.slice(0, 4);
  const logoutIcon = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>`;
  bottomNav.innerHTML = bottomItems.map(item => `
    <a data-route="${item.path}"
       class="nav-link bottom-link flex flex-col items-center gap-1 px-3 py-2 text-xs font-medium transition-colors cursor-pointer min-w-0"
       data-path="${item.path}">
      <span>${item.icon}</span>
      <span class="truncate">${item.label}</span>
    </a>
  `).join('') + `
    <button id="btn-logout-bottom" class="flex flex-col items-center gap-1 px-3 py-2 text-xs font-medium cursor-pointer min-w-0 text-red-500 bg-transparent border-none">
      <span>${logoutIcon}</span>
      <span class="truncate">Salir</span>
    </button>`;

  // Sidebar logout button
  sidebarNav.innerHTML += `
    <button id="btn-logout-sidebar" class="flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium text-red-500 cursor-pointer bg-transparent border-none w-full mt-4">
      ${logoutIcon}
      <span>Cerrar sesión</span>
    </button>`;

  // Wire logout buttons
  document.getElementById('btn-logout-bottom')?.addEventListener('click', doLogout);
  document.getElementById('btn-logout-sidebar')?.addEventListener('click', doLogout);

  // PWA install button in sidebar
  const installBtn = document.createElement('button');
  installBtn.id = 'btn-install-pwa';
  installBtn.style.cssText = 'display:none;align-items:center;gap:8px;width:100%;padding:8px 12px;border-radius:10px;font-size:12px;font-weight:600;color:#0F766E;background:#F0FDFA;border:1px solid #99F6E4;cursor:pointer;margin-top:8px;';
  installBtn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Instalar app';
  installBtn.addEventListener('click', function() { window.__installPWA && window.__installPWA(); });
  const sidebarNav = document.getElementById('sidebar-nav');
  if (sidebarNav) sidebarNav.appendChild(installBtn);
}

function doLogout() {
  if (!confirm('¿Cerrar sesión?')) return;
  // Clear session
  try { window.__firebase?.auth?.signOut(); } catch(e) {}
  sessionStorage.clear();
  localStorage.removeItem('innova_session');
  window.location.href = '/innova-stc/login.html';
}

function updateActiveNav(currentPath) {
  document.querySelectorAll('.nav-link').forEach(link => {
    const path = link.dataset.path;
    const isActive = path === currentPath || (path !== '/' && currentPath.startsWith(path));

    if (link.classList.contains('sidebar-link')) {
      link.classList.toggle('sidebar-link--active', isActive);
    } else {
      link.classList.toggle('bottom-link--active', isActive);
    }
  });
}

function getCurrentPath() {
  return window.location.hash.slice(1) || '/';
}
