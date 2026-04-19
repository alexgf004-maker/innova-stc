/**
 * cambios.js — Módulo Cambio de Medidores
 * Fase 2: Gestión de órdenes, asignación por parejas, vista campo/admin
 */

import {
  collection, doc, getDocs, addDoc, updateDoc, setDoc,
  query, orderBy, where, serverTimestamp, getDoc
} from 'https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js';

import { showToast } from '../ui.js';

// ─────────────────────────────────────────
// CONSTANTES
// ─────────────────────────────────────────
const PAREJAS = ['Pareja 1', 'Pareja 2', 'Pareja 3', 'Pareja 4'];
const PAREJA_COLORS = {
  'Pareja 1': '#1B4F8A',
  'Pareja 2': '#EA580C',
  'Pareja 3': '#7C3AED',
  'Pareja 4': '#DB2777',
};
const COL_ORDENES    = 'cambios_ordenes';
const COL_CALENDARIO = 'cambios_calendario';

// ─────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────
function safeStr(v, def = '') { return (v !== null && v !== undefined) ? String(v).trim() : def; }
function safeNum(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function fmtDate(ts) {
  if (!ts) return '—';
  const d = ts.toDate ? ts.toDate() : new Date(ts);
  return d.toLocaleDateString('es-SV', { day: '2-digit', month: 'short', year: 'numeric' });
}
function loading() {
  return '<div class="flex items-center justify-center py-12 text-gray-400 text-sm gap-2"><svg class="animate-spin w-4 h-4" viewBox="0 0 24 24" fill="none"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/></svg>Cargando...</div>';
}
function errHtml(msg = 'Error al cargar datos') {
  return `<div class="text-center py-12 text-sm text-red-500">${msg}</div>`;
}
function mkOverlay(inner) {
  const ov = document.createElement('div');
  ov.className = 'fixed inset-0 bg-black/50 z-50 flex flex-col items-center justify-center p-4';
  ov.innerHTML = '<div class="bg-white rounded-2xl shadow-xl w-full max-w-md max-h-[90vh] flex flex-col overflow-hidden">' + inner + '</div>';
  document.body.appendChild(ov);
  return ov;
}

// Determinar si una orden está bloqueada por lectura
function esBloqueada(orden, calendarioMap) {
  const ul = safeStr(orden.unidadLectura);
  if (!ul || !calendarioMap[ul]) return false;
  const fecha = calendarioMap[ul].toDate ? calendarioMap[ul].toDate() : new Date(calendarioMap[ul]);
  const ahora = new Date();
  const diff  = (fecha - ahora) / (1000 * 60 * 60 * 24);
  return diff <= 2 && diff >= -2;
}

// ─────────────────────────────────────────
// INIT
// ─────────────────────────────────────────
export async function initCambios(session) {
  const container = document.getElementById('cambios-root');
  if (!container) return;
  const db = window.__firebase.db;

  const isCampo  = session.role === 'campo';
  const isAdmin  = ['admin', 'coordinadora'].includes(session.role);
  const destino  = session.asignacionActual?.destino || null;

  // Campo sin asignación a CAMBIOS
  if (isCampo && (!session.asignacionActual || session.asignacionActual.area !== 'CAMBIOS')) {
    container.innerHTML =
      '<div class="flex flex-col items-center justify-center min-h-64 px-6 text-center space-y-4">' +
        '<div class="w-16 h-16 rounded-2xl flex items-center justify-center" style="background:#F0FDFA">' +
          '<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#0F766E" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M4.22 4.22l2.12 2.12M17.66 17.66l2.12 2.12M2 12h3M19 12h3"/></svg>' +
        '</div>' +
        '<div>' +
          '<p class="font-bold text-gray-900 text-lg">Sin asignación a Cambios</p>' +
          '<p class="text-sm text-gray-500 mt-1">No tienes una pareja asignada en el área de Cambios.</p>' +
          '<p class="text-sm text-gray-500">Contacta a administración.</p>' +
        '</div>' +
      '</div>';
    return;
  }

  renderShell(container, session, db, isCampo, destino);
}

// ─────────────────────────────────────────
// SHELL
// ─────────────────────────────────────────
function renderShell(container, session, db, isCampo, destino) {
  const isAdmin = ['admin', 'coordinadora'].includes(session.role);

  container.innerHTML =
    '<div class="space-y-3">' +
      '<div class="flex items-center justify-between">' +
        '<div>' +
          '<h1 class="text-xl font-semibold text-gray-900">Cambio de Medidores</h1>' +
          (isCampo ? '<p class="text-xs text-gray-400 mt-0.5" style="color:#0F766E;font-weight:600">' + (destino || '') + '</p>' : '<p class="text-xs text-gray-400 mt-0.5">Gestión de órdenes</p>') +
        '</div>' +
        (isAdmin ?
          '<div class="flex gap-2">' +
            '<button id="btn-cargar-ordenes" class="text-xs font-medium px-3 py-2 rounded-lg border border-gray-300 text-gray-700 hover:bg-gray-50">📥 Órdenes</button>' +
            '<button id="btn-cargar-calendario" class="text-xs font-medium px-3 py-2 rounded-lg border border-gray-300 text-gray-700 hover:bg-gray-50">📅 Lectura</button>' +
          '</div>'
        : '') +
      '</div>' +
      // Tabs
      '<div class="flex gap-1 bg-gray-100 rounded-xl p-1">' +
        (isAdmin ? '<button class="ctab flex-1 py-2 text-xs font-medium rounded-lg transition-colors" data-ctab="listado">Listado</button>' : '') +
        (isAdmin ? '<button class="ctab flex-1 py-2 text-xs font-medium rounded-lg transition-colors" data-ctab="mapa">Mapa</button>' : '') +
        (isAdmin ? '<button class="ctab flex-1 py-2 text-xs font-medium rounded-lg transition-colors" data-ctab="seguimiento">Seguimiento</button>' : '') +
        (isCampo ? '<button class="ctab flex-1 py-2 text-xs font-medium rounded-lg transition-colors" data-ctab="listado">Mis órdenes</button>' : '') +
        (isCampo ? '<button class="ctab flex-1 py-2 text-xs font-medium rounded-lg transition-colors" data-ctab="mapa">Mapa</button>' : '') +
      '</div>' +
      '<div id="cambios-content"></div>' +
    '</div>';

  // Bind tabs
  setActiveTab('listado');
  showListado(db, session, isCampo, destino);

  container.querySelectorAll('.ctab').forEach(function(btn) {
    btn.addEventListener('click', async function() {
      const t = btn.dataset.ctab;
      setActiveTab(t);
      if (t === 'listado')     await showListado(db, session, isCampo, destino);
      if (t === 'mapa')        await showMapa(db, session, isCampo, destino);
      if (t === 'seguimiento') await showSeguimiento(db, session);
    });
  });

  if (isAdmin) {
    container.querySelector('#btn-cargar-ordenes')?.addEventListener('click', () => showCargarOrdenes(db));
    container.querySelector('#btn-cargar-calendario')?.addEventListener('click', () => showCargarCalendario(db));
  }
}

function setActiveTab(tab) {
  document.querySelectorAll('.ctab').forEach(function(b) {
    const a = b.dataset.ctab === tab;
    b.style.backgroundColor = a ? 'white' : 'transparent';
    b.style.color            = a ? '#0F766E' : '#6B7280';
    b.style.boxShadow        = a ? '0 1px 3px rgba(0,0,0,0.1)' : 'none';
  });
}

// ─────────────────────────────────────────
// CARGAR CALENDARIO (mapa de unidadLectura → fechaLectura)
// ─────────────────────────────────────────
async function getCalendarioMap(db) {
  const snap = await getDocs(collection(db, COL_CALENDARIO));
  const map  = {};
  snap.docs.forEach(function(d) {
    const data = d.data();
    if (data.unidadLectura && data.fechaLectura) {
      map[data.unidadLectura] = data.fechaLectura;
    }
  });
  return map;
}

// ─────────────────────────────────────────
// LISTADO
// ─────────────────────────────────────────
async function showListado(db, session, isCampo, destino) {
  const content = document.getElementById('cambios-content');
  if (!content) return;
  content.innerHTML = loading();

  try {
    const [snapOrdenes, calendarioMap] = await Promise.all([
      getDocs(collection(db, COL_ORDENES)),
      getCalendarioMap(db),
    ]);

    let ordenes = snapOrdenes.docs.map(function(d) { return Object.assign({ id: d.id }, d.data()); });

    // Campo: solo su pareja
    if (isCampo) {
      ordenes = ordenes.filter(function(o) { return o.pareja === destino; });
    }

    ordenes.sort(function(a, b) { return safeStr(a.cliente).localeCompare(safeStr(b.cliente)); });

    // Filtros para admin
    let filtroActivo = 'todas';
    let filtroPareja = 'todas';

    function renderFiltros() {
      if (isCampo) return '';
      return (
        '<div class="flex flex-wrap gap-2 mb-3">' +
          ['todas','disponibles','bloqueadas','hechas','visitas','sin_actualizar'].map(function(f) {
            const labels = { todas:'Todas', disponibles:'Disponibles', bloqueadas:'Bloqueadas', hechas:'Realizadas', visitas:'Visitas', sin_actualizar:'Sin actualizar' };
            const activo = filtroActivo === f;
            return '<button class="cfiltro text-xs px-2.5 py-1 rounded-full border font-medium ' +
              (activo ? 'text-white border-transparent' : 'border-gray-200 text-gray-600') + '" ' +
              'style="' + (activo ? 'background:#0F766E' : '') + '" data-filtro="' + f + '">' + labels[f] + '</button>';
          }).join('') +
        '</div>' +
        '<div class="flex gap-2 mb-3 flex-wrap">' +
          ['todas', ...PAREJAS].map(function(p) {
            const activo = filtroPareja === p;
            return '<button class="cpareja text-xs px-2.5 py-1 rounded-full border font-medium ' +
              (activo ? 'text-white border-transparent' : 'border-gray-200 text-gray-500') + '" ' +
              'style="' + (activo ? 'background:#1B4F8A' : '') + '" data-pareja="' + p + '">' + (p === 'todas' ? 'Todas las parejas' : p) + '</button>';
          }).join('') +
        '</div>'
      );
    }

    function filtrarOrdenes() {
      return ordenes.filter(function(o) {
        const bloqueada = esBloqueada(o, calendarioMap);
        if (!isCampo) {
          if (filtroPareja !== 'todas' && o.pareja !== filtroPareja) return false;
          if (filtroActivo === 'disponibles')   return !bloqueada && o.estadoCampo !== 'hecha';
          if (filtroActivo === 'bloqueadas')    return bloqueada;
          if (filtroActivo === 'hechas')        return o.estadoCampo === 'hecha';
          if (filtroActivo === 'visitas')       return o.estadoCampo === 'visita';
          if (filtroActivo === 'sin_actualizar') return o.estadoCampo === 'hecha' && !o.actualizadaDelsur;
        }
        return true;
      });
    }

    function renderCards() {
      const lista = filtrarOrdenes();
      const listEl = document.getElementById('cambios-lista');
      if (!listEl) return;

      if (!lista.length) {
        listEl.innerHTML = '<div class="text-center py-10 text-sm text-gray-400">Sin órdenes</div>';
        return;
      }

      listEl.innerHTML = lista.map(function(o) {
        const bloqueada = esBloqueada(o, calendarioMap);
        const hecha     = o.estadoCampo === 'hecha';
        const visita    = o.estadoCampo === 'visita';

        const statusBadge = bloqueada
          ? '<span class="text-xs px-2 py-0.5 rounded-full bg-gray-200 text-gray-500">🔒 Bloqueada</span>'
          : hecha
            ? '<span class="text-xs px-2 py-0.5 rounded-full" style="background:#DCFCE7;color:#166534">✓ Hecha' + (o.actualizadaDelsur ? '' : ' · ⚠ Sin actualizar') + '</span>'
            : visita
              ? '<span class="text-xs px-2 py-0.5 rounded-full" style="background:#FEF3C7;color:#B45309">👁 Visita</span>'
              : '<span class="text-xs px-2 py-0.5 rounded-full" style="background:#F0FDFA;color:#0F766E">Disponible</span>';

        return '<div class="bg-white rounded-xl border border-gray-200 px-4 py-3 ' + (bloqueada ? 'opacity-60' : '') + ' cursor-pointer active:bg-gray-50" data-wo="' + o.wo + '">' +
          '<div class="flex items-start justify-between gap-2 mb-1">' +
            '<div class="min-w-0">' +
              '<p class="text-xs font-mono text-gray-400">' + safeStr(o.wo) + '</p>' +
              '<p class="text-sm font-semibold text-gray-900 leading-tight mt-0.5">' + safeStr(o.cliente) + '</p>' +
            '</div>' +
            statusBadge +
          '</div>' +
          '<p class="text-xs text-gray-500 truncate">' + safeStr(o.direccion) + '</p>' +
          '<div class="flex items-center gap-2 mt-1.5 flex-wrap">' +
            (o.concepto ? '<span class="text-xs px-2 py-0.5 rounded-full bg-gray-100 text-gray-600">' + safeStr(o.concepto) + '</span>' : '') +
            (o.unidadLectura ? '<span class="text-xs px-2 py-0.5 rounded-full bg-gray-50 text-gray-400 font-mono">' + safeStr(o.unidadLectura) + '</span>' : '') +
            (!isCampo && o.pareja ? '<span class="text-xs px-2 py-0.5 rounded-full text-white" style="background:#1B4F8A">' + safeStr(o.pareja) + '</span>' : '') +
          '</div>' +
        '</div>';
      }).join('');

      // Wire card clicks
      listEl.querySelectorAll('[data-wo]').forEach(function(card) {
        card.addEventListener('click', function() {
          const orden = ordenes.find(function(o) { return o.wo === card.dataset.wo; });
          if (orden) showDetalleOrden(db, session, orden, isCampo, calendarioMap);
        });
      });

      // Asignación masiva admin
      if (!isCampo) {
        const btnAsignar = document.getElementById('btn-asignar-pareja');
        if (btnAsignar) {
          btnAsignar.onclick = function() {
            const selCards = listEl.querySelectorAll('[data-wo].selected');
            const wos = Array.from(selCards).map(function(c) { return c.dataset.wo; });
            if (!wos.length) { showToast('Selecciona al menos una orden.', 'error'); return; }
            showAsignarPareja(db, wos, function() { showListado(db, session, isCampo, destino); });
          };
        }
        // Long press to select
        listEl.querySelectorAll('[data-wo]').forEach(function(card) {
          let timer;
          card.addEventListener('pointerdown', function() {
            timer = setTimeout(function() {
              card.classList.toggle('selected');
              card.style.outline = card.classList.contains('selected') ? '2px solid #1B4F8A' : '';
            }, 500);
          });
          card.addEventListener('pointerup', function() { clearTimeout(timer); });
          card.addEventListener('pointercancel', function() { clearTimeout(timer); });
        });
      }
    }

    if (isCampo) {
      // Campo: two sections — pendientes and realizadas
      const pendientes  = ordenes.filter(function(o) { return o.estadoCampo !== 'hecha'; });
      const realizadas  = ordenes.filter(function(o) { return o.estadoCampo === 'hecha'; });

      function campoCard(o) {
        const bloqueada = esBloqueada(o, calendarioMap);
        const realizada = o.estadoCampo === 'hecha';
        const visita    = o.estadoCampo === 'visita';
        const statusBadge = bloqueada
          ? '<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:#f3f4f6;color:#9ca3af">🔒 Bloqueada</span>'
          : realizada
            ? '<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:#DCFCE7;color:#166534">✓ Realizada' + (!o.actualizadaDelsur ? ' · ⚠' : '') + '</span>'
            : visita
              ? '<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:#FEF3C7;color:#B45309">👁 Visita</span>'
              : '<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:#F0FDFA;color:#0F766E">Pendiente</span>';

        return '<div class="bg-white rounded-xl border border-gray-200 px-4 py-3 cursor-pointer active:bg-gray-50 ' + (bloqueada ? 'opacity-60' : '') + '" data-wo="' + o.wo + '">' +
          '<div class="flex items-start justify-between gap-2 mb-1">' +
            '<div class="min-w-0">' +
              '<p style="font-size:10px;color:#9ca3af;font-family:monospace">' + safeStr(o.wo) + '</p>' +
              '<p class="text-sm font-semibold text-gray-900 leading-tight mt-0.5">' + safeStr(o.cliente) + '</p>' +
            '</div>' +
            statusBadge +
          '</div>' +
          '<p class="text-xs text-gray-500 truncate">' + safeStr(o.direccion) + '</p>' +
          (o.concepto ? '<span class="text-xs px-2 py-0.5 rounded-full bg-gray-100 text-gray-600 inline-block mt-1">' + safeStr(o.concepto) + '</span>' : '') +
        '</div>';
      }

      content.innerHTML =
        '<div class="space-y-4">' +
          // Stats
          '<div class="grid grid-cols-3 gap-2">' +
            '<div class="bg-white rounded-xl border border-gray-200 p-3 text-center">' +
              '<p class="text-xl font-black" style="color:#0F766E">' + pendientes.filter(function(o){ return !esBloqueada(o,calendarioMap) && o.estadoCampo !== 'visita'; }).length + '</p>' +
              '<p class="text-xs text-gray-400 mt-0.5">Pendientes</p>' +
            '</div>' +
            '<div class="bg-white rounded-xl border border-gray-200 p-3 text-center">' +
              '<p class="text-xl font-black" style="color:#166534">' + realizadas.length + '</p>' +
              '<p class="text-xs text-gray-400 mt-0.5">Realizadas</p>' +
            '</div>' +
            '<div class="bg-white rounded-xl border border-gray-200 p-3 text-center">' +
              '<p class="text-xl font-black" style="color:#B45309">' + ordenes.filter(function(o){ return o.estadoCampo === 'visita'; }).length + '</p>' +
              '<p class="text-xs text-gray-400 mt-0.5">Visitas</p>' +
            '</div>' +
          '</div>' +
          // Pendientes
          (pendientes.length > 0 ?
            '<div>' +
              '<p class="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Por realizar (' + pendientes.length + ')</p>' +
              '<div id="lista-pendientes" class="space-y-2">' + pendientes.map(campoCard).join('') + '</div>' +
            '</div>'
          : '<div class="text-center py-6"><p class="text-sm font-semibold text-green-700">¡Todo realizado! 🎉</p></div>') +
          // Realizadas
          (realizadas.length > 0 ?
            '<div>' +
              '<p class="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Realizadas (' + realizadas.length + ')</p>' +
              '<div id="lista-realizadas" class="space-y-2">' + realizadas.map(campoCard).join('') + '</div>' +
            '</div>'
          : '') +
        '</div>';

      // Wire card clicks
      content.querySelectorAll('[data-wo]').forEach(function(card) {
        card.addEventListener('click', function() {
          const orden = ordenes.find(function(o) { return o.wo === card.dataset.wo; });
          if (orden) showDetalleOrden(db, session, orden, isCampo, calendarioMap);
        });
      });

    } else {
      // Admin listado
      content.innerHTML =
        '<div class="space-y-2">' +
          renderFiltros() +
          '<button id="btn-asignar-pareja" class="w-full text-xs font-medium py-2 rounded-lg border border-gray-300 text-gray-600 hover:bg-gray-50 mb-1">Mantén presionado para seleccionar · Asignar pareja</button>' +
          '<div id="cambios-lista" class="space-y-2"></div>' +
        '</div>';

      renderCards();
    }

    // Wire filtros
    content.querySelectorAll('.cfiltro').forEach(function(btn) {
      btn.addEventListener('click', function() {
        filtroActivo = btn.dataset.filtro;
        content.querySelector('.space-y-2').innerHTML = renderFiltros() +
          (!isCampo ? '<button id="btn-asignar-pareja" class="w-full text-xs font-medium py-2 rounded-lg border border-gray-300 text-gray-600 hover:bg-gray-50 mb-1">Mantén presionado para seleccionar · Asignar pareja</button>' : '') +
          '<div id="cambios-lista" class="space-y-2"></div>';
        renderCards();
        rebindFiltros();
      });
    });
    content.querySelectorAll('.cpareja').forEach(function(btn) {
      btn.addEventListener('click', function() {
        filtroPareja = btn.dataset.pareja;
        content.querySelector('.space-y-2').innerHTML = renderFiltros() +
          (!isCampo ? '<button id="btn-asignar-pareja" class="w-full text-xs font-medium py-2 rounded-lg border border-gray-300 text-gray-600 hover:bg-gray-50 mb-1">Mantén presionado para seleccionar · Asignar pareja</button>' : '') +
          '<div id="cambios-lista" class="space-y-2"></div>';
        renderCards();
        rebindFiltros();
      });
    });

    function rebindFiltros() {
      content.querySelectorAll('.cfiltro').forEach(function(btn) {
        btn.addEventListener('click', function() {
          filtroActivo = btn.dataset.filtro;
          content.querySelector('.space-y-2').innerHTML = renderFiltros() +
            (!isCampo ? '<button id="btn-asignar-pareja" class="w-full text-xs font-medium py-2 rounded-lg border border-gray-300 text-gray-600 hover:bg-gray-50 mb-1">Mantén presionado para seleccionar · Asignar pareja</button>' : '') +
            '<div id="cambios-lista" class="space-y-2"></div>';
          renderCards(); rebindFiltros();
        });
      });
      content.querySelectorAll('.cpareja').forEach(function(btn) {
        btn.addEventListener('click', function() {
          filtroPareja = btn.dataset.pareja;
          content.querySelector('.space-y-2').innerHTML = renderFiltros() +
            (!isCampo ? '<button id="btn-asignar-pareja" class="w-full text-xs font-medium py-2 rounded-lg border border-gray-300 text-gray-600 hover:bg-gray-50 mb-1">Mantén presionado para seleccionar · Asignar pareja</button>' : '') +
            '<div id="cambios-lista" class="space-y-2"></div>';
          renderCards(); rebindFiltros();
        });
      });
    }

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

// ─────────────────────────────────────────
// MAPA
// ─────────────────────────────────────────
const GMAPS_KEY = 'AIzaSyAjaEXeu_4PDedaZhLfrwWvatu5RN9q1SU';

function loadGoogleMaps() {
  return new Promise(function(resolve) {
    if (window.google && window.google.maps && window.markerClusterer) { resolve(); return; }
    // Load MarkerClusterer first, then Google Maps
    function loadClusterer() {
      if (window.markerClusterer) { resolve(); return; }
      const sc = document.createElement('script');
      sc.src = 'https://unpkg.com/@googlemaps/markerclusterer/dist/index.min.js';
      sc.onload = function() { resolve(); };
      sc.onerror = function() { resolve(); }; // proceed even if clusterer fails
      document.head.appendChild(sc);
    }
    if (window.google && window.google.maps) { loadClusterer(); return; }
    window.__gmapsCallback = loadClusterer;
    const s = document.createElement('script');
    s.src = 'https://maps.googleapis.com/maps/api/js?key=' + GMAPS_KEY + '&callback=__gmapsCallback&libraries=places,geometry,drawing';
    s.async = true;
    document.head.appendChild(s);
  });
}

async function showMapa(db, session, isCampo, destino) {
  const content = document.getElementById('cambios-content');
  if (!content) return;
  content.innerHTML = loading();

  try {
    const [snapOrdenes, calendarioMap] = await Promise.all([
      getDocs(collection(db, COL_ORDENES)),
      getCalendarioMap(db),
    ]);

    let ordenes = snapOrdenes.docs.map(function(d) { return Object.assign({ id: d.id }, d.data()); });
    if (isCampo) {
      ordenes = ordenes.filter(function(o) { return o.pareja === destino && o.estadoCampo !== 'hecha'; });
    }
    const conCoords = ordenes.filter(function(o) { return o.latitud && o.longitud; });

    if (!conCoords.length) {
      content.innerHTML =
        '<div class="text-center py-12 space-y-2">' +
          '<p class="text-gray-400 text-sm">Sin coordenadas disponibles</p>' +
          '<p class="text-xs text-gray-300">Incluye columnas Latitud y Longitud en el Excel</p>' +
        '</div>';
      return;
    }

    // Full-screen map layout with bottom sheet
    // Add spin keyframe for button spinner
    if (!document.getElementById('cambios-spin-style')) {
      const s = document.createElement('style');
      s.id = 'cambios-spin-style';
      s.textContent = '@keyframes spin { to { transform: rotate(360deg); } }';
      document.head.appendChild(s);
    }

    content.innerHTML =
      '<div style="position:relative;width:100%;height:calc(100vh - 220px);min-height:400px;">' +
        // Map container
        '<div id="mapa-contenedor" style="width:100%;height:100%;border-radius:12px;overflow:hidden;border:1px solid #e5e7eb;">' +
          '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#9ca3af;font-size:14px;gap:8px">' +
            '<svg style="width:16px;height:16px;animation:spin 1s linear infinite" viewBox="0 0 24 24" fill="none"><circle opacity=".25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path opacity=".75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/></svg>' +
            'Cargando mapa...' +
          '</div>' +
        '</div>' +
        // Bottom sheet — hidden by default
        '<div id="mapa-sheet" style="position:absolute;bottom:0;left:0;right:0;background:white;border-radius:16px 16px 0 0;box-shadow:0 -4px 24px rgba(0,0,0,0.15);transform:translateY(100%);transition:transform 0.3s ease;z-index:10;max-height:70%;overflow-y:auto;">' +
          '<div style="display:flex;justify-content:center;padding:10px 0 4px">' +
            '<div style="width:36px;height:4px;background:#e5e7eb;border-radius:2px"></div>' +
          '</div>' +
          '<div id="mapa-sheet-content" style="padding:0 16px 24px"></div>' +
        '</div>' +
      '</div>';

    await loadGoogleMaps();
    initMapaCambios(conCoords, calendarioMap, session, isCampo, db);

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

function initMapaCambios(ordenes, calendarioMap, session, isCampo, db) {
  const contenedor = document.getElementById('mapa-contenedor');
  const sheet      = document.getElementById('mapa-sheet');
  const sheetBody  = document.getElementById('mapa-sheet-content');
  if (!contenedor) return;

  const G   = google.maps;
  const map = new G.Map(contenedor, {
    zoom: 13,
    center: { lat: safeNum(ordenes[0].latitud), lng: safeNum(ordenes[0].longitud) },
    mapTypeId: 'hybrid',
    mapTypeControl: false,
    fullscreenControl: false,
    streetViewControl: false,
    zoomControlOptions: { position: G.ControlPosition.RIGHT_CENTER },
  });

  const directionsService  = new G.DirectionsService();
  const directionsRenderer = new G.DirectionsRenderer({
    map,
    suppressMarkers: true,
    polylineOptions: { strokeColor: '#0F766E', strokeWeight: 5, strokeOpacity: 0.85 },
  });

  let userMarker   = null;
  let activeMarker = null;
  let assignMode   = false;
  let selectedWOs  = new Set();
  let drawnShapes  = [];
  let drawingMgr   = null;
  let routeActive  = false;

  // ── helpers ──
  function getMarkerColor(o) {
    if (o.pareja && PAREJA_COLORS[o.pareja]) return PAREJA_COLORS[o.pareja];
    const bl = esBloqueada(o, calendarioMap);
    if (bl) return '#9CA3AF';
    if (o.estadoCampo === 'visita') return '#B45309';
    if (o.estadoCampo === 'hecha')  return '#166534';
    return '#6B7280';
  }

  function buildIcon(color, selected, inAssign) {
    return {
      path: G.SymbolPath.CIRCLE,
      scale: selected ? 13 : 10,
      fillColor: color,
      fillOpacity: 1,
      strokeColor: inAssign && selected ? '#FBBF24' : '#fff',
      strokeWeight: inAssign && selected ? 4 : selected ? 3 : 2,
    };
  }

  // ── close bottom sheet ──
  map.addListener('click', function() { closeSheet(); });

  function closeSheet() {
    if (sheet) sheet.style.transform = 'translateY(100%)';
    if (activeMarker) { activeMarker.setIcon(buildIcon(activeMarker.__color, false, assignMode)); activeMarker = null; }
  }

  function clearRoute() {
    directionsRenderer.setDirections({ routes: [] });
    routeActive = false;
    const cb = document.getElementById('sheet-cancel-ruta');
    if (cb) cb.style.display = 'none';
    const rb = document.getElementById('sheet-ruta');
    if (rb) { rb.disabled = false; rb.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="3 11 22 2 13 21 11 13 3 11"/></svg>Trazar ruta'; }
  }

  // ── open bottom sheet ──
  function openSheet(o, marker) {
    if (!sheet || !sheetBody) return;
    const bloqueada   = esBloqueada(o, calendarioMap);
    const hecha       = o.estadoCampo === 'hecha';
    const isAdminUser = !isCampo;
    const statusColor = bloqueada ? '#9CA3AF' : hecha ? '#166534' : o.estadoCampo === 'visita' ? '#B45309' : '#0F766E';
    const statusLabel = bloqueada ? '🔒 Bloqueada' : hecha ? '✓ Realizada' : o.estadoCampo === 'visita' ? '👁 Visita' : '● Disponible';

    function chip(label, val) {
      if (!val) return '';
      return '<div style="background:#f3f4f6;border-radius:8px;padding:6px 10px">' +
        '<p style="font-size:10px;color:#9ca3af;margin-bottom:2px">' + label + '</p>' +
        '<p style="font-size:12px;font-weight:600;color:#111827;word-break:break-word">' + safeStr(val) + '</p>' +
      '</div>';
    }

    sheetBody.innerHTML =
      '<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:8px;margin-bottom:12px">' +
        '<div style="flex:1;min-width:0">' +
          '<p style="font-size:10px;color:#9ca3af;font-family:monospace">' + safeStr(o.wo) + '</p>' +
          '<p style="font-size:16px;font-weight:800;color:#111827;margin-top:2px;line-height:1.25">' + safeStr(o.cliente) + '</p>' +
        '</div>' +
        '<span style="font-size:11px;font-weight:600;padding:4px 10px;border-radius:20px;background:' + statusColor + '18;color:' + statusColor + ';white-space:nowrap;flex-shrink:0">' + statusLabel + '</span>' +
      '</div>' +
      '<div style="display:flex;gap:8px;align-items:flex-start;background:#f9fafb;border-radius:10px;padding:9px 12px;margin-bottom:10px">' +
        '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;margin-top:1px"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>' +
        '<p style="font-size:12px;color:#374151;line-height:1.4">' + safeStr(o.direccion || '—') + '</p>' +
      '</div>' +
      '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:10px">' +
        chip('Medidor', o.serie) + chip('DS', o.dsct) +
        chip('MRU', o.unidadLectura) + chip('Concepto', o.concepto) +
        chip('NC', o.nc) + chip('Teléfono', o.telefono) +
        (o.pareja && isAdminUser ? chip('Pareja', o.pareja) : '') +
        (o.observaciones ? '<div style="grid-column:1/-1;background:#f3f4f6;border-radius:8px;padding:6px 10px"><p style="font-size:10px;color:#9ca3af;margin-bottom:2px">Observaciones</p><p style="font-size:12px;color:#374151">' + safeStr(o.observaciones) + '</p></div>' : '') +
        (o.observacion ? '<div style="grid-column:1/-1;background:#FEF3C7;border-radius:8px;padding:6px 10px"><p style="font-size:10px;color:#B45309;margin-bottom:2px">Nota visita</p><p style="font-size:12px;color:#374151">' + safeStr(o.observacion) + '</p></div>' : '') +
        (o.hechaPor ? '<div style="grid-column:1/-1;background:#F0FDF4;border-radius:8px;padding:6px 10px"><p style="font-size:10px;color:#166534;margin-bottom:2px">Hecha por</p><p style="font-size:12px;color:#374151">' + safeStr(o.hechaPor) + '</p></div>' : '') +
      '</div>' +
      (bloqueada ? '<div style="background:#FEF2F2;color:#C62828;padding:9px 12px;border-radius:10px;font-size:12px;font-weight:500;margin-bottom:10px">🔒 En período de lectura — no se puede ejecutar.</div>' : '') +
      '<div style="display:flex;flex-direction:column;gap:7px">' +
        '<div style="display:flex;gap:7px">' +
          '<button id="sheet-ruta" style="flex:1;padding:12px;background:#0F766E;color:white;border:none;border-radius:12px;font-size:13px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:5px">' +
            '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="3 11 22 2 13 21 11 13 3 11"/></svg>Trazar ruta' +
          '</button>' +
          '<a href="https://www.google.com/maps/dir/?api=1&destination=' + o.latitud + ',' + o.longitud + '" target="_blank" style="flex:1;padding:12px;border:1.5px solid #e5e7eb;border-radius:12px;font-size:13px;font-weight:500;color:#374151;text-align:center;text-decoration:none;display:flex;align-items:center;justify-content:center">↗ Maps</a>' +
        '</div>' +
        '<button id="sheet-cancel-ruta" style="width:100%;padding:9px;background:#FEF2F2;color:#C62828;border:1.5px solid #FECACA;border-radius:12px;font-size:13px;font-weight:600;cursor:pointer;display:' + (routeActive ? 'flex' : 'none') + ';align-items:center;justify-content:center">✕ Cancelar ruta</button>' +
        (!bloqueada && !hecha && isCampo ?
          '<div style="display:flex;gap:7px">' +
            '<button id="sheet-visita" style="flex:1;padding:11px;border:2px solid #e5e7eb;border-radius:12px;font-size:13px;font-weight:600;color:#374151;background:white;cursor:pointer">Visita</button>' +
            '<button id="sheet-hecha" style="flex:1;padding:11px;background:#166534;color:white;border:none;border-radius:12px;font-size:13px;font-weight:600;cursor:pointer">✓ Realizada</button>' +
          '</div>'
        : '') +
        (isAdminUser ? '<button id="sheet-asignar-1" style="width:100%;padding:10px;border:1.5px solid #e5e7eb;border-radius:12px;font-size:13px;font-weight:500;color:#374151;background:white;cursor:pointer">Asignar pareja</button>' : '') +
      '</div>';

    sheet.style.transform = 'translateY(0)';
    if (activeMarker && activeMarker !== marker) { activeMarker.setIcon(buildIcon(activeMarker.__color, false, false)); }
    marker.setIcon(buildIcon(marker.__color, true, false));
    activeMarker = marker;
    // Store marker ref on orden for hiding after realizada
    marker.__orden.__marker = marker;

    document.getElementById('sheet-ruta')?.addEventListener('click', function() { trazarRuta(safeNum(o.latitud), safeNum(o.longitud)); });
    document.getElementById('sheet-cancel-ruta')?.addEventListener('click', clearRoute);
    document.getElementById('sheet-hecha')?.addEventListener('click', function() { closeSheet(); showConfirmarHecha(db, session, o); });
    document.getElementById('sheet-visita')?.addEventListener('click', function() { closeSheet(); showRegistrarVisita(db, session, o); });
    document.getElementById('sheet-asignar-1')?.addEventListener('click', function() { closeSheet(); showAsignarPareja(db, [o.wo], null); });
  }

  // ── build markers ──
  const markers = ordenes.map(function(o) {
    const color  = getMarkerColor(o);
    const marker = new G.Marker({ position: { lat: safeNum(o.latitud), lng: safeNum(o.longitud) }, title: o.cliente, icon: buildIcon(color, false, false) });
    marker.__color  = color;
    marker.__orden  = o;
    marker.__sel    = false;
    marker.addListener('click', function() {
      if (assignMode) { toggleSel(marker); } else { openSheet(o, marker); }
    });
    return marker;
  });

  // ── clustering ──
  if (window.markerClusterer && window.markerClusterer.MarkerClusterer) {
    new window.markerClusterer.MarkerClusterer({
      map, markers,
      algorithmOptions: { maxZoom: 15 },
      renderer: {
        render: function(cluster) {
          const count = cluster.count;
          const size  = count > 50 ? 48 : count > 10 ? 40 : 34;
          const bg    = count > 50 ? '#0F766E' : count > 10 ? '#0d9488' : '#14b8a6';
          return new G.Marker({
            position: cluster.position,
            icon: { url: 'data:image/svg+xml;charset=UTF-8,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="' + size + '" height="' + size + '"><circle cx="' + (size/2) + '" cy="' + (size/2) + '" r="' + (size/2-2) + '" fill="' + bg + '" stroke="white" stroke-width="2"/><text x="50%" y="50%" text-anchor="middle" dy=".35em" fill="white" font-size="' + (size > 40 ? 14 : 12) + '" font-family="Inter,sans-serif" font-weight="700">' + count + '</text></svg>'),
              scaledSize: new G.Size(size, size), anchor: new G.Point(size/2, size/2) },
            zIndex: 1000,
          });
        },
      },
    });
  } else {
    markers.forEach(function(m) { m.setMap(map); });
  }

  // ── user location ──
  if (navigator.geolocation) {
    navigator.geolocation.getCurrentPosition(function(pos) {
      userMarker = new G.Marker({ position: { lat: pos.coords.latitude, lng: pos.coords.longitude }, map, title: 'Tu ubicación', zIndex: 999, icon: { path: G.SymbolPath.CIRCLE, scale: 9, fillColor: '#2563EB', fillOpacity: 1, strokeColor: '#fff', strokeWeight: 2.5 } });
    }, function() {});
  }

  // ── assign mode ──
  function toggleSel(marker) {
    const wo = marker.__orden.wo;
    if (selectedWOs.has(wo)) { selectedWOs.delete(wo); marker.__sel = false; }
    else { selectedWOs.add(wo); marker.__sel = true; }
    marker.setIcon(buildIcon(marker.__color, marker.__sel, true));
    updateAssignPanel();
  }

  function updateAssignPanel() {
    const el = document.getElementById('assign-count');
    if (el) el.textContent = selectedWOs.size + ' orden' + (selectedWOs.size !== 1 ? 'es' : '') + ' seleccionada' + (selectedWOs.size !== 1 ? 's' : '');
  }

  function clearSel() {
    selectedWOs.clear();
    drawnShapes.forEach(function(s) { s.setMap(null); }); drawnShapes = [];
    markers.forEach(function(m) { m.__sel = false; m.setIcon(buildIcon(m.__color, false, assignMode)); });
    if (drawingMgr) drawingMgr.setDrawingMode(null);
    updateAssignPanel();
  }

  function enterAssignMode() {
    assignMode = true; closeSheet();
    const p = document.getElementById('assign-panel');
    if (p) p.style.display = 'flex';
    const btn = document.getElementById('btn-assign-mode');
    if (btn) { btn.style.background = '#1B4F8A'; btn.style.color = 'white'; btn.textContent = '✕ Salir asignación'; }
    markers.forEach(function(m) { m.setIcon(buildIcon(m.__color, m.__sel, true)); });
    if (!drawingMgr && G.drawing) {
      drawingMgr = new G.drawing.DrawingManager({
        drawingMode: null, drawingControl: false,
        rectangleOptions: { fillColor:'#FBBF24', fillOpacity:0.15, strokeColor:'#FBBF24', strokeWeight:2, clickable:false, editable:false },
        polygonOptions:   { fillColor:'#FBBF24', fillOpacity:0.15, strokeColor:'#FBBF24', strokeWeight:2, clickable:false, editable:false },
      });
      drawingMgr.setMap(map);
      G.event.addListener(drawingMgr, 'overlaycomplete', function(e) {
        drawnShapes.push(e.overlay);
        drawingMgr.setDrawingMode(null);
        markers.forEach(function(m) {
          const pos = m.getPosition();
          let inside = false;
          if (e.type === 'rectangle') inside = e.overlay.getBounds().contains(pos);
          else if (e.type === 'polygon' && G.geometry) inside = G.geometry.poly.containsLocation(pos, e.overlay);
          if (inside) { selectedWOs.add(m.__orden.wo); m.__sel = true; m.setIcon(buildIcon(m.__color, true, true)); }
        });
        updateAssignPanel();
      });
    }
  }

  function exitAssignMode() {
    assignMode = false; clearSel();
    const p = document.getElementById('assign-panel');
    if (p) p.style.display = 'none';
    const btn = document.getElementById('btn-assign-mode');
    if (btn) { btn.style.background = 'white'; btn.style.color = '#1B4F8A'; btn.textContent = '🗂 Asignar'; }
    markers.forEach(function(m) { m.setIcon(buildIcon(m.__color, false, false)); });
  }

  async function doAssign(pareja) {
    if (!selectedWOs.size) { showToast('Selecciona al menos una orden.','error'); return; }
    const btn = document.getElementById('ap-btn-' + pareja.replace(' ','_'));
    if (btn) { btn.textContent = '...'; btn.disabled = true; }
    try {
      const color = PAREJA_COLORS[pareja] || '#0F766E';
      await Promise.all(Array.from(selectedWOs).map(async function(wo) {
        const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo','==',wo)));
        if (!snap.empty) await updateDoc(snap.docs[0].ref, { pareja, asignadoEn: serverTimestamp() });
        const m = markers.find(function(mk) { return mk.__orden.wo === wo; });
        if (m) { m.__orden.pareja = pareja; m.__color = color; m.__sel = false; m.setIcon(buildIcon(color, false, true)); }
      }));
      showToast(selectedWOs.size + ' órdenes → ' + pareja, 'success');
      selectedWOs.clear();
      updateAssignPanel();
    } catch(e) { showToast('Error al asignar.','error'); console.error(e); }
    if (btn) { btn.textContent = pareja; btn.disabled = false; }
  }

  // Inject assign mode button inside map
  const btnAssign = document.createElement('button');
  btnAssign.id = 'btn-assign-mode';
  btnAssign.textContent = '🗂 Asignar';
  btnAssign.style.cssText = 'position:absolute;top:10px;left:10px;z-index:10;background:white;border:1.5px solid #e5e7eb;border-radius:10px;padding:8px 14px;font-size:13px;font-weight:600;color:#1B4F8A;box-shadow:0 2px 8px rgba(0,0,0,0.15);cursor:pointer;';
  contenedor.style.position = 'relative';
  contenedor.appendChild(btnAssign);
  btnAssign.addEventListener('click', function() { assignMode ? exitAssignMode() : enterAssignMode(); });

  // Inject assign panel below map inside wrapper
  const wrapper = contenedor.parentElement;
  const panelEl = document.createElement('div');
  panelEl.id = 'assign-panel';
  panelEl.style.cssText = 'display:none;flex-direction:column;gap:8px;background:white;border:1px solid #e5e7eb;border-radius:12px;padding:12px;margin-top:8px;';
  panelEl.innerHTML =
    '<div style="display:flex;align-items:center;justify-content:space-between">' +
      '<p id="assign-count" style="font-size:13px;font-weight:600;color:#374151">0 órdenes seleccionadas</p>' +
      '<button id="assign-clear" style="font-size:12px;color:#6b7280;background:none;border:1px solid #e5e7eb;border-radius:8px;padding:4px 10px;cursor:pointer">Limpiar</button>' +
    '</div>' +
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px">' +
      PAREJAS.map(function(p) {
        return '<button id="ap-btn-' + p.replace(' ','_') + '" data-pareja="' + p + '" style="padding:10px;background:' + PAREJA_COLORS[p] + ';color:white;border:none;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer">' + p + '</button>';
      }).join('') +
    '</div>' +
    '<div style="display:flex;gap:6px">' +
      '<button id="assign-rect" style="flex:1;padding:9px;border:1.5px solid #e5e7eb;border-radius:10px;font-size:12px;color:#374151;background:white;cursor:pointer">⬜ Rectángulo</button>' +
      '<button id="assign-poly" style="flex:1;padding:9px;border:1.5px solid #e5e7eb;border-radius:10px;font-size:12px;color:#374151;background:white;cursor:pointer">✏️ Polígono</button>' +
    '</div>';
  wrapper.appendChild(panelEl);

  document.getElementById('assign-clear').addEventListener('click', clearSel);
  document.getElementById('assign-rect').addEventListener('click', function() { if (drawingMgr) drawingMgr.setDrawingMode(G.drawing.OverlayType.RECTANGLE); });
  document.getElementById('assign-poly').addEventListener('click', function() { if (drawingMgr) drawingMgr.setDrawingMode(G.drawing.OverlayType.POLYGON); });
  panelEl.querySelectorAll('[data-pareja]').forEach(function(btn) { btn.addEventListener('click', function() { doAssign(btn.dataset.pareja); }); });

  // Hide assign controls for campo
  if (isCampo) { btnAssign.style.display = 'none'; panelEl.style.display = 'none'; }

  // ── route ──
  function trazarRuta(lat, lng) {
    const btn = document.getElementById('sheet-ruta');
    if (!navigator.geolocation) { window.open('https://www.google.com/maps/dir/?api=1&destination=' + lat + ',' + lng, '_blank'); return; }
    if (btn) { btn.innerHTML = '<svg style="animation:spin .8s linear infinite;width:14px;height:14px" viewBox="0 0 24 24" fill="none"><circle opacity=".25" cx="12" cy="12" r="10" stroke="white" stroke-width="4"/><path opacity=".75" fill="white" d="M4 12a8 8 0 018-8v8z"/></svg> Calculando...'; btn.disabled = true; }
    navigator.geolocation.getCurrentPosition(function(pos) {
      const origin = { lat: pos.coords.latitude, lng: pos.coords.longitude };
      const dest   = { lat: lat, lng: lng };
      if (userMarker) userMarker.setPosition(origin);
      directionsService.route({ origin, destination: dest, travelMode: G.TravelMode.DRIVING }, function(result, status) {
        if (btn) btn.disabled = false;
        if (status === 'OK') {
          directionsRenderer.setDirections(result);
          routeActive = true;
          const leg = result.routes[0].legs[0];
          if (btn) btn.innerHTML = '✓ ' + leg.distance.text + ' · ' + leg.duration.text;
          const cb = document.getElementById('sheet-cancel-ruta'); if (cb) cb.style.display = 'flex';
          const bounds = new G.LatLngBounds();
          result.routes[0].overview_path.forEach(function(p) { bounds.extend(p); });
          map.fitBounds(bounds);
        } else {
          if (btn) btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="3 11 22 2 13 21 11 13 3 11"/></svg>Trazar ruta';
          window.open('https://www.google.com/maps/dir/?api=1&destination=' + lat + ',' + lng, '_blank');
        }
      });
    }, function() {
      if (btn) { btn.disabled = false; btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="3 11 22 2 13 21 11 13 3 11"/></svg>Trazar ruta'; }
      window.open('https://www.google.com/maps/dir/?api=1&destination=' + lat + ',' + lng, '_blank');
    });
  }

  window.__trazarRuta = trazarRuta;

  // spin keyframe
  if (!document.getElementById('cambios-spin-style')) {
    const s = document.createElement('style'); s.id = 'cambios-spin-style';
    s.textContent = '@keyframes spin { to { transform: rotate(360deg); } }';
    document.head.appendChild(s);
  }
}


// ─────────────────────────────────────────
// DETALLE DE ORDEN
// ─────────────────────────────────────────
function showDetalleOrden(db, session, orden, isCampo, calendarioMap) {
  const bloqueada = esBloqueada(orden, calendarioMap);
  const hecha     = orden.estadoCampo === 'hecha';
  const isAdmin   = ['admin', 'coordinadora'].includes(session.role);

  const mapsUrl = orden.latitud && orden.longitud
    ? 'https://www.google.com/maps/dir/?api=1&destination=' + orden.latitud + ',' + orden.longitud
    : 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(safeStr(orden.direccion));

  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<div>' +
        '<p class="text-xs font-mono text-gray-400">' + safeStr(orden.wo) + '</p>' +
        '<h2 class="font-semibold text-gray-900 leading-tight">' + safeStr(orden.cliente) + '</h2>' +
      '</div>' +
      '<button id="det-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-4">' +
      // Bloqueo
      (bloqueada ?
        '<div class="rounded-xl px-4 py-3 text-sm font-medium" style="background:#FEF2F2;color:#C62828">🔒 Esta orden está en período de lectura y no puede ejecutarse.</div>'
      : '') +
      // Datos
      '<div class="space-y-2 text-sm">' +
        row('Dirección', orden.direccion) +
        row('NC', orden.nc) +
        row('Serie', orden.serie) +
        row('DS', orden.dsct) +
        row('Concepto', orden.concepto) +
        row('MRU', orden.unidadLectura) +
        row('Teléfono', orden.telefono) +
        row('Observaciones', orden.observaciones) +
        (orden.pareja ? row('Pareja', orden.pareja) : '') +
        (orden.estadoCampo ? row('Estado', orden.estadoCampo + (orden.actualizadaDelsur ? ' · Actualizada' : orden.estadoCampo === 'hecha' ? ' · ⚠ Sin actualizar en Delsur' : '')) : '') +
        (orden.observacion ? row('Observación', orden.observacion) : '') +
        (orden.fechaHecha ? row('Fecha hecha', fmtDate(orden.fechaHecha)) : '') +
        (orden.hechaPor ? row('Registrada por', orden.hechaPor) : '') +
      '</div>' +
    '</div>' +
    '<div class="px-5 py-4 border-t border-gray-100 space-y-2 shrink-0">' +
      // Cómo llegar — siempre visible
      '<a href="' + mapsUrl + '" target="_blank" class="flex items-center justify-center gap-2 w-full border border-gray-300 text-gray-700 font-medium rounded-xl py-2.5 text-sm hover:bg-gray-50">' +
        '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="3 11 22 2 13 21 11 13 3 11"/></svg>' +
        'Cómo llegar' +
      '</a>' +
      // Acciones campo
      (!bloqueada && !hecha && isCampo ?
        '<div class="flex gap-2">' +
          '<button id="btn-visita" class="flex-1 border-2 border-gray-300 text-gray-700 font-semibold rounded-xl py-2.5 text-sm">Visita</button>' +
          '<button id="btn-hecha" class="flex-1 text-white font-semibold rounded-xl py-2.5 text-sm" style="background:#0F766E">✓ Realizada</button>' +
        '</div>'
      : '') +
      // Admin — asignar pareja
      (isAdmin && !isCampo ?
        '<button id="btn-asignar-1" class="w-full border border-gray-300 text-gray-600 font-medium rounded-xl py-2.5 text-sm hover:bg-gray-50">Asignar pareja</button>'
      : '') +
    '</div>'
  );

  ov.querySelector('#det-close').onclick = () => ov.remove();

  // Acción Hecha
  const btnHecha = ov.querySelector('#btn-hecha');
  if (btnHecha) {
    btnHecha.onclick = function() {
      ov.remove();
      showConfirmarHecha(db, session, orden);
    };
  }

  // Acción Visita
  const btnVisita = ov.querySelector('#btn-visita');
  if (btnVisita) {
    btnVisita.onclick = function() {
      ov.remove();
      showRegistrarVisita(db, session, orden);
    };
  }

  // Admin asignar
  const btnAsignar1 = ov.querySelector('#btn-asignar-1');
  if (btnAsignar1) {
    btnAsignar1.onclick = function() {
      ov.remove();
      showAsignarPareja(db, [orden.wo], null);
    };
  }
}

function row(label, value) {
  if (!value) return '';
  return '<div class="flex gap-2"><span class="text-gray-400 shrink-0 w-28">' + label + '</span><span class="text-gray-900 font-medium flex-1">' + safeStr(value) + '</span></div>';
}

// ─────────────────────────────────────────
// CONFIRMAR HECHA
// ─────────────────────────────────────────
function showConfirmarHecha(db, session, orden) {
  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<h2 class="font-semibold text-gray-900">Marcar como realizada</h2>' +
      '<button id="ch-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-4">' +
      '<p class="text-sm text-gray-600">WO <strong>' + safeStr(orden.wo) + '</strong> · ' + safeStr(orden.cliente) + '</p>' +
      '<div class="rounded-xl p-4 space-y-3" style="background:#F0FDFA">' +
        '<p class="text-sm font-semibold text-gray-800">¿Ya está actualizada en Delsur?</p>' +
        '<div class="flex gap-3">' +
          '<button id="ch-si" class="flex-1 py-3 rounded-xl font-bold text-white text-sm" style="background:#0F766E">Sí</button>' +
          '<button id="ch-no" class="flex-1 py-3 rounded-xl font-bold border-2 border-gray-300 text-gray-700 text-sm">No</button>' +
        '</div>' +
      '</div>' +
    '</div>'
  );

  ov.querySelector('#ch-close').onclick = () => ov.remove();

  async function guardarHecha(actualizada) {
    try {
      // Try direct update first, fallback to query by WO
      if (orden.id) {
        await updateDoc(doc(db, COL_ORDENES, orden.id), {
          estadoCampo: 'hecha', actualizadaDelsur: actualizada,
          fechaHecha: serverTimestamp(), hechaPor: session.displayName,
        });
      } else {
        const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo', '==', orden.wo)));
        if (snap.empty) throw new Error('Orden no encontrada');
        await updateDoc(snap.docs[0].ref, {
          estadoCampo: 'hecha', actualizadaDelsur: actualizada,
          fechaHecha: serverTimestamp(), hechaPor: session.displayName,
        });
      }
      ov.remove();
      showToast('Orden marcada como realizada.', 'success');
      // Hide marker from map if campo
      if (orden.__marker) { orden.__marker.setMap(null); }
    } catch(e) {
      showToast('Error al guardar.', 'error');
      console.error(e);
    }
  }

  ov.querySelector('#ch-si').onclick = () => guardarHecha(true);
  ov.querySelector('#ch-no').onclick = () => guardarHecha(false);
}

// ─────────────────────────────────────────
// REGISTRAR VISITA
// ─────────────────────────────────────────
function showRegistrarVisita(db, session, orden) {
  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<h2 class="font-semibold text-gray-900">Registrar visita</h2>' +
      '<button id="rv-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-3">' +
      '<p class="text-sm text-gray-600">WO <strong>' + safeStr(orden.wo) + '</strong> · ' + safeStr(orden.cliente) + '</p>' +
      '<div>' +
        '<label class="block text-xs font-medium text-gray-500 mb-1.5 uppercase tracking-wider">Motivo / Observación</label>' +
        '<textarea id="rv-obs" rows="3" placeholder="Ej. Cliente ausente, medidor inaccesible..." class="w-full border border-gray-200 rounded-xl px-3 py-2.5 text-sm resize-none focus:outline-none focus:ring-2" style="--tw-ring-color:#0F766E"></textarea>' +
      '</div>' +
    '</div>' +
    '<div class="px-5 py-4 border-t border-gray-100 flex gap-2 shrink-0">' +
      '<button id="rv-cancel" class="flex-1 border border-gray-200 text-gray-600 rounded-xl py-2.5 text-sm font-medium">Cancelar</button>' +
      '<button id="rv-save" class="flex-1 text-white rounded-xl py-2.5 text-sm font-semibold" style="background:#0F766E">Guardar visita</button>' +
    '</div>'
  );

  ov.querySelector('#rv-close').onclick = ov.querySelector('#rv-cancel').onclick = () => ov.remove();
  ov.querySelector('#rv-save').onclick = async function() {
    const obs = ov.querySelector('#rv-obs').value.trim();
    try {
      let ref;
      if (orden.id) {
        ref = doc(db, COL_ORDENES, orden.id);
      } else {
        const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo', '==', orden.wo)));
        if (snap.empty) throw new Error('Orden no encontrada');
        ref = snap.docs[0].ref;
      }
      await updateDoc(ref, {
        estadoCampo: 'visita', observacion: obs || 'Sin observación',
        fechaVisita: serverTimestamp(), visitadoPor: session.displayName,
      });
      ov.remove();
      showToast('Visita registrada.', 'success');
    } catch(e) {
      showToast('Error al guardar.', 'error');
      console.error(e);
    }
  };
}

// ─────────────────────────────────────────
// ASIGNAR PAREJA
// ─────────────────────────────────────────
function showAsignarPareja(db, wos, onDone) {
  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<div>' +
        '<h2 class="font-semibold text-gray-900">Asignar pareja</h2>' +
        '<p class="text-xs text-gray-400 mt-0.5">' + wos.length + ' orden(es) seleccionada(s)</p>' +
      '</div>' +
      '<button id="ap-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-2">' +
      PAREJAS.map(function(p) {
        return '<button data-pareja="' + p + '" class="w-full flex items-center justify-between px-4 py-3.5 rounded-xl border-2 border-gray-200 hover:border-blue-300 text-left transition-all">' +
          '<span class="font-semibold text-gray-900">' + p + '</span>' +
          '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#1B4F8A" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>' +
        '</button>';
      }).join('') +
    '</div>'
  );

  ov.querySelector('#ap-close').onclick = () => ov.remove();

  ov.querySelectorAll('[data-pareja]').forEach(function(btn) {
    btn.addEventListener('click', async function() {
      const pareja = btn.dataset.pareja;
      btn.textContent = 'Asignando...';
      btn.disabled = true;
      try {
        await Promise.all(wos.map(async function(wo) {
          const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo', '==', wo)));
          if (!snap.empty) {
            await updateDoc(snap.docs[0].ref, { pareja, asignadoEn: serverTimestamp() });
          }
        }));
        ov.remove();
        showToast(wos.length + ' orden(es) asignada(s) a ' + pareja, 'success');
        if (onDone) onDone();
      } catch(e) {
        showToast('Error al asignar.', 'error');
        console.error(e);
      }
    });
  });
}

// ─────────────────────────────────────────
// SEGUIMIENTO ADMIN
// ─────────────────────────────────────────
async function showSeguimiento(db, session) {
  const content = document.getElementById('cambios-content');
  if (!content) return;
  content.innerHTML = loading();

  try {
    const [snapOrdenes, calendarioMap] = await Promise.all([
      getDocs(collection(db, COL_ORDENES)),
      getCalendarioMap(db),
    ]);

    const ordenes = snapOrdenes.docs.map(function(d) { return Object.assign({ id: d.id }, d.data()); });
    const total        = ordenes.length;
    const hechas       = ordenes.filter(function(o) { return o.estadoCampo === 'hecha'; });
    const sinActualizar = hechas.filter(function(o) { return !o.actualizadaDelsur; });
    const visitas      = ordenes.filter(function(o) { return o.estadoCampo === 'visita'; });
    const bloqueadas   = ordenes.filter(function(o) { return esBloqueada(o, calendarioMap); });
    const pendientes   = ordenes.filter(function(o) { return !o.estadoCampo && !esBloqueada(o, calendarioMap); });

    // Por pareja
    const porPareja = {};
    PAREJAS.forEach(function(p) { porPareja[p] = { total: 0, hechas: 0, sinActualizar: 0 }; });
    ordenes.forEach(function(o) {
      if (o.pareja && porPareja[o.pareja]) {
        porPareja[o.pareja].total++;
        if (o.estadoCampo === 'hecha') porPareja[o.pareja].hechas++;
        if (o.estadoCampo === 'hecha' && !o.actualizadaDelsur) porPareja[o.pareja].sinActualizar++;
      }
    });

    function statCard(val, label, color) {
      return '<div class="bg-white rounded-xl border border-gray-200 p-4 text-center">' +
        '<p class="text-2xl font-black" style="color:' + color + '">' + val + '</p>' +
        '<p class="text-xs text-gray-500 mt-0.5">' + label + '</p>' +
      '</div>';
    }

    content.innerHTML =
      '<div class="space-y-4">' +
        '<div class="grid grid-cols-3 gap-2">' +
          statCard(total, 'Total', '#374151') +
          statCard(hechas.length, 'Hechas', '#0F766E') +
          statCard(pendientes.length, 'Pendientes', '#1B4F8A') +
        '</div>' +
        '<div class="grid grid-cols-3 gap-2">' +
          statCard(sinActualizar.length, 'Sin actualizar', '#C62828') +
          statCard(visitas.length, 'Visitas', '#B45309') +
          statCard(bloqueadas.length, 'Bloqueadas', '#6B7280') +
        '</div>' +
        // Por pareja
        '<div class="bg-white rounded-xl border border-gray-200 overflow-hidden">' +
          '<div class="px-4 py-3 border-b border-gray-100">' +
            '<p class="font-semibold text-sm text-gray-900">Por pareja</p>' +
          '</div>' +
          '<div class="divide-y divide-gray-50">' +
            PAREJAS.map(function(p) {
              const d = porPareja[p];
              const pct = d.total ? Math.round((d.hechas / d.total) * 100) : 0;
              return '<div class="px-4 py-3">' +
                '<div class="flex items-center justify-between mb-1.5">' +
                  '<p class="text-sm font-semibold text-gray-900">' + p + '</p>' +
                  '<p class="text-xs text-gray-400">' + d.hechas + '/' + d.total + ' · ' + (d.sinActualizar ? '<span style="color:#C62828">' + d.sinActualizar + ' sin actualizar</span>' : '✓') + '</p>' +
                '</div>' +
                '<div class="w-full bg-gray-100 rounded-full h-1.5">' +
                  '<div class="h-1.5 rounded-full transition-all" style="width:' + pct + '%;background:#0F766E"></div>' +
                '</div>' +
              '</div>';
            }).join('') +
          '</div>' +
        '</div>' +
        // Sin actualizar detail
        (sinActualizar.length > 0 ?
          '<div class="bg-white rounded-xl border border-gray-200 overflow-hidden">' +
            '<div class="px-4 py-3 border-b border-gray-100 flex items-center justify-between">' +
              '<p class="font-semibold text-sm text-gray-900">Sin actualizar en Delsur</p>' +
              '<span class="text-xs font-bold px-2 py-0.5 rounded-full text-white" style="background:#C62828">' + sinActualizar.length + '</span>' +
            '</div>' +
            '<div class="divide-y divide-gray-50">' +
              sinActualizar.map(function(o) {
                return '<div class="px-4 py-3">' +
                  '<p class="text-sm font-medium text-gray-900">' + safeStr(o.cliente) + '</p>' +
                  '<p class="text-xs text-gray-400 mt-0.5">' + safeStr(o.wo) + ' · ' + safeStr(o.pareja) + ' · Hecha: ' + fmtDate(o.fechaHecha) + '</p>' +
                '</div>';
              }).join('') +
            '</div>' +
          '</div>'
        : '') +
      '</div>';

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

// ─────────────────────────────────────────
// CARGAR ÓRDENES DESDE EXCEL
// ─────────────────────────────────────────
function showCargarOrdenes(db) {
  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<div>' +
        '<h2 class="font-semibold text-gray-900">Cargar órdenes</h2>' +
        '<p class="text-xs text-gray-400 mt-0.5">Excel con columnas: WO, NC, Cliente, Dirección, Serie, Marca, DSCT, Concepto, unidadLectura</p>' +
      '</div>' +
      '<button id="co-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-3">' +
      '<div class="border-2 border-dashed border-gray-200 rounded-xl p-6 text-center">' +
        '<svg class="w-10 h-10 mx-auto text-gray-300 mb-2" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>' +
        '<p class="text-sm text-gray-500 mb-3">Selecciona un archivo Excel (.xlsx)</p>' +
        '<input id="co-file" type="file" accept=".xlsx,.xls" class="hidden"/>' +
        '<label for="co-file" class="inline-block px-4 py-2 rounded-lg text-sm font-medium text-white cursor-pointer" style="background:#1B4F8A">Seleccionar archivo</label>' +
      '</div>' +
      '<div id="co-preview" class="hidden">' +
        '<p id="co-preview-text" class="text-sm text-gray-600 text-center"></p>' +
      '</div>' +
      '<div id="co-err" class="hidden text-sm text-red-500 bg-red-50 rounded-xl px-3 py-2"></div>' +
    '</div>' +
    '<div class="px-5 py-4 border-t border-gray-100 flex gap-2 shrink-0">' +
      '<button id="co-cancel" class="flex-1 border border-gray-200 text-gray-600 rounded-xl py-2.5 text-sm font-medium">Cancelar</button>' +
      '<button id="co-submit" class="flex-1 text-white rounded-xl py-2.5 text-sm font-semibold disabled:opacity-40" style="background:#1B4F8A" disabled>Importar</button>' +
    '</div>'
  );

  ov.querySelector('#co-close').onclick = ov.querySelector('#co-cancel').onclick = () => ov.remove();

  let parsedRows = [];

  ov.querySelector('#co-file').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (!file) return;
    const errEl = ov.querySelector('#co-err');
    errEl.classList.add('hidden');
    try {
      parsedRows = await parseExcelOrdenes(file);
      ov.querySelector('#co-preview-text').textContent = parsedRows.length + ' órdenes encontradas. Se omitirán duplicados por WO.';
      ov.querySelector('#co-preview').classList.remove('hidden');
      ov.querySelector('#co-submit').disabled = parsedRows.length === 0;
    } catch(err) {
      errEl.textContent = 'Error al leer el archivo: ' + err.message;
      errEl.classList.remove('hidden');
    }
  });

  ov.querySelector('#co-submit').addEventListener('click', async function() {
    const btn    = ov.querySelector('#co-submit');
    const errEl  = ov.querySelector('#co-err');
    btn.disabled = true;
    btn.textContent = 'Importando...';
    try {
      const snapExist = await getDocs(collection(db, COL_ORDENES));
      const existWOs  = new Set(snapExist.docs.map(function(d) { return d.data().wo; }));
      const nuevas    = parsedRows.filter(function(r) { return !existWOs.has(r.wo); });
      await Promise.all(nuevas.map(function(r) { return addDoc(collection(db, COL_ORDENES), r); }));
      ov.remove();
      showToast(nuevas.length + ' órdenes importadas. ' + (parsedRows.length - nuevas.length) + ' duplicados omitidos.', 'success');
    } catch(e) {
      errEl.textContent = 'Error al importar. Intenta de nuevo.';
      errEl.classList.remove('hidden');
      btn.disabled = false;
      btn.textContent = 'Importar';
      console.error(e);
    }
  });
}

async function parseExcelOrdenes(file) {
  return new Promise(function(resolve, reject) {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        import('https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs').then(function(XLSX) {
          const data  = new Uint8Array(e.target.result);
          const wb    = XLSX.read(data, { type: 'array' });
          const ws    = wb.Sheets[wb.SheetNames[0]];
          const rows  = XLSX.utils.sheet_to_json(ws, { defval: '' });
          const mapped = rows.map(function(r) {
            return {
              unidadLectura: safeStr(r['MRU']      || r['mru']      || r['Mru']),
              wo:            safeStr(r['WO']        || r['wo']),
              nc:            safeStr(r['NC']        || r['nc']),
              cliente:       safeStr(r['NOMBRE']    || r['Nombre']   || r['nombre']),
              direccion:     safeStr(r['DIRECCIÓN'] || r['DIRECCION']|| r['Dirección'] || r['Direccion']),
              dsct:          safeStr(r['DS']        || r['ds']),
              serie:         safeStr(r['MEDIDOR']   || r['Medidor']  || r['medidor']),
              latitud:       safeNum(r['LATITUD']   || r['Latitud']  || r['latitud']),
              longitud:      safeNum(r['LONGITUD']  || r['Longitud'] || r['longitud']),
              concepto:      safeStr(r['WO Class']  || r['wo class'] || r['WO_Class']),
              telefono:      safeStr(r['TELÉFONO']  || r['TELEFONO'] || r['Teléfono'] || r['telefono']),
              observaciones: safeStr(r['Lecturas y observaciones'] || r['Lecturas_y_observaciones'] || r['observaciones']),
              estadoCampo:   null,
              pareja:        null,
              cargadoEn:     new Date().toISOString(),
            };
          }).filter(function(r) { return r.wo; });
          resolve(mapped);
        }).catch(reject);
      } catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ─────────────────────────────────────────
// CARGAR CALENDARIO DE LECTURA
// ─────────────────────────────────────────
function showCargarCalendario(db) {
  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<div>' +
        '<h2 class="font-semibold text-gray-900">Calendario de lectura</h2>' +
        '<p class="text-xs text-gray-400 mt-0.5">Excel con columnas: unidadLectura, fechaLectura</p>' +
      '</div>' +
      '<button id="cal-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-3">' +
      '<div class="border-2 border-dashed border-gray-200 rounded-xl p-6 text-center">' +
        '<svg class="w-10 h-10 mx-auto text-gray-300 mb-2" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>' +
        '<p class="text-sm text-gray-500 mb-3">Selecciona el archivo Excel</p>' +
        '<input id="cal-file" type="file" accept=".xlsx,.xls" class="hidden"/>' +
        '<label for="cal-file" class="inline-block px-4 py-2 rounded-lg text-sm font-medium text-white cursor-pointer" style="background:#0F766E">Seleccionar archivo</label>' +
      '</div>' +
      '<div id="cal-preview" class="hidden">' +
        '<p id="cal-preview-text" class="text-sm text-gray-600 text-center"></p>' +
      '</div>' +
      '<div id="cal-err" class="hidden text-sm text-red-500 bg-red-50 rounded-xl px-3 py-2"></div>' +
    '</div>' +
    '<div class="px-5 py-4 border-t border-gray-100 flex gap-2 shrink-0">' +
      '<button id="cal-cancel" class="flex-1 border border-gray-200 text-gray-600 rounded-xl py-2.5 text-sm font-medium">Cancelar</button>' +
      '<button id="cal-submit" class="flex-1 text-white rounded-xl py-2.5 text-sm font-semibold disabled:opacity-40" style="background:#0F766E" disabled>Guardar calendario</button>' +
    '</div>'
  );

  ov.querySelector('#cal-close').onclick = ov.querySelector('#cal-cancel').onclick = () => ov.remove();

  let parsedCal = [];

  ov.querySelector('#cal-file').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (!file) return;
    const errEl = ov.querySelector('#cal-err');
    errEl.classList.add('hidden');
    try {
      parsedCal = await parseExcelCalendario(file);
      ov.querySelector('#cal-preview-text').textContent = parsedCal.length + ' unidades de lectura encontradas.';
      ov.querySelector('#cal-preview').classList.remove('hidden');
      ov.querySelector('#cal-submit').disabled = parsedCal.length === 0;
    } catch(err) {
      errEl.textContent = 'Error al leer el archivo: ' + err.message;
      errEl.classList.remove('hidden');
    }
  });

  ov.querySelector('#cal-submit').addEventListener('click', async function() {
    const btn   = ov.querySelector('#cal-submit');
    const errEl = ov.querySelector('#cal-err');
    btn.disabled = true;
    btn.textContent = 'Guardando...';
    try {
      // Overwrite calendar — use setDoc with unidad as key
      await Promise.all(parsedCal.map(function(r) {
        const id = r.unidadLectura.replace(/[^a-zA-Z0-9_]/g, '_');
        return setDoc(doc(db, COL_CALENDARIO, id), r);
      }));
      ov.remove();
      showToast(parsedCal.length + ' unidades de lectura guardadas.', 'success');
    } catch(e) {
      errEl.textContent = 'Error al guardar.';
      errEl.classList.remove('hidden');
      btn.disabled = false;
      btn.textContent = 'Guardar calendario';
      console.error(e);
    }
  });
}

async function parseExcelCalendario(file) {
  return new Promise(function(resolve, reject) {
    const reader = new FileReader();
    reader.onload = function(e) {
      import('https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs').then(function(XLSX) {
        try {
          const data = new Uint8Array(e.target.result);
          const wb   = XLSX.read(data, { type: 'array', cellDates: true });
          const ws   = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
          const mapped = rows.map(function(r) {
            const ul   = safeStr(r['MRU'] || r['mru'] || r['unidadLectura'] || r['UnidadLectura'] || r['UL']);
            const fl   = r['fechaLectura'] || r['FechaLectura'] || r['Fecha_Lectura'] || r['fecha_lectura'];
            const fecha = fl instanceof Date ? fl : new Date(fl);
            if (!ul || isNaN(fecha.getTime())) return null;
            return { unidadLectura: ul, fechaLectura: fecha.toISOString() };
          }).filter(Boolean);
          resolve(mapped);
        } catch(err) { reject(err); }
      }).catch(reject);
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}
