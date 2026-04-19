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
        (isAdmin ? '<button class="ctab flex-1 py-2 text-xs font-medium rounded-lg transition-colors" data-ctab="listado">Órdenes</button>' : '') +
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

    if (isCampo) {
      // ── CAMPO LISTADO ──
      const pendientes = ordenes.filter(function(o) { return o.estadoCampo !== 'hecha' && o.estadoCampo !== 'aprobada'; });
      const realizadas = ordenes.filter(function(o) { return o.estadoCampo === 'hecha' || o.estadoCampo === 'aprobada'; });

      function campoCard(o) {
        const bloqueada = esBloqueada(o, calendarioMap);
        const realizada = o.estadoCampo === 'hecha' || o.estadoCampo === 'aprobada';
        const visita    = o.estadoCampo === 'visita';
        const statusBadge = bloqueada
          ? '<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:#f3f4f6;color:#9ca3af">🔒 Bloqueada</span>'
          : realizada && !o.actualizadaDelsur
            ? '<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:#FEF3C7;color:#B45309">✓ Realizada · Sin actualizar ⚠</span>'
            : realizada
              ? '<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:#DCFCE7;color:#166534">✓ Realizada</span>'
              : visita
                ? '<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:#111827;color:white">👁 Visita</span>'
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
          (pendientes.length > 0 ?
            '<div>' +
              '<p style="font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">Por realizar (' + pendientes.length + ')</p>' +
              '<div class="space-y-2">' + pendientes.map(campoCard).join('') + '</div>' +
            '</div>'
          : '<div class="text-center py-6 bg-white rounded-xl border border-gray-200"><p class="text-sm font-semibold text-green-700">¡Todo realizado! 🎉</p></div>') +
          (realizadas.length > 0 ?
            '<div>' +
              '<p style="font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">Realizadas (' + realizadas.length + ')</p>' +
              '<div class="space-y-2">' + realizadas.map(campoCard).join('') + '</div>' +
            '</div>'
          : '') +
        '</div>';

      content.querySelectorAll('[data-wo]').forEach(function(card) {
        card.addEventListener('click', function() {
          const orden = ordenes.find(function(o) { return o.wo === card.dataset.wo; });
          if (orden) showDetalleOrden(db, session, orden, isCampo, calendarioMap);
        });
      });

    } else {
      // ── ADMIN LISTADO — Órdenes ──
      let busq = '';
      let filtroEstado = 'todas';
      let filtroPareja = 'todas';

      function ordenFiltrada(o) {
        const q = busq.toLowerCase();
        if (q && !safeStr(o.wo).toLowerCase().includes(q) && !safeStr(o.cliente).toLowerCase().includes(q) && !safeStr(o.direccion).toLowerCase().includes(q)) return false;
        if (filtroPareja !== 'todas' && o.pareja !== filtroPareja) return false;
        const bl = esBloqueada(o, calendarioMap);
        if (filtroEstado === 'disponibles') return !bl && !o.estadoCampo;
        if (filtroEstado === 'bloqueadas')  return bl;
        if (filtroEstado === 'realizadas')  return o.estadoCampo === 'hecha';
        if (filtroEstado === 'visitas')     return o.estadoCampo === 'visita';
        if (filtroEstado === 'sin_asignar') return !o.pareja;
        return true;
      }

      function adminOrderCard(o) {
        const bl    = esBloqueada(o, calendarioMap);
        const color = o.pareja ? (PAREJA_COLORS[o.pareja] || '#6B7280') : '#9CA3AF';
        const estado = bl ? 'Bloqueada' : o.estadoCampo === 'hecha' ? 'Realizada' : o.estadoCampo === 'visita' ? 'Visita' : o.estadoCampo === 'aprobada' ? 'Aprobada' : 'Disponible';
        const estadoColor = bl ? '#9CA3AF' : o.estadoCampo === 'hecha' ? '#B45309' : o.estadoCampo === 'visita' ? '#111827' : o.estadoCampo === 'aprobada' ? '#166534' : '#0F766E';
        const estadoBg    = bl ? '#f3f4f6' : o.estadoCampo === 'hecha' ? '#FEF3C7' : o.estadoCampo === 'visita' ? '#111827' : o.estadoCampo === 'aprobada' ? '#DCFCE7' : '#F0FDFA';
        const estadoTxt   = o.estadoCampo === 'visita' ? 'white' : estadoColor;

        return '<div class="bg-white rounded-xl border border-gray-200 px-4 py-3 cursor-pointer active:bg-gray-50" data-wo="' + o.wo + '">' +
          '<div class="flex items-start justify-between gap-2 mb-1.5">' +
            '<div class="min-w-0">' +
              '<p style="font-size:10px;color:#9ca3af;font-family:monospace">' + safeStr(o.wo) + '</p>' +
              '<p class="text-sm font-semibold text-gray-900 leading-tight mt-0.5">' + safeStr(o.cliente) + '</p>' +
            '</div>' +
            '<span style="font-size:10px;font-weight:600;padding:2px 8px;border-radius:20px;background:' + estadoBg + ';color:' + estadoTxt + ';flex-shrink:0">' + estado + '</span>' +
          '</div>' +
          '<p class="text-xs text-gray-500 truncate mb-1.5">' + safeStr(o.direccion) + '</p>' +
          '<div style="display:flex;align-items:center;gap:6px">' +
            (o.pareja ? '<span style="font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px;color:white;background:' + color + '">' + o.pareja + '</span>' : '<span style="font-size:11px;color:#9ca3af">Sin asignar</span>') +
          '</div>' +
        '</div>';
      }

      function renderAdminList() {
        const filtradas = ordenes.filter(ordenFiltrada);
        const lista = document.getElementById('admin-ordenes-lista');
        const count = document.getElementById('admin-ordenes-count');
        if (count) count.textContent = filtradas.length + ' órdenes';
        if (!lista) return;
        lista.innerHTML = filtradas.length
          ? filtradas.map(adminOrderCard).join('')
          : '<div class="text-center py-8 text-sm text-gray-400">Sin resultados</div>';

        lista.querySelectorAll('[data-wo]').forEach(function(card) {
          card.addEventListener('click', function() {
            const orden = ordenes.find(function(o) { return o.wo === card.dataset.wo; });
            if (orden) showDetalleOrden(db, session, orden, false, calendarioMap);
          });
        });

        // Long press to select for pareja assignment
        lista.querySelectorAll('[data-wo]').forEach(function(card) {
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

      content.innerHTML =
        '<div class="space-y-3">' +
          // Buscador
          '<div class="relative">' +
            '<svg class="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>' +
            '<input id="admin-buscar" type="text" placeholder="Buscar por WO, cliente, dirección..." class="w-full border border-gray-200 rounded-xl pl-9 pr-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>' +
          '</div>' +
          // Filtro estado
          '<div class="flex gap-1.5 overflow-x-auto pb-1 hide-scrollbar">' +
            [['todas','Todas'],['disponibles','Disponibles'],['realizadas','Realizadas'],['visitas','Visitas'],['bloqueadas','Bloqueadas'],['sin_asignar','Sin asignar']].map(function(f) {
              return '<button class="fest-btn text-xs font-semibold px-3 py-1.5 rounded-full border whitespace-nowrap flex-shrink-0 ' + (f[0] === 'todas' ? 'border-transparent text-white' : 'border-gray-200 text-gray-600') + '" style="' + (f[0] === 'todas' ? 'background:#1B4F8A' : '') + '" data-fest="' + f[0] + '">' + f[1] + '</button>';
            }).join('') +
          '</div>' +
          // Filtro pareja
          '<div class="flex gap-1.5 overflow-x-auto pb-1 hide-scrollbar">' +
            [['todas','Todas']].concat(PAREJAS.map(function(p) { return [p, p]; })).map(function(p) {
              const col = p[0] !== 'todas' ? PAREJA_COLORS[p[0]] : null;
              return '<button class="fpar-btn text-xs font-semibold px-3 py-1.5 rounded-full border whitespace-nowrap flex-shrink-0 ' + (p[0] === 'todas' ? 'border-gray-300 text-gray-600' : 'border-transparent text-white') + '" style="' + (col ? 'background:' + col : '') + '" data-fpar="' + p[0] + '">' + p[1] + '</button>';
            }).join('') +
          '</div>' +
          // Count + assign btn
          '<div class="flex items-center justify-between">' +
            '<p id="admin-ordenes-count" class="text-xs text-gray-400">' + ordenes.length + ' órdenes</p>' +
            '<button id="btn-asignar-sel" class="text-xs font-medium px-3 py-1.5 rounded-lg border border-gray-300 text-gray-600">Asignar seleccionadas</button>' +
          '</div>' +
          '<div id="admin-ordenes-lista" class="space-y-2"></div>' +
        '</div>';

      renderAdminList();

      // Wire search
      document.getElementById('admin-buscar').addEventListener('input', function(e) {
        busq = e.target.value.trim();
        renderAdminList();
      });

      // Wire estado filters
      content.querySelectorAll('.fest-btn').forEach(function(btn) {
        btn.addEventListener('click', function() {
          filtroEstado = btn.dataset.fest;
          content.querySelectorAll('.fest-btn').forEach(function(b) {
            b.style.background = b.dataset.fest === filtroEstado ? '#1B4F8A' : '';
            b.style.borderColor = b.dataset.fest === filtroEstado ? 'transparent' : '#e5e7eb';
            b.style.color = b.dataset.fest === filtroEstado ? 'white' : '#4b5563';
          });
          renderAdminList();
        });
      });

      // Wire pareja filters
      content.querySelectorAll('.fpar-btn').forEach(function(btn) {
        btn.addEventListener('click', function() {
          filtroPareja = btn.dataset.fpar;
          content.querySelectorAll('.fpar-btn').forEach(function(b) {
            const col = b.dataset.fpar !== 'todas' ? PAREJA_COLORS[b.dataset.fpar] : null;
            const activo = b.dataset.fpar === filtroPareja;
            b.style.background = activo ? (col || '#374151') : '';
            b.style.borderColor = activo ? 'transparent' : '#e5e7eb';
            b.style.color = activo ? 'white' : (col ? col : '#4b5563');
          });
          renderAdminList();
        });
      });

      // Assign selected
      document.getElementById('btn-asignar-sel').addEventListener('click', function() {
        const selCards = document.querySelectorAll('#admin-ordenes-lista [data-wo].selected');
        const wos = Array.from(selCards).map(function(c) { return c.dataset.wo; });
        if (!wos.length) { showToast('Mantén presionado para seleccionar órdenes.', 'error'); return; }
        showAsignarPareja(db, wos, function() { showListado(db, session, false, null); });
      });
    }

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

async function showSeguimiento(db, session) {
  const content = document.getElementById('cambios-content');
  if (!content) return;
  content.innerHTML = loading();

  const META_DIARIA = 15;

  try {
    const [snapOrdenes, calendarioMap] = await Promise.all([
      getDocs(collection(db, COL_ORDENES)),
      getCalendarioMap(db),
    ]);

    const ordenes = snapOrdenes.docs.map(function(d) { return Object.assign({ id: d.id }, d.data()); });

    // Today boundaries
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    const manana = new Date(hoy); manana.setDate(manana.getDate() + 1);

    function esHoy(ts) {
      if (!ts) return false;
      const d = ts.toDate ? ts.toDate() : new Date(ts);
      return d >= hoy && d < manana;
    }

    // Build pareja stats
    const stats = {};
    PAREJAS.forEach(function(p) {
      stats[p] = { total: 0, hechasHoy: 0, visitasHoy: 0, pendConfirm: [], hechasTotal: 0 };
    });

    ordenes.forEach(function(o) {
      if (!o.pareja || !stats[o.pareja]) return;
      if (o.estadoCampo !== 'aprobada') stats[o.pareja].total++;
      if (o.estadoCampo === 'hecha') {
        stats[o.pareja].hechasTotal++;
        if (esHoy(o.fechaHecha)) { stats[o.pareja].hechasHoy++; stats[o.pareja].pendConfirm.push(o); }
      }
      if (o.estadoCampo === 'visita' && esHoy(o.fechaVisita)) stats[o.pareja].visitasHoy++;
    });

    // Pending confirmation across all parejas
    const todasPendConfirm = ordenes.filter(function(o) { return o.estadoCampo === 'hecha'; });

    function parejaCard(p) {
      const s     = stats[p];
      const color = PAREJA_COLORS[p];
      const logros = s.hechasHoy + s.visitasHoy; // visitas count toward daily progress
      const pct    = Math.min(100, Math.round((logros / META_DIARIA) * 100));
      const barColor = pct >= 100 ? '#166534' : pct >= 60 ? '#0F766E' : pct >= 30 ? '#B45309' : '#C62828';

      return '<div class="bg-white rounded-xl border border-gray-200 p-4">' +
        // Header pareja
        '<div class="flex items-center justify-between mb-3">' +
          '<div class="flex items-center gap-2">' +
            '<span style="width:10px;height:10px;border-radius:50%;background:' + color + ';display:inline-block;flex-shrink:0"></span>' +
            '<p class="font-bold text-gray-900">' + p + '</p>' +
          '</div>' +
          '<span class="text-xs text-gray-400">' + s.total + ' asignadas</span>' +
        '</div>' +
        // Meta progress
        '<div class="mb-3">' +
          '<div class="flex items-center justify-between mb-1">' +
            '<p class="text-xs text-gray-500">Meta diaria</p>' +
            '<p class="text-xs font-bold" style="color:' + barColor + '">' + logros + ' / ' + META_DIARIA + '</p>' +
          '</div>' +
          '<div style="height:8px;background:#f3f4f6;border-radius:4px;overflow:hidden">' +
            '<div style="height:100%;width:' + pct + '%;background:' + barColor + ';border-radius:4px;transition:width .3s"></div>' +
          '</div>' +
        '</div>' +
        // Counts
        '<div class="grid grid-cols-3 gap-2 text-center">' +
          '<div style="background:#F0FDFA;border-radius:8px;padding:6px">' +
            '<p class="text-lg font-black" style="color:#0F766E">' + s.hechasHoy + '</p>' +
            '<p class="text-xs text-gray-400">Realizadas hoy</p>' +
          '</div>' +
          '<div style="background:#f9f9f9;border-radius:8px;padding:6px">' +
            '<p class="text-lg font-black text-gray-700">' + s.visitasHoy + '</p>' +
            '<p class="text-xs text-gray-400">Visitas hoy</p>' +
          '</div>' +
          '<div style="background:' + (s.pendConfirm.length ? '#FEF3C7' : '#F0FDFA') + ';border-radius:8px;padding:6px">' +
            '<p class="text-lg font-black" style="color:' + (s.pendConfirm.length ? '#B45309' : '#166534') + '">' + s.pendConfirm.length + '</p>' +
            '<p class="text-xs text-gray-400">Por confirmar</p>' +
          '</div>' +
        '</div>' +
      '</div>';
    }

    function confirmCard(o) {
      const color = o.pareja ? (PAREJA_COLORS[o.pareja] || '#6B7280') : '#6B7280';
      return '<div class="bg-white rounded-xl border ' + (!o.actualizadaDelsur ? 'border-yellow-300' : 'border-gray-200') + ' px-4 py-3 cursor-pointer active:bg-gray-50" data-wo="' + o.wo + '">' +
        '<div class="flex items-start justify-between gap-2 mb-1">' +
          '<div class="min-w-0">' +
            '<p style="font-size:10px;color:#9ca3af;font-family:monospace">' + safeStr(o.wo) + '</p>' +
            '<p class="text-sm font-semibold text-gray-900 leading-tight mt-0.5">' + safeStr(o.cliente) + '</p>' +
          '</div>' +
          (!o.actualizadaDelsur
            ? '<span style="font-size:10px;font-weight:600;padding:2px 8px;border-radius:20px;background:#FEF3C7;color:#B45309;white-space:nowrap">⚠ Sin actualizar</span>'
            : '<span style="font-size:10px;font-weight:600;padding:2px 8px;border-radius:20px;background:#DCFCE7;color:#166534;white-space:nowrap">✓ Actualizada</span>') +
        '</div>' +
        '<div style="display:flex;align-items:center;gap:6px">' +
          (o.pareja ? '<span style="font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px;color:white;background:' + color + '">' + o.pareja + '</span>' : '') +
          (o.hechaPor ? '<span style="font-size:11px;color:#6b7280">· ' + safeStr(o.hechaPor) + '</span>' : '') +
          (o.fechaHecha ? '<span style="font-size:11px;color:#9ca3af">· ' + fmtDate(o.fechaHecha) + '</span>' : '') +
        '</div>' +
      '</div>';
    }

    content.innerHTML =
      '<div class="space-y-4">' +
        // Meta global
        (function() {
          const totalHoy = Object.values(stats).reduce(function(acc, s) { return acc + s.hechasHoy + s.visitasHoy; }, 0);
          const metaTotal = META_DIARIA * PAREJAS.length;
          const pct = Math.min(100, Math.round((totalHoy / metaTotal) * 100));
          return '<div class="bg-white rounded-xl border border-gray-200 p-4">' +
            '<div class="flex items-center justify-between mb-2">' +
              '<p class="font-bold text-gray-900">Progreso global hoy</p>' +
              '<p class="text-sm font-black" style="color:#0F766E">' + totalHoy + ' / ' + metaTotal + '</p>' +
            '</div>' +
            '<div style="height:10px;background:#f3f4f6;border-radius:5px;overflow:hidden">' +
              '<div style="height:100%;width:' + pct + '%;background:#0F766E;border-radius:5px;transition:width .3s"></div>' +
            '</div>' +
            '<p style="font-size:11px;color:#9ca3af;margin-top:4px">Meta: ' + META_DIARIA + ' cambios por pareja · ' + metaTotal + ' total</p>' +
          '</div>';
        })() +
        // Por pareja
        '<p style="font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.05em">Por pareja</p>' +
        '<div class="space-y-3">' + PAREJAS.map(parejaCard).join('') + '</div>' +
        // Pendientes de confirmar
        (todasPendConfirm.length > 0 ?
          '<div>' +
            '<div class="flex items-center justify-between mb-2">' +
              '<p style="font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.05em">Pendientes de confirmación (' + todasPendConfirm.length + ')</p>' +
              '<button id="btn-reload-seg" style="padding:5px 10px;border:1px solid #e5e7eb;border-radius:8px;font-size:12px;color:#374151;background:white;cursor:pointer">↻</button>' +
            '</div>' +
            '<div class="space-y-2">' + todasPendConfirm.map(confirmCard).join('') + '</div>' +
          '</div>'
        : '<div class="text-center py-6 bg-white rounded-xl border border-gray-200"><p class="text-sm font-semibold text-green-700">Sin órdenes pendientes de confirmación ✓</p></div>') +
      '</div>';

    // Wire confirm cards
    content.querySelectorAll('[data-wo]').forEach(function(card) {
      card.addEventListener('click', function() {
        const orden = todasPendConfirm.find(function(o) { return o.wo === card.dataset.wo; }) ||
                      ordenes.find(function(o) { return o.wo === card.dataset.wo; });
        if (orden) showConfirmarOrdenAdmin(db, session, orden, calendarioMap);
      });
    });

    document.getElementById('btn-reload-seg')?.addEventListener('click', function() {
      showSeguimiento(db, session);
    });

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

function showConfirmarOrdenAdmin(db, session, orden, calendarioMap) {
  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<div>' +
        '<p class="text-xs font-mono text-gray-400">' + safeStr(orden.wo) + '</p>' +
        '<h2 class="font-semibold text-gray-900">' + safeStr(orden.cliente) + '</h2>' +
      '</div>' +
      '<button id="coa-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-3">' +
      '<div class="space-y-1.5 text-sm">' +
        (orden.direccion ? '<p class="text-gray-600">' + safeStr(orden.direccion) + '</p>' : '') +
        (orden.serie ? '<p class="text-gray-500">Medidor: <span class="font-mono font-semibold text-gray-800">' + safeStr(orden.serie) + '</span></p>' : '') +
        (orden.pareja ? '<p class="text-gray-500">Pareja: <span class="font-semibold" style="color:' + (PAREJA_COLORS[orden.pareja]||'#374151') + '">' + safeStr(orden.pareja) + '</span></p>' : '') +
        (orden.hechaPor ? '<p class="text-gray-500">Realizada por: <span class="font-semibold text-gray-800">' + safeStr(orden.hechaPor) + '</span></p>' : '') +
        (orden.fechaHecha ? '<p class="text-gray-500">Fecha: <span class="font-semibold text-gray-800">' + fmtDate(orden.fechaHecha) + '</span></p>' : '') +
      '</div>' +
      '<div class="rounded-xl p-4 space-y-1" style="background:#F0FDFA">' +
        '<p class="text-sm font-semibold text-gray-800">¿Confirmás que esta orden fue realizada?</p>' +
        '<p class="text-xs text-gray-500">Si confirmás, el punto desaparece del mapa permanentemente.</p>' +
      '</div>' +
    '</div>' +
    '<div class="px-5 py-4 border-t border-gray-100 flex gap-2 shrink-0">' +
      '<button id="coa-rechazar" class="flex-1 py-3 rounded-xl font-bold border-2 border-gray-300 text-gray-700 text-sm">✗ Rechazar</button>' +
      '<button id="coa-confirmar" class="flex-1 py-3 rounded-xl font-bold text-white text-sm" style="background:#166534">✓ Confirmar</button>' +
    '</div>'
  );

  ov.querySelector('#coa-close').onclick = () => ov.remove();

  ov.querySelector('#coa-confirmar').onclick = async function() {
    const btn = ov.querySelector('#coa-confirmar');
    btn.textContent = 'Confirmando...'; btn.disabled = true;
    try {
      let ref;
      if (orden.id) { ref = doc(db, COL_ORDENES, orden.id); }
      else {
        const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo','==',orden.wo)));
        if (snap.empty) throw new Error('No encontrada');
        ref = snap.docs[0].ref;
      }
      await updateDoc(ref, { estadoCampo: 'aprobada', aprobadoPor: session.displayName, fechaAprobacion: serverTimestamp() });
      ov.remove();
      showToast('Orden confirmada. Punto eliminado del mapa.', 'success');
      showSeguimiento(db, session);
    } catch(e) { showToast('Error.', 'error'); btn.textContent = '✓ Confirmar'; btn.disabled = false; console.error(e); }
  };

  ov.querySelector('#coa-rechazar').onclick = async function() {
    const btn = ov.querySelector('#coa-rechazar');
    btn.textContent = 'Rechazando...'; btn.disabled = true;
    try {
      let ref;
      if (orden.id) { ref = doc(db, COL_ORDENES, orden.id); }
      else {
        const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo','==',orden.wo)));
        if (snap.empty) throw new Error('No encontrada');
        ref = snap.docs[0].ref;
      }
      await updateDoc(ref, { estadoCampo: null, actualizadaDelsur: null, fechaHecha: null, hechaPor: null });
      ov.remove();
      showToast('Orden devuelta a campo como pendiente.', 'success');
      showSeguimiento(db, session);
    } catch(e) { showToast('Error.', 'error'); btn.textContent = '✗ Rechazar'; btn.disabled = false; console.error(e); }
  };
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
