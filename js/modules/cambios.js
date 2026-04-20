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
  if (isCampo) {
    const sArea = session.asignacionActual?.area || (session.usuarioOperativoAsignado ? 'OTC' : null);
    if (sArea && sArea !== 'CAMBIOS') {
      container.innerHTML =
        '<div class="flex flex-col items-center justify-center min-h-64 px-6 text-center space-y-4">' +
          '<div class="w-16 h-16 rounded-2xl flex items-center justify-center" style="background:#FEF2F2">' +
            '<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#C62828" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0110 0v4"/></svg>' +
          '</div>' +
          '<div>' +
            '<p class="font-bold text-gray-900 text-lg">Sin acceso a Cambios</p>' +
            '<p class="text-sm text-gray-500 mt-1">Estás asignado al área ' + (sArea || 'sin asignar') + '.</p>' +
            '<p class="text-sm text-gray-500">Contacta a administración para cambiar tu asignación.</p>' +
          '</div>' +
        '</div>';
      return;
    }
    if (!session.asignacionActual || session.asignacionActual.area !== 'CAMBIOS') {
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
        (isCampo ? '<button class="ctab flex-1 py-2 text-xs font-medium rounded-lg transition-colors" data-ctab="listado">Órdenes</button>' : '') +
        (isCampo ? '<button class="ctab flex-1 py-2 text-xs font-medium rounded-lg transition-colors" data-ctab="mapa">Mapa</button>' : '') +
        (isCampo ? '<button class="ctab flex-1 py-2 text-xs font-medium rounded-lg transition-colors" data-ctab="bodega">Bodega</button>' : '') +
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
      if (t === 'bodega')      await showBodegaCampo(db, session, destino);
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

        const sinActualizar = (o.estadoCampo === 'hecha') && !o.actualizadaDelsur;
        return '<div class="bg-white rounded-xl border ' + (sinActualizar ? 'border-yellow-300' : 'border-gray-200') + ' px-4 py-3 ' + (bloqueada ? 'opacity-60' : '') + '" data-wo="' + o.wo + '">' +
          '<div class="flex items-start justify-between gap-2 mb-1 cursor-pointer" data-wo-tap="' + o.wo + '">' +
            '<div class="min-w-0">' +
              '<p style="font-size:10px;color:#9ca3af;font-family:monospace">' + safeStr(o.wo) + '</p>' +
              '<p class="text-sm font-semibold text-gray-900 leading-tight mt-0.5">' + safeStr(o.cliente) + '</p>' +
            '</div>' +
            statusBadge +
          '</div>' +
          '<p class="text-xs text-gray-500 truncate mb-1.5" data-wo-tap="' + o.wo + '">' + safeStr(o.direccion) + '</p>' +
          (o.concepto ? '<span class="text-xs px-2 py-0.5 rounded-full bg-gray-100 text-gray-600 inline-block mb-1.5">' + safeStr(o.concepto) + '</span>' : '') +
          (sinActualizar ?
            '<button class="btn-actualizar w-full mt-1 py-2 rounded-xl text-xs font-bold border-2 border-yellow-400 text-yellow-700 bg-yellow-50" data-wo="' + o.wo + '" data-id="' + (o.id||'') + '">✓ Ya actualicé en Delsur</button>'
          : '') +
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

      // Wire tap on card (not on actualizar button)
      content.querySelectorAll('[data-wo-tap]').forEach(function(el) {
        el.addEventListener('click', function() {
          const orden = ordenes.find(function(o) { return o.wo === el.dataset.woTap; });
          if (orden) showDetalleOrden(db, session, orden, isCampo, calendarioMap);
        });
      });

      // Wire "ya actualicé" buttons
      content.querySelectorAll('.btn-actualizar').forEach(function(btn) {
        btn.addEventListener('click', async function(e) {
          e.stopPropagation();
          btn.textContent = 'Guardando...'; btn.disabled = true;
          try {
            const wo  = btn.dataset.wo;
            const id  = btn.dataset.id;
            let ref;
            if (id) { ref = doc(db, COL_ORDENES, id); }
            else { const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo','==',wo))); if (snap.empty) throw new Error(); ref = snap.docs[0].ref; }
            await updateDoc(ref, { actualizadaDelsur: true });
            showToast('Orden marcada como actualizada.', 'success');
            showListado(db, session, true, destino);
          } catch(e) { showToast('Error al guardar.','error'); btn.textContent = '✓ Ya actualicé en Delsur'; btn.disabled = false; }
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
            '<div style="display:flex;gap:6px">' +
              '<button id="btn-asignar-sel" class="text-xs font-medium px-3 py-1.5 rounded-lg border border-gray-300 text-gray-600">Asignar</button>' +
              '<button id="btn-desasignar-pareja" class="text-xs font-medium px-3 py-1.5 rounded-lg border border-red-200 text-red-600 bg-red-50">Desasignar pareja</button>' +
            '</div>' +
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

      // Desasignar pareja completa
      document.getElementById('btn-desasignar-pareja').addEventListener('click', function() {
        showDesasignarPareja(db, ordenes, function() { showListado(db, session, false, null); });
      });
    }

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

async function showSeguimiento(db, session) {
  const content = document.getElementById('cambios-content');
  if (!content) return;
  content.innerHTML = loading();

  const META_DIARIA = 15;

  function calcCorte() {
    const hoy   = new Date();
    const corte = new Date(hoy.getFullYear(), hoy.getMonth(), 20);
    if (hoy > corte) corte.setMonth(corte.getMonth() + 1);
    const diff  = Math.ceil((corte - hoy) / (1000*60*60*24));
    return { fecha: corte, dias: diff };
  }

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
        // Alerta de corte
        (function() {
          const { dias, fecha } = calcCorte();
          const sinAct = ordenes.filter(function(o) { return o.estadoCampo === 'hecha' && !o.actualizadaDelsur; });
          const fmtCorte = fecha.toLocaleDateString('es-SV', { day:'2-digit', month:'long' });
          const urgente  = dias <= 3;
          const alertColor = urgente ? '#C62828' : dias <= 7 ? '#B45309' : '#166534';
          const alertBg    = urgente ? '#FEF2F2' : dias <= 7 ? '#FEF3C7' : '#F0FDF4';
          const alertBorder= urgente ? '#FECACA' : dias <= 7 ? '#FDE68A' : '#BBF7D0';

          // Por pareja sin actualizar
          const sinActPorPareja = {};
          PAREJAS.forEach(function(p) { sinActPorPareja[p] = []; });
          sinAct.forEach(function(o) { if (o.pareja && sinActPorPareja[o.pareja]) sinActPorPareja[o.pareja].push(o); });

          return '<div style="background:' + alertBg + ';border:1.5px solid ' + alertBorder + ';border-radius:12px;padding:14px;margin-bottom:4px">' +
            '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">' +
              '<div>' +
                '<p style="font-size:13px;font-weight:700;color:' + alertColor + '">Corte: ' + fmtCorte + '</p>' +
                '<p style="font-size:11px;color:' + alertColor + ';opacity:.8">' + (dias === 0 ? '¡Hoy es el corte!' : dias === 1 ? 'Mañana es el corte' : 'Faltan ' + dias + ' días') + '</p>' +
              '</div>' +
              '<div style="text-align:right">' +
                '<p style="font-size:24px;font-weight:900;color:' + alertColor + ';line-height:1">' + sinAct.length + '</p>' +
                '<p style="font-size:10px;color:' + alertColor + ';opacity:.8">sin actualizar</p>' +
              '</div>' +
            '</div>' +
            (sinAct.length > 0 ?
              '<div style="display:flex;flex-direction:column;gap:4px">' +
                PAREJAS.filter(function(p) { return sinActPorPareja[p].length > 0; }).map(function(p) {
                  return '<div style="display:flex;align-items:center;justify-content:space-between;background:white;border-radius:8px;padding:6px 10px">' +
                    '<div style="display:flex;align-items:center;gap:6px">' +
                      '<span style="width:8px;height:8px;border-radius:50%;background:' + PAREJA_COLORS[p] + ';display:inline-block;flex-shrink:0"></span>' +
                      '<span style="font-size:12px;font-weight:600;color:#374151">' + p + '</span>' +
                    '</div>' +
                    '<span style="font-size:12px;font-weight:700;color:' + alertColor + '">' + sinActPorPareja[p].length + ' sin actualizar</span>' +
                  '</div>';
                }).join('') +
              '</div>'
            : '<p style="font-size:12px;color:' + alertColor + ';text-align:center;font-weight:600">✓ Todas las órdenes están actualizadas</p>') +
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

// ─────────────────────────────────────────
// MAPA
// ─────────────────────────────────────────
const GMAPS_KEY = 'AIzaSyAjaEXeu_4PDedaZhLfrwWvatu5RN9q1SU';

function loadGoogleMaps() {
  return new Promise(function(resolve) {
    if (window.google && window.google.maps) { resolve(); return; }
    window.__gmapsCallback = resolve;
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
      ordenes = ordenes.filter(function(o) { return o.pareja === destino && o.estadoCampo !== 'hecha' && o.estadoCampo !== 'aprobada'; });
    }
    const conCoords = ordenes.filter(function(o) { return o.latitud && o.longitud && o.estadoCampo !== 'aprobada'; });

    if (!conCoords.length) {
      content.innerHTML = '<div class="text-center py-12 space-y-2"><p class="text-gray-400 text-sm">Sin coordenadas disponibles</p><p class="text-xs text-gray-300">Incluye columnas Latitud y Longitud en el Excel</p></div>';
      return;
    }

    if (!document.getElementById('cambios-spin-style')) {
      const s = document.createElement('style'); s.id = 'cambios-spin-style';
      s.textContent = '@keyframes spin { to { transform: rotate(360deg); } }';
      document.head.appendChild(s);
    }

    content.innerHTML =
      '<div style="position:relative;width:100%;height:calc(100vh - 220px);min-height:400px;">' +
        '<div id="mapa-contenedor" style="width:100%;height:100%;border-radius:12px;overflow:hidden;border:1px solid #e5e7eb;">' +
          '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#9ca3af;font-size:14px;gap:8px">' +
            '<svg style="width:16px;height:16px;animation:spin 1s linear infinite" viewBox="0 0 24 24" fill="none"><circle opacity=".25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path opacity=".75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/></svg>Cargando mapa...' +
          '</div>' +
        '</div>' +
        '<div id="mapa-sheet" style="position:absolute;bottom:0;left:0;right:0;background:white;border-radius:16px 16px 0 0;box-shadow:0 -4px 24px rgba(0,0,0,0.15);transform:translateY(100%);transition:transform 0.3s ease;z-index:10;max-height:70%;overflow-y:auto;pointer-events:none;">' +
          '<div style="display:flex;align-items:center;justify-content:space-between;padding:10px 16px 4px">' +
            '<div style="width:36px;height:4px;background:#e5e7eb;border-radius:2px"></div>' +
            '<button id="sheet-close-btn" style="width:28px;height:28px;border-radius:50%;background:#f3f4f6;border:none;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:14px;color:#6b7280;line-height:1">✕</button>' +
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
    mapTypeControl: false, fullscreenControl: false, streetViewControl: false,
    zoomControlOptions: { position: G.ControlPosition.RIGHT_CENTER },
  });

  const directionsService  = new G.DirectionsService();
  const directionsRenderer = new G.DirectionsRenderer({ map, suppressMarkers: true, polylineOptions: { strokeColor: '#0F766E', strokeWeight: 5, strokeOpacity: 0.85 } });

  let userMarker = null, activeMarker = null, assignMode = false, selectedWOs = new Set(), drawnShapes = [], drawingMgr = null, routeActive = false;

  map.addListener('click', function() { closeSheet(); });

  function closeSheet() {
    if (sheet) { sheet.style.display = 'none'; sheet.style.pointerEvents = 'none'; }
    if (activeMarker) { activeMarker.setIcon(buildIcon(activeMarker.__color, false, assignMode)); activeMarker = null; }
  }

  function clearRoute() {
    directionsRenderer.setDirections({ routes: [] }); routeActive = false;
    const cb = document.getElementById('sheet-cancel-ruta'); if (cb) cb.style.display = 'none';
    const rb = document.getElementById('sheet-ruta');
    if (rb) { rb.disabled = false; rb.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="3 11 22 2 13 21 11 13 3 11"/></svg>Trazar ruta'; }
  }

  function getMarkerColor(o) {
    const bl = esBloqueada(o, calendarioMap);
    if (bl) return '#9CA3AF';
    if (o.estadoCampo === 'visita') return '#111827';
    if (o.pareja && PAREJA_COLORS[o.pareja]) return PAREJA_COLORS[o.pareja];
    return '#6B7280';
  }

  function buildIcon(color, selected, inAssign) {
    return { path: G.SymbolPath.CIRCLE, scale: selected ? 13 : 10, fillColor: color, fillOpacity: 1, strokeColor: inAssign && selected ? '#FBBF24' : '#fff', strokeWeight: inAssign && selected ? 4 : selected ? 3 : 2 };
  }

  function openSheet(o, marker) {
    if (!sheet || !sheetBody) return;
    const bloqueada = esBloqueada(o, calendarioMap);
    const hecha = o.estadoCampo === 'hecha' || o.estadoCampo === 'aprobada';
    const isAdminUser = !isCampo;
    const statusColor = bloqueada ? '#9CA3AF' : hecha ? '#166534' : o.estadoCampo === 'visita' ? '#374151' : '#0F766E';
    const statusLabel = bloqueada ? '🔒 Bloqueada' : hecha ? '✓ Realizada' : o.estadoCampo === 'visita' ? '👁 Visita' : '● Disponible';

    function chip(label, val) {
      if (!val) return '';
      return '<div style="background:#f3f4f6;border-radius:8px;padding:6px 10px"><p style="font-size:10px;color:#9ca3af;margin-bottom:2px">' + label + '</p><p style="font-size:12px;font-weight:600;color:#111827;word-break:break-word">' + safeStr(val) + '</p></div>';
    }

    sheetBody.innerHTML =
      '<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:8px;margin-bottom:12px">' +
        '<div style="flex:1;min-width:0"><p style="font-size:10px;color:#9ca3af;font-family:monospace">' + safeStr(o.wo) + '</p><p style="font-size:16px;font-weight:800;color:#111827;margin-top:2px;line-height:1.25">' + safeStr(o.cliente) + '</p></div>' +
        '<span style="font-size:11px;font-weight:600;padding:4px 10px;border-radius:20px;background:' + statusColor + '18;color:' + statusColor + ';white-space:nowrap;flex-shrink:0">' + statusLabel + '</span>' +
      '</div>' +
      '<div style="display:flex;gap:8px;align-items:flex-start;background:#f9fafb;border-radius:10px;padding:9px 12px;margin-bottom:10px">' +
        '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;margin-top:1px"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>' +
        '<p style="font-size:12px;color:#374151;line-height:1.4">' + safeStr(o.direccion || '—') + '</p>' +
      '</div>' +
      '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:10px">' +
        chip('Medidor', o.serie) + chip('DS', o.dsct) + chip('MRU', o.unidadLectura) + chip('Concepto', o.concepto) + chip('NC', o.nc) + chip('Teléfono', o.telefono) +
        (o.pareja && isAdminUser ? chip('Pareja', o.pareja) : '') +
        (o.observaciones ? '<div style="grid-column:1/-1;background:#f3f4f6;border-radius:8px;padding:6px 10px"><p style="font-size:10px;color:#9ca3af;margin-bottom:2px">Observaciones</p><p style="font-size:12px;color:#374151">' + safeStr(o.observaciones) + '</p></div>' : '') +
        (o.observacion ? '<div style="grid-column:1/-1;background:#FEF3C7;border-radius:8px;padding:6px 10px"><p style="font-size:10px;color:#B45309;margin-bottom:2px">Nota visita</p><p style="font-size:12px;color:#374151">' + safeStr(o.observacion) + '</p></div>' : '') +
      '</div>' +
      (bloqueada ? '<div style="background:#F3F4F6;color:#6B7280;padding:12px;border-radius:10px;font-size:13px;font-weight:500;margin-bottom:10px;text-align:center">🔒 Orden en período de lectura<br><span style="font-size:11px;font-weight:400">No se puede ejecutar en este momento</span></div>' : '') +
      '<div style="display:flex;flex-direction:column;gap:7px">' +
        '<div style="display:flex;gap:7px">' +
          '<button id="sheet-ruta" style="flex:1;padding:12px;background:#0F766E;color:white;border:none;border-radius:12px;font-size:13px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:5px"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="3 11 22 2 13 21 11 13 3 11"/></svg>Trazar ruta</button>' +
          '<a href="https://www.google.com/maps/dir/?api=1&destination=' + o.latitud + ',' + o.longitud + '" target="_blank" style="flex:1;padding:12px;border:1.5px solid #e5e7eb;border-radius:12px;font-size:13px;font-weight:500;color:#374151;text-align:center;text-decoration:none;display:flex;align-items:center;justify-content:center">↗ Maps</a>' +
        '</div>' +
        '<button id="sheet-cancel-ruta" style="width:100%;padding:9px;background:#FEF2F2;color:#C62828;border:1.5px solid #FECACA;border-radius:12px;font-size:13px;font-weight:600;cursor:pointer;display:' + (routeActive ? 'flex' : 'none') + ';align-items:center;justify-content:center">✕ Cancelar ruta</button>' +
        (!bloqueada && !hecha && isCampo ? '<div style="display:flex;gap:7px"><button id="sheet-visita" style="flex:1;padding:11px;border:2px solid #e5e7eb;border-radius:12px;font-size:13px;font-weight:600;color:#374151;background:white;cursor:pointer">Visita</button><button id="sheet-hecha" style="flex:1;padding:11px;background:#166534;color:white;border:none;border-radius:12px;font-size:13px;font-weight:600;cursor:pointer">✓ Realizada</button></div>' : '') +
        (isAdminUser ? '<button id="sheet-asignar-1" style="width:100%;padding:10px;border:1.5px solid #e5e7eb;border-radius:12px;font-size:13px;font-weight:500;color:#374151;background:white;cursor:pointer">Asignar pareja</button>' : '') +
        (isAdminUser && o.estadoCampo === 'hecha' ? '<button id="sheet-aprobar" style="width:100%;padding:11px;background:#166534;color:white;border:none;border-radius:12px;font-size:13px;font-weight:600;cursor:pointer">✓ Confirmar realizada</button>' : '') +
      '</div>';

    sheet.style.display = 'block';
    sheet.style.pointerEvents = 'auto';
    document.getElementById('sheet-close-btn')?.addEventListener('click', function(e) {
      e.stopPropagation();
      closeSheet();
    });
    if (activeMarker && activeMarker !== marker) { activeMarker.setIcon(buildIcon(activeMarker.__color, false, false)); }
    marker.setIcon(buildIcon(marker.__color, true, false));
    activeMarker = marker;
    marker.__orden.__marker = marker;

    document.getElementById('sheet-ruta')?.addEventListener('click', function() { trazarRuta(safeNum(o.latitud), safeNum(o.longitud)); });
    document.getElementById('sheet-cancel-ruta')?.addEventListener('click', clearRoute);
    document.getElementById('sheet-hecha')?.addEventListener('click', function() { closeSheet(); showConfirmarHecha(db, session, o); });
    document.getElementById('sheet-visita')?.addEventListener('click', function() { closeSheet(); showRegistrarVisita(db, session, o); });
    document.getElementById('sheet-asignar-1')?.addEventListener('click', function() { closeSheet(); showAsignarPareja(db, [o.wo], null); });
    document.getElementById('sheet-aprobar')?.addEventListener('click', async function() {
      const btn = document.getElementById('sheet-aprobar');
      if (btn) { btn.textContent = 'Confirmando...'; btn.disabled = true; }
      try {
        let ref;
        if (o.id) { ref = doc(db, COL_ORDENES, o.id); }
        else { const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo','==',o.wo))); if (snap.empty) throw new Error('No encontrada'); ref = snap.docs[0].ref; }
        await updateDoc(ref, { estadoCampo: 'aprobada', aprobadoPor: session.displayName, fechaAprobacion: serverTimestamp() });
        closeSheet();
        if (marker) marker.setMap(null);
        showToast('Orden confirmada. Punto eliminado.', 'success');
      } catch(e) { showToast('Error.','error'); if (btn) { btn.disabled=false; btn.textContent='✓ Confirmar realizada'; } }
    });
  }

  const markers = ordenes.map(function(o) {
    const color  = getMarkerColor(o);
    const marker = new G.Marker({ position: { lat: safeNum(o.latitud), lng: safeNum(o.longitud) }, title: o.cliente, icon: buildIcon(color, false, false) });
    marker.__color = color; marker.__orden = o; marker.__sel = false;
    marker.addListener('click', function() { assignMode ? toggleSel(marker) : openSheet(o, marker); });
    return marker;
  });

  markers.forEach(function(m) { m.setMap(map); });

  if (navigator.geolocation) {
    navigator.geolocation.getCurrentPosition(function(pos) {
      userMarker = new G.Marker({ position: { lat: pos.coords.latitude, lng: pos.coords.longitude }, map, title: 'Tu ubicación', zIndex: 999, icon: { path: G.SymbolPath.CIRCLE, scale: 9, fillColor: '#2563EB', fillOpacity: 1, strokeColor: '#fff', strokeWeight: 2.5 } });
    }, function() {});
  }

  function toggleSel(marker) {
    const wo = marker.__orden.wo;
    if (selectedWOs.has(wo)) { selectedWOs.delete(wo); marker.__sel = false; } else { selectedWOs.add(wo); marker.__sel = true; }
    marker.setIcon(buildIcon(marker.__color, marker.__sel, true));
    updateAssignPanel();
  }

  function updateAssignPanel() {
    const el = document.getElementById('assign-count');
    if (el) el.textContent = selectedWOs.size + ' orden' + (selectedWOs.size !== 1 ? 'es' : '') + ' seleccionada' + (selectedWOs.size !== 1 ? 's' : '');
  }

  function clearSel() {
    selectedWOs.clear(); drawnShapes.forEach(function(s) { s.setMap(null); }); drawnShapes = [];
    markers.forEach(function(m) { m.__sel = false; m.setIcon(buildIcon(m.__color, false, assignMode)); });
    if (drawingMgr) drawingMgr.setDrawingMode(null);
    updateAssignPanel();
  }

  function enterAssignMode() {
    assignMode = true; closeSheet();
    const p = document.getElementById('assign-panel'); if (p) p.style.display = 'flex';
    const btn = document.getElementById('btn-assign-mode'); if (btn) { btn.style.background = '#1B4F8A'; btn.style.color = 'white'; btn.textContent = '✕ Salir'; }
    markers.forEach(function(m) { m.setIcon(buildIcon(m.__color, m.__sel, true)); });
    if (!drawingMgr && G.drawing) {
      drawingMgr = new G.drawing.DrawingManager({ drawingMode: null, drawingControl: false, rectangleOptions: { fillColor:'#FBBF24', fillOpacity:0.15, strokeColor:'#FBBF24', strokeWeight:2, clickable:false, editable:false }, polygonOptions: { fillColor:'#FBBF24', fillOpacity:0.15, strokeColor:'#FBBF24', strokeWeight:2, clickable:false, editable:false } });
      drawingMgr.setMap(map);
      G.event.addListener(drawingMgr, 'overlaycomplete', function(e) {
        drawnShapes.push(e.overlay); drawingMgr.setDrawingMode(null);
        markers.forEach(function(m) {
          const pos = m.getPosition(); let inside = false;
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
    const p = document.getElementById('assign-panel'); if (p) p.style.display = 'none';
    const btn = document.getElementById('btn-assign-mode'); if (btn) { btn.style.background = 'white'; btn.style.color = '#1B4F8A'; btn.textContent = '🗂 Asignar'; }
    markers.forEach(function(m) { m.setIcon(buildIcon(m.__color, false, false)); });
  }

  async function doAssign(pareja) {
    if (!selectedWOs.size) { showToast('Selecciona al menos una orden.','error'); return; }
    const isDesasignar = pareja === null;
    const btnId = isDesasignar ? 'ap-btn-desasignar' : 'ap-btn-' + pareja.replace(' ','_');
    const btn = document.getElementById(btnId);
    if (btn) { btn.textContent = '...'; btn.disabled = true; }
    try {
      const color = isDesasignar ? '#6B7280' : (PAREJA_COLORS[pareja] || '#0F766E');
      await Promise.all(Array.from(selectedWOs).map(async function(wo) {
        const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo','==',wo)));
        if (!snap.empty) await updateDoc(snap.docs[0].ref, isDesasignar ? { pareja: null } : { pareja, asignadoEn: serverTimestamp() });
        const m = markers.find(function(mk) { return mk.__orden.wo === wo; });
        if (m) { m.__orden.pareja = isDesasignar ? null : pareja; m.__color = color; m.__sel = false; m.setIcon(buildIcon(color, false, true)); }
      }));
      showToast(isDesasignar ? selectedWOs.size + ' órdenes desasignadas' : selectedWOs.size + ' órdenes → ' + pareja, 'success');
      selectedWOs.clear(); updateAssignPanel();
    } catch(e) { showToast('Error.','error'); console.error(e); }
    if (btn) { btn.textContent = isDesasignar ? '✕ Desasignar seleccionadas' : pareja; btn.disabled = false; }
  }

  const btnAssign = document.createElement('button');
  btnAssign.id = 'btn-assign-mode';
  btnAssign.textContent = '🗂 Asignar';
  btnAssign.style.cssText = 'position:absolute;top:10px;left:10px;z-index:10;background:white;border:1.5px solid #e5e7eb;border-radius:10px;padding:8px 14px;font-size:13px;font-weight:600;color:#1B4F8A;box-shadow:0 2px 8px rgba(0,0,0,0.15);cursor:pointer;';
  contenedor.style.position = 'relative';
  contenedor.appendChild(btnAssign);
  btnAssign.addEventListener('click', function() { assignMode ? exitAssignMode() : enterAssignMode(); });

  const wrapper = contenedor.parentElement;
  const panelEl = document.createElement('div');
  panelEl.id = 'assign-panel';
  panelEl.style.cssText = 'display:none;flex-direction:column;gap:8px;background:white;border:1px solid #e5e7eb;border-radius:12px;padding:12px;margin-top:8px;';
  panelEl.innerHTML =
    '<div style="display:flex;align-items:center;justify-content:space-between"><p id="assign-count" style="font-size:13px;font-weight:600;color:#374151">0 órdenes seleccionadas</p><button id="assign-clear" style="font-size:12px;color:#6b7280;background:none;border:1px solid #e5e7eb;border-radius:8px;padding:4px 10px;cursor:pointer">Limpiar</button></div>' +
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px">' + PAREJAS.map(function(p) { return '<button id="ap-btn-' + p.replace(' ','_') + '" data-pareja="' + p + '" style="padding:10px;background:' + PAREJA_COLORS[p] + ';color:white;border:none;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer">' + p + '</button>'; }).join('') + '</div>' +
    '<button id="ap-btn-desasignar" style="width:100%;padding:9px;background:#FEF2F2;color:#C62828;border:1.5px solid #FECACA;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer">✕ Desasignar seleccionadas</button>' +
    '<div style="display:flex;gap:6px"><button id="assign-rect" style="flex:1;padding:9px;border:1.5px solid #e5e7eb;border-radius:10px;font-size:12px;color:#374151;background:white;cursor:pointer">⬜ Rectángulo</button><button id="assign-poly" style="flex:1;padding:9px;border:1.5px solid #e5e7eb;border-radius:10px;font-size:12px;color:#374151;background:white;cursor:pointer">✏️ Polígono</button></div>';
  wrapper.appendChild(panelEl);

  document.getElementById('assign-clear').addEventListener('click', clearSel);
  document.getElementById('assign-rect').addEventListener('click', function() { if (drawingMgr) drawingMgr.setDrawingMode(G.drawing.OverlayType.RECTANGLE); });
  document.getElementById('assign-poly').addEventListener('click', function() { if (drawingMgr) drawingMgr.setDrawingMode(G.drawing.OverlayType.POLYGON); });
  panelEl.querySelectorAll('[data-pareja]').forEach(function(btn) { btn.addEventListener('click', function() { doAssign(btn.dataset.pareja); }); });
  document.getElementById('ap-btn-desasignar')?.addEventListener('click', function() { doAssign(null); });

  if (isCampo) { btnAssign.style.display = 'none'; panelEl.style.display = 'none'; }

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
          directionsRenderer.setDirections(result); routeActive = true;
          const leg = result.routes[0].legs[0];
          if (btn) btn.innerHTML = '✓ ' + leg.distance.text + ' · ' + leg.duration.text;
          const cb = document.getElementById('sheet-cancel-ruta'); if (cb) cb.style.display = 'flex';
          const bounds = new G.LatLngBounds(); result.routes[0].overview_path.forEach(function(p) { bounds.extend(p); }); map.fitBounds(bounds);
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
}

// ─────────────────────────────────────────
// DETALLE DE ORDEN (listado overlay)
// ─────────────────────────────────────────
function showDetalleOrden(db, session, orden, isCampo, calendarioMap) {
  const bloqueada = esBloqueada(orden, calendarioMap);
  const hecha     = orden.estadoCampo === 'hecha' || orden.estadoCampo === 'aprobada';
  const isAdmin   = ['admin', 'coordinadora'].includes(session.role);

  const mapsUrl = orden.latitud && orden.longitud
    ? 'https://www.google.com/maps/dir/?api=1&destination=' + orden.latitud + ',' + orden.longitud
    : 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(safeStr(orden.direccion));

  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<div><p class="text-xs font-mono text-gray-400">' + safeStr(orden.wo) + '</p><h2 class="font-semibold text-gray-900 leading-tight">' + safeStr(orden.cliente) + '</h2></div>' +
      '<button id="det-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-3 text-sm">' +
      (bloqueada ? '<div class="rounded-xl px-4 py-3 font-medium text-center" style="background:#F3F4F6;color:#6B7280">🔒 Orden en período de lectura — no se puede ejecutar.</div>' : '') +
      '<div class="space-y-1.5">' +
        (orden.direccion ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">Dirección</span><span class="text-gray-900 font-medium">' + safeStr(orden.direccion) + '</span></div>' : '') +
        (orden.serie ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">Medidor</span><span class="font-mono font-semibold text-gray-900">' + safeStr(orden.serie) + '</span></div>' : '') +
        (orden.dsct ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">DS</span><span class="text-gray-900 font-medium">' + safeStr(orden.dsct) + '</span></div>' : '') +
        (orden.unidadLectura ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">MRU</span><span class="text-gray-900 font-medium">' + safeStr(orden.unidadLectura) + '</span></div>' : '') +
        (orden.concepto ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">Concepto</span><span class="text-gray-900 font-medium">' + safeStr(orden.concepto) + '</span></div>' : '') +
        (orden.nc ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">NC</span><span class="text-gray-900 font-medium">' + safeStr(orden.nc) + '</span></div>' : '') +
        (orden.telefono ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">Teléfono</span><span class="text-gray-900 font-medium">' + safeStr(orden.telefono) + '</span></div>' : '') +
        (orden.pareja ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">Pareja</span><span class="font-bold" style="color:' + (PAREJA_COLORS[orden.pareja]||'#374151') + '">' + safeStr(orden.pareja) + '</span></div>' : '') +
        (orden.estadoCampo ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">Estado</span><span class="text-gray-900 font-medium">' + safeStr(orden.estadoCampo) + (orden.actualizadaDelsur ? ' · Actualizada ✓' : orden.estadoCampo === 'hecha' ? ' · ⚠ Sin actualizar' : '') + '</span></div>' : '') +
        (orden.observaciones ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">Observ.</span><span class="text-gray-600">' + safeStr(orden.observaciones) + '</span></div>' : '') +
        (orden.observacion ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">Nota visita</span><span class="text-gray-600">' + safeStr(orden.observacion) + '</span></div>' : '') +
        (orden.hechaPor ? '<div class="flex gap-2"><span class="text-gray-400 w-24 shrink-0">Realizada por</span><span class="text-gray-900">' + safeStr(orden.hechaPor) + '</span></div>' : '') +
      '</div>' +
    '</div>' +
    '<div class="px-5 py-4 border-t border-gray-100 flex flex-col gap-2 shrink-0">' +
      '<a href="' + mapsUrl + '" target="_blank" class="flex items-center justify-center gap-2 w-full border border-gray-300 text-gray-700 font-medium rounded-xl py-2.5 text-sm"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="3 11 22 2 13 21 11 13 3 11"/></svg>Cómo llegar</a>' +
      (!bloqueada && !hecha && isCampo ? '<div class="flex gap-2"><button id="btn-visita" class="flex-1 border-2 border-gray-300 text-gray-700 font-semibold rounded-xl py-2.5 text-sm">Visita</button><button id="btn-hecha" class="flex-1 text-white font-semibold rounded-xl py-2.5 text-sm" style="background:#0F766E">✓ Realizada</button></div>' : '') +
      (isAdmin && !isCampo ? '<button id="btn-asignar-1" class="w-full border border-gray-300 text-gray-600 font-medium rounded-xl py-2.5 text-sm">Asignar pareja</button>' : '') +
      (isAdmin && !isCampo && orden.estadoCampo === 'hecha' ? '<button id="btn-aprobar" class="w-full text-white font-semibold rounded-xl py-2.5 text-sm" style="background:#166534">✓ Confirmar realizada</button>' : '') +
    '</div>'
  );

  ov.querySelector('#det-close').onclick = () => ov.remove();

  ov.querySelector('#btn-hecha')?.addEventListener('click', function() { ov.remove(); showConfirmarHecha(db, session, orden); });
  ov.querySelector('#btn-visita')?.addEventListener('click', function() { ov.remove(); showRegistrarVisita(db, session, orden); });
  ov.querySelector('#btn-asignar-1')?.addEventListener('click', function() { ov.remove(); showAsignarPareja(db, [orden.wo], null); });
  ov.querySelector('#btn-aprobar')?.addEventListener('click', async function() {
    const btn = ov.querySelector('#btn-aprobar'); btn.textContent = 'Confirmando...'; btn.disabled = true;
    try {
      let ref;
      if (orden.id) { ref = doc(db, COL_ORDENES, orden.id); }
      else { const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo','==',orden.wo))); if (snap.empty) throw new Error('No encontrada'); ref = snap.docs[0].ref; }
      await updateDoc(ref, { estadoCampo: 'aprobada', aprobadoPor: session.displayName, fechaAprobacion: serverTimestamp() });
      ov.remove(); showToast('Orden confirmada.', 'success');
      showSeguimiento(db, session);
    } catch(e) { showToast('Error al confirmar.', 'error'); btn.textContent = '✓ Confirmar realizada'; btn.disabled = false; }
  });
}

// ─────────────────────────────────────────
// CONFIRMAR REALIZADA (campo)
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
      let ref;
      if (orden.id) { ref = doc(db, COL_ORDENES, orden.id); }
      else { const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo','==',orden.wo))); if (snap.empty) throw new Error('No encontrada'); ref = snap.docs[0].ref; }
      await updateDoc(ref, { estadoCampo: 'hecha', actualizadaDelsur: actualizada, fechaHecha: serverTimestamp(), hechaPor: session.displayName });
      ov.remove();
      showToast('Orden marcada como realizada.', 'success');
      if (orden.__marker) { orden.__marker.setMap(null); }
    } catch(e) { showToast('Error al guardar.', 'error'); console.error(e); }
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
      '<div><label class="block text-xs font-medium text-gray-500 mb-1.5 uppercase tracking-wider">Motivo / Observación</label>' +
      '<textarea id="rv-obs" rows="3" placeholder="Ej. Cliente ausente, medidor inaccesible..." class="w-full border border-gray-200 rounded-xl px-3 py-2.5 text-sm resize-none focus:outline-none focus:ring-2" style="--tw-ring-color:#0F766E"></textarea></div>' +
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
      if (orden.id) { ref = doc(db, COL_ORDENES, orden.id); }
      else { const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo','==',orden.wo))); if (snap.empty) throw new Error('No encontrada'); ref = snap.docs[0].ref; }
      await updateDoc(ref, { estadoCampo: 'visita', observacion: obs || 'Sin observación', fechaVisita: serverTimestamp(), visitadoPor: session.displayName });
      ov.remove(); showToast('Visita registrada.', 'success');
    } catch(e) { showToast('Error al guardar.', 'error'); console.error(e); }
  };
}

// ─────────────────────────────────────────
// ASIGNAR PAREJA
// ─────────────────────────────────────────
function showAsignarPareja(db, wos, onDone) {
  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<div><h2 class="font-semibold text-gray-900">Asignar pareja</h2><p class="text-xs text-gray-400 mt-0.5">' + wos.length + ' orden(es)</p></div>' +
      '<button id="ap-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-2">' +
      PAREJAS.map(function(p) {
        return '<button data-pareja="' + p + '" class="w-full flex items-center justify-between px-4 py-3.5 rounded-xl border-2 border-gray-200 text-left transition-all">' +
          '<span class="font-semibold text-gray-900">' + p + '</span>' +
          '<span style="width:12px;height:12px;border-radius:50%;background:' + PAREJA_COLORS[p] + ';display:inline-block"></span>' +
        '</button>';
      }).join('') +
    '</div>'
  );
  ov.querySelector('#ap-close').onclick = () => ov.remove();
  ov.querySelectorAll('[data-pareja]').forEach(function(btn) {
    btn.addEventListener('click', async function() {
      const pareja = btn.dataset.pareja; btn.textContent = 'Asignando...'; btn.disabled = true;
      try {
        await Promise.all(wos.map(async function(wo) {
          const snap = await getDocs(query(collection(db, COL_ORDENES), where('wo','==',wo)));
          if (!snap.empty) await updateDoc(snap.docs[0].ref, { pareja, asignadoEn: serverTimestamp() });
        }));
        ov.remove(); showToast(wos.length + ' orden(es) → ' + pareja, 'success');
        if (onDone) onDone();
      } catch(e) { showToast('Error al asignar.','error'); console.error(e); }
    });
  });
}

// ─────────────────────────────────────────
// BODEGA — CAMPO (área CAMBIOS)
// ─────────────────────────────────────────
async function showBodegaCampo(db, session, destino) {
  const content = document.getElementById('cambios-content');
  if (!content) return;

  // Sub-nav state
  let seccion = 'inicio';

  function renderSubNav() {
    return '<div class="flex gap-1 bg-gray-100 rounded-xl p-1 mb-4">' +
      ['inicio','consumo','solicitar','pedidos'].map(function(s) {
        const labels = { inicio:'Inicio', consumo:'Consumo', solicitar:'Solicitar', pedidos:'Pedidos' };
        const activo = seccion === s;
        return '<button class="bodega-sub flex-1 py-2 text-xs font-medium rounded-lg transition-colors" data-sec="' + s + '" ' +
          'style="background:' + (activo ? 'white' : 'transparent') + ';color:' + (activo ? '#0F766E' : '#6B7280') + ';box-shadow:' + (activo ? '0 1px 3px rgba(0,0,0,0.1)' : 'none') + '">' +
          labels[s] + '</button>';
      }).join('') +
    '</div>';
  }

  async function renderSeccion() {
    const inner = document.getElementById('bodega-inner');
    if (!inner) return;
    inner.innerHTML = loading();
    if (seccion === 'inicio')    await bodegaInicio(db, session, destino, inner);
    if (seccion === 'consumo')   await bodegaConsumo(db, session, destino, inner);
    if (seccion === 'solicitar') await bodegaSolicitar(db, session, inner);
    if (seccion === 'pedidos')   await bodegaPedidos(db, session, inner);
  }

  content.innerHTML =
    '<div>' +
      renderSubNav() +
      '<div id="bodega-inner"></div>' +
    '</div>';

  await renderSeccion();

  function rewire() {
    content.querySelectorAll('.bodega-sub').forEach(function(btn) {
      btn.addEventListener('click', async function() {
        seccion = btn.dataset.sec;
        // Update styles
        content.querySelectorAll('.bodega-sub').forEach(function(b) {
          const a = b.dataset.sec === seccion;
          b.style.background  = a ? 'white' : 'transparent';
          b.style.color       = a ? '#0F766E' : '#6B7280';
          b.style.boxShadow   = a ? '0 1px 3px rgba(0,0,0,0.1)' : 'none';
        });
        await renderSeccion();
        rewire();
      });
    });
  }
  rewire();
}

// ── Inicio ──────────────────────────────
async function bodegaInicio(db, session, destino, el) {
  try {
    const usuario = session.asignacionActual?.destino || destino;

    // Material asignado al usuario (salidas de bodega)
    const snapSalidas = await getDocs(
      query(collection(db, 'kardex/movimientos/salidas'),
        where('area', '==', 'CAMBIOS'),
        where('usuarioOperativo', '==', usuario)
      )
    );
    const snapItems = await getDocs(collection(db, 'kardex/inventario/items'));
    const itemsMap  = {};
    snapItems.docs.forEach(function(d) { itemsMap[d.id] = d.data(); });

    // Aggregate assigned material
    const asignado = {};
    snapSalidas.docs.forEach(function(d) {
      const data = d.data();
      (data.materiales || []).forEach(function(m) {
        if (!asignado[m.itemId]) asignado[m.itemId] = { nombre: m.nombre || (itemsMap[m.itemId]?.nombre || m.itemId), unit: m.unit || '', cantidad: 0 };
        asignado[m.itemId].cantidad += (m.cantidad || 0);
      });
    });

    // Subtract consumos
    const snapConsumos = await getDocs(
      query(collection(db, 'kardex/movimientos/consumos'),
        where('area', '==', 'CAMBIOS'),
        where('usuarioOperativo', '==', usuario)
      )
    );
    snapConsumos.docs.forEach(function(d) {
      const data = d.data();
      (data.materiales || []).forEach(function(m) {
        if (asignado[m.itemId]) asignado[m.itemId].cantidad -= (m.cantidad || 0);
      });
    });

    const items = Object.values(asignado).filter(function(i) { return i.cantidad > 0; });

    el.innerHTML =
      '<div class="space-y-4">' +
        // Bienvenida
        '<div class="rounded-2xl p-4 text-white" style="background:#0F766E">' +
          '<p class="text-xs font-medium opacity-80 uppercase tracking-wider">Bodega · CAMBIOS</p>' +
          '<p class="text-2xl font-black mt-1">' + safeStr(usuario) + '</p>' +
          '<p class="text-xs opacity-70 mt-0.5">' + items.length + ' material' + (items.length !== 1 ? 'es' : '') + ' asignado' + (items.length !== 1 ? 's' : '') + '</p>' +
        '</div>' +
        // Accesos rápidos
        '<div class="grid grid-cols-2 gap-3">' +
          '<button class="bodega-quick bg-white border border-gray-200 rounded-xl p-4 text-left" data-sec="consumo">' +
            '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#0F766E" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="mb-2"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>' +
            '<p class="text-sm font-semibold text-gray-900">Registrar consumo</p>' +
            '<p class="text-xs text-gray-400 mt-0.5">Material usado</p>' +
          '</button>' +
          '<button class="bodega-quick bg-white border border-gray-200 rounded-xl p-4 text-left" data-sec="solicitar">' +
            '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#0F766E" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="mb-2"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>' +
            '<p class="text-sm font-semibold text-gray-900">Solicitar material</p>' +
            '<p class="text-xs text-gray-400 mt-0.5">Pedir a bodega</p>' +
          '</button>' +
        '</div>' +
        // Material asignado
        '<div class="bg-white rounded-xl border border-gray-200 overflow-hidden">' +
          '<div class="px-4 py-3 border-b border-gray-100">' +
            '<p class="font-semibold text-sm text-gray-900">Material asignado</p>' +
          '</div>' +
          (items.length ?
            '<div class="divide-y divide-gray-50">' +
              items.map(function(i) {
                return '<div class="px-4 py-3 flex items-center justify-between">' +
                  '<p class="text-sm text-gray-800">' + safeStr(i.nombre) + '</p>' +
                  '<span class="text-sm font-bold text-gray-900">' + i.cantidad + ' <span class="text-xs text-gray-400 font-normal">' + safeStr(i.unit) + '</span></span>' +
                '</div>';
              }).join('') +
            '</div>'
          : '<p class="text-sm text-gray-400 text-center py-6">Sin material asignado</p>') +
        '</div>' +
      '</div>';

    // Wire quick access buttons
    el.querySelectorAll('.bodega-quick').forEach(function(btn) {
      btn.addEventListener('click', function() {
        const sec = btn.dataset.sec;
        // Trigger the sub-nav button
        const subBtn = document.querySelector('.bodega-sub[data-sec="' + sec + '"]');
        if (subBtn) subBtn.click();
      });
    });

  } catch(e) { el.innerHTML = errHtml(); console.error(e); }
}

// ── Consumo ─────────────────────────────
async function bodegaConsumo(db, session, destino, el) {
  try {
    const usuario = session.asignacionActual?.destino || destino;
    const snap    = await getDocs(collection(db, 'kardex/inventario/items'));
    const items   = snap.docs
      .map(function(d) { return Object.assign({ id: d.id }, d.data()); })
      .filter(function(i) { return i.area === 'CAMBIOS' || !i.area; });

    if (!items.length) {
      el.innerHTML = '<div class="text-center py-10 text-sm text-gray-400">No hay materiales configurados en bodega.</div>';
      return;
    }

    let sel = []; // { itemId, name, unit, cantidad }

    function renderConsumo() {
      el.innerHTML =
        '<div class="space-y-3">' +
          '<p class="text-xs text-gray-500">Selecciona los materiales que usaste:</p>' +
          '<div class="space-y-2">' +
            items.map(function(item) {
              const s = sel.find(function(x) { return x.itemId === item.id; });
              return '<div class="bg-white rounded-xl border border-gray-200 px-4 py-3 flex items-center justify-between gap-3">' +
                '<div class="min-w-0">' +
                  '<p class="text-sm font-medium text-gray-900">' + safeStr(item.nombre) + '</p>' +
                  '<p class="text-xs text-gray-400">' + safeStr(item.unit) + '</p>' +
                '</div>' +
                '<div class="flex items-center gap-2 flex-shrink-0">' +
                  '<button class="btn-minus w-7 h-7 rounded-lg border border-gray-200 text-gray-600 font-bold flex items-center justify-center" data-id="' + item.id + '">−</button>' +
                  '<span class="text-sm font-bold w-6 text-center">' + (s ? s.cantidad : 0) + '</span>' +
                  '<button class="btn-plus w-7 h-7 rounded-lg text-white font-bold flex items-center justify-center" style="background:#0F766E" data-id="' + item.id + '" data-name="' + safeStr(item.nombre) + '" data-unit="' + safeStr(item.unit) + '">+</button>' +
                '</div>' +
              '</div>';
            }).join('') +
          '</div>' +
          (sel.filter(function(x) { return x.cantidad > 0; }).length ?
            '<button id="btn-guardar-consumo" class="w-full py-3 rounded-xl text-white font-bold text-sm" style="background:#0F766E">Registrar consumo (' + sel.filter(function(x){ return x.cantidad>0;}).length + ' materiales)</button>'
          : '') +
        '</div>';

      el.querySelectorAll('.btn-plus').forEach(function(btn) {
        btn.addEventListener('click', function() {
          const id = btn.dataset.id; let s = sel.find(function(x){ return x.itemId===id; });
          if (!s) { s = { itemId:id, name:btn.dataset.name, unit:btn.dataset.unit, cantidad:0 }; sel.push(s); }
          s.cantidad++; renderConsumo();
        });
      });
      el.querySelectorAll('.btn-minus').forEach(function(btn) {
        btn.addEventListener('click', function() {
          const id = btn.dataset.id; const s = sel.find(function(x){ return x.itemId===id; });
          if (s && s.cantidad > 0) { s.cantidad--; renderConsumo(); }
        });
      });
      el.querySelector('#btn-guardar-consumo')?.addEventListener('click', async function() {
        const btn = el.querySelector('#btn-guardar-consumo');
        const materiales = sel.filter(function(x){ return x.cantidad>0; });
        btn.textContent = 'Guardando...'; btn.disabled = true;
        try {
          await addDoc(collection(db, 'kardex/movimientos/consumos'), {
            usuarioUid:      session.uid,
            usuarioNombre:   session.displayName,
            usuarioOperativo: usuario,
            area:            'CAMBIOS',
            materiales:      materiales,
            fecha:           serverTimestamp(),
          });
          sel = [];
          showToast('Consumo registrado.', 'success');
          renderConsumo();
        } catch(e) { showToast('Error al guardar.','error'); btn.disabled=false; btn.textContent='Registrar consumo'; }
      });
    }
    renderConsumo();

  } catch(e) { el.innerHTML = errHtml(); console.error(e); }
}

// ── Solicitar ────────────────────────────
async function bodegaSolicitar(db, session, el) {
  try {
    const snap  = await getDocs(collection(db, 'kardex/inventario/items'));
    const items = snap.docs
      .map(function(d) { return Object.assign({ id: d.id }, d.data()); })
      .filter(function(i) { return i.area === 'CAMBIOS' || !i.area; });

    if (!items.length) {
      el.innerHTML = '<div class="text-center py-10 text-sm text-gray-400">No hay materiales configurados en bodega.</div>';
      return;
    }

    let sel = [];

    function renderSolicitar() {
      el.innerHTML =
        '<div class="space-y-3">' +
          '<p class="text-xs text-gray-500">Selecciona el material que necesitas:</p>' +
          '<div class="space-y-2">' +
            items.map(function(item) {
              const s = sel.find(function(x){ return x.itemId===item.id; });
              const stock = item.stockActual || 0;
              return '<div class="bg-white rounded-xl border border-gray-200 px-4 py-3 ' + (stock <= 0 ? 'opacity-50' : '') + '">' +
                '<div class="flex items-center justify-between gap-3">' +
                  '<div class="min-w-0">' +
                    '<p class="text-sm font-medium text-gray-900">' + safeStr(item.nombre) + '</p>' +
                    '<p class="text-xs text-gray-400">Stock: ' + stock + ' ' + safeStr(item.unit) + '</p>' +
                  '</div>' +
                  '<div class="flex items-center gap-2 flex-shrink-0">' +
                    '<button class="btn-sminus w-7 h-7 rounded-lg border border-gray-200 text-gray-600 font-bold flex items-center justify-center" data-id="' + item.id + '"' + (stock<=0?'disabled':'') + '>−</button>' +
                    '<span class="text-sm font-bold w-6 text-center">' + (s ? s.cantidad : 0) + '</span>' +
                    '<button class="btn-splus w-7 h-7 rounded-lg text-white font-bold flex items-center justify-center" style="background:#0F766E" data-id="' + item.id + '" data-name="' + safeStr(item.nombre) + '" data-unit="' + safeStr(item.unit) + '"' + (stock<=0?'disabled':'') + '>+</button>' +
                  '</div>' +
                '</div>' +
              '</div>';
            }).join('') +
          '</div>' +
          (sel.filter(function(x){ return x.cantidad>0; }).length ?
            '<button id="btn-enviar-solicitud" class="w-full py-3 rounded-xl text-white font-bold text-sm" style="background:#1B4F8A">Enviar solicitud (' + sel.filter(function(x){ return x.cantidad>0;}).length + ' materiales)</button>'
          : '') +
        '</div>';

      el.querySelectorAll('.btn-splus').forEach(function(btn) {
        btn.addEventListener('click', function() {
          const id = btn.dataset.id; let s = sel.find(function(x){ return x.itemId===id; });
          if (!s) { s = { itemId:id, name:btn.dataset.name, unit:btn.dataset.unit, cantidad:0 }; sel.push(s); }
          s.cantidad++; renderSolicitar();
        });
      });
      el.querySelectorAll('.btn-sminus').forEach(function(btn) {
        btn.addEventListener('click', function() {
          const id = btn.dataset.id; const s = sel.find(function(x){ return x.itemId===id; });
          if (s && s.cantidad > 0) { s.cantidad--; renderSolicitar(); }
        });
      });
      el.querySelector('#btn-enviar-solicitud')?.addEventListener('click', async function() {
        const btn = el.querySelector('#btn-enviar-solicitud');
        const materiales = sel.filter(function(x){ return x.cantidad>0; });
        btn.textContent = 'Enviando...'; btn.disabled = true;
        try {
          await addDoc(collection(db, 'solicitudes_material'), {
            usuarioUid:    session.uid,
            usuarioNombre: session.displayName,
            usuarioRole:   session.role,
            area:          'CAMBIOS',
            materiales:    materiales.map(function(s){ return { itemId:s.itemId, nombre:s.name, unit:s.unit, cantidad:s.cantidad }; }),
            estado:        'pendiente',
            fecha:         serverTimestamp(),
            notas:         '',
          });
          sel = [];
          showToast('Solicitud enviada.', 'success');
          renderSolicitar();
        } catch(e) { showToast('Error al enviar.','error'); btn.disabled=false; btn.textContent='Enviar solicitud'; }
      });
    }
    renderSolicitar();

  } catch(e) { el.innerHTML = errHtml(); console.error(e); }
}

// ── Pedidos ──────────────────────────────
async function bodegaPedidos(db, session, el) {
  try {
    const snap = await getDocs(
      query(collection(db, 'solicitudes_material'),
        where('usuarioUid', '==', session.uid),
        where('area', '==', 'CAMBIOS'),
        orderBy('fecha', 'desc')
      )
    );

    const pedidos = snap.docs.map(function(d) { return Object.assign({ id: d.id }, d.data()); });

    if (!pedidos.length) {
      el.innerHTML = '<div class="text-center py-10 text-sm text-gray-400">No has hecho solicitudes aún.</div>';
      return;
    }

    const estadoColors = { pendiente: { bg:'#FEF3C7', color:'#B45309', label:'Pendiente' }, aprobada: { bg:'#DCFCE7', color:'#166534', label:'Aprobada' }, rechazada: { bg:'#FEF2F2', color:'#C62828', label:'Rechazada' } };

    el.innerHTML =
      '<div class="space-y-2">' +
        pedidos.map(function(p) {
          const e = estadoColors[p.estado] || estadoColors.pendiente;
          return '<div class="bg-white rounded-xl border border-gray-200 px-4 py-3">' +
            '<div class="flex items-center justify-between mb-1.5">' +
              '<p class="text-xs text-gray-400">' + fmtDate(p.fecha) + '</p>' +
              '<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:' + e.bg + ';color:' + e.color + '">' + e.label + '</span>' +
            '</div>' +
            '<div class="space-y-0.5">' +
              (p.materiales || []).map(function(m) {
                return '<p class="text-sm text-gray-700">' + safeStr(m.nombre) + ' · <span class="font-semibold">' + m.cantidad + ' ' + safeStr(m.unit) + '</span></p>';
              }).join('') +
            '</div>' +
            (p.notas ? '<p class="text-xs text-gray-400 mt-1.5">' + safeStr(p.notas) + '</p>' : '') +
          '</div>';
        }).join('') +
      '</div>';

  } catch(e) { el.innerHTML = errHtml(); console.error(e); }
}

// ─────────────────────────────────────────
// DESASIGNAR PAREJA COMPLETA
// ─────────────────────────────────────────
function showDesasignarPareja(db, ordenes, onDone) {
  // Count pending orders per pareja
  const counts = {};
  PAREJAS.forEach(function(p) { counts[p] = 0; });
  ordenes.forEach(function(o) {
    if (o.pareja && counts[o.pareja] !== undefined && o.estadoCampo !== 'aprobada') counts[o.pareja]++;
  });

  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<div>' +
        '<h2 class="font-semibold text-gray-900">Desasignar pareja</h2>' +
        '<p class="text-xs text-gray-400 mt-0.5">Quita todas las órdenes pendientes de una pareja</p>' +
      '</div>' +
      '<button id="dp-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-2">' +
      PAREJAS.map(function(p) {
        const c = PAREJA_COLORS[p];
        const n = counts[p];
        return '<div class="flex items-center justify-between bg-white border border-gray-200 rounded-xl px-4 py-3">' +
          '<div class="flex items-center gap-3">' +
            '<span style="width:10px;height:10px;border-radius:50%;background:' + c + ';display:inline-block;flex-shrink:0"></span>' +
            '<div>' +
              '<p class="font-semibold text-gray-900 text-sm">' + p + '</p>' +
              '<p class="text-xs text-gray-400">' + n + ' orden' + (n !== 1 ? 'es' : '') + ' asignada' + (n !== 1 ? 's' : '') + '</p>' +
            '</div>' +
          '</div>' +
          (n > 0 ?
            '<button data-pareja="' + p + '" class="btn-dp text-xs font-bold px-3 py-2 rounded-lg" style="background:#FEF2F2;color:#C62828;border:1px solid #FECACA">Desasignar</button>'
          : '<span class="text-xs text-gray-300">Sin órdenes</span>') +
        '</div>';
      }).join('') +
      '<div class="pt-2 border-t border-gray-100">' +
        '<button id="dp-all" class="w-full py-3 rounded-xl text-sm font-bold" style="background:#FEF2F2;color:#C62828;border:1.5px solid #FECACA">✕ Desasignar TODAS las parejas</button>' +
      '</div>' +
    '</div>'
  );

  ov.querySelector('#dp-close').onclick = () => ov.remove();

  async function desasignar(pareja) {
    const btn = pareja
      ? ov.querySelector('[data-pareja="' + pareja + '"].btn-dp')
      : ov.querySelector('#dp-all');
    if (btn) { btn.textContent = 'Desasignando...'; btn.disabled = true; }
    try {
      const toUnassign = ordenes.filter(function(o) {
        return (pareja ? o.pareja === pareja : o.pareja) && o.estadoCampo !== 'aprobada' && o.estadoCampo !== 'hecha';
      });
      await Promise.all(toUnassign.map(function(o) {
        const ref = o.id ? doc(db, COL_ORDENES, o.id) : null;
        if (!ref) return Promise.resolve();
        return updateDoc(ref, { pareja: null });
      }));
      ov.remove();
      showToast(toUnassign.length + ' órdenes desasignadas' + (pareja ? ' de ' + pareja : ''), 'success');
      if (onDone) onDone();
    } catch(e) {
      showToast('Error al desasignar.', 'error');
      if (btn) { btn.disabled = false; btn.textContent = pareja ? 'Desasignar' : '✕ Desasignar TODAS las parejas'; }
      console.error(e);
    }
  }

  ov.querySelectorAll('.btn-dp').forEach(function(btn) {
    btn.addEventListener('click', function() { desasignar(btn.dataset.pareja); });
  });
  ov.querySelector('#dp-all').addEventListener('click', function() { desasignar(null); });
}
